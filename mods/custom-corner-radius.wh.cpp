// ==WindhawkMod==
// @id              custom-corner-radius
// @name            Custom Window Corner Radius
// @description     Customizes window corner radius in Windows 11, making corners more or less rounded
// @version         1.3
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         dwm.exe
// @architecture    x86-64
// @compilerOptions -lgdi32 -lole32 -lwevtapi
// ==/WindhawkMod==

// Source code is published under The GNU General Public License v3.0.
//
// For bug reports and feature requests, please open an issue here:
// https://github.com/ramensoftware/windhawk-mods/issues
//
// For pull requests, development takes place here:
// https://github.com/m417z/my-windhawk-mods

// ==WindhawkModReadme==
/*
# Custom Window Corner Radius

Customizes Windows 11 app window corner radius. Make corners more rounded than
the default 8px, or reduce the radius for less rounded or completely sharp
corners.

The mod was [originally
submitted](https://github.com/ramensoftware/windhawk-mods/pull/3587) by
[Kanak415](https://github.com/kanak-buet19). The option for maximized and
snapped windows is based on a
[mod](https://github.com/ramensoftware/windhawk-mods/pull/5022) by [Alexey
Lavrinenko](https://github.com/leshaalexey).

![Screenshot](https://i.imgur.com/mMGkBwc.png)

## ⚠ Important usage note ⚠

This mod needs to hook into `dwm.exe` to work. Please navigate to Windhawk's
Settings > Advanced settings > More advanced settings > Process inclusion list,
and make sure that `dwm.exe` is in the list.

![Advanced settings screenshot](https://i.imgur.com/LRhREtJ.png)

## Additional notes

- Some elements, such as context menus, use a smaller radius (4px by default).
  This can be customized separately with the "Small corner radius" option.
- Standard tooltips can be customized separately with the "Tooltip corner
  radius" option.
- Windows 11 squares off the corners of maximized and snapped windows. The
  "Rounded corners for maximized and snapped windows" option enables rounding
  for them.
- Some elements, such as the taskbar, the Start menu, and the notification
  center, are unaffected by this mod. Some of them can be customized using other
  mods, such as Windows 11 Taskbar Styler.
- Disabling the mod instantly restores default behavior - no system files are
  modified.

## Compatibility

- When using this mod alongside Translucent Flyouts, set its `CornerType` option
  to `0` ("Don't Change") to prevent conflicts between the two.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- radius: 12
  $name: Corner radius
  $description: >-
    Corner radius in pixels. Default Win11 is 8. Use smaller values (e.g. 4 or
    0) for less rounded or sharp corners, or larger values (e.g. 10-20) for more
    rounded corners. Values above 20 may cause visual artifacts depending on
    your DPI scaling.

    Set to -1 to keep the original radius.
- smallRadius: 6
  $name: Small corner radius
  $description: >-
    Corner radius for elements that use a smaller radius, such as context menus.
    Default Win11 is 4.

    Set to -1 to keep the original radius.
- tooltipRadius: -1
  $name: Tooltip corner radius
  $description: >-
    Corner radius for standard tooltips. Note that this doesn't affect modern
    (WinUI) tooltips. Values above 8 may cause visual artifacts depending on
    your DPI scaling.

    Set to -1 to leave tooltips unchanged.
- roundMaximizedAndSnapped: none
  $name: Rounded corners for maximized and snapped windows
  $description: >-
    Windows 11 squares off the corners of maximized and snapped windows. This
    option enables rounding for them using the "Corner radius" value above.
  $options:
  - none: Leave them square
  - snapped: Round snapped windows
  - snappedAndMaximized: Round snapped and maximized windows
- excludedPrograms: [""]
  $name: Excluded programs
  $description: >-
    Windows of these programs keep their original corner radius.

    Entries can be process names, paths or application IDs, for example:

    mspaint.exe

    C:\Windows\System32\notepad.exe

    Microsoft.WindowsCalculator_8wekyb3d8bbwe!App
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <initguid.h>  // Must appear before propkey.h

#include <dwmapi.h>
#include <propkey.h>
#include <propsys.h>
#include <winevt.h>
#include <winternl.h>

#include <cmath>
#include <limits>
#include <regex>
#include <string>
#include <string_view>
#include <unordered_set>

enum class RoundMaximizedAndSnapped {
    none,
    snapped,
    snappedAndMaximized,
};

struct {
    float radius;
    float smallRadius;
    float tooltipRadius;
    RoundMaximizedAndSnapped roundMaximizedAndSnapped;
    std::unordered_set<std::wstring> excludedPrograms;
} g_settings;

using GetWindowData_t = void*(WINAPI*)(void* pThis);
GetWindowData_t GetWindowData_Original;

using IsMaximizedOrSnapped_t = bool(WINAPI*)(void* pThis);
IsMaximizedOrSnapped_t IsMaximizedOrSnapped_Original;

// Captured (not hooked) address of CWindowData::IsGhostWindow. We disassemble
// its first few instructions at init time to recover the HWND member offset.
void* IsGhostWindow_Func;

// HWND offset within CWindowData. Recovered at init time from the first
// `mov rcx, qword ptr [rcx+disp]` instruction in CWindowData::IsGhostWindow,
// which loads its HWND member as the first argument to GetPropW. Stays at
// SIZE_MAX if the pattern can't be matched, in which case HWND lookup is
// disabled and tooltip-specific behavior gracefully degrades.
size_t g_windowDataHwndOffset = SIZE_MAX;

// Scans the first `limit` instructions of `func` for a match against `regex`
// and returns the value of the first capture group parsed as hex. Mirrors the
// helper used by taskbar-button-scroll / taskbar-icon-size.
size_t OffsetFromAssemblyRegex(void* func,
                               size_t defValue,
                               std::regex regex,
                               int limit = 30) {
    BYTE* p = (BYTE*)func;
    for (int i = 0; i < limit; i++) {
        WH_DISASM_RESULT result;
        if (!Wh_Disasm(p, &result)) {
            break;
        }

        p += result.length;

        std::string_view s = result.text;
        if (s == "ret") {
            break;
        }

        std::match_results<std::string_view::const_iterator> match;
        if (std::regex_match(s.begin(), s.end(), match, regex)) {
            // Wh_Log(L"%S", result.text);
            return std::stoull(match[1], nullptr, 16);
        }
    }

    Wh_Log(L"Failed for %p", func);
    return defValue;
}

HWND HwndFromTopLevelWindow(void* pThis) {
    if (g_windowDataHwndOffset == SIZE_MAX || !GetWindowData_Original) {
        return nullptr;
    }
    void* pData = GetWindowData_Original(pThis);
    if (!pData) {
        return nullptr;
    }
    HWND hwnd = *(HWND*)((BYTE*)pData + g_windowDataHwndOffset);
    return IsWindow(hwnd) ? hwnd : nullptr;
}

bool HwndHasClass(HWND hwnd, PCWSTR className) {
    if (!hwnd) {
        return false;
    }
    WCHAR buf[32];
    return GetClassNameW(hwnd, buf, ARRAYSIZE(buf)) &&
           _wcsicmp(buf, className) == 0;
}

bool IsTopLevelWindowTooltip(void* pThis) {
    return HwndHasClass(HwndFromTopLevelWindow(pThis), L"tooltips_class32");
}

typedef struct _SYSTEM_PROCESS_ID_INFORMATION {
    HANDLE ProcessId;
    UNICODE_STRING ImageName;
} SYSTEM_PROCESS_ID_INFORMATION;

// The image path of a process in NT form, e.g.
// \Device\HarddiskVolume3\Windows\System32\notepad.exe. Unlike
// QueryFullProcessImageName, this needs no handle to the process, which dwm.exe
// can't get: it runs under a virtual account that isn't in the DACL of the
// processes owning the windows it composes.
std::wstring GetProcessImageNtPath(DWORD processId) {
    using NtQuerySystemInformation_t =
        LONG(NTAPI*)(ULONG systemInformationClass, PVOID systemInformation,
                     ULONG systemInformationLength, PULONG returnLength);
    static NtQuerySystemInformation_t pNtQuerySystemInformation = []() {
        HMODULE ntdll = GetModuleHandle(L"ntdll.dll");
        return ntdll ? (NtQuerySystemInformation_t)GetProcAddress(
                           ntdll, "NtQuerySystemInformation")
                     : nullptr;
    }();
    if (!pNtQuerySystemInformation) {
        return std::wstring{};
    }

    constexpr ULONG kSystemProcessIdInformation = 88;
    constexpr LONG kStatusInfoLengthMismatch = 0xC0000004;

    SYSTEM_PROCESS_ID_INFORMATION info{
        .ProcessId = (HANDLE)(ULONG_PTR)processId,
    };

    // An empty ImageName makes the call report the size it needs.
    LONG status = pNtQuerySystemInformation(kSystemProcessIdInformation, &info,
                                            sizeof(info), nullptr);
    if (status != kStatusInfoLengthMismatch) {
        Wh_Log(L"Size query failed for pid=%u: %08X", processId, status);
        return std::wstring{};
    }

    std::wstring path(info.ImageName.MaximumLength / sizeof(WCHAR), L'\0');
    info.ImageName.Buffer = path.data();
    info.ImageName.Length = 0;

    status = pNtQuerySystemInformation(kSystemProcessIdInformation, &info,
                                       sizeof(info), nullptr);
    if (status < 0) {
        Wh_Log(L"Query failed for pid=%u: %08X", processId, status);
        return std::wstring{};
    }

    path.resize(info.ImageName.Length / sizeof(WCHAR));
    return path;
}

// Turns \Device\HarddiskVolume3\Windows\... into C:\Windows\..., or returns an
// empty string if no drive letter maps to the device.
std::wstring NtPathToDosPath(const std::wstring& ntPath) {
    WCHAR drives[512];
    DWORD len = GetLogicalDriveStrings(ARRAYSIZE(drives), drives);
    if (!len || len > ARRAYSIZE(drives)) {
        return std::wstring{};
    }

    for (PCWSTR drive = drives; *drive; drive += wcslen(drive) + 1) {
        WCHAR driveName[] = {drive[0], L':', L'\0'};
        WCHAR target[MAX_PATH];
        if (!QueryDosDevice(driveName, target, ARRAYSIZE(target))) {
            continue;
        }

        size_t targetLen = wcslen(target);
        if (ntPath.length() > targetLen && ntPath[targetLen] == L'\\' &&
            _wcsnicmp(ntPath.c_str(), target, targetLen) == 0) {
            return driveName + ntPath.substr(targetLen);
        }
    }

    return std::wstring{};
}

// The AppUserModelID explicitly set on a window. It's what identifies packaged
// apps, whose windows belong to a shared host process. shell32 is resolved on
// demand to keep it out of dwm.exe unless exclusions are actually used.
std::wstring GetWindowAppId(HWND hWnd) {
    using SHGetPropertyStoreForWindow_t =
        HRESULT(WINAPI*)(HWND hwnd, REFIID riid, void** ppv);
    static SHGetPropertyStoreForWindow_t pSHGetPropertyStoreForWindow = []() {
        HMODULE shell32 = LoadLibraryEx(L"shell32.dll", nullptr,
                                        LOAD_LIBRARY_SEARCH_SYSTEM32);
        return shell32 ? (SHGetPropertyStoreForWindow_t)GetProcAddress(
                             shell32, "SHGetPropertyStoreForWindow")
                       : nullptr;
    }();

    std::wstring result;

    if (!pSHGetPropertyStoreForWindow) {
        Wh_Log(L"SHGetPropertyStoreForWindow isn't available");
        return result;
    }

    IPropertyStore* propertyStore;
    HRESULT hr =
        pSHGetPropertyStoreForWindow(hWnd, IID_PPV_ARGS(&propertyStore));
    if (FAILED(hr)) {
        Wh_Log(L"SHGetPropertyStoreForWindow failed for hwnd=%p: %08X", hWnd,
               hr);
        return result;
    }

    PROPVARIANT pv;
    PropVariantInit(&pv);
    hr = propertyStore->GetValue(PKEY_AppUserModel_ID, &pv);
    if (SUCCEEDED(hr)) {
        if (pv.vt == VT_LPWSTR && pv.pwszVal) {
            result = pv.pwszVal;
        }
        PropVariantClear(&pv);
    } else {
        Wh_Log(L"GetValue failed for hwnd=%p: %08X", hWnd, hr);
    }

    propertyStore->Release();
    return result;
}

void MakeUpper(std::wstring* str) {
    LCMapStringEx(LOCALE_NAME_USER_DEFAULT, LCMAP_UPPERCASE, str->data(),
                  static_cast<int>(str->length()), str->data(),
                  static_cast<int>(str->length()), nullptr, nullptr, 0);
}

bool IsWindowExcluded(HWND hWnd) {
    DWORD dwProcessId = 0;
    GetWindowThreadProcessId(hWnd, &dwProcessId);

    std::wstring processPathUpper;
    if (dwProcessId) {
        std::wstring ntPath = GetProcessImageNtPath(dwProcessId);
        processPathUpper = NtPathToDosPath(ntPath);
        if (processPathUpper.empty()) {
            // Without a drive letter no configured path can match, but the
            // last component is still the file name.
            processPathUpper = std::move(ntPath);
        }
        MakeUpper(&processPathUpper);
    }

    if (!processPathUpper.empty()) {
        if (g_settings.excludedPrograms.contains(processPathUpper)) {
            Wh_Log(L"hwnd=%p excluded by path: %s", hWnd,
                   processPathUpper.c_str());
            return true;
        }

        size_t fileNamePos = processPathUpper.rfind(L'\\');
        if (fileNamePos != std::wstring::npos) {
            std::wstring fileNameUpper =
                processPathUpper.substr(fileNamePos + 1);
            if (!fileNameUpper.empty() &&
                g_settings.excludedPrograms.contains(fileNameUpper)) {
                Wh_Log(L"hwnd=%p excluded by file name: %s", hWnd,
                       fileNameUpper.c_str());
                return true;
            }
        }
    }

    std::wstring appIdUpper = GetWindowAppId(hWnd);
    MakeUpper(&appIdUpper);
    if (!appIdUpper.empty() &&
        g_settings.excludedPrograms.contains(appIdUpper)) {
        Wh_Log(L"hwnd=%p excluded by app id: %s", hWnd, appIdUpper.c_str());
        return true;
    }

    Wh_Log(L"hwnd=%p not excluded, path=[%s], appId=[%s]", hWnd,
           processPathUpper.c_str(), appIdUpper.c_str());
    return false;
}

constexpr WCHAR kWindowExclusionProp[] = L"Windhawk_Excluded_" WH_MOD_ID;

const HANDLE kWindowNotExcluded = (HANDLE)1;
const HANDLE kWindowExcluded = (HANDLE)2;

// Resolving the program of a window is expensive, and the hooks run as part of
// composing every frame, so the verdict is cached in a window property. The
// properties are dropped when the settings change and when the mod is unloaded.
bool IsWindowExcludedCached(HWND hWnd) {
    HANDLE prop = GetProp(hWnd, kWindowExclusionProp);
    if (!prop) {
        prop = IsWindowExcluded(hWnd) ? kWindowExcluded : kWindowNotExcluded;
        if (!SetProp(hWnd, kWindowExclusionProp, prop)) {
            Wh_Log(L"SetProp failed for hwnd=%p: %u", hWnd, GetLastError());
        }
    }

    return prop == kWindowExcluded;
}

// Windows whose HWND can't be recovered are never excluded, same as the other
// HWND-based checks.
bool IsTopLevelWindowExcluded(void* pThis) {
    if (g_settings.excludedPrograms.empty()) {
        return false;
    }

    HWND hwnd = HwndFromTopLevelWindow(pThis);
    if (!hwnd) {
        Wh_Log(L"No hwnd for %p, can't check exclusions", pThis);
        return false;
    }

    return IsWindowExcludedCached(hwnd);
}

void ClearWindowExclusionProps() {
    EnumWindows(
        [](HWND hWnd, LPARAM) -> BOOL {
            RemoveProp(hWnd, kWindowExclusionProp);
            return TRUE;
        },
        0);
}

float RadiusForOriginal(float orig, bool isTooltip) {
    // In new builds, multiple hooks fire in sequence (GetRadiusFromCornerStyle
    // -> GetFloatCornerRadiusForCurrentStyle -> SetBorderParameters), so a
    // downstream hook may see a value already replaced by an upstream hook.
    // Skip replacement if the value already matches a configured radius to keep
    // the function idempotent.
    if (orig == g_settings.radius || orig == g_settings.smallRadius ||
        orig == g_settings.tooltipRadius) {
        return orig;
    }

    if (isTooltip && g_settings.tooltipRadius >= 0.0f) {
        return g_settings.tooltipRadius;
    }

    // Win11 defaults: 4.0 for smaller radius, 8.0 for larger radius. Use middle
    // point as a threshold. Don't override if new value is negative.
    float newValue = orig < 6.0f ? g_settings.smallRadius : g_settings.radius;
    if (newValue < 0.0f) {
        return orig;
    }

    return newValue;
}

// Forces an empty window region on a SysShadow companion HWND so the
// legacy rectangular drop shadow stops being composited. Idempotent: once
// the window already has the (empty) region we own, the call is a no-op.
void HideSysShadowWindow(HWND hwnd) {
    HRGN rgn = CreateRectRgn(0, 0, 0, 0);
    if (!rgn) {
        return;
    }
    if (!SetWindowRgn(hwnd, rgn, FALSE)) {
        DeleteObject(rgn);
    }
    // SetWindowRgn takes ownership of the region on success.
}

// The CTopLevelWindow whose UpdateWindowVisuals call is currently running, or
// null outside of one. SetBorderParameters gets a CWindowBorder, which offers
// no way back to the window it belongs to.
thread_local void* g_updateWindowVisualsTarget;

using UpdateWindowVisuals_t = long(WINAPI*)(void* pThis);
UpdateWindowVisuals_t UpdateWindowVisuals_Original;
long WINAPI UpdateWindowVisuals_Hook(void* pThis) {
    if (g_settings.tooltipRadius >= 0.0f && !IsTopLevelWindowExcluded(pThis)) {
        HWND hwnd = HwndFromTopLevelWindow(pThis);
        if (HwndHasClass(hwnd, L"SysShadow")) {
            Wh_Log(L"> hiding SysShadow hwnd=%p", hwnd);
            HideSysShadowWindow(hwnd);
            return 0;
        }
    }

    void* prevTarget = g_updateWindowVisualsTarget;
    g_updateWindowVisualsTarget = pThis;
    long ret = UpdateWindowVisuals_Original(pThis);
    g_updateWindowVisualsTarget = prevTarget;
    return ret;
}

using GetEffectiveCornerStyle_t = int(WINAPI*)(void* pThis);
GetEffectiveCornerStyle_t GetEffectiveCornerStyle_Original;
int WINAPI GetEffectiveCornerStyle_Hook(void* pThis) {
    int orig = GetEffectiveCornerStyle_Original(pThis);
    // Tooltips report DWMWCP_DONOTROUND, meaning DWM won't round them.
    // Promote them to DWMWCP_ROUNDSMALL so the rounding pipeline kicks in:
    // GetShadowStyle returns a rounded-shadow style, and
    // GetRadiusFromCornerStyle returns a non-zero radius that our hooks
    // override to the configured tooltipRadius via RadiusForOriginal.
    if (orig == DWMWCP_DONOTROUND && g_settings.tooltipRadius >= 0.0f &&
        !IsTopLevelWindowExcluded(pThis) && IsTopLevelWindowTooltip(pThis)) {
        Wh_Log(L"> cornerStyle DONOTROUND -> ROUNDSMALL (tooltip)");
        return DWMWCP_ROUNDSMALL;
    }
    return orig;
}

bool ShouldRoundMaximizedOrSnapped(void* pThis) {
    if (g_settings.roundMaximizedAndSnapped == RoundMaximizedAndSnapped::none ||
        g_settings.radius < 0.0f || !IsMaximizedOrSnapped_Original ||
        !IsMaximizedOrSnapped_Original(pThis)) {
        return false;
    }

    // A zero radius can also come from an app asking for DWMWCP_DONOTROUND.
    // Only the squaring should be undone, so consult the corner style, which
    // the squaring doesn't touch.
    if (GetEffectiveCornerStyle_Original &&
        GetEffectiveCornerStyle_Original(pThis) == DWMWCP_DONOTROUND) {
        return false;
    }

    if (g_settings.roundMaximizedAndSnapped ==
        RoundMaximizedAndSnapped::snappedAndMaximized) {
        return true;
    }

    // Snapped windows only: a maximized window is zoomed, a snapped one isn't.
    // Without an HWND there's no way to tell the two apart, so leave the window
    // alone.
    HWND hwnd = HwndFromTopLevelWindow(pThis);
    return hwnd && !IsZoomed(hwnd);
}

// Only GetFloatCornerRadiusForCurrentStyle squares maximized and snapped
// windows, so only it passes canRoundMaximizedOrSnapped. Elsewhere a zero
// radius comes from the corner style and is left alone.
float AdjustCornerRadius(void* pThis,
                         float orig,
                         bool canRoundMaximizedOrSnapped) {
    if (IsTopLevelWindowExcluded(pThis)) {
        return orig;
    }

    if (orig > 0) {
        bool isTooltip =
            g_settings.tooltipRadius >= 0.0f && IsTopLevelWindowTooltip(pThis);
        Wh_Log(L"> %f isTooltip=%d", orig, isTooltip);
        return RadiusForOriginal(orig, isTooltip);
    }

    if (canRoundMaximizedOrSnapped && ShouldRoundMaximizedOrSnapped(pThis)) {
        Wh_Log(L"> %f -> %f (maximized or snapped)", orig, g_settings.radius);
        return g_settings.radius;
    }

    return orig;
}

using GetRadiusFromCornerStyle_t = float(WINAPI*)(void* pThis);
GetRadiusFromCornerStyle_t GetRadiusFromCornerStyle_Original;
float WINAPI GetRadiusFromCornerStyle_Hook(void* pThis) {
    return AdjustCornerRadius(pThis, GetRadiusFromCornerStyle_Original(pThis),
                              /*canRoundMaximizedOrSnapped=*/false);
}

using GetFloatCornerRadiusForCurrentStyle_t = float(WINAPI*)(void* pThis);
GetFloatCornerRadiusForCurrentStyle_t
    GetFloatCornerRadiusForCurrentStyle_Original;
float WINAPI GetFloatCornerRadiusForCurrentStyle_Hook(void* pThis) {
    return AdjustCornerRadius(
        pThis, GetFloatCornerRadiusForCurrentStyle_Original(pThis),
        /*canRoundMaximizedOrSnapped=*/true);
}

using SetBorderParameters_t = long(WINAPI*)(void* pThis,
                                            const RECT& borderRect,
                                            float cornerRadius,
                                            int dpi,
                                            const void* color,
                                            int borderStyle,
                                            int shadowStyle);
SetBorderParameters_t SetBorderParameters_Original;
long WINAPI SetBorderParameters_Hook(void* pThis,
                                     const RECT& borderRect,
                                     float cornerRadius,
                                     int dpi,
                                     const void* color,
                                     int borderStyle,
                                     int shadowStyle) {
    if (cornerRadius > 0) {
        // pThis here is a CWindowBorder, not a CTopLevelWindow, so there's no
        // straightforward way to recover the HWND for tooltip detection. This
        // path is only used by old builds where the radius is computed inline.
        // The enclosing UpdateWindowVisuals call is what identifies the window
        // for the exclusion check.
        bool excluded = g_updateWindowVisualsTarget &&
                        IsTopLevelWindowExcluded(g_updateWindowVisualsTarget);
        if (!excluded) {
            Wh_Log(L"> %f", cornerRadius);
            cornerRadius = RadiusForOriginal(cornerRadius, false);
        }
    }
    return SetBorderParameters_Original(pThis, borderRect, cornerRadius, dpi,
                                        color, borderStyle, shadowStyle);
}

void LoadSettings() {
    // Use `std::nextafter` to get a value that's just slightly above the
    // integer, for two reasons:
    // 1. The original radius values are integer-based, so if the new value is
    //    exactly the same as the original value, it's impossible to determine
    //    whether the mod should override it or not if the custom value is
    //    identical to one of the original values (see RadiusForOriginal).
    // 2. If the zero value is used, some functions may treat it as a special
    //    case, for example dark mode menus will have a white border.
    g_settings.radius =
        std::nextafter(static_cast<float>(Wh_GetIntSetting(L"radius")),
                       std::numeric_limits<float>::max());
    g_settings.smallRadius =
        std::nextafter(static_cast<float>(Wh_GetIntSetting(L"smallRadius")),
                       std::numeric_limits<float>::max());
    g_settings.tooltipRadius =
        std::nextafter(static_cast<float>(Wh_GetIntSetting(L"tooltipRadius")),
                       std::numeric_limits<float>::max());

    PCWSTR roundMaximizedAndSnapped =
        Wh_GetStringSetting(L"roundMaximizedAndSnapped");
    g_settings.roundMaximizedAndSnapped = RoundMaximizedAndSnapped::none;
    if (wcscmp(roundMaximizedAndSnapped, L"snapped") == 0) {
        g_settings.roundMaximizedAndSnapped = RoundMaximizedAndSnapped::snapped;
    } else if (wcscmp(roundMaximizedAndSnapped, L"snappedAndMaximized") == 0) {
        g_settings.roundMaximizedAndSnapped =
            RoundMaximizedAndSnapped::snappedAndMaximized;
    }
    Wh_FreeStringSetting(roundMaximizedAndSnapped);

    g_settings.excludedPrograms.clear();

    for (int i = 0;; i++) {
        PCWSTR program = Wh_GetStringSetting(L"excludedPrograms[%d]", i);

        bool hasProgram = *program;
        if (hasProgram) {
            std::wstring programUpper = program;
            LCMapStringEx(
                LOCALE_NAME_USER_DEFAULT, LCMAP_UPPERCASE, &programUpper[0],
                static_cast<int>(programUpper.length()), &programUpper[0],
                static_cast<int>(programUpper.length()), nullptr, nullptr, 0);

            Wh_Log(L"Excluded program: [%s]", programUpper.c_str());

            g_settings.excludedPrograms.insert(std::move(programUpper));
        }

        Wh_FreeStringSetting(program);

        if (!hasProgram) {
            break;
        }
    }
}

// Returns true if at least two Dwminit warnings (Level=3) were logged in the
// Application event log within the last 60 seconds. DWM logs warnings here when
// it crashes and is restarted by the session manager, so repeated warnings are
// a strong signal that something in the desktop pipeline is unstable.
bool HasMultipleDwminitWarningsInLastMinute() {
    const WCHAR* queryPath = L"Application";
    const WCHAR* query =
        L"*[System[Provider[@Name='Dwminit'] and (Level=3) and "
        L"TimeCreated[timediff(@SystemTime) <= 60000]]]";

    EVT_HANDLE queryHandle = EvtQuery(nullptr,    // Local machine
                                      queryPath,  // Application log
                                      query, EvtQueryChannelPath);
    if (!queryHandle) {
        Wh_Log(L"EvtQuery failed with error: %u", GetLastError());
        return false;
    }

    EVT_HANDLE events[2] = {};
    DWORD returned = 0;
    constexpr DWORD kTimeout = 1000;
    BOOL ok =
        EvtNext(queryHandle, ARRAYSIZE(events), events, kTimeout, 0, &returned);
    if (!ok && GetLastError() != ERROR_NO_MORE_ITEMS) {
        Wh_Log(L"EvtNext failed with error: %u", GetLastError());
    }
    for (DWORD i = 0; i < returned; i++) {
        EvtClose(events[i]);
    }

    EvtClose(queryHandle);
    return ok && returned >= ARRAYSIZE(events);
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    if (HasMultipleDwminitWarningsInLastMinute()) {
        Wh_Log(L"Refusing to load: multiple recent Dwminit warnings");
        return FALSE;
    }

    LoadSettings();

    HMODULE udwm = GetModuleHandle(L"udwm.dll");
    if (!udwm) {
        Wh_Log(L"udwm.dll isn't loaded");
        return FALSE;
    }

    // Call tree for corner radius in each version:
    //
    // Old builds (e.g. 10.0.22621.6199):
    //   UpdateWindowVisuals
    //     -> GetEffectiveCornerStyle (inlined radius mapping, 8.0/4.0)
    //     -> SetBorderParameters (receives radius as param)
    //   CTopLevelWindow3D::UpdateAnimatedResources
    //     -> GetRadiusFromCornerStyle (DPI scaling inlined)
    //     -> ResourceHelper::CreateRectangleGeometry
    //
    // New builds (e.g. 10.0.26100.7920):
    //   UpdateWindowVisuals
    //     -> GetFloatCornerRadiusForCurrentStyle
    //       -> GetRadiusFromCornerStyle
    //     -> SetBorderParameters (receives radius as param)
    //   CTopLevelWindow3D::UpdateAnimatedResources
    //     -> GetDpiAdjustedFloatCornerRadius
    //       -> GetRadiusFromCornerStyle
    //     -> ResourceHelper::CreateRectangleGeometry
    //
    // In new builds, the squaring of maximized and snapped windows sits in
    // GetFloatCornerRadiusForCurrentStyle, which returns zero under the same
    // condition IsMaximizedOrSnapped tests, without consulting the corner style
    // at all. The animation path isn't squared, so restoring the rounding only
    // takes replacing that zero.

    WindhawkUtils::SYMBOL_HOOK udwmDllHooks[] = {
        // Used to recover the HWND for tooltip detection. Returns the
        // CWindowData* stored on the CTopLevelWindow. Capture only, no hook.
        {
            {LR"(public: class CWindowData * __cdecl CTopLevelWindow::GetWindowData(void)const )"},
            &GetWindowData_Original,
            nullptr,
            true,  // Optional - tooltip detection is skipped if missing.
        },
        // Used to derive the HWND member offset on CWindowData at runtime by
        // disassembling the function's first `mov rcx, [rcx+disp]`. Capture
        // only, no hook.
        {
            {LR"(public: bool __cdecl CWindowData::IsGhostWindow(struct HWND__ * *)const )"},
            &IsGhostWindow_Func,
            nullptr,
            true,  // Optional - tooltip detection is skipped if missing.
        },
        // DWM's own answer about whether the surface being composed is a
        // maximized or snapped window. Every replacement made for such windows
        // is gated on it. Capture only, no hook.
        {
            {LR"(public: bool __cdecl CTopLevelWindow::IsMaximizedOrSnapped(void)const )"},
            &IsMaximizedOrSnapped_Original,
            nullptr,
            true,  // Optional - maximized/snapped rounding is skipped.
        },
        // Skips visual updates for SysShadow companion windows so the legacy
        // rectangular drop shadow doesn't poke out beside the rounded tooltip.
        {
            {LR"(private: long __cdecl CTopLevelWindow::UpdateWindowVisuals(void))"},
            &UpdateWindowVisuals_Original,
            UpdateWindowVisuals_Hook,
            true,  // Optional - SysShadow remains visible if missing.
        },
        // Used to promote tooltips from "no rounding" to "round small" so the
        // full DWM rounding pipeline kicks in (border + shadow + clip), instead
        // of us trying to force a radius onto a window DWM thinks is square.
        // Also consulted by ShouldRoundMaximizedOrSnapped.
        {
            {LR"(private: enum CORNER_STYLE __cdecl CTopLevelWindow::GetEffectiveCornerStyle(void))"},
            &GetEffectiveCornerStyle_Original,
            GetEffectiveCornerStyle_Hook,
            true,  // Optional - tooltip rounding skipped if missing.
        },
        // Covers the 3D animation path in both old and new builds.
        {
            {LR"(private: float __cdecl CTopLevelWindow::GetRadiusFromCornerStyle(void))"},
            &GetRadiusFromCornerStyle_Original,
            GetRadiusFromCornerStyle_Hook,
        },
        // Covers UpdateWindowVisuals in new builds, and with it the squaring of
        // maximized and snapped windows. Calls GetRadiusFromCornerStyle, but is
        // hooked separately since it zeroes the radius of its own accord and in
        // case a future build inlines that call. The access specifier differs
        // between builds, so both are listed.
        {
            {
                LR"(public: float __cdecl CTopLevelWindow::GetFloatCornerRadiusForCurrentStyle(void))",

                // Older Windows 11 builds:
                LR"(private: float __cdecl CTopLevelWindow::GetFloatCornerRadiusForCurrentStyle(void))",
            },
            &GetFloatCornerRadiusForCurrentStyle_Original,
            GetFloatCornerRadiusForCurrentStyle_Hook,
            true,  // Missing in earlier builds (e.g. 10.0.22621.6199).
        },
        // Covers UpdateWindowVisuals in old builds where the radius is
        // computed inline (no call to GetRadiusFromCornerStyle) and passed
        // directly to this function.
        {
            {LR"(public: long __cdecl CWindowBorder::SetBorderParameters(struct tagRECT const &,float,int,struct _D3DCOLORVALUE const &,enum CWindowBorder::BorderStyle,enum CWindowBorder::ShadowStyle))"},
            &SetBorderParameters_Original,
            SetBorderParameters_Hook,
        },
    };

    if (!HookSymbols(udwm, udwmDllHooks, ARRAYSIZE(udwmDllHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
    }

    // Hooks queued by HookSymbols aren't applied until Wh_ModInit returns, so
    // IsGhostWindow's bytes are still original here. Disassemble its prologue
    // to recover the HWND member offset on CWindowData. The function loads
    // this->hwnd as the first argument to GetPropW, so the first load whose
    // base register is the `this` pointer (rcx on x64, x0 on ARM64) is the
    // HWND member - the destination register is left unconstrained because
    // compilers may stage the value through a scratch register first.
    if (IsGhostWindow_Func) {
        g_windowDataHwndOffset = OffsetFromAssemblyRegex(
            IsGhostWindow_Func, SIZE_MAX,
#if defined(_M_X64)
            std::regex(R"(mov \w+, \[rcx\+0x([0-9a-f]+)\])",
                       std::regex_constants::icase),
#elif defined(_M_ARM64)
            std::regex(R"(ldr\s+\w+, \[x0, #0x([0-9a-f]+)\])",
                       std::regex_constants::icase),
#else
#error "Unsupported architecture"
#endif
            10);
        Wh_Log(L"windowDataHwndOffset=0x%zx", g_windowDataHwndOffset);
    } else {
        Wh_Log(L"IsGhostWindow wasn't found, HWND lookup is disabled");
    }

    return TRUE;
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();

    ClearWindowExclusionProps();
}

void Wh_ModUninit() {
    Wh_Log(L">");

    ClearWindowExclusionProps();
}
