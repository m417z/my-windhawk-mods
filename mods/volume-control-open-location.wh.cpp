// ==WindhawkMod==
// @id              volume-control-open-location
// @name            Volume control open location
// @description     Shows the volume control on the monitor where the mouse cursor is located, or on a custom monitor of choice
// @version         1.0
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lshcore
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
# Volume control open location

Shows the volume control on the monitor where the mouse cursor is located, or on
a custom monitor of choice.

## Selecting a monitor

You can select a monitor by its number or by its interface name in the mod
settings.

### By monitor number

Set the **Monitor** setting to the desired monitor number (1, 2, 3, etc.). Note
that this number may differ from the monitor number shown in Windows Display
Settings.

Set the **Monitor** setting to 0 to use the monitor where the mouse cursor is
currently located.

### By interface name

If monitor numbers change frequently (e.g., after locking your PC or
restarting), you can use the monitor's interface name instead. To find the
interface name:

1. Go to the mod's **Advanced** tab.
2. Set **Debug logging** to **Mod logs**.
3. Click on **Show log output**.
4. In the mod settings, enter any text (e.g., `TEST`) in the **Monitor interface
   name** field.
5. Change the volume to trigger the mod's logic.
6. In the log output, look for lines containing `Found display device`. You will
   see one line per monitor, for example:
   ```
   Found display device \\.\DISPLAY1, interface name: \\?\DISPLAY#DELA1D2#5&abc123#0#{e6f07b5f-ee97-4a90-b076-33f57bf4eaa7}
   Found display device \\.\DISPLAY2, interface name: \\?\DISPLAY#GSM5B09#4&def456#0#{e6f07b5f-ee97-4a90-b076-33f57bf4eaa7}
   ```
   Use the interface name that follows the "interface name:" text. You may need
   to experiment to determine which interface name corresponds to which physical
   monitor.
7. Copy the relevant interface name (or a unique substring of it) into the
   **Monitor interface name** setting.
8. Set **Debug logging** back to **None** when done.

The **Monitor interface name** setting takes priority over the **Monitor**
number when both are configured.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- monitor: 0
  $name: Monitor
  $description: >-
    The monitor number that the volume control will appear on. Set to zero to
    use the monitor where the mouse cursor is located.
- monitorInterfaceName: ""
  $name: Monitor interface name
  $description: >-
    If not empty, the given monitor interface name (can also be an interface
    name substring) will be used instead of the monitor number. Can be useful if
    the monitor numbers change often. To see all available interface names, set
    any interface name, enable mod logs, change the volume and look for "Found
    display device" messages.
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <shellscalingapi.h>

#include <string>

struct {
    int monitor;
    WindhawkUtils::StringSetting monitorInterfaceName;
} g_settings;

HMODULE g_hardwareConfirmatorModule;

HMONITOR GetMonitorById(int monitorId) {
    HMONITOR monitorResult = nullptr;
    int currentMonitorId = 0;

    auto monitorEnumProc = [&](HMONITOR hMonitor) -> BOOL {
        if (currentMonitorId == monitorId) {
            monitorResult = hMonitor;
            return FALSE;
        }
        currentMonitorId++;
        return TRUE;
    };

    EnumDisplayMonitors(
        nullptr, nullptr,
        [](HMONITOR hMonitor, HDC hdc, LPRECT lprcMonitor,
           LPARAM dwData) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(monitorEnumProc)*>(dwData);
            return proc(hMonitor);
        },
        reinterpret_cast<LPARAM>(&monitorEnumProc));

    return monitorResult;
}

using MonitorFromPoint_t = decltype(&MonitorFromPoint);
MonitorFromPoint_t MonitorFromPoint_Original;

using EnumDisplayDevicesW_t = decltype(&EnumDisplayDevicesW);
EnumDisplayDevicesW_t EnumDisplayDevicesW_Original;

HMONITOR GetMonitorByInterfaceNameSubstr(PCWSTR interfaceNameSubstr) {
    HMONITOR monitorResult = nullptr;

    auto monitorEnumProc = [&](HMONITOR hMonitor) -> BOOL {
        MONITORINFOEX monitorInfo = {};
        monitorInfo.cbSize = sizeof(monitorInfo);

        if (GetMonitorInfo(hMonitor, &monitorInfo)) {
            DISPLAY_DEVICE displayDevice = {
                .cb = sizeof(displayDevice),
            };

            if (EnumDisplayDevicesW_Original(monitorInfo.szDevice, 0,
                                             &displayDevice,
                                             EDD_GET_DEVICE_INTERFACE_NAME)) {
                Wh_Log(L"Found display device %s, interface name: %s",
                       monitorInfo.szDevice, displayDevice.DeviceID);

                if (wcsstr(displayDevice.DeviceID, interfaceNameSubstr)) {
                    Wh_Log(L"Matched display device");
                    monitorResult = hMonitor;
                    return FALSE;
                }
            }
        }
        return TRUE;
    };

    EnumDisplayMonitors(
        nullptr, nullptr,
        [](HMONITOR hMonitor, HDC hdc, LPRECT lprcMonitor,
           LPARAM dwData) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(monitorEnumProc)*>(dwData);
            return proc(hMonitor);
        },
        reinterpret_cast<LPARAM>(&monitorEnumProc));

    return monitorResult;
}

HMONITOR GetDestMonitor() {
    if (*g_settings.monitorInterfaceName.get()) {
        return GetMonitorByInterfaceNameSubstr(
            g_settings.monitorInterfaceName.get());
    } else if (g_settings.monitor == 0) {
        POINT pt;
        GetCursorPos(&pt);
        return MonitorFromPoint_Original(pt, MONITOR_DEFAULTTONEAREST);
    } else if (g_settings.monitor >= 1) {
        return GetMonitorById(g_settings.monitor - 1);
    }

    return nullptr;
}

HMODULE GetModuleFromAddress(void* address) {
    HMODULE module;
    if (!GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                               GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           (PCWSTR)address, &module)) {
        return nullptr;
    }

    return module;
}

std::wstring GetModulePath(HMODULE module) {
    if (!module) {
        return L"<unknown>";
    }

    std::wstring path(MAX_PATH, L'\0');
    while (true) {
        DWORD len = GetModuleFileName(module, path.data(), path.size());
        if (len == 0) {
            return L"<unknown>";
        }

        // A result equal to the buffer size means the path was truncated.
        if (len == path.size()) {
            path.resize(len * 2);
            continue;
        }

        path.resize(len);
        return path;
    }
}

// Whether the path is under the Windows directory.
bool IsSystemModulePath(PCWSTR path) {
    WCHAR windowsDir[MAX_PATH];
    UINT len = GetSystemWindowsDirectory(windowsDir, ARRAYSIZE(windowsDir));
    if (len == 0 || len >= ARRAYSIZE(windowsDir)) {
        return false;
    }

    return _wcsnicmp(path, windowsDir, len) == 0 && path[len] == L'\\';
}

// Another hook on top of ours makes its hook function the direct caller, so a
// few frames further up the stack are checked as well. A hook isn't in a system
// module, so the search stops at the first frame in one.
[[clang::noinline]] bool IsHookCallerFromModule(void* retAddress,
                                                PCWSTR moduleName) {
    HMODULE expectedModule = GetModuleHandle(moduleName);
    if (!expectedModule) {
        return false;
    }

    HMODULE callerModule = GetModuleFromAddress(retAddress);
    if (callerModule == expectedModule) {
        return true;
    }

    std::wstring callerPath = GetModulePath(callerModule);
    if (IsSystemModulePath(callerPath.c_str())) {
        Wh_Log(L"Skipping caller %p in module %s, expected %s", retAddress,
               callerPath.c_str(), moduleName);
        return false;
    }

    Wh_Log(L"Tracing caller %p in module %s, expected %s", retAddress,
           callerPath.c_str(), moduleName);

    // The backtrace skips the frames of this function, the hook, and the
    // caller.
    void* frames[4];
    WORD count = CaptureStackBackTrace(3, ARRAYSIZE(frames), frames, nullptr);
    for (WORD i = 0; i < count; i++) {
        HMODULE module = GetModuleFromAddress(frames[i]);
        std::wstring modulePath = GetModulePath(module);
        Wh_Log(L"Frame %u: %p in module %s", i + 1, frames[i],
               modulePath.c_str());
        if (module == expectedModule) {
            return true;
        }

        if (IsSystemModulePath(modulePath.c_str())) {
            return false;
        }
    }

    return false;
}

HMONITOR WINAPI MonitorFromPoint_Hook(POINT pt, DWORD dwFlags) {
    auto original = [=] { return MonitorFromPoint_Original(pt, dwFlags); };

    if (pt.x != 0 || pt.y != 0) {
        return original();
    }

    if (!IsHookCallerFromModule(__builtin_return_address(0),
                                L"Windows.Internal.HardwareConfirmator.dll")) {
        return original();
    }

    Wh_Log(L">");

    HMONITOR monitor = GetDestMonitor();
    if (!monitor) {
        return original();
    }

    return monitor;
}

using MonitorFromRect_t = decltype(&MonitorFromRect);
MonitorFromRect_t MonitorFromRect_Original;
HMONITOR WINAPI MonitorFromRect_Hook(LPCRECT lprc, DWORD dwFlags) {
    auto original = [=] { return MonitorFromRect_Original(lprc, dwFlags); };

    if (!lprc || lprc->left != 0 || lprc->top != 0 || lprc->right != 0 ||
        lprc->bottom != 0) {
        return original();
    }

    if (!IsHookCallerFromModule(__builtin_return_address(0),
                                L"Windows.Internal.HardwareConfirmator.dll")) {
        return original();
    }

    Wh_Log(L">");

    HMONITOR monitor = GetDestMonitor();
    if (!monitor) {
        return original();
    }

    return monitor;
}

BOOL WINAPI EnumDisplayDevicesW_Hook(LPCWSTR lpDevice,
                                     DWORD iDevNum,
                                     PDISPLAY_DEVICEW lpDisplayDevice,
                                     DWORD dwFlags) {
    BOOL result = EnumDisplayDevicesW_Original(lpDevice, iDevNum,
                                               lpDisplayDevice, dwFlags);

    if (!result || !lpDisplayDevice || lpDevice) {
        return result;
    }

    if (!IsHookCallerFromModule(__builtin_return_address(0),
                                L"Windows.Internal.HardwareConfirmator.dll")) {
        return result;
    }

    Wh_Log(L">");

    HMONITOR monitor = GetDestMonitor();
    if (!monitor) {
        return result;
    }

    MONITORINFOEX monitorInfo = {};
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (!GetMonitorInfo(monitor, &monitorInfo)) {
        return result;
    }

    if (wcscmp(lpDisplayDevice->DeviceName, monitorInfo.szDevice) == 0) {
        lpDisplayDevice->StateFlags |= DISPLAY_DEVICE_PRIMARY_DEVICE;
    } else {
        lpDisplayDevice->StateFlags &= ~DISPLAY_DEVICE_PRIMARY_DEVICE;
    }

    return result;
}

struct WinrtRect {
    float X;
    float Y;
    float Width;
    float Height;
};

using HardwareConfirmatorHost_GetPositionRect_t =
    WinrtRect*(WINAPI*)(void* pThis, WinrtRect* retval, const WinrtRect* rect);
HardwareConfirmatorHost_GetPositionRect_t
    HardwareConfirmatorHost_GetPositionRect_Original;
WinrtRect* WINAPI
HardwareConfirmatorHost_GetPositionRect_Hook(void* pThis,
                                             WinrtRect* retval,
                                             const WinrtRect* rect) {
    Wh_Log(L">");

    // Shift the input rect to 0,0 since the original function assumes that.
    WinrtRect shiftedRect = *rect;
    float offsetX = shiftedRect.X;
    float offsetY = shiftedRect.Y;
    shiftedRect.X = 0;
    shiftedRect.Y = 0;

    WinrtRect* result = HardwareConfirmatorHost_GetPositionRect_Original(
        pThis, retval, &shiftedRect);

    // Shift the result back.
    result->X += offsetX;
    result->Y += offsetY;

    return result;
}

using ScaleRelativePixelsForDevice_t = int(WINAPI*)(int deviceType,
                                                    float pixels);
ScaleRelativePixelsForDevice_t ScaleRelativePixelsForDevice_Original;
int WINAPI ScaleRelativePixelsForDevice_Hook(int deviceType, float pixels) {
    if (IsHookCallerFromModule(__builtin_return_address(0),
                               L"Windows.Internal.HardwareConfirmator.dll")) {
        HMONITOR monitor = GetDestMonitor();
        if (monitor) {
            Wh_Log(L">");

            UINT dpiX = 0, dpiY = 0;
            if (SUCCEEDED(GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX,
                                           &dpiY))) {
                return (int)(pixels * (float)dpiX / 96.0f);
            }
        }
    }

    return ScaleRelativePixelsForDevice_Original(deviceType, pixels);
}

void LoadSettings() {
    g_settings.monitor = Wh_GetIntSetting(L"monitor");
    g_settings.monitorInterfaceName =
        WindhawkUtils::StringSetting::make(L"monitorInterfaceName");
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    g_hardwareConfirmatorModule =
        LoadLibraryEx(L"Windows.Internal.HardwareConfirmator.dll", nullptr,
                      LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!g_hardwareConfirmatorModule) {
        Wh_Log(L"Couldn't load Windows.Internal.HardwareConfirmator.dll");
        return FALSE;
    }

    // Windows.Internal.HardwareConfirmator.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(private: struct winrt::Windows::Foundation::Rect __cdecl winrt::Windows::Internal::HardwareConfirmator::implementation::HardwareConfirmatorHost::GetPositionRect(struct winrt::Windows::Foundation::Rect const &))"},
            &HardwareConfirmatorHost_GetPositionRect_Original,
            HardwareConfirmatorHost_GetPositionRect_Hook,
        },
    };

    if (!HookSymbols(g_hardwareConfirmatorModule, symbolHooks,
                     ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
    }

    HMODULE shcoreModule = GetModuleHandle(L"Shcore.dll");
    if (shcoreModule) {
        auto pScaleRelativePixelsForDevice =
            (ScaleRelativePixelsForDevice_t)GetProcAddress(
                shcoreModule, MAKEINTRESOURCEA(222));
        if (pScaleRelativePixelsForDevice) {
            WindhawkUtils::SetFunctionHook(
                pScaleRelativePixelsForDevice,
                ScaleRelativePixelsForDevice_Hook,
                &ScaleRelativePixelsForDevice_Original);
        }
    }

    WindhawkUtils::SetFunctionHook(MonitorFromPoint, MonitorFromPoint_Hook,
                                   &MonitorFromPoint_Original);

    WindhawkUtils::SetFunctionHook(MonitorFromRect, MonitorFromRect_Hook,
                                   &MonitorFromRect_Original);

    WindhawkUtils::SetFunctionHook(EnumDisplayDevicesW,
                                   EnumDisplayDevicesW_Hook,
                                   &EnumDisplayDevicesW_Original);

    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L">");
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}
