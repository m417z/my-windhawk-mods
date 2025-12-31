// ==WindhawkMod==
// @id              taskbar-auto-hide-keyboard-only
// @name            Taskbar keyboard-only auto-hide
// @description     When taskbar auto-hide is enabled, the taskbar will only be unhidden with the keyboard, with an optional Ctrl+Esc sticky toggle
// @version         1.1.3
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -ldwmapi -lversion
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
# Taskbar keyboard-only auto-hide

When taskbar auto-hide is enabled, the taskbar will only be unhidden with the
keyboard, hovering the mouse over the taskbar will not unhide it. For example,
the Win key or Ctrl+Esc will unhide the taskbar.

Optional: enable the Ctrl+Esc sticky toggle setting to lock taskbar visibility
until you press Ctrl+Esc again.

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- fullyHide: false
  $name: Fully hide
  $description: >-
    Normally, the taskbar is hidden to a thin line which can be clicked to
    unhide it. This option makes it so that the taskbar is fully hidden on
    auto-hide, leaving no traces at all. With this option, the taskbar can only
    be unhidden via the keyboard.
- toggleOnHotkey: false
  $name: Ctrl+Esc sticky toggle
  $description: >-
    Press Ctrl+Esc to toggle and lock taskbar visibility. The taskbar stays
    shown or hidden until you press Ctrl+Esc again, and other keyboard unhide
    triggers won't override a forced hidden state. This setting requires
    auto-hide to be enabled in Windows settings.
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool).
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <dwmapi.h>
#include <psapi.h>

#include <atomic>
#include <vector>

struct {
    bool fullyHide;
    bool toggleOnHotkey;
    bool oldTaskbarOnWin11;
} g_settings;

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_24H2,
};

WinVersion g_winVersion;

std::atomic<bool> g_initialized;
std::atomic<bool> g_explorerPatcherInitialized;

enum class ToggleState {
    Default,
    ForcedShown,
    ForcedHidden,
};

std::atomic<ToggleState> g_toggleState{ToggleState::Default};
std::atomic<DWORD> g_lastToggleTick{0};

enum {
    kTrayUITimerHide = 2,
    kTrayUITimerUnhide = 3,
};

constexpr DWORD kToggleDebounceMs = 250;

bool IsTaskbarWindow(HWND hWnd) {
    WCHAR szClassName[32];
    if (!GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName))) {
        return false;
    }

    return _wcsicmp(szClassName, L"Shell_TrayWnd") == 0 ||
           _wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0;
}

HWND FindCurrentProcessTaskbarWnd() {
    HWND hTaskbarWnd = nullptr;

    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            DWORD dwProcessId;
            WCHAR className[32];
            if (GetWindowThreadProcessId(hWnd, &dwProcessId) &&
                dwProcessId == GetCurrentProcessId() &&
                GetClassName(hWnd, className, ARRAYSIZE(className)) &&
                _wcsicmp(className, L"Shell_TrayWnd") == 0) {
                *reinterpret_cast<HWND*>(lParam) = hWnd;
                return FALSE;
            }
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&hTaskbarWnd));

    return hTaskbarWnd;
}

HWND FindTaskbarWindows(std::vector<HWND>* secondaryTaskbarWindows) {
    secondaryTaskbarWindows->clear();

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (!hTaskbarWnd) {
        return nullptr;
    }

    DWORD taskbarThreadId = GetWindowThreadProcessId(hTaskbarWnd, nullptr);
    if (!taskbarThreadId) {
        return nullptr;
    }

    auto enumWindowsProc = [&secondaryTaskbarWindows](HWND hWnd) -> BOOL {
        WCHAR szClassName[32];
        if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
            return TRUE;
        }

        if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
            secondaryTaskbarWindows->push_back(hWnd);
        }

        return TRUE;
    };

    EnumThreadWindows(
        taskbarThreadId,
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(enumWindowsProc)*>(lParam);
            return proc(hWnd);
        },
        reinterpret_cast<LPARAM>(&enumWindowsProc));

    return hTaskbarWnd;
}

void CloakWindow(HWND hWnd, BOOL cloak) {
    DwmSetWindowAttribute(hWnd, DWMWA_CLOAK, &cloak, sizeof(cloak));
}

bool IsTaskbarShown(HWND hWnd) {
    if (!hWnd) {
        return false;
    }

    if (g_settings.fullyHide) {
        BOOL cloaked = FALSE;
        if (SUCCEEDED(DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, &cloaked,
                                            sizeof(cloaked))) &&
            cloaked) {
            return false;
        }
    }

    RECT rect{};
    if (!GetWindowRect(hWnd, &rect)) {
        return false;
    }

    HMONITOR monitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
    MONITORINFO monitorInfo{
        .cbSize = sizeof(monitorInfo),
    };
    if (!GetMonitorInfo(monitor, &monitorInfo)) {
        return false;
    }

    RECT intersect{};
    if (!IntersectRect(&intersect, &rect, &monitorInfo.rcMonitor)) {
        return false;
    }

    int width = intersect.right - intersect.left;
    int height = intersect.bottom - intersect.top;
    if (width <= 0 || height <= 0) {
        return false;
    }

    int taskbarWidth = rect.right - rect.left;
    int taskbarHeight = rect.bottom - rect.top;
    int taskbarThickness =
        taskbarWidth < taskbarHeight ? taskbarWidth : taskbarHeight;
    if (taskbarThickness <= 0) {
        return false;
    }

    int visibleThickness = width < height ? width : height;
    return visibleThickness * 2 >= taskbarThickness;
}

bool AreAnyTaskbarsShown() {
    std::vector<HWND> secondaryTaskbarWindows;
    HWND taskbarWindow = FindTaskbarWindows(&secondaryTaskbarWindows);
    if (!taskbarWindow) {
        return false;
    }

    if (IsTaskbarShown(taskbarWindow)) {
        return true;
    }

    for (HWND hWnd : secondaryTaskbarWindows) {
        if (IsTaskbarShown(hWnd)) {
            return true;
        }
    }

    return false;
}

void HideAllTaskbars() {
    std::vector<HWND> secondaryTaskbarWindows;
    HWND taskbarWindow = FindTaskbarWindows(&secondaryTaskbarWindows);
    if (taskbarWindow) {
        SetTimer(taskbarWindow, kTrayUITimerHide, 0, nullptr);
    }

    for (HWND hWnd : secondaryTaskbarWindows) {
        SetTimer(hWnd, kTrayUITimerHide, 0, nullptr);
    }
}

void CancelHideTimers() {
    std::vector<HWND> secondaryTaskbarWindows;
    HWND taskbarWindow = FindTaskbarWindows(&secondaryTaskbarWindows);
    if (taskbarWindow) {
        KillTimer(taskbarWindow, kTrayUITimerHide);
    }

    for (HWND hWnd : secondaryTaskbarWindows) {
        KillTimer(hWnd, kTrayUITimerHide);
    }
}

bool IsCtrlEscPressed() {
    bool ctrlDown =
        (GetAsyncKeyState(VK_CONTROL) & 0x8000) ||
        (GetAsyncKeyState(VK_LCONTROL) & 0x8000) ||
        (GetAsyncKeyState(VK_RCONTROL) & 0x8000);
    return ctrlDown && (GetAsyncKeyState(VK_ESCAPE) & 0x8000);
}

using SetTimer_t = decltype(&SetTimer);
SetTimer_t SetTimer_Original;
UINT_PTR WINAPI SetTimer_Hook(HWND hWnd,
                              UINT_PTR nIDEvent,
                              UINT uElapse,
                              TIMERPROC lpTimerFunc) {
    if (nIDEvent == kTrayUITimerHide && IsTaskbarWindow(hWnd) &&
        g_settings.toggleOnHotkey &&
        g_toggleState.load(std::memory_order_relaxed) ==
            ToggleState::ForcedShown) {
        Wh_Log(L">");
        return 1;
    }

    if (nIDEvent == kTrayUITimerUnhide && IsTaskbarWindow(hWnd)) {
        Wh_Log(L">");
        return 1;
    }

    UINT_PTR ret = SetTimer_Original(hWnd, nIDEvent, uElapse, lpTimerFunc);

    return ret;
}

enum class ToggleAction {
    Show,
    Hide,
};

struct ToggleDecision {
    ToggleAction action;
    bool isNewToggle;
};

ToggleDecision DecideToggleAction(DWORD now) {
    ToggleState state = g_toggleState.load(std::memory_order_relaxed);
    DWORD lastTick = g_lastToggleTick.load(std::memory_order_relaxed);
    if (lastTick && now - lastTick < kToggleDebounceMs) {
        if (state == ToggleState::ForcedShown) {
            return {ToggleAction::Show, false};
        }
        return {ToggleAction::Hide, false};
    }

    bool shouldHide = false;
    if (state == ToggleState::ForcedShown) {
        shouldHide = true;
    } else if (state == ToggleState::ForcedHidden) {
        shouldHide = false;
    } else {
        shouldHide = AreAnyTaskbarsShown();
    }

    return {shouldHide ? ToggleAction::Hide : ToggleAction::Show, true};
}

using TrayUI_Unhide_t = void(WINAPI*)(void* pThis,
                                      int trayUnhideFlags,
                                      int unhideRequest);
TrayUI_Unhide_t TrayUI_Unhide_Original;
void WINAPI TrayUI_Unhide_Hook(void* pThis,
                               int trayUnhideFlags,
                               int unhideRequest) {
    if (!g_settings.toggleOnHotkey) {
        TrayUI_Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
        return;
    }

    if (!IsCtrlEscPressed()) {
        if (g_toggleState.load(std::memory_order_relaxed) ==
            ToggleState::ForcedHidden) {
            return;
        }

        TrayUI_Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
        return;
    }

    DWORD now = GetTickCount();
    ToggleDecision decision = DecideToggleAction(now);
    if (decision.action == ToggleAction::Hide) {
        if (decision.isNewToggle) {
            g_toggleState.store(ToggleState::ForcedHidden,
                                std::memory_order_relaxed);
            g_lastToggleTick.store(now, std::memory_order_relaxed);
        }
        HideAllTaskbars();
        return;
    }

    if (decision.isNewToggle) {
        g_toggleState.store(ToggleState::ForcedShown,
                            std::memory_order_relaxed);
        g_lastToggleTick.store(now, std::memory_order_relaxed);
    }

    CancelHideTimers();
    TrayUI_Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
}

using CSecondaryTray__Unhide_t = void(WINAPI*)(void* pThis,
                                               int trayUnhideFlags,
                                               int unhideRequest);
CSecondaryTray__Unhide_t CSecondaryTray__Unhide_Original;
void WINAPI CSecondaryTray__Unhide_Hook(void* pThis,
                                        int trayUnhideFlags,
                                        int unhideRequest) {
    if (!g_settings.toggleOnHotkey) {
        CSecondaryTray__Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
        return;
    }

    if (!IsCtrlEscPressed()) {
        if (g_toggleState.load(std::memory_order_relaxed) ==
            ToggleState::ForcedHidden) {
            return;
        }

        CSecondaryTray__Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
        return;
    }

    DWORD now = GetTickCount();
    ToggleDecision decision = DecideToggleAction(now);
    if (decision.action == ToggleAction::Hide) {
        if (decision.isNewToggle) {
            g_toggleState.store(ToggleState::ForcedHidden,
                                std::memory_order_relaxed);
            g_lastToggleTick.store(now, std::memory_order_relaxed);
        }
        HideAllTaskbars();
        return;
    }

    if (decision.isNewToggle) {
        g_toggleState.store(ToggleState::ForcedShown,
                            std::memory_order_relaxed);
        g_lastToggleTick.store(now, std::memory_order_relaxed);
    }

    CancelHideTimers();
    CSecondaryTray__Unhide_Original(pThis, trayUnhideFlags, unhideRequest);
}

using TrayUI__Hide_t = void(WINAPI*)(void* pThis);
TrayUI__Hide_t TrayUI__Hide_Original;
void WINAPI TrayUI__Hide_Hook(void* pThis) {
    if (g_settings.toggleOnHotkey &&
        g_toggleState.load(std::memory_order_relaxed) ==
            ToggleState::ForcedShown) {
        CancelHideTimers();
        return;
    }

    TrayUI__Hide_Original(pThis);
}

using CSecondaryTray__AutoHide_t = void(WINAPI*)(void* pThis, bool param1);
CSecondaryTray__AutoHide_t CSecondaryTray__AutoHide_Original;
void WINAPI CSecondaryTray__AutoHide_Hook(void* pThis, bool param1) {
    if (g_settings.toggleOnHotkey &&
        g_toggleState.load(std::memory_order_relaxed) ==
            ToggleState::ForcedShown) {
        CancelHideTimers();
        return;
    }

    CSecondaryTray__AutoHide_Original(pThis, param1);
}

using TrayUI_SlideWindow_t = void(WINAPI*)(void* pThis,
                                           HWND hWnd,
                                           const RECT* rc,
                                           HMONITOR monitor,
                                           bool show,
                                           bool flag);
TrayUI_SlideWindow_t TrayUI_SlideWindow_Original;
void WINAPI TrayUI_SlideWindow_Hook(void* pThis,
                                    HWND hWnd,
                                    const RECT* rc,
                                    HMONITOR monitor,
                                    bool show,
                                    bool flag) {
    Wh_Log(L">");

    if (show && g_settings.fullyHide) {
        CloakWindow(hWnd, FALSE);
    }

    TrayUI_SlideWindow_Original(pThis, hWnd, rc, monitor, show, flag);

    if (!show && g_settings.fullyHide) {
        CloakWindow(hWnd, TRUE);
    }
}

bool HookTaskbarSymbols() {
    HMODULE module;
    if (g_winVersion <= WinVersion::Win10) {
        module = GetModuleHandle(nullptr);
    } else {
        module = LoadLibraryEx(L"taskbar.dll", nullptr,
                               LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (!module) {
            Wh_Log(L"Couldn't load taskbar.dll");
            return false;
        }
    }

    // Taskbar.dll, explorer.exe
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: virtual void __cdecl TrayUI::SlideWindow(struct HWND__ *,struct tagRECT const *,struct HMONITOR__ *,bool,bool))"},
            &TrayUI_SlideWindow_Original,
            TrayUI_SlideWindow_Hook,
        },
        {
            {LR"(public: void __cdecl TrayUI::_Hide(void))"},
            &TrayUI__Hide_Original,
            TrayUI__Hide_Hook,
            true,
        },
        {
            {LR"(private: void __cdecl CSecondaryTray::_AutoHide(bool))"},
            &CSecondaryTray__AutoHide_Original,
            CSecondaryTray__AutoHide_Hook,
            true,
        },
        {
            {LR"(public: virtual void __cdecl TrayUI::Unhide(enum TrayCommon::TrayUnhideFlags,enum TrayCommon::UnhideRequest))"},
            &TrayUI_Unhide_Original,
            TrayUI_Unhide_Hook,
            true,
        },
        {
            {LR"(private: void __cdecl CSecondaryTray::_Unhide(enum TrayCommon::TrayUnhideFlags,enum TrayCommon::UnhideRequest))"},
            &CSecondaryTray__Unhide_Original,
            CSecondaryTray__Unhide_Hook,
            true,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
}

VS_FIXEDFILEINFO* GetModuleVersionInfo(HMODULE hModule, UINT* puPtrLen) {
    void* pFixedFileInfo = nullptr;
    UINT uPtrLen = 0;

    HRSRC hResource =
        FindResource(hModule, MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION);
    if (hResource) {
        HGLOBAL hGlobal = LoadResource(hModule, hResource);
        if (hGlobal) {
            void* pData = LockResource(hGlobal);
            if (pData) {
                if (!VerQueryValue(pData, L"\\", &pFixedFileInfo, &uPtrLen) ||
                    uPtrLen == 0) {
                    pFixedFileInfo = nullptr;
                    uPtrLen = 0;
                }
            }
        }
    }

    if (puPtrLen) {
        *puPtrLen = uPtrLen;
    }

    return (VS_FIXEDFILEINFO*)pFixedFileInfo;
}

WinVersion GetExplorerVersion() {
    VS_FIXEDFILEINFO* fixedFileInfo = GetModuleVersionInfo(nullptr, nullptr);
    if (!fixedFileInfo) {
        return WinVersion::Unsupported;
    }

    WORD major = HIWORD(fixedFileInfo->dwFileVersionMS);
    WORD minor = LOWORD(fixedFileInfo->dwFileVersionMS);
    WORD build = HIWORD(fixedFileInfo->dwFileVersionLS);
    WORD qfe = LOWORD(fixedFileInfo->dwFileVersionLS);

    Wh_Log(L"Version: %u.%u.%u.%u", major, minor, build, qfe);

    switch (major) {
        case 10:
            if (build < 22000) {
                return WinVersion::Win10;
            } else if (build < 26100) {
                return WinVersion::Win11;
            } else {
                return WinVersion::Win11_24H2;
            }
            break;
    }

    return WinVersion::Unsupported;
}

struct EXPLORER_PATCHER_HOOK {
    PCSTR symbol;
    void** pOriginalFunction;
    void* hookFunction = nullptr;
    bool optional = false;

    template <typename Prototype>
    EXPLORER_PATCHER_HOOK(
        PCSTR symbol,
        Prototype** originalFunction,
        std::type_identity_t<Prototype*> hookFunction = nullptr,
        bool optional = false)
        : symbol(symbol),
          pOriginalFunction(reinterpret_cast<void**>(originalFunction)),
          hookFunction(reinterpret_cast<void*>(hookFunction)),
          optional(optional) {}
};

bool HookExplorerPatcherSymbols(HMODULE explorerPatcherModule) {
    if (g_explorerPatcherInitialized.exchange(true)) {
        return true;
    }

    if (g_winVersion >= WinVersion::Win11) {
        g_winVersion = WinVersion::Win10;
    }

    EXPLORER_PATCHER_HOOK hooks[] = {
        {R"(?SlideWindow@TrayUI@@UEAAXPEAUHWND__@@PEBUtagRECT@@PEAUHMONITOR__@@_N3@Z)",
         &TrayUI_SlideWindow_Original, TrayUI_SlideWindow_Hook},
        {R"(?_Hide@TrayUI@@QEAAXXZ)", &TrayUI__Hide_Original,
         TrayUI__Hide_Hook, true},
        {R"(?_AutoHide@CSecondaryTray@@AEAAX_N@Z)",
         &CSecondaryTray__AutoHide_Original, CSecondaryTray__AutoHide_Hook,
         true},
        {R"(?Unhide@TrayUI@@UEAAXW4TrayUnhideFlags@TrayCommon@@W4UnhideRequest@3@@Z)",
         &TrayUI_Unhide_Original, TrayUI_Unhide_Hook, true},
        {R"(?_Unhide@CSecondaryTray@@AEAAXW4TrayUnhideFlags@TrayCommon@@W4UnhideRequest@3@@Z)",
         &CSecondaryTray__Unhide_Original, CSecondaryTray__Unhide_Hook, true},
    };

    bool succeeded = true;

    for (const auto& hook : hooks) {
        void* ptr = (void*)GetProcAddress(explorerPatcherModule, hook.symbol);
        if (!ptr) {
            Wh_Log(L"ExplorerPatcher symbol%s doesn't exist: %S",
                   hook.optional ? L" (optional)" : L"", hook.symbol);
            if (!hook.optional) {
                succeeded = false;
            }
            continue;
        }

        if (hook.hookFunction) {
            Wh_SetFunctionHook(ptr, hook.hookFunction, hook.pOriginalFunction);
        } else {
            *hook.pOriginalFunction = ptr;
        }
    }

    if (!succeeded) {
        Wh_Log(L"HookExplorerPatcherSymbols failed");
    } else if (g_initialized) {
        Wh_ApplyHookOperations();
    }

    return succeeded;
}

bool IsExplorerPatcherModule(HMODULE module) {
    WCHAR moduleFilePath[MAX_PATH];
    switch (
        GetModuleFileName(module, moduleFilePath, ARRAYSIZE(moduleFilePath))) {
        case 0:
        case ARRAYSIZE(moduleFilePath):
            return false;
    }

    PCWSTR moduleFileName = wcsrchr(moduleFilePath, L'\\');
    if (!moduleFileName) {
        return false;
    }

    moduleFileName++;

    if (_wcsnicmp(L"ep_taskbar.", moduleFileName, sizeof("ep_taskbar.") - 1) ==
        0) {
        Wh_Log(L"ExplorerPatcher taskbar module: %s", moduleFileName);
        return true;
    }

    return false;
}

bool HandleLoadedExplorerPatcher() {
    HMODULE hMods[1024];
    DWORD cbNeeded;
    if (EnumProcessModules(GetCurrentProcess(), hMods, sizeof(hMods),
                           &cbNeeded)) {
        for (size_t i = 0; i < cbNeeded / sizeof(HMODULE); i++) {
            if (IsExplorerPatcherModule(hMods[i])) {
                return HookExplorerPatcherSymbols(hMods[i]);
            }
        }
    }

    return true;
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module && !((ULONG_PTR)module & 3) && !g_explorerPatcherInitialized) {
        if (IsExplorerPatcherModule(module)) {
            HookExplorerPatcherSymbols(module);
        }
    }

    return module;
}

void LoadSettings() {
    g_settings.fullyHide = Wh_GetIntSetting(L"fullyHide");
    g_settings.toggleOnHotkey = Wh_GetIntSetting(L"toggleOnHotkey");
    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();
    g_toggleState.store(ToggleState::Default, std::memory_order_relaxed);
    g_lastToggleTick.store(0, std::memory_order_relaxed);

    g_winVersion = GetExplorerVersion();
    if (g_winVersion == WinVersion::Unsupported) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
    }

    if (g_settings.oldTaskbarOnWin11) {
        bool hasWin10Taskbar = g_winVersion < WinVersion::Win11_24H2;

        if (g_winVersion >= WinVersion::Win11) {
            g_winVersion = WinVersion::Win10;
        }

        if (hasWin10Taskbar && !HookTaskbarSymbols()) {
            return FALSE;
        }
    } else if (!HookTaskbarSymbols()) {
        return FALSE;
    }

    if (!HandleLoadedExplorerPatcher()) {
        Wh_Log(L"HandleLoadedExplorerPatcher failed");
        return FALSE;
    }

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(
        kernelBaseModule, "LoadLibraryExW");
    WindhawkUtils::Wh_SetFunctionHookT(pKernelBaseLoadLibraryExW,
                                       LoadLibraryExW_Hook,
                                       &LoadLibraryExW_Original);

    WindhawkUtils::Wh_SetFunctionHookT(SetTimer, SetTimer_Hook,
                                       &SetTimer_Original);

    g_initialized = true;

    return TRUE;
}

void CloakAllTaskbars(BOOL cloak) {
    std::vector<HWND> secondaryTaskbarWindows;
    HWND taskbarWindow = FindTaskbarWindows(&secondaryTaskbarWindows);
    CloakWindow(taskbarWindow, cloak);
    for (HWND hWnd : secondaryTaskbarWindows) {
        CloakWindow(hWnd, cloak);
    }
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    // Try again in case there's a race between the previous attempt and the
    // LoadLibraryExW hook.
    if (!g_explorerPatcherInitialized) {
        HandleLoadedExplorerPatcher();
    }

    if (g_settings.fullyHide) {
        CloakAllTaskbars(TRUE);
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");

    if (g_settings.fullyHide) {
        CloakAllTaskbars(FALSE);
    }
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;
    bool prevFullyHide = g_settings.fullyHide;
    bool prevToggleOnHotkey = g_settings.toggleOnHotkey;

    LoadSettings();

    if (g_settings.toggleOnHotkey != prevToggleOnHotkey) {
        g_toggleState.store(ToggleState::Default,
                            std::memory_order_relaxed);
        g_lastToggleTick.store(0, std::memory_order_relaxed);
    }

    if (g_settings.fullyHide != prevFullyHide) {
        if (g_settings.fullyHide) {
            CloakAllTaskbars(TRUE);
        } else {
            CloakAllTaskbars(FALSE);
        }
    }

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11;

    return TRUE;
}
