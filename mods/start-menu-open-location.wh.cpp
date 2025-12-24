// ==WindhawkMod==
// @id              start-menu-open-location
// @name            Start menu open location
// @description     When clicking the Start button, opens the Start Menu on the monitor where the mouse cursor is located, or in a custom monitor of choice
// @version         1.0
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
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
# Start menu open location

When clicking the Start button, opens the Start Menu on the monitor where the
mouse cursor is located, or in a custom monitor of choice.

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
5. Open the Start menu to trigger the mod's logic.
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
    The monitor number that the start menu will appear on. Set to zero to use
    the monitor where the mouse cursor is located.
- monitorInterfaceName: ""
  $name: Monitor interface name
  $description: >-
    If not empty, the given monitor interface name (can also be an interface
    name substring) will be used instead of the monitor number. Can be useful if
    the monitor numbers change often. To see all available interface names, set
    any interface name, enable mod logs, open the start menu and look for "Found
    display device" messages.
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

struct {
    int monitor;
    WindhawkUtils::StringSetting monitorInterfaceName;
} g_settings;

thread_local bool g_inShowStartView;

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

HMONITOR GetMonitorByInterfaceNameSubstr(PCWSTR interfaceNameSubstr) {
    HMONITOR monitorResult = nullptr;

    auto monitorEnumProc = [&](HMONITOR hMonitor) -> BOOL {
        MONITORINFOEX monitorInfo = {};
        monitorInfo.cbSize = sizeof(monitorInfo);

        if (GetMonitorInfo(hMonitor, &monitorInfo)) {
            DISPLAY_DEVICE displayDevice = {
                .cb = sizeof(displayDevice),
            };

            if (EnumDisplayDevices(monitorInfo.szDevice, 0, &displayDevice,
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

using ImmersiveMonitorHelper_ConnectToMonitor_t = bool(WINAPI*)(void* pThis,
                                                                HWND hWnd,
                                                                POINT point);
ImmersiveMonitorHelper_ConnectToMonitor_t
    ImmersiveMonitorHelper_ConnectToMonitor_Original;

using XamlLauncher_ShowStartView_t =
    HRESULT(WINAPI*)(void* pThis,
                     int immersiveLauncherShowMethod,
                     int immersiveLauncherShowFlags);
XamlLauncher_ShowStartView_t XamlLauncher_ShowStartView_Original;
HRESULT WINAPI XamlLauncher_ShowStartView_Hook(void* pThis,
                                               int immersiveLauncherShowMethod,
                                               int immersiveLauncherShowFlags) {
    Wh_Log(L">");

    g_inShowStartView = true;

    HRESULT ret = XamlLauncher_ShowStartView_Original(
        pThis, immersiveLauncherShowMethod, immersiveLauncherShowFlags);

    g_inShowStartView = false;

    return ret;
}

using ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_t =
    HRESULT(WINAPI*)(void* pThis);
ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_t
    ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_Original;
HRESULT WINAPI
ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_Hook(void* pThis) {
    Wh_Log(L">");

    auto original = [=]() {
        return ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_Original(
            pThis);
    };

    if (!g_inShowStartView) {
        return original();
    }

    HMONITOR destMonitor = nullptr;

    if (*g_settings.monitorInterfaceName.get()) {
        destMonitor = GetMonitorByInterfaceNameSubstr(
            g_settings.monitorInterfaceName.get());
    } else if (g_settings.monitor == 0) {
        POINT pt;
        GetCursorPos(&pt);
        destMonitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    } else if (g_settings.monitor >= 1) {
        destMonitor = GetMonitorById(g_settings.monitor - 1);
    }

    if (!destMonitor) {
        return original();
    }

    MONITORINFO monitorInfo{
        .cbSize = sizeof(MONITORINFO),
    };
    GetMonitorInfo(destMonitor, &monitorInfo);

    RECT rc = monitorInfo.rcMonitor;

    POINT pt = {
        rc.left + (rc.right - rc.left) / 2,
        rc.top + (rc.bottom - rc.top) / 2,
    };

    ImmersiveMonitorHelper_ConnectToMonitor_Original(pThis, nullptr, pt);

    return S_OK;
}

void LoadSettings() {
    g_settings.monitor = Wh_GetIntSetting(L"monitor");
    g_settings.monitorInterfaceName =
        WindhawkUtils::StringSetting::make(L"monitorInterfaceName");
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    // twinui.pcshell.dll
    WindhawkUtils::SYMBOL_HOOK twinuiPcshellSymbolHooks[] = {
        {
            {LR"(public: bool __cdecl ImmersiveMonitorHelper::ConnectToMonitor(struct HWND__ *,struct tagPOINT))"},
            &ImmersiveMonitorHelper_ConnectToMonitor_Original,
        },
        {
            {LR"(public: virtual long __cdecl XamlLauncher::ShowStartView(enum IMMERSIVELAUNCHERSHOWMETHOD,enum IMMERSIVELAUNCHERSHOWFLAGS))"},
            &XamlLauncher_ShowStartView_Original,
            XamlLauncher_ShowStartView_Hook,
        },
        {
            {LR"(public: long __cdecl ImmersiveMonitorHelper::AdjustMonitorConnectedIfNeeded(void))"},
            &ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_Original,
            ImmersiveMonitorHelper_AdjustMonitorConnectedIfNeeded_Hook,
        },
    };

    HMODULE twinuiPcshellModule = LoadLibraryEx(L"twinui.pcshell.dll", nullptr,
                                                LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!twinuiPcshellModule) {
        Wh_Log(L"Couldn't load twinui.pcshell.dll");
        return FALSE;
    }

    if (!HookSymbols(twinuiPcshellModule, twinuiPcshellSymbolHooks,
                     ARRAYSIZE(twinuiPcshellSymbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
    }

    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L">");
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}
