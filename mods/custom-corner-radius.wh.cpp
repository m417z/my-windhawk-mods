// ==WindhawkMod==
// @id              custom-corner-radius
// @name            Custom Window Corner Radius
// @description     Customizes window corner radius in Windows 11, making corners more or less rounded
// @version         1.2
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         dwm.exe
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
# Custom Window Corner Radius

Customizes Windows 11 app window corner radius. Make corners more rounded than
the default 8px, or reduce the radius for less rounded or completely sharp
corners.

The mod was [originally
submitted](https://github.com/ramensoftware/windhawk-mods/pull/3587) by
[Kanak415](https://github.com/kanak-buet19).

![Screenshot](https://i.imgur.com/mMGkBwc.png)

## ⚠ Important usage note ⚠

This mod needs to hook into `dwm.exe` to work. Please navigate to Windhawk's
Settings > Advanced settings > More advanced settings > Process inclusion list,
and make sure that `dwm.exe` is in the list.

![Advanced settings screenshot](https://i.imgur.com/LRhREtJ.png)

## Additional notes

- Some elements, such as context menus and tooltips, use a smaller radius (4px
  by default). These can be customized separately with the "Small corner
  radius" and "Tooltip corner radius" options.
- Some elements, such as the taskbar, the Start menu, and the notification
  center, are unaffected by this mod. Some of them can be customized using other
  mods, such as Windows 11 Taskbar Styler.
- Disabling the mod instantly restores default behavior — no system files are
  modified.
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
- smallRadius: 6
  $name: Small corner radius
  $description: >-
    Corner radius for elements that use a smaller radius, such as context menus.
    Default Win11 is 4.
- tooltipRadius: 4
  $name: Tooltip corner radius
  $description: >-
    Corner radius for elements that use a smaller radius, such as tooltips.
    Default Win11 is 4.
*/
// ==/WindhawkModSettings==

#include <windows.h>
#include <windhawk_utils.h>

#include <algorithm>
#include <vector>

struct {
    float radius;
    float smallRadius;
    float tooltipRadius;
} g_settings;

thread_local int g_getFloatCornerRadiusHookDepth = 0;
thread_local bool g_getFloatCornerRadiusCalledGetRadius = false;
thread_local int g_skipNextSetBorderParametersRemap = 0;

enum class SmallPopupKind {
    Unknown,
    ContextMenu,
    Tooltip,
};

struct TrackedSmallPopup {
    HWND hwnd;
    SmallPopupKind kind;
};

SRWLOCK g_smallPopupLock = SRWLOCK_INIT;
std::vector<TrackedSmallPopup> g_smallPopups;
HANDLE g_smallPopupTrackerThread = nullptr;
HANDLE g_smallPopupTrackerStopEvent = nullptr;
HWINEVENTHOOK g_smallPopupTrackerHook = nullptr;

static SmallPopupKind SmallPopupKindFromHwnd(HWND hwnd) {
    wchar_t className[256] = {};
    if (!GetClassNameW(hwnd, className, ARRAYSIZE(className))) {
        return SmallPopupKind::Unknown;
    }

    if (wcscmp(className, L"#32768") == 0) {
        return SmallPopupKind::ContextMenu;
    }

    if (wcscmp(className, L"tooltips_class32") == 0) {
        return SmallPopupKind::Tooltip;
    }

    return SmallPopupKind::Unknown;
}

static void TrackSmallPopup(HWND hwnd, SmallPopupKind kind) {
    if (!hwnd || kind == SmallPopupKind::Unknown) {
        return;
    }

    AcquireSRWLockExclusive(&g_smallPopupLock);

    for (TrackedSmallPopup& popup : g_smallPopups) {
        if (popup.hwnd == hwnd) {
            popup.kind = kind;
            ReleaseSRWLockExclusive(&g_smallPopupLock);
            return;
        }
    }

    if (g_smallPopups.size() >= 64) {
        g_smallPopups.erase(g_smallPopups.begin());
    }

    g_smallPopups.push_back({hwnd, kind});

    ReleaseSRWLockExclusive(&g_smallPopupLock);
}

static void UntrackSmallPopup(HWND hwnd) {
    if (!hwnd) {
        return;
    }

    AcquireSRWLockExclusive(&g_smallPopupLock);

    g_smallPopups.erase(
        std::remove_if(g_smallPopups.begin(), g_smallPopups.end(),
                       [hwnd](const TrackedSmallPopup& popup) {
                           return popup.hwnd == hwnd;
                       }),
        g_smallPopups.end());

    ReleaseSRWLockExclusive(&g_smallPopupLock);
}

static bool SameWindowSizeAsBorder(HWND hwnd, const RECT& borderRect) {
    RECT rc = {};
    if (!GetWindowRect(hwnd, &rc)) {
        return false;
    }

    LONG width = rc.right - rc.left;
    LONG height = rc.bottom - rc.top;
    LONG borderWidth = borderRect.right - borderRect.left;
    LONG borderHeight = borderRect.bottom - borderRect.top;

    if (width <= 0 || height <= 0 || borderWidth <= 0 || borderHeight <= 0) {
        return false;
    }

    LONG dw = width - borderWidth;
    if (dw < 0) dw = -dw;
    LONG dh = height - borderHeight;
    if (dh < 0) dh = -dh;

    return dw <= 1 && dh <= 1;
}

static bool HasMatchingTrackedPopup(SmallPopupKind kind,
                                    const RECT& borderRect) {
    std::vector<TrackedSmallPopup> popups;

    AcquireSRWLockShared(&g_smallPopupLock);
    popups = g_smallPopups;
    ReleaseSRWLockShared(&g_smallPopupLock);

    for (const TrackedSmallPopup& popup : popups) {
        if (popup.kind != kind) {
            continue;
        }

        if (!IsWindow(popup.hwnd) || !IsWindowVisible(popup.hwnd)) {
            continue;
        }

        if (SmallPopupKindFromHwnd(popup.hwnd) != kind) {
            continue;
        }

        if (SameWindowSizeAsBorder(popup.hwnd, borderRect)) {
            return true;
        }
    }

    return false;
}

static bool HasMatchingPopupClass(const wchar_t* className,
                                  SmallPopupKind kind,
                                  const RECT& borderRect) {
    for (HWND hwnd = nullptr;
         (hwnd = FindWindowExW(nullptr, hwnd, className, nullptr)) != nullptr;) {
        if (!IsWindowVisible(hwnd)) {
            continue;
        }

        if (SameWindowSizeAsBorder(hwnd, borderRect)) {
            TrackSmallPopup(hwnd, kind);
            return true;
        }
    }

    return false;
}

static SmallPopupKind ClassifySmallPopup(const RECT& borderRect) {
    bool hasContextMenu = HasMatchingTrackedPopup(SmallPopupKind::ContextMenu,
                                                 borderRect);
    bool hasTooltip = HasMatchingTrackedPopup(SmallPopupKind::Tooltip,
                                             borderRect);

    if (!hasContextMenu) {
        hasContextMenu = HasMatchingPopupClass(L"#32768",
                                               SmallPopupKind::ContextMenu,
                                               borderRect);
    }

    if (!hasTooltip) {
        hasTooltip = HasMatchingPopupClass(L"tooltips_class32",
                                           SmallPopupKind::Tooltip,
                                           borderRect);
    }

    if (hasTooltip == hasContextMenu) {
        return SmallPopupKind::Unknown;
    }

    return hasTooltip ? SmallPopupKind::Tooltip : SmallPopupKind::ContextMenu;
}

void CALLBACK SmallPopupWinEventProc(HWINEVENTHOOK hook,
                                     DWORD event,
                                     HWND hwnd,
                                     LONG idObject,
                                     LONG idChild,
                                     DWORD eventThread,
                                     DWORD eventTime) {
    if (!hwnd || idObject != OBJID_WINDOW || idChild != CHILDID_SELF) {
        return;
    }

    if (event == EVENT_OBJECT_DESTROY || event == EVENT_OBJECT_HIDE) {
        UntrackSmallPopup(hwnd);
        return;
    }

    if (event != EVENT_OBJECT_CREATE && event != EVENT_OBJECT_SHOW) {
        return;
    }

    SmallPopupKind kind = SmallPopupKindFromHwnd(hwnd);
    if (kind != SmallPopupKind::Unknown) {
        TrackSmallPopup(hwnd, kind);
    }
}

DWORD WINAPI SmallPopupTrackerThreadProc(void*) {
    g_smallPopupTrackerHook = SetWinEventHook(
        EVENT_OBJECT_CREATE,
        EVENT_OBJECT_HIDE,
        nullptr,
        SmallPopupWinEventProc,
        0,
        0,
        WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

    if (!g_smallPopupTrackerHook) {
        return 0;
    }

    while (true) {
        DWORD result = MsgWaitForMultipleObjects(
            1, &g_smallPopupTrackerStopEvent, FALSE, INFINITE, QS_ALLINPUT);

        if (result == WAIT_OBJECT_0) {
            break;
        }

        if (result == WAIT_OBJECT_0 + 1) {
            MSG msg;
            while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    }

    UnhookWinEvent(g_smallPopupTrackerHook);
    g_smallPopupTrackerHook = nullptr;

    return 0;
}

static bool StartSmallPopupTracker() {
    if (g_smallPopupTrackerThread) {
        return true;
    }

    g_smallPopupTrackerStopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!g_smallPopupTrackerStopEvent) {
        return false;
    }

    g_smallPopupTrackerThread = CreateThread(
        nullptr, 0, SmallPopupTrackerThreadProc, nullptr, 0, nullptr);
    if (!g_smallPopupTrackerThread) {
        CloseHandle(g_smallPopupTrackerStopEvent);
        g_smallPopupTrackerStopEvent = nullptr;
        return false;
    }

    return true;
}

static void StopSmallPopupTracker() {
    if (g_smallPopupTrackerStopEvent) {
        SetEvent(g_smallPopupTrackerStopEvent);
    }

    if (g_smallPopupTrackerThread) {
        WaitForSingleObject(g_smallPopupTrackerThread, 1000);
        CloseHandle(g_smallPopupTrackerThread);
        g_smallPopupTrackerThread = nullptr;
    }

    if (g_smallPopupTrackerStopEvent) {
        CloseHandle(g_smallPopupTrackerStopEvent);
        g_smallPopupTrackerStopEvent = nullptr;
    }

    AcquireSRWLockExclusive(&g_smallPopupLock);
    g_smallPopups.clear();
    ReleaseSRWLockExclusive(&g_smallPopupLock);
}

float RadiusForOriginal(float orig) {
    // Win11 defaults: 4.0 for smaller radius, 8.0 for larger radius. Use middle
    // point as a threshold.
    if (orig <= 6.0f) {
        return g_settings.smallRadius;
    }
    return g_settings.radius;
}

using GetRadiusFromCornerStyle_t = float(WINAPI*)(void* pThis);
GetRadiusFromCornerStyle_t GetRadiusFromCornerStyle_Original;
float WINAPI GetRadiusFromCornerStyle_Hook(void* pThis) {
    float orig = GetRadiusFromCornerStyle_Original(pThis);
    if (orig > 0) {
        Wh_Log(L"> %f", orig);

        if (g_getFloatCornerRadiusHookDepth > 0) {
            g_getFloatCornerRadiusCalledGetRadius = true;
        }

        return RadiusForOriginal(orig);
    }
    return orig;
}

using GetFloatCornerRadiusForCurrentStyle_t = float(WINAPI*)(void* pThis);
GetFloatCornerRadiusForCurrentStyle_t
    GetFloatCornerRadiusForCurrentStyle_Original;
float WINAPI GetFloatCornerRadiusForCurrentStyle_Hook(void* pThis) {
    ++g_getFloatCornerRadiusHookDepth;
    bool prevCalledGetRadius = g_getFloatCornerRadiusCalledGetRadius;
    g_getFloatCornerRadiusCalledGetRadius = false;

    float orig = GetFloatCornerRadiusForCurrentStyle_Original(pThis);
    bool calledGetRadius = g_getFloatCornerRadiusCalledGetRadius;

    g_getFloatCornerRadiusCalledGetRadius = prevCalledGetRadius;
    --g_getFloatCornerRadiusHookDepth;

    if (orig > 0) {
        Wh_Log(L"> %f", orig);

        float result = orig;
        if (!calledGetRadius) {
            result = RadiusForOriginal(orig);
        }

        ++g_skipNextSetBorderParametersRemap;
        return result;
    }
    return orig;
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
    bool alreadyRemappedByGetFloat = false;
    if (g_skipNextSetBorderParametersRemap > 0) {
        alreadyRemappedByGetFloat = true;
        --g_skipNextSetBorderParametersRemap;
    }

    if (cornerRadius > 0) {
        Wh_Log(L"> %f", cornerRadius);

        SmallPopupKind popupKind = SmallPopupKind::Unknown;
        if (borderStyle == 0 && shadowStyle == 1) {
            popupKind = ClassifySmallPopup(borderRect);
        }

        if (popupKind == SmallPopupKind::Tooltip) {
            cornerRadius = g_settings.tooltipRadius;
        } else if (!alreadyRemappedByGetFloat) {
            cornerRadius = RadiusForOriginal(cornerRadius);
        }
    }

    return SetBorderParameters_Original(pThis, borderRect, cornerRadius, dpi,
                                        color, borderStyle, shadowStyle);
}

void LoadSettings() {
    g_settings.radius = static_cast<float>(Wh_GetIntSetting(L"radius"));
    if (g_settings.radius < 0) {
        g_settings.radius = 0;
    }

    g_settings.smallRadius =
        static_cast<float>(Wh_GetIntSetting(L"smallRadius"));
    if (g_settings.smallRadius < 0) {
        g_settings.smallRadius = 0;
    }

    g_settings.tooltipRadius =
        static_cast<float>(Wh_GetIntSetting(L"tooltipRadius"));
    if (g_settings.tooltipRadius < 0) {
        g_settings.tooltipRadius = 0;
    }
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    HMODULE udwm = GetModuleHandle(L"udwm.dll");
    if (!udwm) {
        Wh_Log(L"udwm.dll isn't loaded");
        return FALSE;
    }

    if (!StartSmallPopupTracker()) {
        Wh_Log(L"Failed to start small popup tracker");
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

    WindhawkUtils::SYMBOL_HOOK udwmDllHooks[] = {
        // Covers the 3D animation path in both old and new builds.
        {
            {LR"(private: float __cdecl CTopLevelWindow::GetRadiusFromCornerStyle(void))"},
            &GetRadiusFromCornerStyle_Original,
            GetRadiusFromCornerStyle_Hook,
        },
        // Covers UpdateWindowVisuals in new builds (calls
        // GetRadiusFromCornerStyle, but hooked separately in case a future
        // build inlines that call).
        {
            {LR"(private: float __cdecl CTopLevelWindow::GetFloatCornerRadiusForCurrentStyle(void))"},
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
        StopSmallPopupTracker();
        return FALSE;
    }

    return TRUE;
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}

void Wh_ModUninit() {
    Wh_Log(L">");

    StopSmallPopupTracker();
}
