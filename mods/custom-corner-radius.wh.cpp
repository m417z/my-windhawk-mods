// ==WindhawkMod==
// @id              custom-corner-radius
// @name            Custom Window Corner Radius
// @description     Customizes window corner radius in Windows 11, making corners more or less rounded
// @version         1.0
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

- Popups and flyouts (taskbar, wifi, volume, context menus) use a smaller radius
  (4px) and are unaffected by this mod. Some of them can be customized using
  other mods, such as Windows 11 Taskbar Styler.
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
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

struct {
    int radius;
} g_settings;

using GetRadiusFromCornerStyle_t = float(WINAPI*)(void* pThis);
GetRadiusFromCornerStyle_t GetRadiusFromCornerStyle_Original;
float WINAPI GetRadiusFromCornerStyle_Hook(void* pThis) {
    float orig = GetRadiusFromCornerStyle_Original(pThis);
    if (orig > 0) {
        Wh_Log(L"> %f", orig);
        return (float)g_settings.radius;
    }
    return orig;
}

using GetFloatCornerRadiusForCurrentStyle_t = float(WINAPI*)(void* pThis);
GetFloatCornerRadiusForCurrentStyle_t
    GetFloatCornerRadiusForCurrentStyle_Original;
float WINAPI GetFloatCornerRadiusForCurrentStyle_Hook(void* pThis) {
    float orig = GetFloatCornerRadiusForCurrentStyle_Original(pThis);
    if (orig > 0) {
        Wh_Log(L"> %f", orig);
        return (float)g_settings.radius;
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
    if (cornerRadius > 0) {
        Wh_Log(L"> %f", cornerRadius);
        cornerRadius = (float)g_settings.radius;
    }
    return SetBorderParameters_Original(pThis, borderRect, cornerRadius, dpi,
                                        color, borderStyle, shadowStyle);
}

void LoadSettings() {
    g_settings.radius = Wh_GetIntSetting(L"radius");
    if (g_settings.radius < 0) {
        g_settings.radius = 0;
    } else if (g_settings.radius > 60) {
        g_settings.radius = 60;
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

    return HookSymbols(udwm, udwmDllHooks, ARRAYSIZE(udwmDllHooks));
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}

void Wh_ModUninit() {
    Wh_Log(L">");
}
