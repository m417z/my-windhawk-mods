// ==WindhawkMod==
// @id              windows-start-icon
// @name            Windows Start icon
// @description     Recolor, rotate, animate, or replace the Windows 11 Start button icon
// @version         1.0
// @author          Smurf Customs
// @github          https://github.com/smurfcustoms
// @homepage        https://smurfcustoms.co.uk
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lole32 -loleaut32 -lruntimeobject
// @license         GPL-3.0-only
// ==/WindhawkMod==

// Based on Start button colorizer by m417z (GPL-3.0):
// https://github.com/ramensoftware/windhawk-mods
//
// Source code is published under the GNU General Public License v3.0 only.

// ==WindhawkModReadme==
/*
# Windows Start icon

Customize the Windows 11 Start button icon on the taskbar.

* **Color** and **effects**: a preset, the accent color, or hue, saturation,
  brightness, and opacity. A white or black logo is tinted so the color shows.
* **Size**: scales the icon without moving the button.
* **Rotation**: a fixed angle, a continuous spin, or both. Positive spin turns
  clockwise.
* **Live colors**: the logo cycles through hues. Gradient mode spreads that
  color across the icon and moves it.
* **Custom image**: PNG, JPG, GIF, BMP, or WEBP, from a file path or an
  http(s) URL. `%USERPROFILE%\Pictures\start.png` works. GIFs play on their own.

The Windows logo keeps the color, effects, and live colors. A custom image
replaces that logo. Size, opacity, angle, and spin still apply to the image.

Disable **Start button colorizer** and **Start Button Replacer** while this
mod is on. They edit the same button and will fight this one.

After replacing a file at the same path, save this mod's settings again so
the picture reloads.

Only Windows 11 is supported.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- color: none
  $name: Color
  $description: >-
    Recolors the Windows logo. The system accent color also uses that color's
    saturation. Original keeps the logo's own hue.
  $options:
  - none: Original
  - accent: System accent color
  - red: Red
  - orange: Orange
  - yellow: Yellow
  - green: Green
  - teal: Teal
  - blue: Blue
  - purple: Purple
  - pink: Pink
- effects:
  - hue: 0
    $name: Hue shift
    $description: Extra hue rotation in degrees, applied on top of Color.
  - saturation: 100
    $name: Saturation
    $description: Percent of the working color. Zero makes the icon gray.
  - brightness: 100
    $name: Brightness
    $description: >-
      Percent of the working lightness. Zero is black. Values above 100 lighten
      toward white.
  - opacity: 100
    $name: Opacity
    $description: Percent. Also applies to a custom image.
  $name: Effects
- size: 100
  $name: Size
  $description: Icon scale in percent. 100 is the normal size.
- rotation:
  - angle: 0
    $name: Angle
    $description: Fixed rotation in degrees. Positive turns clockwise.
  - spin: 0
    $name: Spin speed
    $description: >-
      Extra rotation in degrees per second. 0 holds the angle still. 90 is a
      slow turn, 360 is one full turn per second. Negative turns the other way.
  $name: Rotation
- live:
  - enabled: false
    $name: Enabled
    $description: >-
      Animate the Windows logo's colors. Hidden while a custom image is showing.
  - style: gradient
    $name: Style
    $description: >-
      Whole icon shifts one color. Gradient spreads the hue across the logo
      and moves that spread.
    $options:
    - cycle: Whole icon
    - gradient: Gradient across the icon
  - speed: 45
    $name: Speed
    $description: >-
      Hue change in degrees per second. 0 holds the live color still. Negative
      runs the cycle backward.
  - spread: 90
    $name: Spread
    $description: >-
      Hue difference from one side of the logo to the other in Gradient style.
      0 uses one color.
  $name: Live colors
- image:
  - path: ""
    $name: Image file or URL
    $description: >-
      Leave empty to keep the Windows logo. Otherwise use a full path
      (%USERPROFILE%\Pictures\start.png) or an http(s) URL.
  - size: 0
    $name: Image size
    $description: >-
      Display size in pixels before the Size percent is applied. 0 matches the
      Windows icon.
  $name: Custom image
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <cwctype>
#include <functional>
#include <list>
#include <mutex>
#include <optional>
#include <string>
#include <vector>

#undef GetCurrentTime

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Foundation.Numerics.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.System.h>
#include <winrt/Windows.UI.Composition.h>
#include <winrt/Windows.UI.ViewManagement.h>
#include <winrt/Windows.UI.Xaml.Automation.Peers.h>
#include <winrt/Windows.UI.Xaml.Automation.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Hosting.h>
#include <winrt/Windows.UI.Xaml.Media.Imaging.h>
#include <winrt/Windows.UI.Xaml.Media.h>
#include <winrt/Windows.UI.Xaml.h>
#include <winrt/base.h>

using namespace winrt::Windows::UI::Xaml;

namespace Composition = winrt::Windows::UI::Composition;
namespace Controls = winrt::Windows::UI::Xaml::Controls;
namespace Foundation = winrt::Windows::Foundation;
namespace Imaging = winrt::Windows::UI::Xaml::Media::Imaging;
namespace ViewManagement = winrt::Windows::UI::ViewManagement;

struct Settings {
    int colorHue = -1;
    bool accentColor = false;
    int hue = 0;
    double saturation = 1;
    double brightness = 1;
    double opacity = 1;
    double size = 1;
    double angle = 0;
    double spin = 0;
    bool liveEnabled = false;
    bool liveGradient = true;
    double liveSpeed = 45;
    double liveSpread = 90;
    std::wstring imagePath;
    double imageSize = 0;
    bool customizeColors = false;
};

Settings g_settings;
std::mutex g_settingsMutex;

std::atomic<bool> g_taskbarViewDllLoaded;
std::atomic<bool> g_unloading;
std::atomic<int> g_imageGeneration{0};

thread_local const Settings* t_activeSettings = nullptr;
thread_local Settings t_fallbackSettings;

struct SettingsScope {
    const Settings* previous;

    explicit SettingsScope(const Settings& settings)
        : previous(t_activeSettings) {
        t_activeSettings = &settings;
    }

    ~SettingsScope() { t_activeSettings = previous; }

    SettingsScope(const SettingsScope&) = delete;
    SettingsScope& operator=(const SettingsScope&) = delete;
};

Settings GetSettings() {
    std::lock_guard lock(g_settingsMutex);
    return g_settings;
}

const Settings& ActiveSettings() {
    if (t_activeSettings) {
        return *t_activeSettings;
    }

    t_fallbackSettings = GetSettings();
    return t_fallbackSettings;
}

[[clang::no_destroy]] std::optional<std::list<FrameworkElement::Loaded_revoker>>
    g_iconLoadedRevokers{std::in_place};

struct IconTarget {
    winrt::weak_ref<FrameworkElement> icon;
    winrt::weak_ref<FrameworkElement> image;
};

struct LiveState {
    winrt::Windows::System::DispatcherQueueTimer timer{nullptr};
    Imaging::BitmapImage bitmap{nullptr};
    int bitmapGeneration = -1;
    bool imageReady = false;
    bool imageFailed = false;
    double spinPhase = 0;
    double huePhase = 0;
    bool hasTick = false;
    int colorTick = 0;
    std::chrono::steady_clock::time_point lastTick{};
    std::vector<IconTarget> icons;
};

[[clang::no_destroy]] LiveState g_live{};

bool g_inApply = false;
bool g_inLiveTick = false;
bool g_applyingStyle = false;
int g_sourcesAppliedGeneration = -1;

constexpr PCWSTR kCustomImageHostName = L"WindhawkStartIconHost";

struct BoolFlag {
    bool& flag;

    explicit BoolFlag(bool& flag) : flag(flag) { flag = true; }

    ~BoolFlag() { flag = false; }

    BoolFlag(const BoolFlag&) = delete;
    BoolFlag& operator=(const BoolFlag&) = delete;
};

void ApplySettingsFromTaskbarThread();

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

FrameworkElement EnumChildElements(
    FrameworkElement element,
    std::function<bool(FrameworkElement)> enumCallback) {
    int childrenCount = Media::VisualTreeHelper::GetChildrenCount(element);

    for (int i = 0; i < childrenCount; i++) {
        auto child = Media::VisualTreeHelper::GetChild(element, i)
                         .try_as<FrameworkElement>();
        if (!child) {
            Wh_Log(L"Failed to get child %d of %d", i + 1, childrenCount);
            continue;
        }

        if (enumCallback(child)) {
            return child;
        }
    }

    return nullptr;
}

FrameworkElement FindChildByName(FrameworkElement element, PCWSTR name) {
    return EnumChildElements(element, [name](FrameworkElement child) {
        return child.Name() == name;
    });
}

FrameworkElement FindChildByClassName(FrameworkElement element,
                                      PCWSTR className) {
    return EnumChildElements(element, [className](FrameworkElement child) {
        return winrt::get_class_name(child) == className;
    });
}

FrameworkElement FindDescendant(
    FrameworkElement element,
    const std::function<bool(FrameworkElement)>& predicate) {
    FrameworkElement result = nullptr;

    EnumChildElements(element, [&predicate, &result](FrameworkElement child) {
        result = predicate(child) ? child : FindDescendant(child, predicate);
        return !!result;
    });

    return result;
}

FrameworkElement FindDescendantByName(FrameworkElement element, PCWSTR name) {
    return FindDescendant(element, [name](FrameworkElement child) {
        return child.Name() == name;
    });
}

bool IsStartButton(FrameworkElement element) {
    return winrt::get_class_name(element) == L"Taskbar.ExperienceToggleButton" &&
           Automation::AutomationProperties::GetAutomationId(element) ==
               L"StartButton";
}

FrameworkElement FindAncestorStartButton(FrameworkElement element) {
    DependencyObject current = element;

    for (int depth = 0; current && depth < 16; depth++) {
        auto frameworkElement = current.try_as<FrameworkElement>();
        if (frameworkElement && IsStartButton(frameworkElement)) {
            return frameworkElement;
        }

        try {
            current = Media::VisualTreeHelper::GetParent(current);
        } catch (...) {
            return nullptr;
        }
    }

    return nullptr;
}

struct HslColor {
    double hue;
    double saturation;
    double lightness;
};

HslColor RgbToHsl(winrt::Windows::UI::Color color) {
    double r = color.R / 255.0;
    double g = color.G / 255.0;
    double b = color.B / 255.0;

    double max = std::max({r, g, b});
    double min = std::min({r, g, b});
    double delta = max - min;

    HslColor hsl{
        .hue = 0,
        .saturation = 0,
        .lightness = (max + min) / 2,
    };

    if (delta > 0) {
        hsl.saturation = delta / (1 - std::abs(2 * hsl.lightness - 1));

        if (max == r) {
            hsl.hue = (g - b) / delta;
        } else if (max == g) {
            hsl.hue = (b - r) / delta + 2;
        } else {
            hsl.hue = (r - g) / delta + 4;
        }

        hsl.hue *= 60;
        if (hsl.hue < 0) {
            hsl.hue += 360;
        }
    }

    return hsl;
}

winrt::Windows::UI::Color HslToRgb(HslColor hsl, uint8_t alpha) {
    double c = (1 - std::abs(2 * hsl.lightness - 1)) * hsl.saturation;
    double h = hsl.hue / 60;
    double x = c * (1 - std::abs(std::fmod(h, 2) - 1));
    double m = hsl.lightness - c / 2;

    double r = 0;
    double g = 0;
    double b = 0;
    if (h < 1) {
        r = c;
        g = x;
    } else if (h < 2) {
        r = x;
        g = c;
    } else if (h < 3) {
        g = c;
        b = x;
    } else if (h < 4) {
        g = x;
        b = c;
    } else if (h < 5) {
        r = x;
        b = c;
    } else {
        r = c;
        b = x;
    }

    auto toByte = [m](double value) {
        return static_cast<uint8_t>(
            std::lround(std::clamp(value + m, 0.0, 1.0) * 255));
    };

    return {alpha, toByte(r), toByte(g), toByte(b)};
}

double WrapDegrees(double degrees) {
    degrees = std::fmod(degrees, 360.0);
    if (degrees < 0) {
        degrees += 360.0;
    }
    return degrees;
}

[[clang::no_destroy]] ViewManagement::UISettings g_uiSettings{nullptr};
winrt::event_token g_colorValuesChangedToken;
HslColor g_accentColor;

HslColor QueryAccentColor() {
    return RgbToHsl(
        g_uiSettings.GetColorValue(ViewManagement::UIColorType::Accent));
}

HslColor AccentColor() {
    if (g_uiSettings) {
        return g_accentColor;
    }

    g_uiSettings = ViewManagement::UISettings();
    g_accentColor = QueryAccentColor();

    g_colorValuesChangedToken = g_uiSettings.ColorValuesChanged(
        [dispatcherQueue =
             winrt::Windows::System::DispatcherQueue::GetForCurrentThread()](
            auto&&, auto&&) {
            dispatcherQueue.TryEnqueue([] {
                Wh_Log(L">");

                if (g_uiSettings) {
                    g_accentColor = QueryAccentColor();
                    ApplySettingsFromTaskbarThread();
                }
            });
        });

    return g_accentColor;
}

void StopTrackingAccentColor() {
    if (g_uiSettings) {
        g_uiSettings.ColorValuesChanged(g_colorValuesChangedToken);
        g_uiSettings = nullptr;
    }
}

winrt::Windows::UI::Color TransformColor(const Settings& settings,
                                        winrt::Windows::UI::Color color,
                                        double spatialHue) {
    HslColor hsl = RgbToHsl(color);

    bool wantsHue = settings.accentColor || settings.colorHue >= 0 ||
                    (settings.liveEnabled && !g_unloading);
    if (wantsHue && hsl.saturation < 0.08) {
        // The idle Windows logo is often plain white or black, which has no
        // hue to rotate. Give those brushes a color the sliders can change.
        hsl.saturation = 1;
        if (hsl.lightness >= 0.94) {
            hsl.lightness = 0.60;
        } else if (hsl.lightness <= 0.06) {
            hsl.lightness = 0.45;
        }
    }

    if (settings.accentColor) {
        HslColor accentColor = AccentColor();
        hsl.hue = accentColor.hue;
        hsl.saturation *= accentColor.saturation;
    } else if (settings.colorHue >= 0) {
        hsl.hue = settings.colorHue;
    }

    double hue = hsl.hue + settings.hue + spatialHue;
    if (settings.liveEnabled && !g_unloading) {
        hue += g_live.huePhase;
    }
    hsl.hue = WrapDegrees(hue);

    hsl.saturation =
        std::clamp(hsl.saturation * settings.saturation, 0.0, 1.0);
    hsl.lightness =
        std::clamp(hsl.lightness * settings.brightness, 0.0, 1.0);

    double alpha = std::clamp(color.A / 255.0 * settings.opacity, 0.0, 1.0);
    return HslToRgb(hsl, static_cast<uint8_t>(std::lround(alpha * 255)));
}

thread_local bool g_creatingIconAnimations;

using IColorKeyFrameAnimation_InsertKeyFrame_t =
    HRESULT(WINAPI*)(void* pThis,
                     float normalizedProgressKey,
                     winrt::Windows::UI::Color value);
IColorKeyFrameAnimation_InsertKeyFrame_t
    IColorKeyFrameAnimation_InsertKeyFrame_Original;
HRESULT WINAPI
IColorKeyFrameAnimation_InsertKeyFrame_Hook(void* pThis,
                                            float normalizedProgressKey,
                                            winrt::Windows::UI::Color value) {
    if (g_creatingIconAnimations) {
        value = TransformColor(ActiveSettings(), value, 0);
    }

    return IColorKeyFrameAnimation_InsertKeyFrame_Original(
        pThis, normalizedProgressKey, value);
}

using IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_t =
    HRESULT(WINAPI*)(void* pThis,
                     float normalizedProgressKey,
                     winrt::Windows::UI::Color value,
                     void* easingFunction);
IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_t
    IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_Original;
HRESULT WINAPI IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_Hook(
    void* pThis,
    float normalizedProgressKey,
    winrt::Windows::UI::Color value,
    void* easingFunction) {
    if (g_creatingIconAnimations) {
        value = TransformColor(ActiveSettings(), value, 0);
    }

    return IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_Original(
        pThis, normalizedProgressKey, value, easingFunction);
}

constexpr int kInsertKeyFrameVtableIndex = 8;
constexpr int kInsertKeyFrameWithEasingFunctionVtableIndex = 9;

void HookColorKeyFrameAnimation(Composition::Compositor compositor) {
    void** vtable =
        *(void***)winrt::get_abi(compositor.CreateColorKeyFrameAnimation());

    WindhawkUtils::SetFunctionHook(
        (IColorKeyFrameAnimation_InsertKeyFrame_t)
            vtable[kInsertKeyFrameVtableIndex],
        IColorKeyFrameAnimation_InsertKeyFrame_Hook,
        &IColorKeyFrameAnimation_InsertKeyFrame_Original);

    WindhawkUtils::SetFunctionHook(
        (IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_t)
            vtable[kInsertKeyFrameWithEasingFunctionVtableIndex],
        IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_Hook,
        &IColorKeyFrameAnimation_InsertKeyFrameWithEasingFunction_Original);

    Wh_ApplyHookOperations();
}

void EnsureColorKeyFrameAnimationHooked(UIElement element) {
    [[maybe_unused]] static bool hooked = [&element] {
        HookColorKeyFrameAnimation(
            Hosting::ElementCompositionPreview::GetElementVisual(element)
                .Compositor());
        return true;
    }();
}

constexpr PCWSTR kOriginalColorProperty = L"WindhawkOriginalColor";

winrt::Windows::UI::Color CustomizedColor(
    Composition::CompositionObject colorOwner,
    winrt::Windows::UI::Color currentColor,
    double spatialHue) {
    auto properties = colorOwner.Properties();

    winrt::Windows::UI::Color originalColor;
    if (properties.TryGetColor(kOriginalColorProperty, originalColor) !=
        Composition::CompositionGetValueStatus::Succeeded) {
        originalColor = currentColor;
        properties.InsertColor(kOriginalColorProperty, originalColor);
    }

    const Settings& settings = ActiveSettings();
    return g_unloading || !settings.customizeColors
               ? originalColor
               : TransformColor(settings, originalColor, spatialHue);
}

void ApplyBrushColors(Composition::CompositionBrush brush, double spatialHue) {
    if (auto colorBrush = brush.try_as<Composition::CompositionColorBrush>()) {
        colorBrush.Color(CustomizedColor(colorBrush, colorBrush.Color(), spatialHue));
        return;
    }

    if (auto gradientBrush =
            brush.try_as<Composition::CompositionGradientBrush>()) {
        for (const auto& stop : gradientBrush.ColorStops()) {
            stop.Color(CustomizedColor(stop, stop.Color(), spatialHue));
        }
    }
}

struct BrushSample {
    Composition::CompositionBrush brush{nullptr};
    float x = 0;
    float y = 0;
};

void CollectShapeBrushes(Composition::CompositionShape shape,
                         float x,
                         float y,
                         std::vector<BrushSample>& out) {
    auto offset = shape.Offset();
    x += offset.x;
    y += offset.y;

    if (auto containerShape =
            shape.try_as<Composition::CompositionContainerShape>()) {
        for (const auto& child : containerShape.Shapes()) {
            CollectShapeBrushes(child, x, y, out);
        }
        return;
    }

    auto spriteShape = shape.try_as<Composition::CompositionSpriteShape>();
    if (!spriteShape) {
        return;
    }

    for (const auto& brush :
         {spriteShape.FillBrush(), spriteShape.StrokeBrush()}) {
        if (brush) {
            out.push_back(BrushSample{brush, x, y});
        }
    }
}

void CollectVisualBrushes(Composition::Visual visual,
                          float x,
                          float y,
                          std::vector<BrushSample>& out) {
    auto offset = visual.Offset();
    x += offset.x;
    y += offset.y;

    if (auto shapeVisual = visual.try_as<Composition::ShapeVisual>()) {
        for (const auto& shape : shapeVisual.Shapes()) {
            CollectShapeBrushes(shape, x, y, out);
        }
    }

    if (auto containerVisual = visual.try_as<Composition::ContainerVisual>()) {
        for (const auto& child : containerVisual.Children()) {
            CollectVisualBrushes(child, x, y, out);
        }
    }
}

double SpatialHueForSample(const Settings& settings,
                           const BrushSample& sample,
                           float minX,
                           float maxX,
                           float minY,
                           float maxY,
                           bool usePositions) {
    if (!settings.liveEnabled || !settings.liveGradient || g_unloading ||
        !usePositions) {
        return 0;
    }

    double nx = (maxX - minX) > 0.01f ? (sample.x - minX) / (maxX - minX) : 0.5;
    double ny = (maxY - minY) > 0.01f ? (sample.y - minY) / (maxY - minY) : 0.5;
    return (nx * 0.65 + ny * 0.35 - 0.5) * settings.liveSpread;
}

void ApplyIconColors(FrameworkElement icon, bool quiet) {
    try {
        EnsureColorKeyFrameAnimationHooked(icon);

        auto visual =
            Hosting::ElementCompositionPreview::GetElementChildVisual(icon);
        if (!visual) {
            if (!quiet) {
                Wh_Log(L"The icon has no animated visual");
            }
            return;
        }

        std::vector<BrushSample> samples;
        samples.reserve(64);
        CollectVisualBrushes(visual, 0, 0, samples);

        float minX = 0;
        float maxX = 0;
        float minY = 0;
        float maxY = 0;
        if (!samples.empty()) {
            minX = maxX = samples[0].x;
            minY = maxY = samples[0].y;
            for (const auto& sample : samples) {
                minX = std::min(minX, sample.x);
                maxX = std::max(maxX, sample.x);
                minY = std::min(minY, sample.y);
                maxY = std::max(maxY, sample.y);
            }
        }

        const Settings& settings = ActiveSettings();
        bool usePositions =
            settings.liveEnabled && settings.liveGradient &&
            ((maxX - minX) > 0.01f || (maxY - minY) > 0.01f);
        if (settings.liveEnabled && settings.liveGradient && !usePositions &&
            !quiet) {
            static bool logged = false;
            if (!logged) {
                logged = true;
                Wh_Log(L"Live gradient has no brush positions; the whole icon "
                       L"cycles together");
            }
        }

        for (const auto& sample : samples) {
            double spatial = SpatialHueForSample(settings, sample, minX, maxX,
                                                 minY, maxY, usePositions);
            try {
                ApplyBrushColors(sample.brush, spatial);
            } catch (const winrt::hresult_error& e) {
                if (!quiet) {
                    Wh_Log(L"Brush update failed: %s", e.message().c_str());
                }
            }
        }

        if (!quiet) {
            Wh_Log(L"Applied the colors of %d brushes",
                   static_cast<int>(samples.size()));
        }
    } catch (const winrt::hresult_error& e) {
        if (!quiet) {
            Wh_Log(L"ApplyIconColors failed: %s", e.message().c_str());
        }
    } catch (...) {
        if (!quiet) {
            Wh_Log(L"ApplyIconColors failed");
        }
    }
}

bool WantsTransform(const Settings& settings) {
    return !g_unloading && (std::abs(settings.size - 1.0) > 0.001 ||
                            std::abs(settings.angle) > 0.01 ||
                            std::abs(settings.spin) > 0.01);
}

double RotationDegrees(const Settings& settings) {
    double angle = settings.angle;
    if (!g_unloading && std::abs(settings.spin) > 0.01) {
        angle += g_live.spinPhase;
    }
    return WrapDegrees(angle);
}

void ApplyIconTransform(FrameworkElement element) {
    const Settings& settings = ActiveSettings();
    if (!WantsTransform(settings)) {
        element.ClearValue(UIElement::RenderTransformProperty());
        element.ClearValue(UIElement::RenderTransformOriginProperty());
        return;
    }

    Media::CompositeTransform transform{nullptr};
    if (auto renderTransform = element.RenderTransform()) {
        transform = renderTransform.try_as<Media::CompositeTransform>();
    }

    if (!transform) {
        transform = Media::CompositeTransform();
        element.RenderTransform(transform);
    }

    element.RenderTransformOrigin({0.5f, 0.5f});
    transform.ScaleX(settings.size);
    transform.ScaleY(settings.size);
    transform.Rotation(RotationDegrees(settings));
}

bool WantsLiveTimer(const Settings& settings) {
    return !g_unloading &&
           (std::abs(settings.spin) > 0.01 ||
            (settings.liveEnabled && std::abs(settings.liveSpeed) > 0.01));
}

void AdvanceClock(const Settings& settings) {
    auto now = std::chrono::steady_clock::now();
    double dt = 0;
    if (g_live.hasTick) {
        dt = std::chrono::duration<double>(now - g_live.lastTick).count();
        if (dt < 0) {
            dt = 0;
        } else if (dt > 0.25) {
            dt = 0.25;
        }
    }

    g_live.lastTick = now;
    g_live.hasTick = true;

    if (std::abs(settings.spin) > 0.01) {
        g_live.spinPhase = WrapDegrees(g_live.spinPhase + settings.spin * dt);
    }
    if (settings.liveEnabled && std::abs(settings.liveSpeed) > 0.01) {
        g_live.huePhase =
            WrapDegrees(g_live.huePhase + settings.liveSpeed * dt);
    }
}

void OnLiveTick();

void StopLiveTimer() {
    if (g_live.timer) {
        g_live.timer.Stop();
        g_live.timer = nullptr;
    }
    g_live.hasTick = false;
}

void StartLiveTimer() {
    if (g_live.timer) {
        return;
    }

    auto queue = winrt::Windows::System::DispatcherQueue::GetForCurrentThread();
    if (!queue) {
        Wh_Log(L"No dispatcher queue; spin and live colors will not animate");
        return;
    }

    try {
        g_live.timer = queue.CreateTimer();
        g_live.timer.Interval(std::chrono::milliseconds(16));
        g_live.timer.IsRepeating(true);
        g_live.timer.Tick([](auto&&, auto&&) { OnLiveTick(); });
        g_live.timer.Start();
    } catch (const winrt::hresult_error& e) {
        g_live.timer = nullptr;
        Wh_Log(L"Live timer failed: %s", e.message().c_str());
    }
}

void OnLiveTick() {
    if (g_unloading || g_inApply || g_inLiveTick) {
        return;
    }

    BoolFlag tickGuard(g_inLiveTick);

    Settings settings = GetSettings();
    SettingsScope scope(settings);
    AdvanceClock(settings);

    bool colorTick = false;
    if (settings.liveEnabled && std::abs(settings.liveSpeed) > 0.01 &&
        !g_live.imageReady) {
        g_live.colorTick = (g_live.colorTick + 1) % 2;
        colorTick = g_live.colorTick == 0;
    }

    bool spin = std::abs(settings.spin) > 0.01;

    for (auto it = g_live.icons.begin(); it != g_live.icons.end();) {
        auto icon = it->icon.get();
        auto image = it->image.get();
        if (!icon && !image) {
            it = g_live.icons.erase(it);
            continue;
        }

        try {
            if (icon && !(image && g_live.imageReady)) {
                if (colorTick) {
                    ApplyIconColors(icon, true);
                }
                if (spin) {
                    ApplyIconTransform(icon);
                }
            }
            if (image && spin) {
                ApplyIconTransform(image);
            }
        } catch (...) {
            // A taskbar rebuild can invalidate an element between ticks.
        }

        ++it;
    }
}

void TrackIcon(FrameworkElement icon, FrameworkElement image) {
    for (auto& target : g_live.icons) {
        if (target.icon.get() == icon) {
            target.image = image ? winrt::make_weak(image)
                                 : winrt::weak_ref<FrameworkElement>{};
            return;
        }
    }

    IconTarget target;
    target.icon = winrt::make_weak(icon);
    if (image) {
        target.image = winrt::make_weak(image);
    }
    g_live.icons.push_back(std::move(target));
}

std::wstring TrimCopy(std::wstring value) {
    auto isSpace = [](wchar_t c) {
        return iswspace(static_cast<wint_t>(c)) != 0;
    };
    while (!value.empty() && isSpace(value.front())) {
        value.erase(value.begin());
    }
    while (!value.empty() && isSpace(value.back())) {
        value.pop_back();
    }
    if (value.size() >= 2 &&
        ((value.front() == L'"' && value.back() == L'"') ||
         (value.front() == L'\'' && value.back() == L'\''))) {
        value = value.substr(1, value.size() - 2);
    }
    return value;
}

bool StartsWithIgnoreCase(const std::wstring& value, const wchar_t* prefix) {
    size_t length = wcslen(prefix);
    return value.size() >= length &&
           _wcsnicmp(value.c_str(), prefix, length) == 0;
}

std::wstring MakeFileUri(const std::wstring& path) {
    std::wstring result;
    size_t start = 0;
    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\') {
        result = L"file://";
        start = 2;
    } else {
        result = L"file:///";
    }

    for (size_t i = start; i < path.size(); i++) {
        wchar_t c = path[i];
        switch (c) {
            case L'\\':
                result += L'/';
                break;
            case L' ':
                result += L"%20";
                break;
            case L'#':
                result += L"%23";
                break;
            case L'%':
                result += L"%25";
                break;
            default:
                result += c;
                break;
        }
    }

    return result;
}

std::wstring ImageSourceToUri(const std::wstring& source) {
    std::wstring trimmed = TrimCopy(source);
    if (trimmed.empty()) {
        return {};
    }

    if (StartsWithIgnoreCase(trimmed, L"http://") ||
        StartsWithIgnoreCase(trimmed, L"https://") ||
        StartsWithIgnoreCase(trimmed, L"file:")) {
        return trimmed;
    }

    DWORD required = ExpandEnvironmentStringsW(trimmed.c_str(), nullptr, 0);
    if (!required) {
        return {};
    }

    std::wstring expanded(required, L'\0');
    DWORD written =
        ExpandEnvironmentStringsW(trimmed.c_str(), expanded.data(), required);
    if (!written || written > required) {
        return {};
    }
    expanded.resize(written - 1);
    if (expanded.empty()) {
        return {};
    }

    return MakeFileUri(expanded);
}

void EnsureCustomBitmap(const Settings& settings) {
    int generation = g_imageGeneration.load(std::memory_order_acquire);
    if (g_live.bitmapGeneration == generation &&
        (g_live.bitmap || g_live.imageFailed)) {
        return;
    }

    g_live.bitmapGeneration = generation;
    g_live.imageReady = false;
    g_live.imageFailed = false;
    g_live.bitmap = nullptr;

    std::wstring uri = ImageSourceToUri(settings.imagePath);
    if (uri.empty()) {
        g_live.imageFailed = true;
        Wh_Log(L"Custom image path is empty or invalid");
        return;
    }

    try {
        Imaging::BitmapImage bitmap;
        bitmap.CreateOptions(Imaging::BitmapCreateOptions::IgnoreImageCache);
        bitmap.AutoPlay(true);
        bitmap.ImageOpened([](auto&&, auto&&) {
            g_live.imageReady = true;
            g_live.imageFailed = false;
            Wh_Log(L"Custom Start image loaded");
            if (!g_unloading) {
                ApplySettingsFromTaskbarThread();
            }
        });
        bitmap.ImageFailed(
            [](auto&&, ExceptionRoutedEventArgs const& args) {
                g_live.imageReady = false;
                g_live.imageFailed = true;
                Wh_Log(L"Custom Start image failed: %s",
                       args.ErrorMessage().c_str());
                if (!g_unloading) {
                    ApplySettingsFromTaskbarThread();
                }
            });

        Wh_Log(L"Loading custom image: %s", uri.c_str());
        bitmap.UriSource(Foundation::Uri{uri});
        if (!g_live.imageFailed) {
            g_live.bitmap = bitmap;
        }
    } catch (const winrt::hresult_error& e) {
        g_live.imageFailed = true;
        g_live.bitmap = nullptr;
        Wh_Log(L"Custom image URI failed: %s", e.message().c_str());
    }
}

double IconSide(FrameworkElement icon) {
    double width = icon.ActualWidth();
    double height = icon.ActualHeight();
    if (width > 1 && height > 1) {
        return std::max(width, height);
    }

    width = icon.Width();
    height = icon.Height();
    if (width > 1 && height > 1 && !std::isnan(width) && !std::isnan(height)) {
        return std::max(width, height);
    }

    return 24;
}

Controls::Panel FindHostPanel(FrameworkElement icon,
                              FrameworkElement startButton) {
    DependencyObject current = icon;
    for (int depth = 0; current && depth < 8; depth++) {
        try {
            current = Media::VisualTreeHelper::GetParent(current);
        } catch (...) {
            return nullptr;
        }
        if (!current) {
            return nullptr;
        }

        if (auto panel = current.try_as<Controls::Panel>()) {
            return panel;
        }

        auto frameworkElement = current.try_as<FrameworkElement>();
        if (frameworkElement && frameworkElement == startButton) {
            break;
        }
    }

    return nullptr;
}

void RemoveNamedChild(Controls::Panel panel, PCWSTR name) {
    if (!panel) {
        return;
    }

    auto children = panel.Children();
    for (uint32_t i = children.Size(); i-- > 0;) {
        auto child = children.GetAt(i).try_as<FrameworkElement>();
        if (child && child.Name() == name) {
            children.RemoveAt(i);
        }
    }
}

FrameworkElement UpdateCustomImage(FrameworkElement startButton,
                                   FrameworkElement icon) {
    const Settings& settings = ActiveSettings();
    auto panel = FindHostPanel(icon, startButton);

    if (g_unloading || settings.imagePath.empty()) {
        if (panel) {
            RemoveNamedChild(panel, kCustomImageHostName);
        }
        icon.Opacity(1);
        g_live.bitmap = nullptr;
        g_live.imageReady = false;
        g_live.imageFailed = false;
        g_live.bitmapGeneration = -1;
        return nullptr;
    }

    if (!panel) {
        Wh_Log(L"Couldn't find a panel for the custom Start image");
        return nullptr;
    }

    EnsureCustomBitmap(settings);

    Controls::Grid host{nullptr};
    auto children = panel.Children();
    for (uint32_t i = 0; i < children.Size(); i++) {
        auto child = children.GetAt(i).try_as<Controls::Grid>();
        if (child && child.Name() == kCustomImageHostName) {
            host = child;
            break;
        }
    }

    if (!host) {
        host = Controls::Grid();
        host.Name(kCustomImageHostName);
        host.IsHitTestVisible(false);
        host.HorizontalAlignment(HorizontalAlignment::Center);
        host.VerticalAlignment(VerticalAlignment::Center);
        Controls::Canvas::SetZIndex(host, 10000);
        Automation::AutomationProperties::SetAccessibilityView(
            host, Automation::Peers::AccessibilityView::Raw);

        Controls::Image image;
        image.Stretch(Media::Stretch::Uniform);
        image.HorizontalAlignment(HorizontalAlignment::Stretch);
        image.VerticalAlignment(VerticalAlignment::Stretch);
        image.IsHitTestVisible(false);
        Automation::AutomationProperties::SetAccessibilityView(
            image, Automation::Peers::AccessibilityView::Raw);
        host.Children().Append(image);
        panel.Children().Append(host);
    }

    double side = settings.imageSize > 0 ? settings.imageSize : IconSide(icon);
    host.Width(side);
    host.Height(side);

    Controls::Image image{nullptr};
    if (host.Children().Size() > 0) {
        image = host.Children().GetAt(0).try_as<Controls::Image>();
    }

    int generation = g_imageGeneration.load(std::memory_order_acquire);
    if (image && g_live.bitmap && !g_live.imageFailed &&
        (!image.Source() || g_sourcesAppliedGeneration != generation)) {
        image.Source(g_live.bitmap);
    }
    if (image) {
        image.Opacity(std::clamp(settings.opacity, 0.0, 1.0));
    }

    icon.Opacity(g_live.imageReady ? 0 : 1);
    return host;
}

FrameworkElement FindStartButtonIcon(FrameworkElement startButton) {
    FrameworkElement icon = FindDescendantByName(startButton, L"Icon");
    if (!icon) {
        Wh_Log(L"Failed to find the Start button icon");
    }
    return icon;
}

void ApplyStartButtonStyle(FrameworkElement startButton) {
    if (g_applyingStyle || g_inLiveTick) {
        return;
    }

    BoolFlag styleGuard(g_applyingStyle);

    try {
        auto icon = FindStartButtonIcon(startButton);
        if (!icon) {
            return;
        }

        FrameworkElement image = UpdateCustomImage(startButton, icon);
        bool logoVisible = !(image && g_live.imageReady);
        if (logoVisible) {
            ApplyIconColors(icon, false);
            ApplyIconTransform(icon);
        } else {
            icon.ClearValue(UIElement::RenderTransformProperty());
            icon.ClearValue(UIElement::RenderTransformOriginProperty());
        }
        if (image) {
            ApplyIconTransform(image);
        }
        TrackIcon(icon, image);
    } catch (const winrt::hresult_error& e) {
        Wh_Log(L"ApplyStartButtonStyle failed: %s", e.message().c_str());
    } catch (...) {
        Wh_Log(L"ApplyStartButtonStyle failed");
    }
}

bool ApplyStyle(XamlRoot xamlRoot) {
    FrameworkElement child = xamlRoot.Content().try_as<FrameworkElement>();
    if (!child ||
        !(child = FindChildByClassName(child, L"Taskbar.TaskbarFrame")) ||
        !(child = FindChildByName(child, L"RootGrid")) ||
        !(child = FindChildByName(child, L"TaskbarFrameRepeater"))) {
        return false;
    }

    FrameworkElement startButton = FindDescendant(child, IsStartButton);
    if (!startButton) {
        Wh_Log(L"Failed to find the Start button");
        return false;
    }

    ApplyStartButtonStyle(startButton);
    return true;
}

void* CTaskBand_ITaskListWndSite_vftable;
void* CSecondaryTaskBand_ITaskListWndSite_vftable;

using CTaskBand_GetTaskbarHost_t = void*(WINAPI*)(void* pThis, void** result);
CTaskBand_GetTaskbarHost_t CTaskBand_GetTaskbarHost_Original;
void* TaskbarHost_FrameHeight_Original;

using CSecondaryTaskBand_GetTaskbarHost_t = void*(WINAPI*)(void* pThis,
                                                           void** result);
CSecondaryTaskBand_GetTaskbarHost_t CSecondaryTaskBand_GetTaskbarHost_Original;

using std__Ref_count_base__Decref_t = void(WINAPI*)(void* pThis);
std__Ref_count_base__Decref_t std__Ref_count_base__Decref_Original;

XamlRoot XamlRootFromTaskbarHostSharedPtr(void* taskbarHostSharedPtr[2]) {
    if (!taskbarHostSharedPtr[0] && !taskbarHostSharedPtr[1]) {
        return nullptr;
    }

    size_t taskbarElementIUnknownOffset = 0x10;

#if defined(_M_X64)
    {
        const BYTE* b = (const BYTE*)TaskbarHost_FrameHeight_Original;
        if (b[0] == 0x48 && b[1] == 0x83 && b[2] == 0xEC && b[4] == 0x48 &&
            b[5] == 0x83 && b[6] == 0xC1 && b[7] <= 0x7F) {
            taskbarElementIUnknownOffset = b[7];
        } else {
            Wh_Log(L"Unsupported TaskbarHost::FrameHeight");
        }
    }
#elif defined(_M_ARM64)
    {
        const DWORD* p = (const DWORD*)TaskbarHost_FrameHeight_Original;
        if (p[0] == 0xD503237F && (p[1] & 0xFFC07FFF) == 0xA9807BFD &&
            p[2] == 0x910003FD && (p[3] & 0xFFF00FE0) == 0xF8400C00) {
            taskbarElementIUnknownOffset = (p[3] >> 12) & 0xFF;
        } else {
            Wh_Log(L"Unsupported TaskbarHost::FrameHeight");
        }
    }
#else
#error "Unsupported architecture"
#endif

    auto* taskbarElementIUnknown =
        *(IUnknown**)((BYTE*)taskbarHostSharedPtr[0] +
                      taskbarElementIUnknownOffset);

    FrameworkElement taskbarElement = nullptr;
    taskbarElementIUnknown->QueryInterface(winrt::guid_of<FrameworkElement>(),
                                           winrt::put_abi(taskbarElement));

    auto result = taskbarElement ? taskbarElement.XamlRoot() : nullptr;
    std__Ref_count_base__Decref_Original(taskbarHostSharedPtr[1]);
    return result;
}

XamlRoot GetTaskbarXamlRoot(HWND hTaskbarWnd) {
    HWND hTaskSwWnd = (HWND)GetProp(hTaskbarWnd, L"TaskbandHWND");
    if (!hTaskSwWnd) {
        return nullptr;
    }

    void* taskBand = (void*)GetWindowLongPtr(hTaskSwWnd, 0);
    void* taskBandForTaskListWndSite = taskBand;
    for (int i = 0; *(void**)taskBandForTaskListWndSite !=
                    CTaskBand_ITaskListWndSite_vftable;
         i++) {
        if (i == 20) {
            return nullptr;
        }
        taskBandForTaskListWndSite = (void**)taskBandForTaskListWndSite + 1;
    }

    void* taskbarHostSharedPtr[2]{};
    CTaskBand_GetTaskbarHost_Original(taskBandForTaskListWndSite,
                                      taskbarHostSharedPtr);
    return XamlRootFromTaskbarHostSharedPtr(taskbarHostSharedPtr);
}

XamlRoot GetSecondaryTaskbarXamlRoot(HWND hSecondaryTaskbarWnd) {
    HWND hTaskSwWnd = (HWND)FindWindowEx(hSecondaryTaskbarWnd, nullptr,
                                         L"WorkerW", nullptr);
    if (!hTaskSwWnd) {
        return nullptr;
    }

    void* taskBand = (void*)GetWindowLongPtr(hTaskSwWnd, 0);
    void* taskBandForTaskListWndSite = taskBand;
    for (int i = 0; *(void**)taskBandForTaskListWndSite !=
                    CSecondaryTaskBand_ITaskListWndSite_vftable;
         i++) {
        if (i == 20) {
            return nullptr;
        }
        taskBandForTaskListWndSite = (void**)taskBandForTaskListWndSite + 1;
    }

    void* taskbarHostSharedPtr[2]{};
    CSecondaryTaskBand_GetTaskbarHost_Original(taskBandForTaskListWndSite,
                                               taskbarHostSharedPtr);
    return XamlRootFromTaskbarHostSharedPtr(taskbarHostSharedPtr);
}

using RunFromWindowThreadProc_t = void(WINAPI*)(void* parameter);

bool RunFromWindowThread(HWND hWnd,
                         RunFromWindowThreadProc_t proc,
                         void* procParam) {
    static const UINT runFromWindowThreadRegisteredMsg =
        RegisterWindowMessage(L"Windhawk_RunFromWindowThread_" WH_MOD_ID);

    struct RUN_FROM_WINDOW_THREAD_PARAM {
        RunFromWindowThreadProc_t proc;
        void* procParam;
    };

    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    if (dwThreadId == 0) {
        return false;
    }

    if (dwThreadId == GetCurrentThreadId()) {
        proc(procParam);
        return true;
    }

    HHOOK hook = SetWindowsHookEx(
        WH_CALLWNDPROC,
        [](int nCode, WPARAM wParam, LPARAM lParam) -> LRESULT {
            if (nCode == HC_ACTION) {
                const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
                if (cwp->message == runFromWindowThreadRegisteredMsg) {
                    RUN_FROM_WINDOW_THREAD_PARAM* param =
                        (RUN_FROM_WINDOW_THREAD_PARAM*)cwp->lParam;
                    param->proc(param->procParam);
                }
            }

            return CallNextHookEx(nullptr, nCode, wParam, lParam);
        },
        nullptr, dwThreadId);
    if (!hook) {
        return false;
    }

    RUN_FROM_WINDOW_THREAD_PARAM param;
    param.proc = proc;
    param.procParam = procParam;
    SendMessage(hWnd, runFromWindowThreadRegisteredMsg, 0, (LPARAM)&param);
    UnhookWindowsHookEx(hook);
    return true;
}

void ApplySettingsFromTaskbarThread() {
    if (g_inApply) {
        return;
    }

    BoolFlag applyGuard(g_inApply);
    Wh_Log(L"Applying settings");

    Settings settings = GetSettings();
    SettingsScope scope(settings);

    if (g_unloading || !WantsLiveTimer(settings)) {
        StopLiveTimer();
    }

    if (g_unloading) {
        g_iconLoadedRevokers.reset();
    }

    if (g_unloading || !settings.accentColor) {
        StopTrackingAccentColor();
    }

    g_live.icons.clear();

    EnumThreadWindows(
        GetCurrentThreadId(),
        [](HWND hWnd, LPARAM) -> BOOL {
            WCHAR szClassName[32];
            if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
                return TRUE;
            }

            XamlRoot xamlRoot = nullptr;
            if (_wcsicmp(szClassName, L"Shell_TrayWnd") == 0) {
                xamlRoot = GetTaskbarXamlRoot(hWnd);
            } else if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
                xamlRoot = GetSecondaryTaskbarXamlRoot(hWnd);
            } else {
                return TRUE;
            }

            if (!xamlRoot) {
                Wh_Log(L"Getting XamlRoot failed");
                return TRUE;
            }

            if (!ApplyStyle(xamlRoot)) {
                Wh_Log(L"ApplyStyle failed");
            }

            return TRUE;
        },
        0);

    if (!g_live.imageFailed && g_live.bitmap) {
        g_sourcesAppliedGeneration =
            g_imageGeneration.load(std::memory_order_acquire);
    }

    if (g_unloading) {
        // Drop these on the taskbar thread. The timer callback and the image
        // decode events also run here, so releasing them elsewhere races.
        g_live.icons.clear();
        g_live.bitmap = nullptr;
        g_live.imageReady = false;
        return;
    }

    if (WantsLiveTimer(settings)) {
        StartLiveTimer();
    }
}

void ApplySettings(HWND hTaskbarWnd) {
    RunFromWindowThread(
        hTaskbarWnd, [](void*) { ApplySettingsFromTaskbarThread(); }, nullptr);
}

FrameworkElement GetStartButtonElement(void* pThis) {
    FrameworkElement element = nullptr;
    ((IUnknown**)pThis)[1]->QueryInterface(winrt::guid_of<FrameworkElement>(),
                                           winrt::put_abi(element));
    if (!element || !IsStartButton(element)) {
        return nullptr;
    }
    return element;
}

using ExperienceToggleButton_UpdateVisualStates_t = void(WINAPI*)(void* pThis);
ExperienceToggleButton_UpdateVisualStates_t
    ExperienceToggleButton_UpdateVisualStates_Original;
void WINAPI ExperienceToggleButton_UpdateVisualStates_Hook(void* pThis) {
    Wh_Log(L">");
    ExperienceToggleButton_UpdateVisualStates_Original(pThis);

    if (auto startButton = GetStartButtonElement(pThis)) {
        Settings settings = GetSettings();
        SettingsScope scope(settings);
        ApplyStartButtonStyle(startButton);
    }
}

void ApplyIconStyleOnceLoaded(FrameworkElement icon) {
    g_iconLoadedRevokers->emplace_back();
    auto revokerIt = std::prev(g_iconLoadedRevokers->end());

    *revokerIt = icon.Loaded(
        winrt::auto_revoke,
        [revokerIt](winrt::Windows::Foundation::IInspectable const& sender,
                    RoutedEventArgs const&) {
            Wh_Log(L">");
            g_iconLoadedRevokers->erase(revokerIt);

            auto iconElement = sender.try_as<FrameworkElement>();
            if (!iconElement) {
                return;
            }

            if (auto startButton = FindAncestorStartButton(iconElement)) {
                Settings settings = GetSettings();
                SettingsScope scope(settings);
                ApplyStartButtonStyle(startButton);
            }
        });
}

using ExperienceToggleButton_InitializeAnimatedVisualPlayer_t =
    void(WINAPI*)(void* pThis);
ExperienceToggleButton_InitializeAnimatedVisualPlayer_t
    ExperienceToggleButton_InitializeAnimatedVisualPlayer_Original;
void WINAPI
ExperienceToggleButton_InitializeAnimatedVisualPlayer_Hook(void* pThis) {
    Wh_Log(L">");

    FrameworkElement startButton = GetStartButtonElement(pThis);
    if (startButton) {
        EnsureColorKeyFrameAnimationHooked(startButton);
    }

    ExperienceToggleButton_InitializeAnimatedVisualPlayer_Original(pThis);

    if (!startButton || g_unloading) {
        return;
    }

    FrameworkElement icon = FindStartButtonIcon(startButton);
    if (!icon) {
        return;
    }

    if (!Hosting::ElementCompositionPreview::GetElementChildVisual(icon)) {
        Wh_Log(L"Waiting for the icon to load");
        ApplyIconStyleOnceLoaded(icon);
        return;
    }

    Settings settings = GetSettings();
    SettingsScope scope(settings);
    ApplyStartButtonStyle(startButton);
}

using StartIconAnimatedVisual_CreateAnimations_t = void(WINAPI*)(void* pThis);

void CreateIconAnimations(void* pThis,
                          StartIconAnimatedVisual_CreateAnimations_t original) {
    Settings settings = GetSettings();
    SettingsScope scope(settings);

    struct FlagGuard {
        bool saved = g_creatingIconAnimations;
        ~FlagGuard() { g_creatingIconAnimations = saved; }
    } guard;

    g_creatingIconAnimations = !g_unloading && settings.customizeColors;
    original(pThis);
}

StartIconAnimatedVisual_CreateAnimations_t
    StartDarkAnimatedVisual_CreateAnimations_Original;
void WINAPI StartDarkAnimatedVisual_CreateAnimations_Hook(void* pThis) {
    Wh_Log(L">");
    CreateIconAnimations(pThis, StartDarkAnimatedVisual_CreateAnimations_Original);
}

StartIconAnimatedVisual_CreateAnimations_t
    StartLightAnimatedVisual_CreateAnimations_Original;
void WINAPI StartLightAnimatedVisual_CreateAnimations_Hook(void* pThis) {
    Wh_Log(L">");
    CreateIconAnimations(pThis,
                         StartLightAnimatedVisual_CreateAnimations_Original);
}

bool HookTaskbarViewDllSymbols(HMODULE module) {
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(protected: virtual void __cdecl winrt::Taskbar::implementation::ExperienceToggleButton::UpdateVisualStates(void))"},
            &ExperienceToggleButton_UpdateVisualStates_Original,
            ExperienceToggleButton_UpdateVisualStates_Hook,
        },
        {
            {LR"(private: void __cdecl winrt::Taskbar::implementation::ExperienceToggleButton::InitializeAnimatedVisualPlayer(void))"},
            &ExperienceToggleButton_InitializeAnimatedVisualPlayer_Original,
            ExperienceToggleButton_InitializeAnimatedVisualPlayer_Hook,
        },
        {
            {LR"(public: void __cdecl winrt::AnimatedVisuals::implementation::StartDark_AnimatedVisual::CreateAnimations(void))"},
            &StartDarkAnimatedVisual_CreateAnimations_Original,
            StartDarkAnimatedVisual_CreateAnimations_Hook,
            true,
        },
        {
            {LR"(public: void __cdecl winrt::AnimatedVisuals::implementation::StartLight_AnimatedVisual::CreateAnimations(void))"},
            &StartLightAnimatedVisual_CreateAnimations_Original,
            StartLightAnimatedVisual_CreateAnimations_Hook,
            true,
        },
    };

    return HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks));
}

bool HookTaskbarDllSymbols() {
    HMODULE module =
        LoadLibraryEx(L"taskbar.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!module) {
        Wh_Log(L"Failed to load taskbar.dll");
        return false;
    }

    WindhawkUtils::SYMBOL_HOOK taskbarDllHooks[] = {
        {
            {LR"(const CTaskBand::`vftable'{for `ITaskListWndSite'})"},
            &CTaskBand_ITaskListWndSite_vftable,
        },
        {
            {LR"(const CSecondaryTaskBand::`vftable'{for `ITaskListWndSite'})"},
            &CSecondaryTaskBand_ITaskListWndSite_vftable,
        },
        {
            {LR"(public: virtual class std::shared_ptr<class TaskbarHost> __cdecl CTaskBand::GetTaskbarHost(void)const )"},
            &CTaskBand_GetTaskbarHost_Original,
        },
        {
            {LR"(public: int __cdecl TaskbarHost::FrameHeight(void)const )"},
            &TaskbarHost_FrameHeight_Original,
        },
        {
            {LR"(public: virtual class std::shared_ptr<class TaskbarHost> __cdecl CSecondaryTaskBand::GetTaskbarHost(void)const )"},
            &CSecondaryTaskBand_GetTaskbarHost_Original,
        },
        {
            {LR"(public: void __cdecl std::_Ref_count_base::_Decref(void))"},
            &std__Ref_count_base__Decref_Original,
        },
    };

    return HookSymbols(module, taskbarDllHooks, ARRAYSIZE(taskbarDllHooks));
}

HMODULE GetTaskbarViewModuleHandle() {
    HMODULE module = GetModuleHandle(L"Taskbar.View.dll");
    if (!module) {
        module = GetModuleHandle(L"ExplorerExtensions.dll");
    }
    return module;
}

void HandleLoadedModuleIfTaskbarView(HMODULE module, LPCWSTR lpLibFileName) {
    if (!g_taskbarViewDllLoaded && GetTaskbarViewModuleHandle() == module &&
        !g_taskbarViewDllLoaded.exchange(true)) {
        Wh_Log(L"Loaded %s", lpLibFileName);

        if (HookTaskbarViewDllSymbols(module)) {
            Wh_ApplyHookOperations();
        }
    }
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module) {
        HandleLoadedModuleIfTaskbarView(module, lpLibFileName);
    }
    return module;
}

std::wstring ReadStringSetting(PCWSTR name) {
    PCWSTR value = Wh_GetStringSetting(name);
    std::wstring copy = value ? value : L"";
    Wh_FreeStringSetting(value);
    return copy;
}

void LoadSettings() {
    Settings settings;

    static const struct {
        PCWSTR name;
        int hue;
    } colors[] = {
        {L"red", 0},    {L"orange", 30}, {L"yellow", 55},  {L"green", 120},
        {L"teal", 175}, {L"blue", 215},  {L"purple", 275}, {L"pink", 325},
    };

    std::wstring color = ReadStringSetting(L"color");
    settings.accentColor = color == L"accent";
    settings.colorHue = -1;
    for (const auto& [name, hue] : colors) {
        if (color == name) {
            settings.colorHue = hue;
            break;
        }
    }

    settings.hue = std::clamp(Wh_GetIntSetting(L"effects.hue"), -360, 360);
    settings.saturation = Wh_GetIntSetting(L"effects.saturation") / 100.0;
    settings.brightness = Wh_GetIntSetting(L"effects.brightness") / 100.0;
    settings.opacity = Wh_GetIntSetting(L"effects.opacity") / 100.0;
    settings.size = Wh_GetIntSetting(L"size") / 100.0;

    settings.angle = std::clamp(Wh_GetIntSetting(L"rotation.angle"), -3600, 3600);
    settings.spin = std::clamp(Wh_GetIntSetting(L"rotation.spin"), -1080, 1080);

    settings.liveEnabled = Wh_GetIntSetting(L"live.enabled") != 0;
    std::wstring liveStyle = ReadStringSetting(L"live.style");
    settings.liveGradient = liveStyle != L"cycle";
    settings.liveSpeed =
        std::clamp(Wh_GetIntSetting(L"live.speed"), -720, 720);
    settings.liveSpread =
        std::clamp(Wh_GetIntSetting(L"live.spread"), 0, 360);

    settings.imagePath = TrimCopy(ReadStringSetting(L"image.path"));
    settings.imageSize =
        std::clamp(Wh_GetIntSetting(L"image.size"), 0, 256);

    settings.customizeColors =
        settings.accentColor || settings.colorHue >= 0 || settings.hue != 0 ||
        settings.saturation != 1 || settings.brightness != 1 ||
        settings.opacity != 1 || settings.liveEnabled;

    std::lock_guard lock(g_settingsMutex);
    g_settings = std::move(settings);
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    if (!HookTaskbarDllSymbols()) {
        return FALSE;
    }

    if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
        g_taskbarViewDllLoaded = true;
        if (!HookTaskbarViewDllSymbols(taskbarViewModule)) {
            return FALSE;
        }
    } else {
        Wh_Log(L"Taskbar view module not loaded yet");

        HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
        auto pKernelBaseLoadLibraryExW =
            (decltype(&LoadLibraryExW))GetProcAddress(kernelBaseModule,
                                                      "LoadLibraryExW");
        WindhawkUtils::SetFunctionHook(pKernelBaseLoadLibraryExW,
                                       LoadLibraryExW_Hook,
                                       &LoadLibraryExW_Original);
    }

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (!g_taskbarViewDllLoaded) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            if (!g_taskbarViewDllLoaded.exchange(true)) {
                Wh_Log(L"Got Taskbar.View.dll");

                if (HookTaskbarViewDllSymbols(taskbarViewModule)) {
                    Wh_ApplyHookOperations();
                }
            }
        }
    }

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        ApplySettings(hTaskbarWnd);
    }
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");

    g_unloading = true;

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        ApplySettings(hTaskbarWnd);
    }
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    g_imageGeneration.fetch_add(1, std::memory_order_acq_rel);
    LoadSettings();

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        ApplySettings(hTaskbarWnd);
    }
}
