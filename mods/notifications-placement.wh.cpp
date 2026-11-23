// ==WindhawkMod==
// @id              notifications-placement
// @name            Customize Windows notifications placement
// @description     Move notifications to another monitor or another corner of the screen
// @version         1.2.4
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @include         ShellExperienceHost.exe
// @architecture    x86-64
// @compilerOptions -lole32 -loleaut32 -lruntimeobject -lshcore
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
# Customize Windows notifications placement

Move notifications to another monitor or another corner of the screen.

Only Windows 10 64-bit and Windows 11 are supported.

![Screenshot](https://i.imgur.com/4PxMvLg.png)

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
5. Trigger a notification to appear.
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
- monitor: 1
  $name: Monitor
  $description: >-
    The monitor number that notifications will appear on, set to zero to use the
    monitor where the mouse cursor is located
- monitorInterfaceName: ""
  $name: Monitor interface name
  $description: >-
    If not empty, the given monitor interface name (can also be an interface
    name substring) will be used instead of the monitor number. Can be useful if
    the monitor numbers change often. To see all available interface names, set
    any interface name, enable mod logs, trigger a notification and look for
    "Found display device" messages.
- horizontalPlacement: right
  $name: Horizontal placement on the screen
  $options:
  - right: Right
  - left: Left
  - center: Center
- horizontalDistanceFromScreenEdge: 0
  $name: Distance from the right/left side of the screen
- verticalPlacement: bottom
  $name: Vertical placement on the screen
  $options:
  - bottom: Bottom
  - top: Top
  - center: Center
- verticalDistanceFromScreenEdge: 0
  $name: Distance from the bottom/top side of the screen
- animationDirection: automatic
  $name: Notification appearance animation direction
  $options:
  - automatic: Automatic
  - fromLeft: From left
  - fromRight: From right
  - fromTop: From top
  - fromBottom: From bottom
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#undef GetCurrentTime

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.UI.Input.h>
#include <winrt/Windows.UI.Xaml.Input.h>
#include <winrt/Windows.UI.Xaml.Media.Animation.h>
#include <winrt/Windows.UI.Xaml.Media.h>
#include <winrt/Windows.UI.Xaml.h>
#include <winrt/base.h>

#include <algorithm>
#include <atomic>
#include <cmath>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

using namespace winrt::Windows::UI::Xaml;

std::atomic<bool> g_unloading;

enum class HorizontalPlacement {
    right,
    left,
    center,
};

enum class VerticalPlacement {
    bottom,
    top,
    center,
};

enum class AnimationDirection {
    automatic,
    fromLeft,
    fromRight,
    fromTop,
    fromBottom,
};

struct {
    int monitor;
    WindhawkUtils::StringSetting monitorInterfaceName;
    HorizontalPlacement horizontalPlacement;
    int horizontalDistanceFromScreenEdge;
    VerticalPlacement verticalPlacement;
    int verticalDistanceFromScreenEdge;
    AnimationDirection animationDirection;
} g_settings;

enum class Target {
    Explorer,
    ShellExperienceHost,
};

Target g_target;

bool g_inCToastCenterExperienceManager_PositionView;

// Only accessed on the thread of the notification window.
std::optional<VerticalAlignment> g_originalToastCenterViewAlignment;

WINUSERAPI UINT WINAPI GetDpiForWindow(HWND hwnd);
typedef enum MONITOR_DPI_TYPE {
    MDT_EFFECTIVE_DPI = 0,
    MDT_ANGULAR_DPI = 1,
    MDT_RAW_DPI = 2,
    MDT_DEFAULT = MDT_EFFECTIVE_DPI
} MONITOR_DPI_TYPE;
STDAPI GetDpiForMonitor(HMONITOR hmonitor,
                        MONITOR_DPI_TYPE dpiType,
                        UINT* dpiX,
                        UINT* dpiY);

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

bool GetMonitorWorkArea(HMONITOR monitor, RECT* rc) {
    MONITORINFO monitorInfo{
        .cbSize = sizeof(MONITORINFO),
    };
    return GetMonitorInfo(monitor, &monitorInfo) &&
           CopyRect(rc, &monitorInfo.rcWork);
}

std::wstring GetProcessFileName(DWORD dwProcessId) {
    HANDLE hProcess =
        OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dwProcessId);
    if (!hProcess) {
        return std::wstring{};
    }

    WCHAR processPath[MAX_PATH];

    DWORD dwSize = ARRAYSIZE(processPath);
    if (!QueryFullProcessImageName(hProcess, 0, processPath, &dwSize)) {
        CloseHandle(hProcess);
        return std::wstring{};
    }

    CloseHandle(hProcess);

    PCWSTR processFileName = wcsrchr(processPath, L'\\');
    if (!processFileName) {
        return std::wstring{};
    }

    processFileName++;
    return processFileName;
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

bool IsTargetCoreWindow(HWND hWnd) {
    DWORD processId = 0;
    if (!hWnd || !GetWindowThreadProcessId(hWnd, &processId)) {
        return false;
    }

    if (_wcsicmp(GetProcessFileName(processId).c_str(),
                 L"ShellExperienceHost.exe") != 0) {
        return false;
    }

    WCHAR szClassName[32];
    if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0 ||
        _wcsicmp(szClassName, L"Windows.UI.Core.CoreWindow") != 0) {
        return false;
    }

    // The window title is locale-dependent, and unfortunately I didn't find a
    // simpler way to identify the target window.
    // String source: Windows.UI.ShellCommon.<locale>.pri
    // String resource: \ActionCenter\AC_ToastCenter_Title
    // The strings were collected from here:
    // https://github.com/m417z/windows-language-files

    // clang-format off
    static const std::unordered_set<std::wstring> newNotificationStrings = {
        L"Nuwe kennisgewing", // AF-ZA
        L"አዲስ ማሳወቂያ", // AM-ET
        L"নতুন জাননী", // AS-IN
        L"Yeni bildiriş", // AZ-LATN-AZ
        L"Новае апавяшчэнне", // BE-BY
        L"নতুন বিজ্ঞপ্তি", // BN-IN
        L"Novo obavještenje", // BS-LATN-BA
        L"Notificació nova", // CA-ES-VALENCIA
        L"ᎢᏤᎢ ᎧᏃᎮᏓ", // CHR-CHER-US
        L"Hysbysiad newydd", // CY-GB
        L"اعلان جدید", // FA-IR
        L"Bagong notification", // FIL-PH
        L"Fógra nua", // GA-IE
        L"Brath ùr", // GD-GB
        L"નવી સૂચના", // GU-IN
        L"नई अधिसूचना", // HI-IN
        L"Նոր ծանուցում", // HY-AM
        L"Ný tilkynning", // IS-IS
        L"ახალი შეტყობინება", // KA-GE
        L"Жаңа хабарландыру", // KK-KZ
        L"ការ\u200bជូន\u200bដំណឹង\u200bថ្មី", // KM-KH
        L"ಹೊಸ ಪ್ರಕಟಣೆ", // KN-IN
        L"नवी अधिसुचोवणी", // KOK-IN
        L"Nei Notifikatioun", // LB-LU
        L"ການແຈ້ງເຕືອນໃໝ່", // LO-LA
        L"Whakamōhiotanga hōu", // MI-NZ
        L"Ново известување", // MK-MK
        L"പുതിയ അറിയിപ്പ്", // ML-IN
        L"नवीन सूचना", // MR-IN
        L"Pemberitahuan baharu", // MS-MY
        L"Notifika ġdida", // MT-MT
        L"नयाँ सूचना", // NE-NP
        L"Nytt varsel", // NN-NO
        L"ନୂତନ ବିଜ୍ଞପ୍ତି", // OR-IN
        L"ਨਵੀਂ ਸੂਚਨਾ", // PA-IN
        L"Musuq willana", // QUZ-PE
        L"Njoftim i ri", // SQ-AL
        L"Ново обавјештење", // SR-CYRL-BA
        L"Ново обавештење", // SR-CYRL-RS
        L"புதிய அறிவிப்பு", // TA-IN
        L"కొత్త నోటిఫికేషన్", // TE-IN
        L"Яңа белдерү", // TT-RU
        L"يېڭى ئۇقتۇرۇش", // UG-CN
        L"نئی اطلاع", // UR-PK
        L"Yangi xabarnoma", // UZ-LATN-UZ
        L"\u200f\u200fإعلام جديد", // AR-SA
        L"Ново известие", // BG-BG
        L"Notificació nova", // CA-ES
        L"Nové oznámení", // CS-CZ
        L"Ny meddelelse", // DA-DK
        L"Neue Benachrichtigung", // DE-DE
        L"Νέα ειδοποίηση", // EL-GR
        L"New notification", // EN-GB
        L"New notification", // EN-US
        L"Notificación nueva", // ES-ES
        L"Nueva notificación", // ES-MX
        L"Uus teatis", // ET-EE
        L"Jakinarazpen berria", // EU-ES
        L"Uusi ilmoitus", // FI-FI
        L"Nouvelle notification", // FR-CA
        L"Nouvelle notification", // FR-FR
        L"Nova notificación", // GL-ES
        L"הודעה חדשה", // HE-IL
        L"Nova obavijesti", // HR-HR
        L"Új értesítés", // HU-HU
        L"Pemberitahuan baru", // ID-ID
        L"Nuova notifica", // IT-IT
        L"新しい通知", // JA-JP
        L"새 알림", // KO-KR
        L"Naujas pranešimas", // LT-LT
        L"Jauns paziņojums", // LV-LV
        L"Ny varsling", // NB-NO
        L"Nieuwe melding", // NL-NL
        L"Nowe powiadomienie", // PL-PL
        L"Nova notificação", // PT-BR
        L"Nova notificação", // PT-PT
        L"Notificare nouă", // RO-RO
        L"Новое уведомление", // RU-RU
        L"Nové oznámenie", // SK-SK
        L"Novo obvestilo", // SL-SI
        L"Novo obaveštenje", // SR-LATN-RS
        L"Nytt meddelande", // SV-SE
        L"การแจ้งให้ทราบใหม่", // TH-TH
        L"Yeni bildirim", // TR-TR
        L"Нове сповіщення", // UK-UA
        L"Thông báo mới", // VI-VN
        L"新通知", // ZH-CN
        L"新通知", // ZH-TW
    };
    // clang-format on

    WCHAR szWindowText[256]{};
    if (GetWindowText(hWnd, szWindowText, ARRAYSIZE(szWindowText)) == 0 ||
        !newNotificationStrings.contains(szWindowText)) {
        Wh_Log(L"Not targeting CoreWindow, window text: %s", szWindowText);
        return false;
    }

    return true;
}

std::vector<HWND> GetCoreWindows() {
    struct ENUM_WINDOWS_PARAM {
        std::vector<HWND>* hWnds;
    };

    std::vector<HWND> hWnds;
    ENUM_WINDOWS_PARAM param = {&hWnds};
    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            ENUM_WINDOWS_PARAM& param = *(ENUM_WINDOWS_PARAM*)lParam;

            if (IsTargetCoreWindow(hWnd)) {
                param.hWnds->push_back(hWnd);
            }

            return TRUE;
        },
        (LPARAM)&param);

    return hWnds;
}

void AdjustCoreWindowPos(int* x, int* y, int* cx, int* cy) {
    Wh_Log(L"Before: %dx%d %dx%d", *x, *y, *cx, *cy);

    RECT rc{
        .left = *x,
        .top = *y,
        .right = *x + *cx,
        .bottom = *y + *cy,
    };
    HMONITOR srcMonitor = MonitorFromRect(&rc, MONITOR_DEFAULTTONEAREST);

    UINT srcMonitorDpiX = 96;
    UINT srcMonitorDpiY = 96;
    GetDpiForMonitor(srcMonitor, MDT_DEFAULT, &srcMonitorDpiX, &srcMonitorDpiY);

    RECT srcMonitorWorkArea;
    if (!GetMonitorWorkArea(srcMonitor, &srcMonitorWorkArea)) {
        return;
    }

    HMONITOR destMonitor = nullptr;

    if (!g_unloading) {
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
    }

    if (!destMonitor) {
        HMONITOR primaryMonitor =
            MonitorFromPoint({0, 0}, MONITOR_DEFAULTTONEAREST);

        destMonitor = primaryMonitor;
    }

    RECT destMonitorWorkArea;
    int horizontalDistanceFromScreenEdge = 0;
    int verticalDistanceFromScreenEdge = 0;

    Wh_Log(L"Monitor %p->%p", srcMonitor, destMonitor);

    if (destMonitor != srcMonitor) {
        UINT destMonitorDpiX = 96;
        UINT destMonitorDpiY = 96;
        GetDpiForMonitor(destMonitor, MDT_DEFAULT, &destMonitorDpiX,
                         &destMonitorDpiY);

        if (!GetMonitorWorkArea(destMonitor, &destMonitorWorkArea)) {
            return;
        }

        *cx = MulDiv(*cx, destMonitorDpiX, srcMonitorDpiX);
        if (*y + *cy == srcMonitorWorkArea.bottom) {
            *y = destMonitorWorkArea.bottom -
                 MulDiv(*cy, destMonitorDpiY, srcMonitorDpiY);
            *cy = MulDiv(*cy, destMonitorDpiY, srcMonitorDpiY);
        } else {
            *cy = MulDiv(*cy, destMonitorDpiY, srcMonitorDpiY);
        }

        if (*y == destMonitorWorkArea.top &&
            *y + *cy > destMonitorWorkArea.bottom) {
            *cy = destMonitorWorkArea.bottom - destMonitorWorkArea.top;
        }

        if (!g_unloading) {
            horizontalDistanceFromScreenEdge =
                MulDiv(g_settings.horizontalDistanceFromScreenEdge,
                       destMonitorDpiX, 96);
            verticalDistanceFromScreenEdge = MulDiv(
                g_settings.verticalDistanceFromScreenEdge, destMonitorDpiY, 96);
        }
    } else {
        CopyRect(&destMonitorWorkArea, &srcMonitorWorkArea);

        if (!g_unloading) {
            horizontalDistanceFromScreenEdge =
                MulDiv(g_settings.horizontalDistanceFromScreenEdge,
                       srcMonitorDpiX, 96);
            verticalDistanceFromScreenEdge = MulDiv(
                g_settings.verticalDistanceFromScreenEdge, srcMonitorDpiY, 96);
        }
    }

    switch (g_unloading ? HorizontalPlacement::right
                        : g_settings.horizontalPlacement) {
        case HorizontalPlacement::right:
            *x = destMonitorWorkArea.right - *cx -
                 horizontalDistanceFromScreenEdge;
            break;

        case HorizontalPlacement::left:
            *x = destMonitorWorkArea.left + horizontalDistanceFromScreenEdge;
            break;

        case HorizontalPlacement::center:
            *x = destMonitorWorkArea.left +
                 (destMonitorWorkArea.right - destMonitorWorkArea.left - *cx) /
                     2 +
                 horizontalDistanceFromScreenEdge;
            break;
    }

    switch (g_unloading ? VerticalPlacement::bottom
                        : g_settings.verticalPlacement) {
        case VerticalPlacement::bottom:
            *y = destMonitorWorkArea.bottom - *cy -
                 verticalDistanceFromScreenEdge;
            break;

        case VerticalPlacement::top:
            *y = destMonitorWorkArea.top + verticalDistanceFromScreenEdge;
            break;

        case VerticalPlacement::center:
            *y = destMonitorWorkArea.top +
                 (destMonitorWorkArea.bottom - destMonitorWorkArea.top - *cy) /
                     2 +
                 verticalDistanceFromScreenEdge;
            break;
    }

    Wh_Log(L"After: %dx%d %dx%d", *x, *y, *cx, *cy);
}

using CToastCenterExperienceManager_PositionView_t =
    HRESULT(WINAPI*)(void* pThis);
CToastCenterExperienceManager_PositionView_t
    CToastCenterExperienceManager_PositionView_Original;
HRESULT WINAPI CToastCenterExperienceManager_PositionView_Hook(void* pThis) {
    Wh_Log(L">");

    g_inCToastCenterExperienceManager_PositionView = true;
    HRESULT ret = CToastCenterExperienceManager_PositionView_Original(pThis);
    g_inCToastCenterExperienceManager_PositionView = false;

    return ret;
}

using MonitorFromPoint_t = decltype(&MonitorFromPoint);
MonitorFromPoint_t MonitorFromPoint_Original;
HMONITOR WINAPI MonitorFromPoint_Hook(POINT pt, DWORD dwFlags) {
    Wh_Log(L">");

    if (g_inCToastCenterExperienceManager_PositionView && !g_unloading &&
        pt.x == 0 && pt.y == 0) {
        HMONITOR monitor = nullptr;

        if (*g_settings.monitorInterfaceName.get()) {
            monitor = GetMonitorByInterfaceNameSubstr(
                g_settings.monitorInterfaceName.get());
        } else if (g_settings.monitor == 0) {
            POINT cursorPt;
            GetCursorPos(&cursorPt);
            monitor =
                MonitorFromPoint_Original(cursorPt, MONITOR_DEFAULTTONEAREST);
        } else if (g_settings.monitor >= 1) {
            monitor = GetMonitorById(g_settings.monitor - 1);
        }

        if (monitor) {
            return monitor;
        }
    }

    return MonitorFromPoint_Original(pt, dwFlags);
}

// Toasts slide along the X axis of MainGrid's CompositeTransform. The slide
// animations are storyboards of the transitions of the StateMachine visual
// state group: they're retargeted to the chosen direction, so layout, clipping
// and shadows stay native.
//
// A drag is handled natively too: the toast reads the drag translation from
// manipulation events and scrubs a storyboard that it builds in code and that
// moves the toast view to the right. The drag vector is rotated onto that axis
// and the storyboard is retargeted, so the toast is swiped out toward the edge
// it slid in from.
struct TranslatePropertyPaths {
    PCWSTR x;
    PCWSTR y;
};

constexpr TranslatePropertyPaths kSlidePropertyPaths{
    L"(UIElement.RenderTransform).(CompositeTransform.TranslateX)",
    L"(UIElement.RenderTransform).(CompositeTransform.TranslateY)",
};

constexpr TranslatePropertyPaths kDragPropertyPaths{
    L"CompositeTransform.TranslateX",
    L"CompositeTransform.TranslateY",
};

// Vtable slots of the XAML methods hooked for dragging, after the 6 slots of
// IInspectable. See the abi declarations in
// winrt/impl/Windows.UI.Xaml.Media.Animation.0.h and
// winrt/impl/Windows.UI.Xaml.Input.0.h.
constexpr int kStoryboardBeginSlot = 9;
constexpr int kManipulationDeltaArgsDeltaSlot = 9;
constexpr int kManipulationDeltaArgsCumulativeSlot = 10;
constexpr int kManipulationDeltaArgsVelocitiesSlot = 11;
constexpr int kManipulationCompletedArgsCumulativeSlot = 9;
constexpr int kManipulationCompletedArgsVelocitiesSlot = 10;
constexpr int kManipulationStartedArgsCumulativeSlot = 11;

// The pointer can't move past the screen edge, so a swipe toward an edge with
// little room left is shorter than the native drag track. Don't let the room
// used for scaling drop below this, in DIPs.
constexpr double kMinDragTravel = 24;

std::atomic<DWORD> g_notificationThreadId;
std::atomic<bool> g_xamlHooksInstalled;

// Only accessed on the thread of the notification window. Set when a drag
// starts, see RetargetDragStoryboard.
double g_dragScale = 1;

AnimationDirection GetAnimationDirection() {
    if (g_unloading) {
        return AnimationDirection::fromRight;
    }

    if (g_settings.animationDirection != AnimationDirection::automatic) {
        return g_settings.animationDirection;
    }

    if (g_settings.horizontalPlacement == HorizontalPlacement::center) {
        return g_settings.verticalPlacement == VerticalPlacement::bottom
                   ? AnimationDirection::fromBottom
                   : AnimationDirection::fromTop;
    }

    if (g_settings.horizontalPlacement == HorizontalPlacement::left) {
        return AnimationDirection::fromLeft;
    }

    return AnimationDirection::fromRight;
}

// The states whose transitions slide the toast in (Appearing) or out (timeout,
// close button, click).
bool IsSlideState(std::wstring_view state) {
    return state == L"Appearing" || state == L"Dismissed" ||
           state == L"Decayed" || state.starts_with(L"Activated");
}

bool IsNotificationThread() {
    DWORD threadId = g_notificationThreadId;
    return threadId && threadId == GetCurrentThreadId();
}

bool IsVerticalDirection(AnimationDirection direction) {
    return direction == AnimationDirection::fromTop ||
           direction == AnimationDirection::fromBottom;
}

// The sign of a translation away from the screen center, toward the edge the
// toast slides in from.
double DirectionSign(AnimationDirection direction) {
    return direction == AnimationDirection::fromLeft ||
                   direction == AnimationDirection::fromTop
               ? -1
               : 1;
}

// Native translate values are non-negative, so a value's magnitude is its
// native value. This makes retargeting repeatable and lets fromRight restore
// it.
void RetargetTranslateAnimations(const Media::Animation::Storyboard& storyboard,
                                 const TranslatePropertyPaths& paths,
                                 AnimationDirection direction) {
    double sign = DirectionSign(direction);
    PCWSTR targetProperty = IsVerticalDirection(direction) ? paths.y : paths.x;

    auto retargetValue = [sign](const winrt::Windows::Foundation::IReference<
                                    double>& value,
                                auto setValue) {
        if (!value) {
            return;
        }
        double newValue = std::abs(value.Value()) * sign;
        if (value.Value() != newValue) {
            setValue(winrt::Windows::Foundation::IReference<double>(newValue));
        }
    };

    for (const auto& timeline : storyboard.Children()) {
        winrt::hstring property =
            Media::Animation::Storyboard::GetTargetProperty(timeline);
        if (property != paths.x && property != paths.y) {
            continue;
        }

        auto animation = timeline.try_as<Media::Animation::DoubleAnimation>();
        auto keyFrameAnimation =
            timeline.try_as<Media::Animation::DoubleAnimationUsingKeyFrames>();
        if (!animation && !keyFrameAnimation) {
            continue;
        }

        if (property != targetProperty) {
            Media::Animation::Storyboard::SetTargetProperty(timeline,
                                                            targetProperty);
        }

        if (animation) {
            retargetValue(animation.From(),
                          [&](const auto& value) { animation.From(value); });
            retargetValue(animation.To(),
                          [&](const auto& value) { animation.To(value); });
            retargetValue(animation.By(),
                          [&](const auto& value) { animation.By(value); });
        } else {
            for (const auto& keyFrame : keyFrameAnimation.KeyFrames()) {
                double value = std::abs(keyFrame.Value()) * sign;
                if (keyFrame.Value() != value) {
                    keyFrame.Value(value);
                }
            }
        }
    }
}

void ApplySlideDirection(const FrameworkElement& mainGrid,
                         AnimationDirection direction) {
    for (const auto& group : VisualStateManager::GetVisualStateGroups(mainGrid)) {
        if (group.Name() != L"StateMachine") {
            continue;
        }

        for (const auto& transition : group.Transitions()) {
            if (!transition.From().empty()) {
                continue;
            }

            winrt::hstring state = transition.To();
            if (!IsSlideState(state)) {
                continue;
            }

            auto storyboard = transition.Storyboard();
            if (!storyboard) {
                continue;
            }

            // Fails while the storyboard runs, the next call retries.
            try {
                RetargetTranslateAnimations(storyboard, kSlidePropertyPaths,
                                            direction);
            } catch (const winrt::hresult_error& e) {
                Wh_Log(L"Retargeting %s failed: %08X", state.c_str(),
                       e.code().value);
            }
        }
    }
}

bool IsToastView(const FrameworkElement& element) {
    auto name = element.Name();
    return name.starts_with(L"FlexibleNormalToastView") ||
           name.starts_with(L"FlexiblePriorityToastView");
}

// Runs on the thread of the notification window.
bool FindToastCenterElements(FrameworkElement* toastCenterView,
                             FrameworkElement* toastCenterGrid) {
    auto window = Window::Current();
    if (!window) {
        Wh_Log(L"Failed to get current window");
        return false;
    }

    FrameworkElement child = window.Content().try_as<FrameworkElement>();
    if (!child) {
        Wh_Log(L"Failed to get window content");
        return false;
    }

    if (!(child = FindChildByClassName(
              child, L"Windows.UI.Xaml.Controls.ContentPresenter")) ||
        !(child =
              FindChildByClassName(child, L"ActionCenter.ToastCenterPage")) ||
        !(child = FindChildByName(child, L"ToastCenterMainGrid")) ||
        !(child = FindChildByName(child, L"ToastCenterView"))) {
        return false;
    }

    *toastCenterView = child;

    if (!(child = FindChildByName(child, L"ToastCenterScrollViewer")) ||
        !(child = FindChildByName(child, L"Root")) ||
        !(child =
              FindChildByClassName(child, L"Windows.UI.Xaml.Controls.Grid")) ||
        !(child = FindChildByName(child, L"ScrollContentPresenter")) ||
        !(child = FindChildByName(child, L"ToastCenterGrid"))) {
        return false;
    }

    *toastCenterGrid = child;
    return true;
}

// The visible toast that is pressed or dragged, else the first visible one.
FrameworkElement FindDraggedToastView() {
    FrameworkElement toastCenterView = nullptr;
    FrameworkElement toastCenterGrid = nullptr;
    if (!FindToastCenterElements(&toastCenterView, &toastCenterGrid)) {
        return nullptr;
    }

    FrameworkElement firstVisible = nullptr;
    FrameworkElement dragged =
        EnumChildElements(toastCenterGrid, [&](FrameworkElement toastView) {
            if (!IsToastView(toastView) ||
                toastView.Visibility() != Visibility::Visible) {
                return false;  // continue enumeration
            }

            if (!firstVisible) {
                firstVisible = toastView;
            }

            FrameworkElement mainGrid = FindChildByName(toastView, L"MainGrid");
            if (!mainGrid) {
                return false;  // continue enumeration
            }

            for (const auto& group :
                 VisualStateManager::GetVisualStateGroups(mainGrid)) {
                if (group.Name() != L"StateMachine") {
                    continue;
                }

                auto state = group.CurrentState();
                if (state && (state.Name() == L"Pressed" ||
                              state.Name() == L"Dragging")) {
                    return true;
                }
            }

            return false;  // continue enumeration
        });

    return dragged ? dragged : firstVisible;
}

// In DIPs: the distance that takes the dragged toast out of its window toward
// the slide edge, and the distance the pointer can still move toward the
// screen edge.
bool GetDragDistances(AnimationDirection direction,
                      double* exitDistance,
                      double* travel) {
    FrameworkElement toastView = FindDraggedToastView();
    if (!toastView) {
        return false;
    }

    auto xamlRoot = toastView.XamlRoot();
    if (!xamlRoot) {
        return false;
    }

    double scale = xamlRoot.RasterizationScale();
    winrt::Windows::Foundation::Size rootSize = xamlRoot.Size();
    winrt::Windows::Foundation::Rect bounds =
        toastView.TransformToVisual(nullptr).TransformBounds(
            winrt::Windows::Foundation::Rect{
                0, 0, static_cast<float>(toastView.ActualWidth()),
                static_cast<float>(toastView.ActualHeight())});

    POINT cursor;
    MONITORINFO monitorInfo{
        .cbSize = sizeof(MONITORINFO),
    };
    if (!GetCursorPos(&cursor) ||
        !GetMonitorInfo(MonitorFromPoint(cursor, MONITOR_DEFAULTTONEAREST),
                        &monitorInfo)) {
        return false;
    }

    const RECT& monitor = monitorInfo.rcMonitor;

    switch (direction) {
        case AnimationDirection::fromLeft:
            *exitDistance = bounds.X + bounds.Width;
            *travel = (cursor.x - monitor.left) / scale;
            break;

        case AnimationDirection::fromTop:
            *exitDistance = bounds.Y + bounds.Height;
            *travel = (cursor.y - monitor.top) / scale;
            break;

        case AnimationDirection::fromBottom:
            *exitDistance = rootSize.Height - bounds.Y;
            *travel = (monitor.bottom - 1 - cursor.y) / scale;
            break;

        default:
            return false;
    }

    *exitDistance = std::max(*exitDistance, 1.0);
    *travel = std::max(*travel, kMinDragTravel);
    return true;
}

// The toast builds this storyboard in code when a drag starts, and scrubs it
// along the drag: it moves the toast view by the native track length (the
// animation's To) to the right. It's pointed at the slide edge instead, moving
// the toast just out of its window. The drag is scaled so that the room the
// pointer has left covers the whole track: with enough room the toast follows
// the pointer, and reaching the screen edge always dismisses it.
void RetargetDragStoryboard(const Media::Animation::Storyboard& storyboard,
                            AnimationDirection direction) {
    for (const auto& timeline : storyboard.Children()) {
        if (Media::Animation::Storyboard::GetTargetProperty(timeline) !=
            kDragPropertyPaths.x) {
            continue;
        }

        auto animation = timeline.try_as<Media::Animation::DoubleAnimation>();
        if (!animation || !animation.To()) {
            continue;
        }

        double nativeDistance = std::abs(animation.To().Value());
        double exitDistance = nativeDistance;
        double travel = nativeDistance;
        if (!GetDragDistances(direction, &exitDistance, &travel)) {
            Wh_Log(L"Failed to measure the drag, using the native distance");
            exitDistance = nativeDistance;
            travel = nativeDistance;
        }

        double sign = DirectionSign(direction);

        if (IsVerticalDirection(direction)) {
            Media::Animation::Storyboard::SetTargetProperty(
                timeline, kDragPropertyPaths.y);
        }

        if (auto from = animation.From()) {
            animation.From(winrt::Windows::Foundation::IReference<double>(
                std::abs(from.Value()) * sign));
        }
        animation.To(
            winrt::Windows::Foundation::IReference<double>(exitDistance * sign));

        g_dragScale = nativeDistance / std::min(exitDistance, travel);

        Wh_Log(L"Drag: native %f, exit %f, travel %f, scale %f",
               nativeDistance, exitDistance, travel, g_dragScale);
    }
}

// Maps a drag along the slide direction onto the native rightward axis that
// the toast's drag handling reads, scaled by the drag scale. The
// perpendicular axis is kept for the native logic.
winrt::Windows::Foundation::Point ToNativeDragVector(
    winrt::Windows::Foundation::Point vector,
    AnimationDirection direction) {
    winrt::Windows::Foundation::Point rotated = vector;
    switch (direction) {
        case AnimationDirection::fromLeft:
            rotated = {-vector.X, -vector.Y};
            break;

        case AnimationDirection::fromTop:
            rotated = {-vector.Y, vector.X};
            break;

        case AnimationDirection::fromBottom:
            rotated = {vector.Y, -vector.X};
            break;

        default:
            break;
    }

    float scale = static_cast<float>(g_dragScale);
    return {rotated.X * scale, rotated.Y * scale};
}

bool ShouldMapDrag(AnimationDirection* direction) {
    if (g_unloading || !IsNotificationThread()) {
        return false;
    }

    *direction = GetAnimationDirection();
    return *direction != AnimationDirection::fromRight;
}

using Storyboard_Begin_t = int32_t(WINAPI*)(void* pThis);
Storyboard_Begin_t Storyboard_Begin_Original;
int32_t WINAPI Storyboard_Begin_Hook(void* pThis) {
    AnimationDirection direction;
    if (ShouldMapDrag(&direction)) {
        try {
            Media::Animation::Storyboard storyboard{nullptr};
            winrt::copy_from_abi(storyboard, pThis);
            RetargetDragStoryboard(storyboard, direction);
        } catch (const winrt::hresult_error& e) {
            Wh_Log(L"Retargeting the drag storyboard failed: %08X",
                   e.code().value);
        }
    }

    return Storyboard_Begin_Original(pThis);
}

using ManipulationDeltaGetter_t =
    int32_t(WINAPI*)(void* pThis,
                     winrt::Windows::UI::Input::ManipulationDelta* value);

using ManipulationVelocitiesGetter_t =
    int32_t(WINAPI*)(void* pThis,
                     winrt::Windows::UI::Input::ManipulationVelocities* value);

// One instance per hooked slot, each with its own original function.
template <int Slot>
struct ManipulationDeltaGetterHook {
    static inline ManipulationDeltaGetter_t original;

    static int32_t WINAPI Hook(
        void* pThis,
        winrt::Windows::UI::Input::ManipulationDelta* value) {
        int32_t result = original(pThis, value);

        AnimationDirection direction;
        if (result >= 0 && value && ShouldMapDrag(&direction)) {
            value->Translation =
                ToNativeDragVector(value->Translation, direction);
        }

        return result;
    }
};

template <int Slot>
struct ManipulationVelocitiesGetterHook {
    static inline ManipulationVelocitiesGetter_t original;

    static int32_t WINAPI Hook(
        void* pThis,
        winrt::Windows::UI::Input::ManipulationVelocities* value) {
        int32_t result = original(pThis, value);

        AnimationDirection direction;
        if (result >= 0 && value && ShouldMapDrag(&direction)) {
            value->Linear = ToNativeDragVector(value->Linear, direction);
        }

        return result;
    }
};

void** GetVtable(const winrt::Windows::Foundation::IInspectable& object) {
    return *reinterpret_cast<void***>(winrt::get_abi(object));
}

// Identical getters of different classes may share one function, which is
// hooked once: every getter here maps its value the same way.
template <typename Prototype>
void HookVtableSlot(void** vtable,
                    int slot,
                    Prototype* hook,
                    Prototype** original,
                    std::vector<void*>* hookedTargets) {
    void* target = vtable[slot];
    for (void* hookedTarget : *hookedTargets) {
        if (hookedTarget == target) {
            Wh_Log(L"Slot %d shares an already hooked function", slot);
            return;
        }
    }

    hookedTargets->push_back(target);
    WindhawkUtils::SetFunctionHook(reinterpret_cast<Prototype*>(target), hook,
                                   original);
}

// Runs on the thread of the notification window: the objects whose vtables
// are read can only be created on a XAML thread.
void EnsureXamlHooks() {
    if (g_xamlHooksInstalled || g_unloading) {
        return;
    }

    Media::Animation::Storyboard storyboard;
    Input::ManipulationDeltaRoutedEventArgs deltaArgs;
    Input::ManipulationCompletedRoutedEventArgs completedArgs;
    Input::ManipulationStartedRoutedEventArgs startedArgs;

    std::vector<void*> hookedTargets;

    HookVtableSlot(GetVtable(storyboard), kStoryboardBeginSlot,
                   Storyboard_Begin_Hook, &Storyboard_Begin_Original,
                   &hookedTargets);

    void** vtable = GetVtable(deltaArgs);
    HookVtableSlot(vtable, kManipulationDeltaArgsDeltaSlot,
                   ManipulationDeltaGetterHook<0>::Hook,
                   &ManipulationDeltaGetterHook<0>::original, &hookedTargets);
    HookVtableSlot(vtable, kManipulationDeltaArgsCumulativeSlot,
                   ManipulationDeltaGetterHook<1>::Hook,
                   &ManipulationDeltaGetterHook<1>::original, &hookedTargets);
    HookVtableSlot(vtable, kManipulationDeltaArgsVelocitiesSlot,
                   ManipulationVelocitiesGetterHook<0>::Hook,
                   &ManipulationVelocitiesGetterHook<0>::original,
                   &hookedTargets);

    vtable = GetVtable(completedArgs);
    HookVtableSlot(vtable, kManipulationCompletedArgsCumulativeSlot,
                   ManipulationDeltaGetterHook<2>::Hook,
                   &ManipulationDeltaGetterHook<2>::original, &hookedTargets);
    HookVtableSlot(vtable, kManipulationCompletedArgsVelocitiesSlot,
                   ManipulationVelocitiesGetterHook<1>::Hook,
                   &ManipulationVelocitiesGetterHook<1>::original,
                   &hookedTargets);

    HookVtableSlot(GetVtable(startedArgs),
                   kManipulationStartedArgsCumulativeSlot,
                   ManipulationDeltaGetterHook<3>::Hook,
                   &ManipulationDeltaGetterHook<3>::original, &hookedTargets);

    g_xamlHooksInstalled = true;
    Wh_ApplyHookOperations();
}

// Windows bottom-aligns the toasts in a window that is taller than them, with a
// larger gap above the toast than below it. Aligning them to the placement
// mirrors that gap, so the toast keeps the native distance from the screen edge
// and the spare room stays on the side the shadow falls.
void ApplyVerticalAlignment(const FrameworkElement& toastCenterView) {
    if (!g_originalToastCenterViewAlignment) {
        g_originalToastCenterViewAlignment = toastCenterView.VerticalAlignment();
    }

    VerticalAlignment alignment = *g_originalToastCenterViewAlignment;
    if (!g_unloading) {
        switch (g_settings.verticalPlacement) {
            case VerticalPlacement::bottom:
                break;

            case VerticalPlacement::top:
                alignment = VerticalAlignment::Top;
                break;

            case VerticalPlacement::center:
                alignment = VerticalAlignment::Center;
                break;
        }
    }

    if (toastCenterView.VerticalAlignment() != alignment) {
        toastCenterView.VerticalAlignment(alignment);
    }
}

// Versions up to 1.2.4 rotated these elements to change the animation
// direction and didn't undo it on unload.
void RemoveLegacyRotation(const FrameworkElement& element) {
    if (element && element.RenderTransform().try_as<Media::RotateTransform>()) {
        element.RenderTransform(nullptr);
    }
}

// Runs on the thread of the notification window.
void ApplyToastXamlTweaks() {
    FrameworkElement toastCenterView = nullptr;
    FrameworkElement toastCenterGrid = nullptr;
    if (!FindToastCenterElements(&toastCenterView, &toastCenterGrid)) {
        Wh_Log(L"Failed to find the toast center elements");
        return;
    }

    g_notificationThreadId = GetCurrentThreadId();
    EnsureXamlHooks();

    ApplyVerticalAlignment(toastCenterView);
    RemoveLegacyRotation(toastCenterGrid);

    AnimationDirection direction = GetAnimationDirection();

    EnumChildElements(toastCenterGrid, [&](FrameworkElement toastView) {
        if (!IsToastView(toastView)) {
            return false;  // continue enumeration
        }

        FrameworkElement mainGrid = FindChildByName(toastView, L"MainGrid");
        if (!mainGrid) {
            return false;  // continue enumeration
        }

        RemoveLegacyRotation(FindChildByName(mainGrid, L"RevealGrid"));
        RemoveLegacyRotation(FindChildByName(mainGrid, L"RevealGrid2"));

        ApplySlideDirection(mainGrid, direction);

        return false;  // continue enumeration to handle all toast slots
    });
}

void ApplyToastXamlTweaksSafe() {
    try {
        ApplyToastXamlTweaks();
    } catch (const winrt::hresult_error& e) {
        Wh_Log(L"Applying the toast tweaks failed: %08X", e.code().value);
    }
}

using RunFromWindowThreadProc_t = void (*)(void* param);

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

using SetWindowPos_t = decltype(&SetWindowPos);
SetWindowPos_t SetWindowPos_Original;
BOOL WINAPI SetWindowPos_Hook(HWND hWnd,
                              HWND hWndInsertAfter,
                              int X,
                              int Y,
                              int cx,
                              int cy,
                              UINT uFlags) {
    auto original = [&]() {
        return SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy,
                                     uFlags);
    };

    if (!IsTargetCoreWindow(hWnd)) {
        return original();
    }

    Wh_Log(L"%08X %08X", (DWORD)(ULONG_PTR)hWnd, uFlags);

    RECT rc{};
    GetWindowRect(hWnd, &rc);

    // Skip if no size.
    if (uFlags & SWP_NOSIZE) {
        Wh_Log(L"Skipping");
        uFlags |= SWP_NOMOVE | SWP_NOSIZE;
        return original();
    }

    if (uFlags & SWP_NOMOVE) {
        uFlags &= ~SWP_NOMOVE;
        X = rc.left;
        Y = rc.top;
    }

    if (uFlags & SWP_NOSIZE) {
        uFlags &= ~SWP_NOSIZE;
        cx = rc.right - rc.left;
        cy = rc.bottom - rc.top;
    }

    AdjustCoreWindowPos(&X, &Y, &cx, &cy);

    BOOL ret =
        SetWindowPos_Original(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);

    if (g_target == Target::ShellExperienceHost &&
        GetWindowThreadProcessId(hWnd, nullptr) == GetCurrentThreadId()) {
        ApplyToastXamlTweaksSafe();
    }

    return ret;
}

namespace ShellExperienceHost {

void ApplySettings() {
    for (HWND hCoreWnd : GetCoreWindows()) {
        Wh_Log(L"Adjusting core window %08X", (DWORD)(ULONG_PTR)hCoreWnd);

        RECT rc;
        if (!GetWindowRect(hCoreWnd, &rc)) {
            continue;
        }

        int x = rc.left;
        int y = rc.top;
        int cx = rc.right - rc.left;
        int cy = rc.bottom - rc.top;

        AdjustCoreWindowPos(&x, &y, &cx, &cy);

        SetWindowPos_Original(hCoreWnd, nullptr, x, y, cx, cy,
                              SWP_NOZORDER | SWP_NOACTIVATE);

        // XAML objects can only be accessed from the window's own thread.
        RunFromWindowThread(
            hCoreWnd, [](void*) { ApplyToastXamlTweaksSafe(); }, nullptr);
    }
}

}  // namespace ShellExperienceHost

bool HookTwinuiPcshellSymbols() {
    HMODULE module = LoadLibraryEx(L"twinui.pcshell.dll", nullptr,
                                   LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!module) {
        Wh_Log(L"Loading twinui.pcshell.dll failed");
        return false;
    }

    // twinui.pcshell.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(private: long __cdecl CToastCenterExperienceManager::PositionView(void))"},
            &CToastCenterExperienceManager_PositionView_Original,
            CToastCenterExperienceManager_PositionView_Hook,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
}

void LoadSettings() {
    g_settings.monitor = Wh_GetIntSetting(L"monitor");
    g_settings.monitorInterfaceName =
        WindhawkUtils::StringSetting::make(L"monitorInterfaceName");

    PCWSTR horizontalPlacement = Wh_GetStringSetting(L"horizontalPlacement");
    g_settings.horizontalPlacement = HorizontalPlacement::right;
    if (wcscmp(horizontalPlacement, L"left") == 0) {
        g_settings.horizontalPlacement = HorizontalPlacement::left;
    } else if (wcscmp(horizontalPlacement, L"center") == 0) {
        g_settings.horizontalPlacement = HorizontalPlacement::center;
    }
    Wh_FreeStringSetting(horizontalPlacement);

    g_settings.horizontalDistanceFromScreenEdge =
        Wh_GetIntSetting(L"horizontalDistanceFromScreenEdge");

    PCWSTR verticalPlacement = Wh_GetStringSetting(L"verticalPlacement");
    g_settings.verticalPlacement = VerticalPlacement::bottom;
    if (wcscmp(verticalPlacement, L"top") == 0) {
        g_settings.verticalPlacement = VerticalPlacement::top;
    } else if (wcscmp(verticalPlacement, L"center") == 0) {
        g_settings.verticalPlacement = VerticalPlacement::center;
    }
    Wh_FreeStringSetting(verticalPlacement);

    g_settings.verticalDistanceFromScreenEdge =
        Wh_GetIntSetting(L"verticalDistanceFromScreenEdge");

    PCWSTR animationDirection = Wh_GetStringSetting(L"animationDirection");
    g_settings.animationDirection = AnimationDirection::automatic;
    if (wcscmp(animationDirection, L"fromLeft") == 0) {
        g_settings.animationDirection = AnimationDirection::fromLeft;
    } else if (wcscmp(animationDirection, L"fromRight") == 0) {
        g_settings.animationDirection = AnimationDirection::fromRight;
    } else if (wcscmp(animationDirection, L"fromTop") == 0) {
        g_settings.animationDirection = AnimationDirection::fromTop;
    } else if (wcscmp(animationDirection, L"fromBottom") == 0) {
        g_settings.animationDirection = AnimationDirection::fromBottom;
    }
    Wh_FreeStringSetting(animationDirection);
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    g_target = Target::Explorer;

    WCHAR moduleFilePath[MAX_PATH];
    switch (
        GetModuleFileName(nullptr, moduleFilePath, ARRAYSIZE(moduleFilePath))) {
        case 0:
        case ARRAYSIZE(moduleFilePath):
            Wh_Log(L"GetModuleFileName failed");
            return FALSE;

        default:
            if (PCWSTR moduleFileName = wcsrchr(moduleFilePath, L'\\')) {
                moduleFileName++;
                if (_wcsicmp(moduleFileName, L"ShellExperienceHost.exe") == 0) {
                    g_target = Target::ShellExperienceHost;
                }
            } else {
                Wh_Log(L"GetModuleFileName returned an unsupported path");
                return FALSE;
            }
            break;
    }

    WindhawkUtils::SetFunctionHook(SetWindowPos, SetWindowPos_Hook,
                                   &SetWindowPos_Original);

    if (g_target == Target::Explorer) {
        if (!HookTwinuiPcshellSymbols()) {
            return FALSE;
        }

        WindhawkUtils::SetFunctionHook(MonitorFromPoint, MonitorFromPoint_Hook,
                                       &MonitorFromPoint_Original);
    }

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (g_target == Target::ShellExperienceHost) {
        ShellExperienceHost::ApplySettings();
    }
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");

    g_unloading = true;

    if (g_target == Target::ShellExperienceHost) {
        ShellExperienceHost::ApplySettings();
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();

    if (g_target == Target::ShellExperienceHost) {
        ShellExperienceHost::ApplySettings();
    }
}
