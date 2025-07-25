// ==WindhawkMod==
// @id              taskbar-volume-control
// @name            Taskbar Volume Control
// @description     Control the system volume by scrolling over the taskbar
// @version         1.2.2
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @compilerOptions -DWINVER=0x0A00 -lcomctl32 -ldwmapi -lgdi32 -lole32 -lversion
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
# Taskbar Volume Control

Control the system volume by scrolling over the taskbar.

**Note:** Some laptop touchpads might not support scrolling over the taskbar. A
workaround is to use the "pinch to zoom" gesture. For details, check out [a
relevant
issue](https://tweaker.userecho.com/topics/826-scroll-on-trackpadtouchpad-doesnt-trigger-mouse-wheel-options).

![Demonstration](https://i.imgur.com/B6mtUj9.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- volumeIndicator: win11
  $name: Volume control indicator
  $options:
  - win11: Windows 11
  - modern: Windows 10
  - classic: Windows 7
  - none: None
- scrollArea: taskbar
  $name: Scroll area
  $options:
  - taskbar: The taskbar
  - notification_area: The tray area
  - taskbarWithoutNotificationArea: The taskbar without the tray area
- middleClickToMute: true
  $name: Middle click to mute
  $description: >-
    With this option enabled, middle clicking the volume tray icon will
    mute/unmute the system volume (Windows 11 version 22H2 or newer).
- ctrlScrollVolumeChange: false
  $name: Ctrl + Scroll to change volume
  $description: >-
    When enabled, holding the Ctrl key and scrolling the mouse wheel will change
    the system volume.
- noAutomaticMuteToggle: false
  $name: No automatic mute toggle
  $description: >-
    For the Windows 11 indicator, this option causes volume scrolling to be
    disabled when the volume is muted. For the None control indicator: By
    default, the output device is muted once the volume reaches zero, and is
    unmuted on any change to a non-zero volume. Enabling this option turns off
    this functionality, such that the device mute status is not changed.
- volumeChangeStep: 2
  $name: Volume change step
  $description: >-
    Allows to configure the volume change that will occur with each notch of
    mouse wheel movement. This option has effect only for the Windows 11, None
    control indicators. For the Windows 11 indicator, must be a multiple of 2.
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool).
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <commctrl.h>
#include <dwmapi.h>
#include <endpointvolume.h>
#include <mmdeviceapi.h>
#include <objbase.h>
#include <psapi.h>
#include <windowsx.h>

#include <atomic>
#include <unordered_set>

enum class VolumeIndicator {
    None,
    Classic,
    Modern,
    Win11,
};

enum class ScrollArea {
    taskbar,
    notificationArea,
    taskbarWithoutNotificationArea,
};

struct {
    VolumeIndicator volumeIndicator;
    ScrollArea scrollArea;
    bool middleClickToMute;
    bool ctrlScrollVolumeChange;
    bool noAutomaticMuteToggle;
    int volumeChangeStep;
    bool oldTaskbarOnWin11;
} g_settings;

std::atomic<bool> g_taskbarViewDllLoaded;
std::atomic<bool> g_initialized;
bool g_inputSiteProcHooked;
std::unordered_set<HWND> g_secondaryTaskbarWindows;

enum {
    WIN_VERSION_UNSUPPORTED = 0,
    WIN_VERSION_7,
    WIN_VERSION_8,
    WIN_VERSION_81,
    WIN_VERSION_811,
    WIN_VERSION_10_T1,        // 1507
    WIN_VERSION_10_T2,        // 1511
    WIN_VERSION_10_R1,        // 1607
    WIN_VERSION_10_R2,        // 1703
    WIN_VERSION_10_R3,        // 1709
    WIN_VERSION_10_R4,        // 1803
    WIN_VERSION_10_R5,        // 1809
    WIN_VERSION_10_19H1,      // 1903, 1909
    WIN_VERSION_10_20H1,      // 2004, 20H2, 21H1, 21H2
    WIN_VERSION_SERVER_2022,  // Server 2022
    WIN_VERSION_11_21H2,
    WIN_VERSION_11_22H2,
};

#if defined(__GNUC__) && __GNUC__ > 8
#define WINAPI_LAMBDA_RETURN(return_t) ->return_t WINAPI
#elif defined(__GNUC__)
#define WINAPI_LAMBDA_RETURN(return_t) WINAPI->return_t
#else
#define WINAPI_LAMBDA_RETURN(return_t) ->return_t
#endif

#ifndef WM_POINTERWHEEL
#define WM_POINTERWHEEL 0x024E
#endif

int g_nWinVersion;
int g_nExplorerVersion;
HWND g_hTaskbarWnd;
DWORD g_dwTaskbarThreadId;

#pragma region functions

UINT GetDpiForWindowWithFallback(HWND hWnd) {
    using GetDpiForWindow_t = UINT(WINAPI*)(HWND hwnd);
    static GetDpiForWindow_t pGetDpiForWindow = []() {
        HMODULE hUser32 = GetModuleHandle(L"user32.dll");
        if (hUser32) {
            return (GetDpiForWindow_t)GetProcAddress(hUser32,
                                                     "GetDpiForWindow");
        }

        return (GetDpiForWindow_t) nullptr;
    }();

    int iDpi = 96;
    if (pGetDpiForWindow) {
        iDpi = pGetDpiForWindow(hWnd);
    } else {
        HDC hdc = GetDC(NULL);
        if (hdc) {
            iDpi = GetDeviceCaps(hdc, LOGPIXELSX);
            ReleaseDC(NULL, hdc);
        }
    }

    return iDpi;
}

bool IsTaskbarWindow(HWND hWnd) {
    WCHAR szClassName[32];
    if (!GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName))) {
        return false;
    }

    return _wcsicmp(szClassName, L"Shell_TrayWnd") == 0 ||
           _wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0;
}

bool GetNotificationAreaRect(HWND hMMTaskbarWnd, RECT* rcResult) {
    if (hMMTaskbarWnd == g_hTaskbarWnd) {
        HWND hTrayNotifyWnd =
            FindWindowEx(hMMTaskbarWnd, NULL, L"TrayNotifyWnd", NULL);
        if (hTrayNotifyWnd && GetWindowRect(hTrayNotifyWnd, rcResult) &&
            !IsRectEmpty(rcResult)) {
            return true;
        }

        // When attaching an external monitor, it was observed that the rect can
        // be empty. Use fallback in this case.
    } else if (g_nExplorerVersion >= WIN_VERSION_11_21H2) {
        RECT rcTaskbar;
        if (GetWindowRect(hMMTaskbarWnd, &rcTaskbar)) {
            HWND hBridgeWnd = FindWindowEx(
                hMMTaskbarWnd, NULL,
                L"Windows.UI.Composition.DesktopWindowContentBridge", NULL);
            while (hBridgeWnd) {
                RECT rcBridge;
                if (!GetWindowRect(hBridgeWnd, &rcBridge)) {
                    break;
                }

                if (!EqualRect(&rcBridge, &rcTaskbar)) {
                    if (IsRectEmpty(&rcBridge)) {
                        break;
                    }

                    CopyRect(rcResult, &rcBridge);
                    return true;
                }

                hBridgeWnd = FindWindowEx(
                    hMMTaskbarWnd, hBridgeWnd,
                    L"Windows.UI.Composition.DesktopWindowContentBridge", NULL);
            }
        }

        // On newer Win11 versions, the clock on secondary taskbars is difficult
        // to detect without either UI Automation or UWP UI APIs. Use fallback.
    } else if (g_nExplorerVersion >= WIN_VERSION_10_R1) {
        HWND hClockButtonWnd =
            FindWindowEx(hMMTaskbarWnd, NULL, L"ClockButton", NULL);
        if (hClockButtonWnd && GetWindowRect(hClockButtonWnd, rcResult) &&
            !IsRectEmpty(rcResult)) {
            return true;
        }
    } else {
        // In older Windows versions, there's no clock on the secondary taskbar.
        SetRectEmpty(rcResult);
        return true;
    }

    RECT rcTaskbar;
    if (!GetWindowRect(hMMTaskbarWnd, &rcTaskbar)) {
        return false;
    }

    // Just consider the last pixels as a fallback, not accurate, but better
    // than nothing.
    int lastPixels = MulDiv(50, GetDpiForWindowWithFallback(hMMTaskbarWnd), 96);
    CopyRect(rcResult, &rcTaskbar);
    if (rcResult->right - rcResult->left > lastPixels) {
        if (GetWindowLong(hMMTaskbarWnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) {
            rcResult->right = rcResult->left + lastPixels;
        } else {
            rcResult->left = rcResult->right - lastPixels;
        }
    }

    return true;
}

bool IsPointInsideScrollArea(HWND hMMTaskbarWnd, POINT pt) {
    switch (g_settings.scrollArea) {
        case ScrollArea::taskbar: {
            RECT rc;
            return GetWindowRect(hMMTaskbarWnd, &rc) && PtInRect(&rc, pt);
        }

        case ScrollArea::notificationArea: {
            RECT rc;
            return GetNotificationAreaRect(hMMTaskbarWnd, &rc) &&
                   PtInRect(&rc, pt);
        }

        case ScrollArea::taskbarWithoutNotificationArea: {
            RECT rc;
            return GetWindowRect(hMMTaskbarWnd, &rc) && PtInRect(&rc, pt) &&
                   (!GetNotificationAreaRect(hMMTaskbarWnd, &rc) ||
                    !PtInRect(&rc, pt));
        }
    }

    return false;
}

VS_FIXEDFILEINFO* GetModuleVersionInfo(HMODULE hModule, UINT* puPtrLen) {
    HRSRC hResource;
    HGLOBAL hGlobal;
    void* pData;
    void* pFixedFileInfo;
    UINT uPtrLen;

    pFixedFileInfo = NULL;
    uPtrLen = 0;

    hResource =
        FindResource(hModule, MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION);
    if (hResource != NULL) {
        hGlobal = LoadResource(hModule, hResource);
        if (hGlobal != NULL) {
            pData = LockResource(hGlobal);
            if (pData != NULL) {
                if (!VerQueryValue(pData, L"\\", &pFixedFileInfo, &uPtrLen) ||
                    uPtrLen == 0) {
                    pFixedFileInfo = NULL;
                    uPtrLen = 0;
                }
            }
        }
    }

    if (puPtrLen)
        *puPtrLen = uPtrLen;

    return (VS_FIXEDFILEINFO*)pFixedFileInfo;
}

BOOL WindowsVersionInit() {
    g_nWinVersion = WIN_VERSION_UNSUPPORTED;

    VS_FIXEDFILEINFO* pFixedFileInfo = GetModuleVersionInfo(NULL, NULL);
    if (!pFixedFileInfo)
        return FALSE;

    WORD nMajor = HIWORD(pFixedFileInfo->dwFileVersionMS);
    WORD nMinor = LOWORD(pFixedFileInfo->dwFileVersionMS);
    WORD nBuild = HIWORD(pFixedFileInfo->dwFileVersionLS);
    WORD nQFE = LOWORD(pFixedFileInfo->dwFileVersionLS);

    switch (nMajor) {
        case 6:
            switch (nMinor) {
                case 1:
                    g_nWinVersion = WIN_VERSION_7;
                    break;

                case 2:
                    g_nWinVersion = WIN_VERSION_8;
                    break;

                case 3:
                    if (nQFE < 17000)
                        g_nWinVersion = WIN_VERSION_81;
                    else
                        g_nWinVersion = WIN_VERSION_811;
                    break;

                case 4:
                    g_nWinVersion = WIN_VERSION_10_T1;
                    break;
            }
            break;

        case 10:
            if (nBuild <= 10240)
                g_nWinVersion = WIN_VERSION_10_T1;
            else if (nBuild <= 10586)
                g_nWinVersion = WIN_VERSION_10_T2;
            else if (nBuild <= 14393)
                g_nWinVersion = WIN_VERSION_10_R1;
            else if (nBuild <= 15063)
                g_nWinVersion = WIN_VERSION_10_R2;
            else if (nBuild <= 16299)
                g_nWinVersion = WIN_VERSION_10_R3;
            else if (nBuild <= 17134)
                g_nWinVersion = WIN_VERSION_10_R4;
            else if (nBuild <= 17763)
                g_nWinVersion = WIN_VERSION_10_R5;
            else if (nBuild <= 18362)
                g_nWinVersion = WIN_VERSION_10_19H1;
            else if (nBuild <= 19041)
                g_nWinVersion = WIN_VERSION_10_20H1;
            else if (nBuild <= 20348)
                g_nWinVersion = WIN_VERSION_SERVER_2022;
            else if (nBuild <= 22000)
                g_nWinVersion = WIN_VERSION_11_21H2;
            else
                g_nWinVersion = WIN_VERSION_11_22H2;
            break;
    }

    if (g_nWinVersion == WIN_VERSION_UNSUPPORTED)
        return FALSE;

    return TRUE;
}

#pragma endregion  // functions

#pragma region volume_functions

const static GUID XIID_IMMDeviceEnumerator = {
    0xA95664D2,
    0x9614,
    0x4F35,
    {0xA7, 0x46, 0xDE, 0x8D, 0xB6, 0x36, 0x17, 0xE6}};
const static GUID XIID_MMDeviceEnumerator = {
    0xBCDE0395,
    0xE52F,
    0x467C,
    {0x8E, 0x3D, 0xC4, 0x57, 0x92, 0x91, 0x69, 0x2E}};
const static GUID XIID_IAudioEndpointVolume = {
    0x5CDF2C82,
    0x841E,
    0x4546,
    {0x97, 0x22, 0x0C, 0xF7, 0x40, 0x78, 0x22, 0x9A}};

static IMMDeviceEnumerator* g_pDeviceEnumerator;

BOOL IsDefaultAudioEndpointAvailable() {
    IMMDevice* defaultDevice = NULL;
    HRESULT hr;
    BOOL bSuccess = FALSE;

    if (g_pDeviceEnumerator) {
        hr = g_pDeviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole,
                                                          &defaultDevice);
        if (SUCCEEDED(hr)) {
            bSuccess = TRUE;

            defaultDevice->Release();
        }
    }

    return bSuccess;
}

BOOL IsVolMuted(BOOL* pbMuted) {
    IMMDevice* defaultDevice = NULL;
    IAudioEndpointVolume* endpointVolume = NULL;
    HRESULT hr;
    BOOL bSuccess = FALSE;

    if (g_pDeviceEnumerator) {
        hr = g_pDeviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole,
                                                          &defaultDevice);
        if (SUCCEEDED(hr)) {
            hr = defaultDevice->Activate(XIID_IAudioEndpointVolume,
                                         CLSCTX_INPROC_SERVER, NULL,
                                         (LPVOID*)&endpointVolume);
            if (SUCCEEDED(hr)) {
                if (SUCCEEDED(endpointVolume->GetMute(pbMuted)))
                    bSuccess = TRUE;

                endpointVolume->Release();
            }

            defaultDevice->Release();
        }
    }

    return bSuccess;
}

BOOL ToggleVolMuted() {
    IMMDevice* defaultDevice = NULL;
    IAudioEndpointVolume* endpointVolume = NULL;
    HRESULT hr;
    BOOL bMuted;
    BOOL bSuccess = FALSE;

    if (g_pDeviceEnumerator) {
        hr = g_pDeviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole,
                                                          &defaultDevice);
        if (SUCCEEDED(hr)) {
            hr = defaultDevice->Activate(XIID_IAudioEndpointVolume,
                                         CLSCTX_INPROC_SERVER, NULL,
                                         (LPVOID*)&endpointVolume);
            if (SUCCEEDED(hr)) {
                if (SUCCEEDED(endpointVolume->GetMute(&bMuted))) {
                    if (SUCCEEDED(endpointVolume->SetMute(!bMuted, NULL)))
                        bSuccess = TRUE;
                }

                endpointVolume->Release();
            }

            defaultDevice->Release();
        }
    }

    return bSuccess;
}

BOOL AddMasterVolumeLevelScalar(float fMasterVolumeAdd) {
    IMMDevice* defaultDevice = NULL;
    IAudioEndpointVolume* endpointVolume = NULL;
    HRESULT hr;
    float fMasterVolume;
    BOOL bSuccess = FALSE;

    if (g_pDeviceEnumerator) {
        hr = g_pDeviceEnumerator->GetDefaultAudioEndpoint(eRender, eConsole,
                                                          &defaultDevice);
        if (SUCCEEDED(hr)) {
            hr = defaultDevice->Activate(XIID_IAudioEndpointVolume,
                                         CLSCTX_INPROC_SERVER, NULL,
                                         (LPVOID*)&endpointVolume);
            if (SUCCEEDED(hr)) {
                if (SUCCEEDED(endpointVolume->GetMasterVolumeLevelScalar(
                        &fMasterVolume))) {
                    fMasterVolume += fMasterVolumeAdd;

                    if (fMasterVolume < 0.0)
                        fMasterVolume = 0.0;
                    else if (fMasterVolume > 1.0)
                        fMasterVolume = 1.0;

                    if (SUCCEEDED(endpointVolume->SetMasterVolumeLevelScalar(
                            fMasterVolume, NULL))) {
                        bSuccess = TRUE;

                        if (!g_settings.noAutomaticMuteToggle) {
                            // Windows displays the volume rounded to the
                            // nearest percentage. The range [0, 0.005) is
                            // displayed as 0%, [0.005, 0.015) as 1%, etc. It
                            // also mutes the volume when it becomes zero, we do
                            // the same.

                            if (fMasterVolume < 0.005)
                                endpointVolume->SetMute(TRUE, NULL);
                            else
                                endpointVolume->SetMute(FALSE, NULL);
                        }
                    }
                }

                endpointVolume->Release();
            }

            defaultDevice->Release();
        }
    }

    return bSuccess;
}

void SndVolInit() {
    HRESULT hr = CoCreateInstance(
        XIID_MMDeviceEnumerator, NULL, CLSCTX_INPROC_SERVER,
        XIID_IMMDeviceEnumerator, (LPVOID*)&g_pDeviceEnumerator);
    if (FAILED(hr))
        g_pDeviceEnumerator = NULL;
}

void SndVolUninit() {
    if (g_pDeviceEnumerator) {
        g_pDeviceEnumerator->Release();
        g_pDeviceEnumerator = NULL;
    }
}

#pragma endregion  // volume_functions

#pragma region sndvol_win11

UINT g_uShellHookMsg = RegisterWindowMessage(L"SHELLHOOK");
DWORD g_lastScrollTime;
short g_lastScrollDeltaRemainder;

bool PostAppCommand(SHORT appCommand, int count) {
    if (!g_hTaskbarWnd) {
        return false;
    }

    HWND hReBarWindow32 =
        FindWindowEx(g_hTaskbarWnd, nullptr, L"ReBarWindow32", nullptr);
    if (!hReBarWindow32) {
        return false;
    }

    HWND hMSTaskSwWClass =
        FindWindowEx(hReBarWindow32, nullptr, L"MSTaskSwWClass", nullptr);
    if (!hMSTaskSwWClass) {
        return false;
    }

    for (int i = 0; i < count; i++) {
        PostMessage(hMSTaskSwWClass, g_uShellHookMsg, HSHELL_APPCOMMAND,
                    MAKELPARAM(0, appCommand));
    }

    return true;
}

bool Win11IndicatorAdjustVolumeLevelWithMouseWheel(short delta) {
    BOOL muted;
    if (g_settings.noAutomaticMuteToggle && IsVolMuted(&muted) && muted) {
        return true;
    }

    if (GetTickCount() - g_lastScrollTime < 1000 * 5) {
        delta += g_lastScrollDeltaRemainder;
    }

    int clicks = delta / WHEEL_DELTA;
    Wh_Log(L"%d clicks (delta=%d)", clicks, delta);

    SHORT appCommand = APPCOMMAND_VOLUME_UP;
    if (clicks < 0) {
        clicks = -clicks;
        appCommand = APPCOMMAND_VOLUME_DOWN;
    }

    if (g_settings.volumeChangeStep) {
        clicks *= g_settings.volumeChangeStep / 2;
    }

    if (!PostAppCommand(appCommand, clicks)) {
        return false;
    }

    g_lastScrollTime = GetTickCount();
    g_lastScrollDeltaRemainder = delta % WHEEL_DELTA;

    return true;
}

#pragma endregion  // sndvol_win11

#pragma region sndvol

BOOL OpenScrollSndVol(WPARAM wParam, LPARAM lMousePosParam);
BOOL ScrollSndVol(WPARAM wParam, LPARAM lMousePosParam);
void SetSndVolTimer();
void KillSndVolTimer();
void CleanupSndVol();

static BOOL AdjustVolumeLevelWithMouseWheel(int nWheelDelta, int nStep);
static BOOL OpenScrollSndVolInternal(WPARAM wParam,
                                     LPARAM lMousePosParam,
                                     HWND hVolumeAppWnd,
                                     BOOL* pbOpened);
static BOOL ValidateSndVolProcess();
static BOOL ValidateSndVolWnd();
static void CALLBACK CloseSndVolTimerProc(HWND hWnd,
                                          UINT uMsg,
                                          UINT_PTR idEvent,
                                          DWORD dwTime);
static HWND GetSndVolDlg(HWND hVolumeAppWnd);
static BOOL CALLBACK EnumThreadFindSndVolWnd(HWND hWnd, LPARAM lParam);
static BOOL IsSndVolWndInitialized(HWND hWnd);
static BOOL MoveSndVolCenterMouse(HWND hWnd);

// Modern indicator functions
static BOOL CanUseModernIndicator();
static BOOL ShowSndVolModernIndicator();
static BOOL HideSndVolModernIndicator();
static void EndSndVolModernIndicatorSession();
static HWND GetOpenSndVolModernIndicatorWnd();
static HWND GetSndVolTrayControlWnd();
static BOOL CALLBACK EnumThreadFindSndVolTrayControlWnd(HWND hWnd,
                                                        LPARAM lParam);

static HANDLE hSndVolProcess;
static HWND hSndVolWnd;
static UINT_PTR nCloseSndVolTimer;
static int nCloseSndVolTimerCount;
static HWND hSndVolModernPreviousForegroundWnd;
static BOOL bSndVolModernLaunched;
static BOOL bSndVolModernAppeared;

BOOL OpenScrollSndVol(WPARAM wParam, LPARAM lMousePosParam) {
    HANDLE hMutex;
    HWND hVolumeAppWnd;
    DWORD dwProcessId;
    WCHAR szCommandLine[sizeof("SndVol.exe -f 4294967295")];
    STARTUPINFO si;
    PROCESS_INFORMATION pi;

    if (g_settings.volumeIndicator == VolumeIndicator::Win11 &&
        g_nWinVersion >= WIN_VERSION_11_22H2) {
        return Win11IndicatorAdjustVolumeLevelWithMouseWheel(
            GET_WHEEL_DELTA_WPARAM(wParam));
    }

    if (g_settings.volumeIndicator == VolumeIndicator::None) {
        return AdjustVolumeLevelWithMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam),
                                               0);
    }

    if (CanUseModernIndicator()) {
        if (!AdjustVolumeLevelWithMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam), 2))
            return FALSE;

        ShowSndVolModernIndicator();
        SetSndVolTimer();
        return TRUE;
    }

    if (!IsDefaultAudioEndpointAvailable())
        return FALSE;

    if (ValidateSndVolProcess()) {
        if (WaitForInputIdle(hSndVolProcess, 0) == 0)  // If not initializing
        {
            if (ValidateSndVolWnd()) {
                ScrollSndVol(wParam, lMousePosParam);

                return FALSE;  // False because we didn't open it, it was open
            } else {
                hVolumeAppWnd = FindWindow(L"Windows Volume App Window",
                                           L"Windows Volume App Window");
                if (hVolumeAppWnd) {
                    GetWindowThreadProcessId(hVolumeAppWnd, &dwProcessId);

                    if (GetProcessId(hSndVolProcess) == dwProcessId) {
                        BOOL bOpened;
                        if (OpenScrollSndVolInternal(wParam, lMousePosParam,
                                                     hVolumeAppWnd, &bOpened)) {
                            if (bOpened)
                                SetSndVolTimer();

                            return bOpened;
                        }
                    }
                }
            }
        }

        return FALSE;
    }

    hMutex = OpenMutex(SYNCHRONIZE, FALSE, L"Windows Volume App Window");
    if (hMutex) {
        CloseHandle(hMutex);

        hVolumeAppWnd = FindWindow(L"Windows Volume App Window",
                                   L"Windows Volume App Window");
        if (hVolumeAppWnd) {
            GetWindowThreadProcessId(hVolumeAppWnd, &dwProcessId);

            hSndVolProcess =
                OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE,
                            FALSE, dwProcessId);
            if (hSndVolProcess) {
                if (WaitForInputIdle(hSndVolProcess, 0) ==
                    0)  // if not initializing
                {
                    if (ValidateSndVolWnd()) {
                        ScrollSndVol(wParam, lMousePosParam);

                        return FALSE;  // False because we didn't open it, it
                                       // was open
                    } else {
                        BOOL bOpened;
                        if (OpenScrollSndVolInternal(wParam, lMousePosParam,
                                                     hVolumeAppWnd, &bOpened)) {
                            if (bOpened)
                                SetSndVolTimer();

                            return bOpened;
                        }
                    }
                }
            }
        }

        return FALSE;
    }

    wsprintf(szCommandLine, L"SndVol.exe -f %u", (DWORD)lMousePosParam);

    ZeroMemory(&si, sizeof(STARTUPINFO));
    si.cb = sizeof(STARTUPINFO);

    if (!CreateProcess(NULL, szCommandLine, NULL, NULL, FALSE,
                       ABOVE_NORMAL_PRIORITY_CLASS | CREATE_SUSPENDED, NULL,
                       NULL, &si, &pi))
        return FALSE;

    if (g_nExplorerVersion <= WIN_VERSION_7)
        SendMessage(g_hTaskbarWnd, WM_USER + 12, 0, 0);  // Close start menu

    AllowSetForegroundWindow(pi.dwProcessId);
    ResumeThread(pi.hThread);

    CloseHandle(pi.hThread);
    hSndVolProcess = pi.hProcess;

    SetSndVolTimer();

    return TRUE;
}

BOOL ScrollSndVol(WPARAM wParam, LPARAM lMousePosParam) {
    GUITHREADINFO guithreadinfo;

    guithreadinfo.cbSize = sizeof(GUITHREADINFO);

    if (!GetGUIThreadInfo(GetWindowThreadProcessId(hSndVolWnd, NULL),
                          &guithreadinfo))
        return FALSE;

    PostMessage(guithreadinfo.hwndFocus, WM_MOUSEWHEEL, wParam, lMousePosParam);
    return TRUE;
}

void SetSndVolTimer() {
    nCloseSndVolTimer =
        SetTimer(NULL, nCloseSndVolTimer, 100, CloseSndVolTimerProc);
    nCloseSndVolTimerCount = 0;
}

void KillSndVolTimer() {
    if (nCloseSndVolTimer != 0) {
        KillTimer(NULL, nCloseSndVolTimer);
        nCloseSndVolTimer = 0;
    }
}

void CleanupSndVol() {
    KillSndVolTimer();

    if (hSndVolProcess) {
        CloseHandle(hSndVolProcess);
        hSndVolProcess = NULL;
        hSndVolWnd = NULL;
    }
}

static BOOL AdjustVolumeLevelWithMouseWheel(int nWheelDelta, int nStep) {
    if (!nStep) {
        nStep = g_settings.volumeChangeStep;
        if (!nStep)
            nStep = 2;
    }

    return AddMasterVolumeLevelScalar((float)nWheelDelta * nStep *
                                      ((float)0.01 / 120));
}

static BOOL OpenScrollSndVolInternal(WPARAM wParam,
                                     LPARAM lMousePosParam,
                                     HWND hVolumeAppWnd,
                                     BOOL* pbOpened) {
    HWND hSndVolDlg = GetSndVolDlg(hVolumeAppWnd);
    if (hSndVolDlg) {
        if (GetWindowTextLength(hSndVolDlg) == 0)  // Volume control
        {
            if (IsSndVolWndInitialized(hSndVolDlg) &&
                MoveSndVolCenterMouse(hSndVolDlg)) {
                if (g_nExplorerVersion <= WIN_VERSION_7)
                    SendMessage(g_hTaskbarWnd, WM_USER + 12, 0,
                                0);  // Close start menu

                SetForegroundWindow(hVolumeAppWnd);
                PostMessage(hVolumeAppWnd, WM_USER + 35, 0, 0);

                *pbOpened = TRUE;
                return TRUE;
            }
        } else if (IsWindowVisible(
                       hSndVolDlg))  // Another dialog, e.g. volume mixer
        {
            if (g_nExplorerVersion <= WIN_VERSION_7)
                SendMessage(g_hTaskbarWnd, WM_USER + 12, 0,
                            0);  // Close start menu

            SetForegroundWindow(hVolumeAppWnd);
            PostMessage(hVolumeAppWnd, WM_USER + 35, 0, 0);

            *pbOpened = FALSE;
            return TRUE;
        }
    }

    return FALSE;
}

static BOOL ValidateSndVolProcess() {
    if (!hSndVolProcess)
        return FALSE;

    if (WaitForSingleObject(hSndVolProcess, 0) != WAIT_TIMEOUT) {
        CloseHandle(hSndVolProcess);
        hSndVolProcess = NULL;
        hSndVolWnd = NULL;

        return FALSE;
    }

    return TRUE;
}

static BOOL ValidateSndVolWnd() {
    HWND hForegroundWnd;
    DWORD dwProcessId;
    WCHAR szClass[sizeof("#32770") + 1];

    hForegroundWnd = GetForegroundWindow();

    if (hSndVolWnd == hForegroundWnd)
        return TRUE;

    GetWindowThreadProcessId(hForegroundWnd, &dwProcessId);

    if (GetProcessId(hSndVolProcess) == dwProcessId) {
        GetClassName(hForegroundWnd, szClass, sizeof("#32770") + 1);

        if (lstrcmp(szClass, L"#32770") == 0) {
            hSndVolWnd = hForegroundWnd;

            return TRUE;
        }
    }

    hSndVolWnd = NULL;

    return FALSE;
}

static void CALLBACK CloseSndVolTimerProc(HWND hWnd,
                                          UINT uMsg,
                                          UINT_PTR idEvent,
                                          DWORD dwTime) {
    if (CanUseModernIndicator()) {
        HWND hSndVolModernIndicatorWnd = GetOpenSndVolModernIndicatorWnd();
        if (!bSndVolModernAppeared) {
            if (hSndVolModernIndicatorWnd) {
                bSndVolModernAppeared = TRUE;
                nCloseSndVolTimerCount = 1;

                // Windows 11 shows an ugly focus border. Make it go away by
                // making the window think it becomes inactive.
                if (g_nWinVersion >= WIN_VERSION_11_21H2) {
                    PostMessage(hSndVolModernIndicatorWnd, WM_ACTIVATE,
                                MAKEWPARAM(WA_INACTIVE, FALSE), 0);
                }

                return;
            } else {
                nCloseSndVolTimerCount++;
                if (nCloseSndVolTimerCount < 10)
                    return;
            }
        } else {
            if (hSndVolModernIndicatorWnd) {
                POINT pt;
                GetCursorPos(&pt);
                HWND hPointWnd = GetAncestor(WindowFromPoint(pt), GA_ROOT);

                if (!hPointWnd)
                    nCloseSndVolTimerCount++;
                else if (hPointWnd == hSndVolModernIndicatorWnd)
                    nCloseSndVolTimerCount = 0;
                else if (IsTaskbarWindow(hPointWnd) &&
                         IsPointInsideScrollArea(hPointWnd, pt))
                    nCloseSndVolTimerCount = 0;
                else
                    nCloseSndVolTimerCount++;

                if (nCloseSndVolTimerCount < 10)
                    return;

                HideSndVolModernIndicator();
            }
        }

        EndSndVolModernIndicatorSession();
    } else {
        if (ValidateSndVolProcess()) {
            if (WaitForInputIdle(hSndVolProcess, 0) != 0)
                return;

            if (ValidateSndVolWnd()) {
                POINT pt;
                GetCursorPos(&pt);
                HWND hPointWnd = GetAncestor(WindowFromPoint(pt), GA_ROOT);

                if (!hPointWnd)
                    nCloseSndVolTimerCount++;
                else if (hPointWnd == hSndVolWnd)
                    nCloseSndVolTimerCount = 0;
                else if (IsTaskbarWindow(hPointWnd) &&
                         IsPointInsideScrollArea(hPointWnd, pt))
                    nCloseSndVolTimerCount = 0;
                else
                    nCloseSndVolTimerCount++;

                if (nCloseSndVolTimerCount < 10)
                    return;

                if (hPointWnd != hSndVolWnd)
                    PostMessage(hSndVolWnd, WM_ACTIVATE,
                                MAKEWPARAM(WA_INACTIVE, FALSE), (LPARAM)NULL);
            }
        }
    }

    KillTimer(NULL, nCloseSndVolTimer);
    nCloseSndVolTimer = 0;
}

static HWND GetSndVolDlg(HWND hVolumeAppWnd) {
    HWND hWnd = NULL;
    EnumThreadWindows(GetWindowThreadProcessId(hVolumeAppWnd, NULL),
                      EnumThreadFindSndVolWnd, (LPARAM)&hWnd);
    return hWnd;
}

static BOOL CALLBACK EnumThreadFindSndVolWnd(HWND hWnd, LPARAM lParam) {
    WCHAR szClass[16];

    GetClassName(hWnd, szClass, _countof(szClass));
    if (lstrcmp(szClass, L"#32770") == 0) {
        *(HWND*)lParam = hWnd;
        return FALSE;
    }

    return TRUE;
}

static BOOL IsSndVolWndInitialized(HWND hWnd) {
    HWND hChildDlg;

    hChildDlg = FindWindowEx(hWnd, NULL, L"#32770", NULL);
    if (!hChildDlg)
        return FALSE;

    if (!(GetWindowLong(hChildDlg, GWL_STYLE) & WS_VISIBLE))
        return FALSE;

    return TRUE;
}

static BOOL MoveSndVolCenterMouse(HWND hWnd) {
    NOTIFYICONIDENTIFIER notifyiconidentifier;
    BOOL bCompositionEnabled;
    POINT pt;
    SIZE size;
    RECT rc, rcExclude, rcInflate;
    int nInflate;

    ZeroMemory(&notifyiconidentifier, sizeof(NOTIFYICONIDENTIFIER));
    notifyiconidentifier.cbSize = sizeof(NOTIFYICONIDENTIFIER);
    memcpy(&notifyiconidentifier.guidItem,
           "\x73\xAE\x20\x78\xE3\x23\x29\x42\x82\xC1\xE4\x1C\xB6\x7D\x5B\x9C",
           sizeof(GUID));

    if (Shell_NotifyIconGetRect(&notifyiconidentifier, &rcExclude) != S_OK)
        SetRectEmpty(&rcExclude);

    GetCursorPos(&pt);
    GetWindowRect(hWnd, &rc);

    nInflate = 0;

    if (DwmIsCompositionEnabled(&bCompositionEnabled) == S_OK &&
        bCompositionEnabled) {
        memcpy(
            &notifyiconidentifier.guidItem,
            "\x43\x65\x4B\x96\xAD\xBB\xEE\x44\x84\x8A\x3A\x95\xD8\x59\x51\xEA",
            sizeof(GUID));

        if (Shell_NotifyIconGetRect(&notifyiconidentifier, &rcInflate) ==
            S_OK) {
            nInflate = rcInflate.bottom - rcInflate.top;
            InflateRect(&rc, nInflate, nInflate);
        }
    }

    size.cx = rc.right - rc.left;
    size.cy = rc.bottom - rc.top;

    if (!CalculatePopupWindowPosition(
            &pt, &size,
            TPM_CENTERALIGN | TPM_VCENTERALIGN | TPM_VERTICAL | TPM_WORKAREA,
            &rcExclude, &rc))
        return FALSE;

    SetWindowPos(
        hWnd, NULL, rc.left + nInflate, rc.top + nInflate, 0, 0,
        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);

    return TRUE;
}

// Modern indicator functions

static BOOL CanUseModernIndicator() {
    if (g_nWinVersion < WIN_VERSION_10_T1 ||
        g_settings.volumeIndicator == VolumeIndicator::Classic)
        return FALSE;

    DWORD dwEnabled = 1;
    DWORD dwValueSize = sizeof(dwEnabled);
    DWORD dwError = RegGetValue(
        HKEY_LOCAL_MACHINE,
        L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\MTCUVC",
        L"EnableMTCUVC", RRF_RT_REG_DWORD, NULL, &dwEnabled, &dwValueSize);

    // We don't check dwError just like Microsoft doesn't at SndVolSSO.dll.
    if (!dwError)
        Wh_Log(L"%u", dwError);

    return dwEnabled != 0;
}

static BOOL ShowSndVolModernIndicator() {
    if (bSndVolModernLaunched)
        return TRUE;  // already launched

    HWND hSndVolModernIndicatorWnd = GetOpenSndVolModernIndicatorWnd();
    if (hSndVolModernIndicatorWnd)
        return TRUE;  // already shown

    HWND hForegroundWnd = GetForegroundWindow();
    if (hForegroundWnd && hForegroundWnd != g_hTaskbarWnd)
        hSndVolModernPreviousForegroundWnd = hForegroundWnd;

    HWND hSndVolTrayControlWnd = GetSndVolTrayControlWnd();
    if (!hSndVolTrayControlWnd)
        return FALSE;

    if (!PostMessage(hSndVolTrayControlWnd, 0x460, 0,
                     MAKELPARAM(NIN_SELECT, 100)))
        return FALSE;

    bSndVolModernLaunched = TRUE;
    return TRUE;
}

static BOOL HideSndVolModernIndicator() {
    HWND hSndVolModernIndicatorWnd = GetOpenSndVolModernIndicatorWnd();
    if (hSndVolModernIndicatorWnd) {
        if (!hSndVolModernPreviousForegroundWnd ||
            !SetForegroundWindow(hSndVolModernPreviousForegroundWnd))
            SetForegroundWindow(g_hTaskbarWnd);
    }

    return TRUE;
}

static void EndSndVolModernIndicatorSession() {
    hSndVolModernPreviousForegroundWnd = NULL;
    bSndVolModernLaunched = FALSE;
    bSndVolModernAppeared = FALSE;
}

static HWND GetOpenSndVolModernIndicatorWnd() {
    HWND hForegroundWnd = GetForegroundWindow();
    if (!hForegroundWnd)
        return NULL;

    // Check class name
    WCHAR szBuffer[32];
    if (!GetClassName(hForegroundWnd, szBuffer, 32) ||
        wcscmp(szBuffer, L"Windows.UI.Core.CoreWindow") != 0)
        return NULL;

    // Check that the MtcUvc prop exists
    WCHAR szVerifyPropName[sizeof(
        "ApplicationView_CustomWindowTitle#1234567890#MtcUvc")];
    wsprintf(szVerifyPropName, L"ApplicationView_CustomWindowTitle#%u#MtcUvc",
             (DWORD)(DWORD_PTR)hForegroundWnd);

    SetLastError(0);
    GetProp(hForegroundWnd, szVerifyPropName);
    if (GetLastError() != 0)
        return NULL;

    return hForegroundWnd;
}

static HWND GetSndVolTrayControlWnd() {
    // The window we're looking for has a class name similar to
    // "ATL:00007FFAECBBD280". It shares a thread with the bluetooth window,
    // which is easier to find by class, so we use that.

    HWND hBluetoothNotificationWnd =
        FindWindow(L"BluetoothNotificationAreaIconWindowClass", NULL);
    if (!hBluetoothNotificationWnd)
        return NULL;

    HWND hWnd = NULL;
    EnumThreadWindows(GetWindowThreadProcessId(hBluetoothNotificationWnd, NULL),
                      EnumThreadFindSndVolTrayControlWnd, (LPARAM)&hWnd);
    return hWnd;
}

static BOOL CALLBACK EnumThreadFindSndVolTrayControlWnd(HWND hWnd,
                                                        LPARAM lParam) {
    HMODULE hInstance = (HMODULE)GetWindowLongPtr(hWnd, GWLP_HINSTANCE);
    if (hInstance && hInstance == GetModuleHandle(L"sndvolsso.dll")) {
        *(HWND*)lParam = hWnd;
        return FALSE;
    }

    return TRUE;
}

#pragma endregion  // sndvol

////////////////////////////////////////////////////////////

// wParam - TRUE to subclass, FALSE to unsubclass
// lParam - subclass data
UINT g_subclassRegisteredMsg = RegisterWindowMessage(
    L"Windhawk_SetWindowSubclassFromAnyThread_" WH_MOD_ID);

BOOL SetWindowSubclassFromAnyThread(HWND hWnd,
                                    SUBCLASSPROC pfnSubclass,
                                    UINT_PTR uIdSubclass,
                                    DWORD_PTR dwRefData) {
    struct SET_WINDOW_SUBCLASS_FROM_ANY_THREAD_PARAM {
        SUBCLASSPROC pfnSubclass;
        UINT_PTR uIdSubclass;
        DWORD_PTR dwRefData;
        BOOL result;
    };

    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    if (dwThreadId == 0) {
        return FALSE;
    }

    if (dwThreadId == GetCurrentThreadId()) {
        return SetWindowSubclass(hWnd, pfnSubclass, uIdSubclass, dwRefData);
    }

    HHOOK hook = SetWindowsHookEx(
        WH_CALLWNDPROC,
        [](int nCode, WPARAM wParam,
           LPARAM lParam) WINAPI_LAMBDA_RETURN(LRESULT) {
            if (nCode == HC_ACTION) {
                const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
                if (cwp->message == g_subclassRegisteredMsg && cwp->wParam) {
                    SET_WINDOW_SUBCLASS_FROM_ANY_THREAD_PARAM* param =
                        (SET_WINDOW_SUBCLASS_FROM_ANY_THREAD_PARAM*)cwp->lParam;
                    param->result =
                        SetWindowSubclass(cwp->hwnd, param->pfnSubclass,
                                          param->uIdSubclass, param->dwRefData);
                }
            }

            return CallNextHookEx(nullptr, nCode, wParam, lParam);
        },
        nullptr, dwThreadId);
    if (!hook) {
        return FALSE;
    }

    SET_WINDOW_SUBCLASS_FROM_ANY_THREAD_PARAM param;
    param.pfnSubclass = pfnSubclass;
    param.uIdSubclass = uIdSubclass;
    param.dwRefData = dwRefData;
    param.result = FALSE;
    SendMessage(hWnd, g_subclassRegisteredMsg, TRUE, (LPARAM)&param);

    UnhookWindowsHookEx(hook);

    return param.result;
}

bool OnMouseWheel(HWND hWnd, WPARAM wParam, LPARAM lParam) {
    if (GetCapture()) {
        return false;
    }

    if (g_settings.ctrlScrollVolumeChange && GetKeyState(VK_CONTROL) >= 0) {
        return false;
    }

    POINT pt;
    pt.x = GET_X_LPARAM(lParam);
    pt.y = GET_Y_LPARAM(lParam);

    if (!IsPointInsideScrollArea(hWnd, pt)) {
        return false;
    }

    // Allows to steal focus.
    INPUT input;
    ZeroMemory(&input, sizeof(INPUT));
    SendInput(1, &input, sizeof(INPUT));

    OpenScrollSndVol(wParam, lParam);

    return true;
}

LRESULT CALLBACK TaskbarWindowSubclassProc(_In_ HWND hWnd,
                                           _In_ UINT uMsg,
                                           _In_ WPARAM wParam,
                                           _In_ LPARAM lParam,
                                           _In_ UINT_PTR uIdSubclass,
                                           _In_ DWORD_PTR dwRefData) {
    if (uMsg == WM_NCDESTROY || (uMsg == g_subclassRegisteredMsg && !wParam)) {
        RemoveWindowSubclass(hWnd, TaskbarWindowSubclassProc, 0);
    }

    LRESULT result = 0;

    switch (uMsg) {
        case WM_COPYDATA: {
            result = DefSubclassProc(hWnd, uMsg, wParam, lParam);

            typedef struct _notifyiconidentifier_internal {
                DWORD dwMagic;    // 0x34753423
                DWORD dwRequest;  // 1 for (x,y) | 2 for (w,h)
                DWORD cbSize;     // 0x20
                DWORD hWndHigh;
                DWORD hWndLow;
                UINT uID;
                GUID guidItem;
            } NOTIFYICONIDENTIFIER_INTERNAL;

            COPYDATASTRUCT* p_copydata = (COPYDATASTRUCT*)lParam;

            // Change Shell_NotifyIconGetRect handling result for the volume
            // icon. In case it's not visible, or in Windows 11, it returns 0,
            // which causes sndvol.exe to ignore the command line position.
            if (result == 0 && p_copydata->dwData == 0x03 &&
                p_copydata->cbData == sizeof(NOTIFYICONIDENTIFIER_INTERNAL)) {
                NOTIFYICONIDENTIFIER_INTERNAL* p_icon_ident =
                    (NOTIFYICONIDENTIFIER_INTERNAL*)p_copydata->lpData;
                if (p_icon_ident->dwMagic == 0x34753423 &&
                    (p_icon_ident->dwRequest == 0x01 ||
                     p_icon_ident->dwRequest == 0x02) &&
                    p_icon_ident->cbSize == 0x20 &&
                    memcmp(&p_icon_ident->guidItem,
                           "\x73\xAE\x20\x78\xE3\x23\x29\x42\x82\xC1\xE4"
                           "\x1C\xB6\x7D\x5B\x9C",
                           sizeof(GUID)) == 0) {
                    RECT rc;
                    GetWindowRect(hWnd, &rc);

                    if (p_icon_ident->dwRequest == 0x01)
                        result = MAKEWORD(rc.left, rc.top);
                    else
                        result =
                            MAKEWORD(rc.right - rc.left, rc.bottom - rc.top);
                }
            }
            break;
        }

        case WM_MOUSEWHEEL:
            if (g_nExplorerVersion < WIN_VERSION_11_21H2 &&
                OnMouseWheel(hWnd, wParam, lParam)) {
                result = 0;
            } else {
                result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
            }
            break;

        case WM_NCDESTROY:
            result = DefSubclassProc(hWnd, uMsg, wParam, lParam);

            if (hWnd != g_hTaskbarWnd) {
                g_secondaryTaskbarWindows.erase(hWnd);
            }
            break;

        default:
            result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
            break;
    }

    return result;
}

WNDPROC InputSiteWindowProc_Original;
LRESULT CALLBACK InputSiteWindowProc_Hook(HWND hWnd,
                                          UINT uMsg,
                                          WPARAM wParam,
                                          LPARAM lParam) {
    switch (uMsg) {
        case WM_POINTERWHEEL:
            if (HWND hRootWnd = GetAncestor(hWnd, GA_ROOT);
                IsTaskbarWindow(hRootWnd) &&
                OnMouseWheel(hRootWnd, wParam, lParam)) {
                return 0;
            }
            break;
    }

    return InputSiteWindowProc_Original(hWnd, uMsg, wParam, lParam);
}

void SubclassTaskbarWindow(HWND hWnd) {
    SetWindowSubclassFromAnyThread(hWnd, TaskbarWindowSubclassProc, 0, 0);
}

void UnsubclassTaskbarWindow(HWND hWnd) {
    SendMessage(hWnd, g_subclassRegisteredMsg, FALSE, 0);
}

void HandleIdentifiedInputSiteWindow(HWND hWnd) {
    if (!g_dwTaskbarThreadId ||
        GetWindowThreadProcessId(hWnd, nullptr) != g_dwTaskbarThreadId) {
        return;
    }

    HWND hParentWnd = GetParent(hWnd);
    WCHAR szClassName[64];
    if (!hParentWnd ||
        !GetClassName(hParentWnd, szClassName, ARRAYSIZE(szClassName)) ||
        _wcsicmp(szClassName,
                 L"Windows.UI.Composition.DesktopWindowContentBridge") != 0) {
        return;
    }

    hParentWnd = GetParent(hParentWnd);
    if (!hParentWnd || !IsTaskbarWindow(hParentWnd)) {
        return;
    }

    // At first, I tried to subclass the window instead of hooking its wndproc,
    // but the inputsite.dll code checks that the value wasn't changed, and
    // crashes otherwise.
    auto wndProc = (WNDPROC)GetWindowLongPtr(hWnd, GWLP_WNDPROC);
    WindhawkUtils::Wh_SetFunctionHookT(wndProc, InputSiteWindowProc_Hook,
                                       &InputSiteWindowProc_Original);

    if (g_initialized) {
        Wh_ApplyHookOperations();
    }

    Wh_Log(L"Hooked InputSite wndproc %p", wndProc);
    g_inputSiteProcHooked = true;
}

void HandleIdentifiedTaskbarWindow(HWND hWnd) {
    g_hTaskbarWnd = hWnd;
    g_dwTaskbarThreadId = GetWindowThreadProcessId(hWnd, nullptr);
    SndVolInit();
    SubclassTaskbarWindow(hWnd);
    for (HWND hSecondaryWnd : g_secondaryTaskbarWindows) {
        SubclassTaskbarWindow(hSecondaryWnd);
    }

    if (g_nExplorerVersion >= WIN_VERSION_11_21H2 && !g_inputSiteProcHooked) {
        HWND hXamlIslandWnd = FindWindowEx(
            hWnd, nullptr, L"Windows.UI.Composition.DesktopWindowContentBridge",
            nullptr);
        if (hXamlIslandWnd) {
            HWND hInputSiteWnd = FindWindowEx(
                hXamlIslandWnd, nullptr,
                L"Windows.UI.Input.InputSite.WindowClass", nullptr);
            if (hInputSiteWnd) {
                HandleIdentifiedInputSiteWindow(hInputSiteWnd);
            }
        }
    }
}

void HandleIdentifiedSecondaryTaskbarWindow(HWND hWnd) {
    if (!g_dwTaskbarThreadId ||
        GetWindowThreadProcessId(hWnd, nullptr) != g_dwTaskbarThreadId) {
        return;
    }

    g_secondaryTaskbarWindows.insert(hWnd);
    SubclassTaskbarWindow(hWnd);

    if (g_nExplorerVersion >= WIN_VERSION_11_21H2 && !g_inputSiteProcHooked) {
        HWND hXamlIslandWnd = FindWindowEx(
            hWnd, nullptr, L"Windows.UI.Composition.DesktopWindowContentBridge",
            nullptr);
        if (hXamlIslandWnd) {
            HWND hInputSiteWnd = FindWindowEx(
                hXamlIslandWnd, nullptr,
                L"Windows.UI.Input.InputSite.WindowClass", nullptr);
            if (hInputSiteWnd) {
                HandleIdentifiedInputSiteWindow(hInputSiteWnd);
            }
        }
    }
}

HWND FindCurrentProcessTaskbarWindows(
    std::unordered_set<HWND>* secondaryTaskbarWindows) {
    struct ENUM_WINDOWS_PARAM {
        HWND* hWnd;
        std::unordered_set<HWND>* secondaryTaskbarWindows;
    };

    HWND hWnd = nullptr;
    ENUM_WINDOWS_PARAM param = {&hWnd, secondaryTaskbarWindows};
    EnumWindows(
        [](HWND hWnd, LPARAM lParam) WINAPI_LAMBDA_RETURN(BOOL) {
            ENUM_WINDOWS_PARAM& param = *(ENUM_WINDOWS_PARAM*)lParam;

            DWORD dwProcessId = 0;
            if (!GetWindowThreadProcessId(hWnd, &dwProcessId) ||
                dwProcessId != GetCurrentProcessId())
                return TRUE;

            WCHAR szClassName[32];
            if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0)
                return TRUE;

            if (_wcsicmp(szClassName, L"Shell_TrayWnd") == 0) {
                *param.hWnd = hWnd;
            } else if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
                param.secondaryTaskbarWindows->insert(hWnd);
            }

            return TRUE;
        },
        (LPARAM)&param);

    return hWnd;
}

using CreateWindowExW_t = decltype(&CreateWindowExW);
CreateWindowExW_t CreateWindowExW_Original;
HWND WINAPI CreateWindowExW_Hook(DWORD dwExStyle,
                                 LPCWSTR lpClassName,
                                 LPCWSTR lpWindowName,
                                 DWORD dwStyle,
                                 int X,
                                 int Y,
                                 int nWidth,
                                 int nHeight,
                                 HWND hWndParent,
                                 HMENU hMenu,
                                 HINSTANCE hInstance,
                                 LPVOID lpParam) {
    HWND hWnd = CreateWindowExW_Original(dwExStyle, lpClassName, lpWindowName,
                                         dwStyle, X, Y, nWidth, nHeight,
                                         hWndParent, hMenu, hInstance, lpParam);
    if (!hWnd)
        return hWnd;

    BOOL bTextualClassName = ((ULONG_PTR)lpClassName & ~(ULONG_PTR)0xffff) != 0;

    if (bTextualClassName && _wcsicmp(lpClassName, L"Shell_TrayWnd") == 0) {
        Wh_Log(L"Taskbar window created: %08X", (DWORD)(ULONG_PTR)hWnd);
        HandleIdentifiedTaskbarWindow(hWnd);
    } else if (bTextualClassName &&
               _wcsicmp(lpClassName, L"Shell_SecondaryTrayWnd") == 0) {
        Wh_Log(L"Secondary taskbar window created: %08X",
               (DWORD)(ULONG_PTR)hWnd);
        HandleIdentifiedSecondaryTaskbarWindow(hWnd);
    }

    return hWnd;
}

using CreateWindowInBand_t = HWND(WINAPI*)(DWORD dwExStyle,
                                           LPCWSTR lpClassName,
                                           LPCWSTR lpWindowName,
                                           DWORD dwStyle,
                                           int X,
                                           int Y,
                                           int nWidth,
                                           int nHeight,
                                           HWND hWndParent,
                                           HMENU hMenu,
                                           HINSTANCE hInstance,
                                           LPVOID lpParam,
                                           DWORD dwBand);
CreateWindowInBand_t CreateWindowInBand_Original;
HWND WINAPI CreateWindowInBand_Hook(DWORD dwExStyle,
                                    LPCWSTR lpClassName,
                                    LPCWSTR lpWindowName,
                                    DWORD dwStyle,
                                    int X,
                                    int Y,
                                    int nWidth,
                                    int nHeight,
                                    HWND hWndParent,
                                    HMENU hMenu,
                                    HINSTANCE hInstance,
                                    LPVOID lpParam,
                                    DWORD dwBand) {
    HWND hWnd = CreateWindowInBand_Original(
        dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight,
        hWndParent, hMenu, hInstance, lpParam, dwBand);
    if (!hWnd)
        return hWnd;

    BOOL bTextualClassName = ((ULONG_PTR)lpClassName & ~(ULONG_PTR)0xffff) != 0;

    if (bTextualClassName &&
        _wcsicmp(lpClassName, L"Windows.UI.Input.InputSite.WindowClass") == 0) {
        Wh_Log(L"InputSite window created: %08X", (DWORD)(ULONG_PTR)hWnd);
        if (g_nExplorerVersion >= WIN_VERSION_11_21H2 &&
            !g_inputSiteProcHooked) {
            HandleIdentifiedInputSiteWindow(hWnd);
        }
    }

    return hWnd;
}

void LoadSettings() {
    PCWSTR volumeIndicator = Wh_GetStringSetting(L"volumeIndicator");
    g_settings.volumeIndicator = VolumeIndicator::Win11;
    if (wcscmp(volumeIndicator, L"modern") == 0) {
        g_settings.volumeIndicator = VolumeIndicator::Modern;
    } else if (wcscmp(volumeIndicator, L"classic") == 0) {
        g_settings.volumeIndicator = VolumeIndicator::Classic;
    } else if (wcscmp(volumeIndicator, L"none") == 0) {
        g_settings.volumeIndicator = VolumeIndicator::None;
    } else {
        // Old option for compatibility.
        if (wcscmp(volumeIndicator, L"tooltip") == 0) {
            g_settings.volumeIndicator = VolumeIndicator::None;
        }
    }
    Wh_FreeStringSetting(volumeIndicator);

    PCWSTR scrollArea = Wh_GetStringSetting(L"scrollArea");
    g_settings.scrollArea = ScrollArea::taskbar;
    if (wcscmp(scrollArea, L"notification_area") == 0) {
        g_settings.scrollArea = ScrollArea::notificationArea;
    } else if (wcscmp(scrollArea, L"taskbarWithoutNotificationArea") == 0) {
        g_settings.scrollArea = ScrollArea::taskbarWithoutNotificationArea;
    }
    Wh_FreeStringSetting(scrollArea);

    g_settings.middleClickToMute = Wh_GetIntSetting(L"middleClickToMute");
    g_settings.ctrlScrollVolumeChange =
        Wh_GetIntSetting(L"ctrlScrollVolumeChange");
    g_settings.noAutomaticMuteToggle =
        Wh_GetIntSetting(L"noAutomaticMuteToggle");
    g_settings.volumeChangeStep = Wh_GetIntSetting(L"volumeChangeStep");
    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
}

using VolumeSystemTrayIconDataModel_OnIconClicked_t =
    void(WINAPI*)(void* pThis, void* iconClickedEventArgs);
VolumeSystemTrayIconDataModel_OnIconClicked_t
    VolumeSystemTrayIconDataModel_OnIconClicked_Original;
void WINAPI
VolumeSystemTrayIconDataModel_OnIconClicked_Hook(void* pThis,
                                                 void* iconClickedEventArgs) {
    Wh_Log(L">");

    if (g_settings.middleClickToMute && GetKeyState(VK_MBUTTON) < 0) {
        ToggleVolMuted();
        return;
    }

    VolumeSystemTrayIconDataModel_OnIconClicked_Original(pThis,
                                                         iconClickedEventArgs);
}

bool HookTaskbarViewDllSymbols(HMODULE module) {
    // Taskbar.View.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: void __cdecl winrt::SystemTray::implementation::VolumeSystemTrayIconDataModel::OnIconClicked(struct winrt::SystemTray::IconClickedEventArgs const &))"},
            &VolumeSystemTrayIconDataModel_OnIconClicked_Original,
            VolumeSystemTrayIconDataModel_OnIconClicked_Hook,
            true,
        },
    };

    return HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks));
}

HMODULE GetTaskbarViewModuleHandle() {
    HMODULE module = GetModuleHandle(L"Taskbar.View.dll");
    if (!module) {
        module = GetModuleHandle(L"ExplorerExtensions.dll");
    }

    return module;
}

bool ShouldHookTaskbarViewDllSymbols() {
    return g_nWinVersion >= WIN_VERSION_11_22H2 && g_settings.middleClickToMute;
}

void HandleLoadedModuleIfTaskbarView(HMODULE module, LPCWSTR lpLibFileName) {
    if (ShouldHookTaskbarViewDllSymbols() && !g_taskbarViewDllLoaded &&
        GetTaskbarViewModuleHandle() == module &&
        !g_taskbarViewDllLoaded.exchange(true)) {
        Wh_Log(L"Loaded %s", lpLibFileName);

        if (HookTaskbarViewDllSymbols(module)) {
            Wh_ApplyHookOperations();
        }
    }
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

void HandleLoadedExplorerPatcher() {
    HMODULE hMods[1024];
    DWORD cbNeeded;
    if (EnumProcessModules(GetCurrentProcess(), hMods, sizeof(hMods),
                           &cbNeeded)) {
        for (size_t i = 0; i < cbNeeded / sizeof(HMODULE); i++) {
            if (IsExplorerPatcherModule(hMods[i])) {
                if (g_nExplorerVersion >= WIN_VERSION_11_21H2) {
                    g_nExplorerVersion = WIN_VERSION_10_20H1;
                }
                break;
            }
        }
    }
}

void HandleLoadedModuleIfExplorerPatcher(HMODULE module) {
    if (module && !((ULONG_PTR)module & 3)) {
        if (IsExplorerPatcherModule(module)) {
            if (g_nExplorerVersion >= WIN_VERSION_11_21H2) {
                g_nExplorerVersion = WIN_VERSION_10_20H1;
            }
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
        HandleLoadedModuleIfExplorerPatcher(module);
        HandleLoadedModuleIfTaskbarView(module, lpLibFileName);
    }

    return module;
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    if (!WindowsVersionInit()) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
    }

    g_nExplorerVersion = g_nWinVersion;
    if (g_nExplorerVersion >= WIN_VERSION_11_21H2 &&
        g_settings.oldTaskbarOnWin11) {
        g_nExplorerVersion = WIN_VERSION_10_20H1;
    }

    if (ShouldHookTaskbarViewDllSymbols()) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            g_taskbarViewDllLoaded = true;
            if (!HookTaskbarViewDllSymbols(taskbarViewModule)) {
                return FALSE;
            }
        } else {
            Wh_Log(L"Taskbar view module not loaded yet");
        }
    }

    WindhawkUtils::Wh_SetFunctionHookT(CreateWindowExW, CreateWindowExW_Hook,
                                       &CreateWindowExW_Original);

    HMODULE user32Module =
        LoadLibraryEx(L"user32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (user32Module) {
        auto pCreateWindowInBand = (CreateWindowInBand_t)GetProcAddress(
            user32Module, "CreateWindowInBand");
        if (pCreateWindowInBand) {
            WindhawkUtils::Wh_SetFunctionHookT(pCreateWindowInBand,
                                               CreateWindowInBand_Hook,
                                               &CreateWindowInBand_Original);
        }
    }

    HandleLoadedExplorerPatcher();

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(
        kernelBaseModule, "LoadLibraryExW");
    WindhawkUtils::Wh_SetFunctionHookT(pKernelBaseLoadLibraryExW,
                                       LoadLibraryExW_Hook,
                                       &LoadLibraryExW_Original);

    g_initialized = true;

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (ShouldHookTaskbarViewDllSymbols() && !g_taskbarViewDllLoaded) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            if (!g_taskbarViewDllLoaded.exchange(true)) {
                Wh_Log(L"Got Taskbar.View.dll");

                if (HookTaskbarViewDllSymbols(taskbarViewModule)) {
                    Wh_ApplyHookOperations();
                }
            }
        }
    }

    // Try again in case there's a race between the previous attempt and the
    // LoadLibraryExW hook.
    HandleLoadedExplorerPatcher();

    WNDCLASS wndclass;
    if (GetClassInfo(GetModuleHandle(NULL), L"Shell_TrayWnd", &wndclass)) {
        HWND hWnd =
            FindCurrentProcessTaskbarWindows(&g_secondaryTaskbarWindows);
        if (hWnd) {
            HandleIdentifiedTaskbarWindow(hWnd);
        }
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");

    if (g_hTaskbarWnd) {
        UnsubclassTaskbarWindow(g_hTaskbarWnd);

        for (HWND hSecondaryWnd : g_secondaryTaskbarWindows) {
            UnsubclassTaskbarWindow(hSecondaryWnd);
        }
    }

    CleanupSndVol();
    SndVolUninit();
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;
    bool prevMiddleClickToMute = g_settings.middleClickToMute;

    LoadSettings();

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11 ||
               g_settings.middleClickToMute != prevMiddleClickToMute;

    return TRUE;
}
