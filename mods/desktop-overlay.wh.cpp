// ==WindhawkMod==
// @id              desktop-overlay
// @name            Desktop Overlay
// @description     Display custom content on the desktop behind icons
// @version         0.1
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lcomctl32 -ldxgi -ld2d1 -ldwrite -ld3d11 -ldcomp
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
# Desktop Overlay

Display custom content on the desktop behind icons, such as text. Uses the
WorkerW technique combined with Direct2D and DirectComposition for proper
per-pixel alpha transparency.

### Features

* Customizable text content
* Font family, size, weight, and style options
* Text color with transparency support (ARGB format)
* Percentage-based positioning
* Multi-monitor support

### Acknowledgements

* Based on the technique from the [weebp](https://github.com/Francesco149/weebp)
project.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- text: Hello World
  $name: Text content
  $description: The text to display on the desktop
- fontSize: 48
  $name: Font size
  $description: Size of the text in points
- textColor: "#80FFFFFF"
  $name: Text color
  $description: >-
    Color in ARGB hex format. Examples: #80FFFFFF (semi-transparent white),
    #FF0000 (red), #80000000 (semi-transparent black)
- fontFamily: Segoe UI
  $name: Font family
  $description: >-
    Font family name. For a list of fonts shipped with Windows, see:
    https://learn.microsoft.com/en-us/typography/fonts/windows_11_font_list
- fontWeight: ""
  $name: Font weight
  $options:
  - "": Default
  - Thin: Thin
  - Light: Light
  - Normal: Normal
  - Medium: Medium
  - SemiBold: Semi bold
  - Bold: Bold
  - ExtraBold: Extra bold
- fontStyle: ""
  $name: Font style
  $options:
  - "": Default
  - Normal: Normal
  - Italic: Italic
- verticalPosition: 20
  $name: Vertical position
  $description: Position in percentage (0 = top, 50 = center, 100 = bottom)
- horizontalPosition: 50
  $name: Horizontal position
  $description: Position in percentage (0 = left, 50 = center, 100 = right)
- monitor: 1
  $name: Monitor
  $description: The monitor number to display text on (1-based)
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <commctrl.h>
#include <d2d1_1.h>
#include <d2d1helper.h>
#include <d3d11.h>
#include <dcomp.h>
#include <dwrite.h>
#include <dxgi1_3.h>
#include <wrl/client.h>

using Microsoft::WRL::ComPtr;

////////////////////////////////////////////////////////////////////////////////
// Types

struct Settings {
    WindhawkUtils::StringSetting text;
    int fontSize;
    BYTE colorA;
    BYTE colorR;
    BYTE colorG;
    BYTE colorB;
    WindhawkUtils::StringSetting fontFamily;
    int fontWeight;
    bool fontItalic;
    int verticalPosition;
    int horizontalPosition;
    int monitor;
};

////////////////////////////////////////////////////////////////////////////////
// Globals

Settings g_settings;

// DirectX device objects (shared).
ComPtr<ID3D11Device> g_d3dDevice;
ComPtr<IDXGIDevice> g_dxgiDevice;
ComPtr<IDXGIFactory2> g_dxgiFactory;
ComPtr<ID2D1Factory1> g_d2dFactory;
ComPtr<ID2D1Device> g_d2dDevice;
ComPtr<IDWriteFactory> g_dwriteFactory;

// Overlay window and resources.
HWND g_overlayWnd;
ComPtr<IDXGISwapChain1> g_swapChain;
ComPtr<ID2D1DeviceContext> g_dc;
ComPtr<IDCompositionDevice> g_compositionDevice;
ComPtr<IDCompositionTarget> g_compositionTarget;
ComPtr<IDCompositionVisual> g_compositionVisual;
ComPtr<IDWriteTextFormat> g_textFormat;
ComPtr<ID2D1SolidColorBrush> g_textBrush;

////////////////////////////////////////////////////////////////////////////////
// Utility functions

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

HMONITOR GetMonitorById(int monitorId) {
    HMONITOR monitorResult = nullptr;
    int currentMonitorId = 0;

    auto monitorEnumProc = [&monitorResult, &currentMonitorId,
                            monitorId](HMONITOR hMonitor) -> BOOL {
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

bool ParseColor(PCWSTR colorStr, BYTE* a, BYTE* r, BYTE* g, BYTE* b) {
    if (!colorStr || !*colorStr) {
        return false;
    }

    if (*colorStr == L'#') {
        colorStr++;
    }

    size_t len = wcslen(colorStr);
    unsigned int value = 0;

    for (size_t i = 0; i < len; i++) {
        WCHAR c = colorStr[i];
        int digit;
        if (c >= L'0' && c <= L'9') {
            digit = c - L'0';
        } else if (c >= L'A' && c <= L'F') {
            digit = c - L'A' + 10;
        } else if (c >= L'a' && c <= L'f') {
            digit = c - L'a' + 10;
        } else {
            return false;
        }
        value = (value << 4) | digit;
    }

    if (len == 6) {
        *a = 255;
        *r = (value >> 16) & 0xFF;
        *g = (value >> 8) & 0xFF;
        *b = value & 0xFF;
    } else if (len == 8) {
        *a = (value >> 24) & 0xFF;
        *r = (value >> 16) & 0xFF;
        *g = (value >> 8) & 0xFF;
        *b = value & 0xFF;
    } else {
        return false;
    }

    return true;
}

DWRITE_FONT_WEIGHT GetDWriteFontWeight() {
    if (g_settings.fontWeight <= 100)
        return DWRITE_FONT_WEIGHT_THIN;
    if (g_settings.fontWeight <= 200)
        return DWRITE_FONT_WEIGHT_EXTRA_LIGHT;
    if (g_settings.fontWeight <= 300)
        return DWRITE_FONT_WEIGHT_LIGHT;
    if (g_settings.fontWeight <= 400)
        return DWRITE_FONT_WEIGHT_NORMAL;
    if (g_settings.fontWeight <= 500)
        return DWRITE_FONT_WEIGHT_MEDIUM;
    if (g_settings.fontWeight <= 600)
        return DWRITE_FONT_WEIGHT_SEMI_BOLD;
    if (g_settings.fontWeight <= 700)
        return DWRITE_FONT_WEIGHT_BOLD;
    if (g_settings.fontWeight <= 800)
        return DWRITE_FONT_WEIGHT_EXTRA_BOLD;
    return DWRITE_FONT_WEIGHT_BLACK;
}

////////////////////////////////////////////////////////////////////////////////
// Desktop window detection

bool IsFolderViewWnd(HWND hWnd) {
    WCHAR buffer[64];

    if (!GetClassName(hWnd, buffer, ARRAYSIZE(buffer)) ||
        _wcsicmp(buffer, L"SysListView32")) {
        return false;
    }

    if (!GetWindowText(hWnd, buffer, ARRAYSIZE(buffer)) ||
        _wcsicmp(buffer, L"FolderView")) {
        return false;
    }

    HWND hParentWnd = GetAncestor(hWnd, GA_PARENT);
    if (!hParentWnd) {
        return false;
    }

    if (!GetClassName(hParentWnd, buffer, ARRAYSIZE(buffer)) ||
        _wcsicmp(buffer, L"SHELLDLL_DefView")) {
        return false;
    }

    if (GetWindowTextLength(hParentWnd) > 0) {
        return false;
    }

    HWND hParentWnd2 = GetAncestor(hParentWnd, GA_PARENT);
    if (!hParentWnd2) {
        return false;
    }

    if ((!GetClassName(hParentWnd2, buffer, ARRAYSIZE(buffer)) ||
         _wcsicmp(buffer, L"Progman")) &&
        hParentWnd2 != GetShellWindow()) {
        return false;
    }

    return true;
}

// Find the WorkerW window behind desktop icons.
// Based on weebp: https://github.com/Francesco149/weebp
HWND GetWorkerW() {
    HWND hProgman = FindWindow(L"Progman", nullptr);
    if (!hProgman) {
        return nullptr;
    }

    // Send undocumented message to spawn WorkerW windows.
    SendMessage(hProgman, 0x052C, 0xD, 0);
    SendMessage(hProgman, 0x052C, 0xD, 1);

    // Find window with SHELLDLL_DefView, then get the next WorkerW sibling.
    HWND hWorkerW = nullptr;
    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            if (!FindWindowEx(hWnd, nullptr, L"SHELLDLL_DefView", nullptr)) {
                return TRUE;
            }
            HWND hWorker = FindWindowEx(nullptr, hWnd, L"WorkerW", nullptr);
            if (hWorker) {
                *(HWND*)lParam = hWorker;
                return FALSE;
            }
            return TRUE;
        },
        (LPARAM)&hWorkerW);

    // Fallback with alternative message parameters.
    if (!hWorkerW) {
        SendMessage(hProgman, 0x052C, 0, 0);
        EnumWindows(
            [](HWND hWnd, LPARAM lParam) -> BOOL {
                if (!FindWindowEx(hWnd, nullptr, L"SHELLDLL_DefView",
                                  nullptr)) {
                    return TRUE;
                }
                HWND hWorker = FindWindowEx(nullptr, hWnd, L"WorkerW", nullptr);
                if (hWorker) {
                    *(HWND*)lParam = hWorker;
                    return FALSE;
                }
                return TRUE;
            },
            (LPARAM)&hWorkerW);
    }

    // Fallback: WorkerW as child of Progman.
    if (!hWorkerW) {
        hWorkerW = FindWindowEx(hProgman, nullptr, L"WorkerW", nullptr);
    }

    // Final fallback: use Progman itself.
    if (!hWorkerW) {
        hWorkerW = hProgman;
    }

    return hWorkerW;
}

////////////////////////////////////////////////////////////////////////////////
// DirectX initialization

bool InitDirectX() {
    HRESULT hr;

    UINT flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
    hr = D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, flags,
                           nullptr, 0, D3D11_SDK_VERSION, &g_d3dDevice, nullptr,
                           nullptr);
    if (FAILED(hr)) {
        Wh_Log(L"D3D11CreateDevice failed: 0x%08X", hr);
        return false;
    }

    hr = g_d3dDevice.As(&g_dxgiDevice);
    if (FAILED(hr)) {
        Wh_Log(L"QueryInterface IDXGIDevice failed: 0x%08X", hr);
        return false;
    }

    hr = CreateDXGIFactory2(0, IID_PPV_ARGS(&g_dxgiFactory));
    if (FAILED(hr)) {
        Wh_Log(L"CreateDXGIFactory2 failed: 0x%08X", hr);
        return false;
    }

    hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED,
                           IID_PPV_ARGS(&g_d2dFactory));
    if (FAILED(hr)) {
        Wh_Log(L"D2D1CreateFactory failed: 0x%08X", hr);
        return false;
    }

    hr = g_d2dFactory->CreateDevice(g_dxgiDevice.Get(), &g_d2dDevice);
    if (FAILED(hr)) {
        Wh_Log(L"D2D CreateDevice failed: 0x%08X", hr);
        return false;
    }

    hr = DWriteCreateFactory(
        DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory),
        reinterpret_cast<IUnknown**>(g_dwriteFactory.GetAddressOf()));
    if (FAILED(hr)) {
        Wh_Log(L"DWriteCreateFactory failed: 0x%08X", hr);
        return false;
    }

    return true;
}

void UninitDirectX() {
    g_dwriteFactory.Reset();
    g_d2dDevice.Reset();
    g_d2dFactory.Reset();
    g_dxgiFactory.Reset();
    g_dxgiDevice.Reset();
    g_d3dDevice.Reset();
}

////////////////////////////////////////////////////////////////////////////////
// Overlay rendering

bool CreateSwapChainResources(UINT width, UINT height) {
    HRESULT hr;

    // Create swap chain for composition with premultiplied alpha.
    DXGI_SWAP_CHAIN_DESC1 scd = {};
    scd.Width = width;
    scd.Height = height;
    scd.Format = DXGI_FORMAT_B8G8R8A8_UNORM;
    scd.SampleDesc.Count = 1;
    scd.BufferUsage = DXGI_USAGE_RENDER_TARGET_OUTPUT;
    scd.BufferCount = 2;
    scd.Scaling = DXGI_SCALING_STRETCH;
    scd.SwapEffect = DXGI_SWAP_EFFECT_FLIP_DISCARD;
    scd.AlphaMode = DXGI_ALPHA_MODE_PREMULTIPLIED;

    hr = g_dxgiFactory->CreateSwapChainForComposition(g_dxgiDevice.Get(), &scd,
                                                      nullptr, &g_swapChain);
    if (FAILED(hr)) {
        Wh_Log(L"CreateSwapChainForComposition failed: 0x%08X", hr);
        return false;
    }

    // Create D2D device context.
    hr = g_d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE,
                                          &g_dc);
    if (FAILED(hr)) {
        Wh_Log(L"CreateDeviceContext failed: 0x%08X", hr);
        return false;
    }

    // Create bitmap target from swap chain surface.
    ComPtr<IDXGISurface2> surface;
    hr = g_swapChain->GetBuffer(0, IID_PPV_ARGS(&surface));
    if (FAILED(hr)) {
        Wh_Log(L"GetBuffer failed: 0x%08X", hr);
        return false;
    }

    D2D1_BITMAP_PROPERTIES1 bitmapProperties = {};
    bitmapProperties.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
    bitmapProperties.pixelFormat.format = DXGI_FORMAT_B8G8R8A8_UNORM;
    bitmapProperties.bitmapOptions =
        D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW;

    ComPtr<ID2D1Bitmap1> targetBitmap;
    hr = g_dc->CreateBitmapFromDxgiSurface(surface.Get(), bitmapProperties,
                                           &targetBitmap);
    if (FAILED(hr)) {
        Wh_Log(L"CreateBitmapFromDxgiSurface failed: 0x%08X", hr);
        return false;
    }

    g_dc->SetTarget(targetBitmap.Get());

    // Create DirectComposition device and visual tree.
    hr = DCompositionCreateDevice(g_dxgiDevice.Get(),
                                  IID_PPV_ARGS(&g_compositionDevice));
    if (FAILED(hr)) {
        Wh_Log(L"DCompositionCreateDevice failed: 0x%08X", hr);
        return false;
    }

    hr = g_compositionDevice->CreateTargetForHwnd(g_overlayWnd, TRUE,
                                                  &g_compositionTarget);
    if (FAILED(hr)) {
        Wh_Log(L"CreateTargetForHwnd failed: 0x%08X", hr);
        return false;
    }

    hr = g_compositionDevice->CreateVisual(&g_compositionVisual);
    if (FAILED(hr)) {
        Wh_Log(L"CreateVisual failed: 0x%08X", hr);
        return false;
    }

    hr = g_compositionVisual->SetContent(g_swapChain.Get());
    if (FAILED(hr)) {
        Wh_Log(L"SetContent failed: 0x%08X", hr);
        return false;
    }

    hr = g_compositionTarget->SetRoot(g_compositionVisual.Get());
    if (FAILED(hr)) {
        Wh_Log(L"SetRoot failed: 0x%08X", hr);
        return false;
    }

    hr = g_compositionDevice->Commit();
    if (FAILED(hr)) {
        Wh_Log(L"Commit failed: 0x%08X", hr);
        return false;
    }

    // Create text format.
    PCWSTR fontFamilyName = g_settings.fontFamily.get();
    if (!fontFamilyName || !*fontFamilyName) {
        fontFamilyName = L"Segoe UI";
    }

    hr = g_dwriteFactory->CreateTextFormat(
        fontFamilyName, nullptr, GetDWriteFontWeight(),
        g_settings.fontItalic ? DWRITE_FONT_STYLE_ITALIC
                              : DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL, (FLOAT)g_settings.fontSize, L"",
        &g_textFormat);
    if (FAILED(hr)) {
        Wh_Log(L"CreateTextFormat failed: 0x%08X", hr);
        return false;
    }

    // Create text brush. Use full opacity here since transparency is applied
    // via a layer during rendering (to also affect color emoji).
    D2D1_COLOR_F textColor =
        D2D1::ColorF(g_settings.colorR / 255.0f, g_settings.colorG / 255.0f,
                     g_settings.colorB / 255.0f, 1.0f);
    hr = g_dc->CreateSolidColorBrush(textColor, &g_textBrush);
    if (FAILED(hr)) {
        Wh_Log(L"CreateSolidColorBrush failed: 0x%08X", hr);
        return false;
    }

    return true;
}

void ReleaseSwapChainResources() {
    g_textBrush.Reset();
    g_textFormat.Reset();
    g_compositionVisual.Reset();
    g_compositionTarget.Reset();
    g_compositionDevice.Reset();
    g_dc.Reset();
    g_swapChain.Reset();
}

void RenderOverlay() {
    if (!g_dc || !g_swapChain) {
        return;
    }

    RECT rc;
    GetClientRect(g_overlayWnd, &rc);
    UINT width = rc.right - rc.left;
    UINT height = rc.bottom - rc.top;

    g_dc->BeginDraw();
    g_dc->Clear(D2D1::ColorF(0, 0, 0, 0));

    PCWSTR text = g_settings.text.get();
    if (text && *text && g_textFormat && g_textBrush) {
        HMONITOR monitor = GetMonitorById(g_settings.monitor - 1);
        if (!monitor) {
            monitor = MonitorFromPoint({0, 0}, MONITOR_DEFAULTTONEAREST);
        }

        MONITORINFO monitorInfo{.cbSize = sizeof(monitorInfo)};
        if (GetMonitorInfo(monitor, &monitorInfo)) {
            RECT workArea = monitorInfo.rcWork;

            ComPtr<IDWriteTextLayout> textLayout;
            g_dwriteFactory->CreateTextLayout(text, (UINT32)wcslen(text),
                                              g_textFormat.Get(), (FLOAT)width,
                                              (FLOAT)height, &textLayout);

            if (textLayout) {
                DWRITE_TEXT_METRICS metrics;
                textLayout->GetMetrics(&metrics);

                float textWidth = metrics.width;
                float textHeight = metrics.height;
                float workWidth = (float)(workArea.right - workArea.left);
                float workHeight = (float)(workArea.bottom - workArea.top);

                // Calculate position based on percentage (0-100).
                // The text is centered at the percentage point.
                float x = workArea.left +
                          (workWidth - textWidth) *
                              (g_settings.horizontalPosition / 100.0f);
                float y =
                    workArea.top + (workHeight - textHeight) *
                                       (g_settings.verticalPosition / 100.0f);

                // Use a layer with opacity so color emoji also respect alpha.
                float opacity = g_settings.colorA / 255.0f;
                g_dc->PushLayer(
                    D2D1::LayerParameters(D2D1::InfiniteRect(), nullptr,
                                          D2D1_ANTIALIAS_MODE_PER_PRIMITIVE,
                                          D2D1::IdentityMatrix(), opacity),
                    nullptr);

                g_dc->DrawTextLayout(D2D1::Point2F(x, y), textLayout.Get(),
                                     g_textBrush.Get(),
                                     D2D1_DRAW_TEXT_OPTIONS_ENABLE_COLOR_FONT);

                g_dc->PopLayer();
            }
        }
    }

    g_dc->EndDraw();
    g_swapChain->Present(1, 0);
}

////////////////////////////////////////////////////////////////////////////////
// Overlay window management

void CreateOverlayWindow() {
    HWND hWorkerW = GetWorkerW();
    if (!hWorkerW) {
        Wh_Log(L"Failed to find WorkerW");
        return;
    }

    WNDCLASS wc = {};
    wc.lpfnWndProc = DefWindowProc;
    wc.hInstance = GetModuleHandle(nullptr);
    wc.lpszClassName = L"DesktopOverlay_" WH_MOD_ID;
    RegisterClass(&wc);

    RECT rc;
    GetWindowRect(hWorkerW, &rc);
    int width = rc.right - rc.left;
    int height = rc.bottom - rc.top;

    g_overlayWnd = CreateWindowEx(
        WS_EX_NOREDIRECTIONBITMAP | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE,
        wc.lpszClassName, nullptr, WS_CHILD | WS_VISIBLE, 0, 0, width, height,
        hWorkerW, nullptr, wc.hInstance, nullptr);

    if (!g_overlayWnd) {
        Wh_Log(L"Failed to create overlay window: %u", GetLastError());
        return;
    }

    if (CreateSwapChainResources(width, height)) {
        RenderOverlay();
    }
}

void DestroyOverlayWindow() {
    ReleaseSwapChainResources();

    if (g_overlayWnd) {
        DestroyWindow(g_overlayWnd);
        g_overlayWnd = nullptr;
    }
}

////////////////////////////////////////////////////////////////////////////////
// Hooks

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
                                 PVOID lpParam) {
    HWND hWnd = CreateWindowExW_Original(dwExStyle, lpClassName, lpWindowName,
                                         dwStyle, X, Y, nWidth, nHeight,
                                         hWndParent, hMenu, hInstance, lpParam);
    if (!hWnd || !IsFolderViewWnd(hWnd)) {
        return hWnd;
    }

    Wh_Log(L"FolderView window created");

    // Delay overlay creation to let the desktop fully initialize.
    static UINT_PTR s_timer = 0;
    s_timer = SetTimer(nullptr, s_timer, 1000,
                       [](HWND, UINT, UINT_PTR idEvent, DWORD) {
                           KillTimer(nullptr, idEvent);
                           CreateOverlayWindow();
                       });

    return hWnd;
}

////////////////////////////////////////////////////////////////////////////////
// Settings

void LoadSettings() {
    g_settings.text = WindhawkUtils::StringSetting::make(L"text");

    g_settings.fontSize = Wh_GetIntSetting(L"fontSize");
    if (g_settings.fontSize <= 0) {
        g_settings.fontSize = 48;
    }

    PCWSTR textColor = Wh_GetStringSetting(L"textColor");
    if (!ParseColor(textColor, &g_settings.colorA, &g_settings.colorR,
                    &g_settings.colorG, &g_settings.colorB)) {
        g_settings.colorA = 0x80;
        g_settings.colorR = 0xFF;
        g_settings.colorG = 0xFF;
        g_settings.colorB = 0xFF;
    }
    Wh_FreeStringSetting(textColor);

    g_settings.fontFamily = WindhawkUtils::StringSetting::make(L"fontFamily");

    PCWSTR fontWeight = Wh_GetStringSetting(L"fontWeight");
    g_settings.fontWeight = 400;
    if (wcscmp(fontWeight, L"Thin") == 0) {
        g_settings.fontWeight = 100;
    } else if (wcscmp(fontWeight, L"Light") == 0) {
        g_settings.fontWeight = 300;
    } else if (wcscmp(fontWeight, L"Normal") == 0) {
        g_settings.fontWeight = 400;
    } else if (wcscmp(fontWeight, L"Medium") == 0) {
        g_settings.fontWeight = 500;
    } else if (wcscmp(fontWeight, L"SemiBold") == 0) {
        g_settings.fontWeight = 600;
    } else if (wcscmp(fontWeight, L"Bold") == 0) {
        g_settings.fontWeight = 700;
    } else if (wcscmp(fontWeight, L"ExtraBold") == 0) {
        g_settings.fontWeight = 800;
    }
    Wh_FreeStringSetting(fontWeight);

    PCWSTR fontStyle = Wh_GetStringSetting(L"fontStyle");
    g_settings.fontItalic = wcscmp(fontStyle, L"Italic") == 0;
    Wh_FreeStringSetting(fontStyle);

    g_settings.verticalPosition = Wh_GetIntSetting(L"verticalPosition");
    g_settings.horizontalPosition = Wh_GetIntSetting(L"horizontalPosition");

    g_settings.monitor = Wh_GetIntSetting(L"monitor");
    if (g_settings.monitor <= 0) {
        g_settings.monitor = 1;
    }
}

////////////////////////////////////////////////////////////////////////////////
// Mod lifecycle

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    if (!InitDirectX()) {
        Wh_Log(L"InitDirectX failed");
        return FALSE;
    }

    Wh_SetFunctionHook((void*)CreateWindowExW, (void*)CreateWindowExW_Hook,
                       (void**)&CreateWindowExW_Original);

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    HWND hWorkerW = GetWorkerW();
    if (hWorkerW) {
        RunFromWindowThread(
            hWorkerW, [](void*) { CreateOverlayWindow(); }, nullptr);
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");

    if (g_overlayWnd) {
        HWND hWorkerW = GetParent(g_overlayWnd);
        if (hWorkerW) {
            RunFromWindowThread(
                hWorkerW, [](void*) { DestroyOverlayWindow(); }, nullptr);
        } else {
            DestroyOverlayWindow();
        }
    }

    UninitDirectX();
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");
    *bReload = TRUE;
    return TRUE;
}
