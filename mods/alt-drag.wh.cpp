// ==WindhawkMod==
// @id              alt-drag
// @name            AltDrag (WIP)
// @description     AltDrag allows you to move any window by holding Alt and dragging it with your mouse
// @version         0.1
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         *
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
# AltDrag

AltDrag allows you to move any window by holding Alt and dragging it with your
mouse.

The idea was inspired by [the original AltDrag
tool](https://stefansundin.github.io/altdrag/).

## WIP

This is a limited work-in-progress version. Unsupported programs include:
* Chrome and other Chromium-based browsers.
* Electron apps.
* Some UWP/WinUI elements, e.g. parts of File Explorer.
*/
// ==/WindhawkModReadme==

#include <atomic>
#include <memory>
#include <mutex>
#include <unordered_set>

std::atomic<bool> g_uninitializing;
std::atomic<int> g_hookRefCount;

thread_local HHOOK g_callWndProcHook;
std::mutex g_allCallWndProcHooksMutex;
std::unordered_set<HHOOK> g_allCallWndProcHooks;

constexpr WCHAR kOriginalWndProcProp[] = L"WindhawkOriginalWndProc_" WH_MOD_ID;

auto HookRefCountScope() {
    g_hookRefCount++;
    return std::unique_ptr<decltype(g_hookRefCount),
                           void (*)(decltype(g_hookRefCount)*)>{
        &g_hookRefCount, [](auto hookRefCount) { (*hookRefCount)--; }};
}

LRESULT CALLBACK CustomWndProc(HWND hWnd,
                               UINT uMsg,
                               WPARAM wParam,
                               LPARAM lParam) {
    WNDPROC originalWndProc = (WNDPROC)GetProp(hWnd, kOriginalWndProcProp);
    SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)originalWndProc);

    switch (uMsg) {
        case WM_NCHITTEST: {
            HWND hParentWnd = GetAncestor(hWnd, GA_PARENT);
            bool hasParentWnd = hParentWnd && hParentWnd != GetDesktopWindow();
            return hasParentWnd ? HTTRANSPARENT : HTCAPTION;
        }

        case WM_SETCURSOR: {
            SetCursor(LoadCursor(nullptr, IDC_SIZEALL));
            return TRUE;
        }
    }

    return originalWndProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK CallWndProc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto hookScope = HookRefCountScope();

    if (nCode != HC_ACTION) {
        return CallNextHookEx(nullptr, nCode, wParam, lParam);
    }

    const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;

    if (cwp->message == WM_NCHITTEST || cwp->message == WM_SETCURSOR) {
        WCHAR className[32];
        if (GetClassName(GetAncestor(cwp->hwnd, GA_ROOT), className,
                         ARRAYSIZE(className)) &&
            (wcsicmp(className, L"Shell_TrayWnd") == 0 ||
             wcsicmp(className, L"Shell_SecondaryTrayWnd") == 0)) {
            // Skip the taskbar.
        } else if (!g_uninitializing && GetKeyState(VK_MENU) < 0) {
            WNDPROC oldWndProc = (WNDPROC)SetWindowLongPtr(
                cwp->hwnd, GWLP_WNDPROC, (LONG_PTR)CustomWndProc);
            SetProp(cwp->hwnd, kOriginalWndProcProp, (HANDLE)oldWndProc);
        }
    }

    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

void SetWindowHookForUiThreadIfNeeded(HWND hWnd) {
    if (!g_callWndProcHook && IsWindowVisible(GetAncestor(hWnd, GA_ROOT))) {
        std::lock_guard<std::mutex> guard(g_allCallWndProcHooksMutex);
        if (!g_uninitializing) {
            DWORD dwThreadId = GetCurrentThreadId();
            HHOOK callWndProcHook = SetWindowsHookEx(
                WH_CALLWNDPROC, CallWndProc, nullptr, dwThreadId);
            if (callWndProcHook) {
                Wh_Log(L"SetWindowsHookEx succeeded for thread %u", dwThreadId);
                g_callWndProcHook = callWndProcHook;
                g_allCallWndProcHooks.insert(callWndProcHook);
            } else {
                Wh_Log(L"SetWindowsHookEx error for thread %u: %u", dwThreadId,
                       GetLastError());
            }
        }
    }
}

using DispatchMessageA_t = decltype(&DispatchMessageA);
DispatchMessageA_t pOriginalDispatchMessageA;
LRESULT WINAPI DispatchMessageAHook(CONST MSG* lpMsg) {
    auto hookScope = HookRefCountScope();

    if (lpMsg && lpMsg->hwnd) {
        SetWindowHookForUiThreadIfNeeded(lpMsg->hwnd);
    }

    return pOriginalDispatchMessageA(lpMsg);
}

using DispatchMessageW_t = decltype(&DispatchMessageW);
DispatchMessageW_t pOriginalDispatchMessageW;
LRESULT WINAPI DispatchMessageWHook(CONST MSG* lpMsg) {
    auto hookScope = HookRefCountScope();

    if (lpMsg && lpMsg->hwnd) {
        SetWindowHookForUiThreadIfNeeded(lpMsg->hwnd);
    }

    return pOriginalDispatchMessageW(lpMsg);
}

using IsDialogMessageA_t = decltype(&IsDialogMessageA);
IsDialogMessageA_t pOriginalIsDialogMessageA;
LRESULT WINAPI IsDialogMessageAHook(HWND hDlg, LPMSG lpMsg) {
    auto hookScope = HookRefCountScope();

    if (hDlg) {
        SetWindowHookForUiThreadIfNeeded(hDlg);
    }

    return pOriginalIsDialogMessageA(hDlg, lpMsg);
}

using IsDialogMessageW_t = decltype(&IsDialogMessageW);
IsDialogMessageW_t pOriginalIsDialogMessageW;
LRESULT WINAPI IsDialogMessageWHook(HWND hDlg, LPMSG lpMsg) {
    auto hookScope = HookRefCountScope();

    if (hDlg) {
        SetWindowHookForUiThreadIfNeeded(hDlg);
    }

    return pOriginalIsDialogMessageW(hDlg, lpMsg);
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    switch (fdwReason) {
        case DLL_PROCESS_ATTACH:
            break;

        case DLL_THREAD_ATTACH:
            break;

        case DLL_THREAD_DETACH:
            if (g_callWndProcHook) {
                std::lock_guard<std::mutex> guard(g_allCallWndProcHooksMutex);

                auto it = g_allCallWndProcHooks.find(g_callWndProcHook);
                if (it != g_allCallWndProcHooks.end()) {
                    UnhookWindowsHookEx(g_callWndProcHook);
                    g_allCallWndProcHooks.erase(it);
                }
            }
            break;

        case DLL_PROCESS_DETACH:
            break;
    }

    return TRUE;
}

void LoadSettings() {
    // None for now.
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    // DispatchMessageA, DispatchMessageW could hopefully be enough to detect a
    // message loop, but DispatchMessageWorker, which implements
    // DispatchMessageA, DispatchMessageW, is sometimes called directly by
    // functions such as DialogBoxParam.
    Wh_SetFunctionHook((void*)DispatchMessageA, (void*)DispatchMessageAHook,
                       (void**)&pOriginalDispatchMessageA);
    Wh_SetFunctionHook((void*)DispatchMessageW, (void*)DispatchMessageWHook,
                       (void**)&pOriginalDispatchMessageW);
    Wh_SetFunctionHook((void*)IsDialogMessageA, (void*)IsDialogMessageAHook,
                       (void**)&pOriginalIsDialogMessageA);
    Wh_SetFunctionHook((void*)IsDialogMessageW, (void*)IsDialogMessageWHook,
                       (void**)&pOriginalIsDialogMessageW);

    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L">");

    g_uninitializing = true;

    {
        std::lock_guard<std::mutex> guard(g_allCallWndProcHooksMutex);

        for (HHOOK hook : g_allCallWndProcHooks) {
            UnhookWindowsHookEx(hook);
        }

        g_allCallWndProcHooks.clear();
    }

    while (g_hookRefCount > 0) {
        Sleep(200);
    }
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}
