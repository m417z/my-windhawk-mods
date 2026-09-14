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
// @compilerOptions -lcomctl32
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
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- dragWindowsWithoutTitleBar: false
  $name: Drag windows without a title bar
  $description: >-
    Also drag windows without a title bar, such as popup menus, tooltips
    and flyouts
*/
// ==/WindhawkModSettings==

// The drag is started by taking the button press away from the target window
// as it's retrieved from the message queue, and posting a move request to the
// root window, which the mod instance in the root window's thread turns into
// the system command a title bar drag generates. Intercepting at retrieval time
// means the window procedure never sees the press, so it doesn't matter how the
// program handles mouse input, and the system move loop provides snapping, DWM
// animations and maximized window handling the same way a title bar drag does.
// The root window may belong to another process, e.g. ApplicationFrameHost.exe
// hosting a UWP app's CoreWindow, so the request is a registered message which
// is allowed through the UIPI message filter.
//
// The interception is done with a WH_GETMESSAGE hook per thread. Threads which
// retrieve messages are discovered by hooking the win32u syscall stubs, which
// every message loop goes through, including user32's internal modal loops such
// as DialogBox2 which bypass the exported PeekMessage. The stub hooks only
// install the thread hooks and then tail call the original function, so no mod
// frame is left on the stack while a thread waits for a message, which would
// otherwise prevent unloading the mod.
//
// Pointer input (WM_POINTER*, used by XAML and other mouse-in-pointer windows)
// is delivered straight to the window procedure and is never seen by
// WH_GETMESSAGE or WH_CALLWNDPROC hooks, so it's intercepted by subclassing.
// A WH_CALLWNDPROC hook subclasses the window under the pointer while Alt is
// held, and the subclass routes an Alt-initiated contact to DefWindowProc,
// which promotes it to the legacy mouse messages the retrieval hook handles.
//
// Content hosted in a composition input sink, such as a WinUI XAML island,
// receives its pointer input over a side channel and produces no window message
// at all, so neither hook nor subclass sees the press, and it reaches the
// island, which turns the release into a click. Only input which hasn't been
// routed yet can be taken away from it, so while Alt is held a low level mouse
// hook swallows the press over composition hosted content, recognized by the
// class of the hosting window, and moves the window itself with SetWindowPos.
// The cursor shown over such content is chosen by the content, on every move it
// sees, from a thread of its own, so nothing set from outside sticks. For the
// duration of the drag an invisible topmost window of the hook's thread covers
// the screen instead: it receives the moves, and with them the right to choose
// the cursor.
// Windows which deliver the press as a message are left to the paths above and
// keep the system move loop, which is of no use here: it retrieves no mouse
// input while the sink owns the contact, so it starts and then tracks nothing.
// The hook exists only while Alt is held, keeping it out of the input path the
// rest of the time.
//
// The hook is global, and the island of a window which isn't focused belongs to
// a process which never saw Alt go down and so has no hook of its own. The
// process which does hold the hook therefore handles any window, not just its
// own, and moves it asynchronously since it doesn't own it.
//
// A swallowed press never reaches the input queue, so as far as the system is
// concerned Alt was tapped on its own, and DefWindowProc turns the release into
// SC_KEYMENU, activating the menu bar. The Alt release which ends such a drag
// is therefore taken as well.

#include <commctrl.h>

#include <atomic>
#include <cstdlib>
#include <memory>
#include <mutex>
#include <unordered_set>

struct {
    bool dragWindowsWithoutTitleBar;
} g_settings;

std::atomic<bool> g_uninitializing;
std::atomic<int> g_hookRefCount;

UINT g_moveRequestMessage =
    RegisterWindowMessage(L"Windhawk_MoveRequest_" WH_MOD_ID);
UINT g_unsubclassRegisteredMessage =
    RegisterWindowMessage(L"Windhawk_Unsubclass_" WH_MOD_ID);

thread_local bool g_threadHooksAttempted;
thread_local HHOOK g_getMessageHook;
thread_local HHOOK g_callWndProcHook;
std::mutex g_allThreadHooksMutex;
std::unordered_set<HHOOK> g_allThreadHooks;

thread_local HWND g_lastSubclassedWnd;
std::mutex g_subclassedWindowsMutex;
std::unordered_set<HWND> g_subclassedWindows;

// The pointer contact being routed to DefWindowProc, if any.
thread_local HWND g_contactWnd;
thread_local UINT g_contactPointerId;

// The low level mouse hook, installed only while Alt is held. Its state is
// touched only on the thread which installed it.
HHOOK g_lowLevelMouseHook;
DWORD g_lowLevelMouseHookThreadId;
// Set when a press was swallowed during the current Alt hold.
std::atomic<bool> g_swallowedPress;

bool g_llCandidate;
HWND g_llCandidateRoot;
POINT g_llDownPt;
HWND g_llDragRoot;
POINT g_llDragGrab;
HWND g_llDragOverlayWnd;

constexpr WCHAR kDragOverlayClassName[] = L"Windhawk_AltDragOverlay_" WH_MOD_ID;

void InstallLowLevelMouseHookIfNeeded();
void RemoveLowLevelMouseHookIfIdle();

auto HookRefCountScope() {
    g_hookRefCount++;
    return std::unique_ptr<decltype(g_hookRefCount),
                           void (*)(decltype(g_hookRefCount)*)>{
        &g_hookRefCount, [](auto hookRefCount) { (*hookRefCount)--; }};
}

bool IsExcludedRootWindow(HWND hRootWnd) {
    WCHAR className[64];
    if (!GetClassName(hRootWnd, className, ARRAYSIZE(className))) {
        return true;
    }

    // The taskbar and the desktop.
    if (_wcsicmp(className, L"Shell_TrayWnd") == 0 ||
        _wcsicmp(className, L"Shell_SecondaryTrayWnd") == 0 ||
        _wcsicmp(className, L"Progman") == 0 ||
        _wcsicmp(className, L"WorkerW") == 0) {
        return true;
    }

    if (!g_settings.dragWindowsWithoutTitleBar) {
        LONG style = GetWindowLong(hRootWnd, GWL_STYLE);
        if ((style & WS_CAPTION) != WS_CAPTION) {
            return true;
        }
    }

    return false;
}

// Windows which host their content in a composition input sink, such as a WinUI
// XAML island. Their pointer input never becomes a window message.
bool IsCompositionHostedWindow(HWND hWnd) {
    WCHAR className[64];
    if (!GetClassName(hWnd, className, ARRAYSIZE(className))) {
        return false;
    }

    return wcsncmp(className, L"Microsoft.UI.Content.",
                   ARRAYSIZE(L"Microsoft.UI.Content.") - 1) == 0 ||
           wcsncmp(className, L"Windows.UI.Composition.",
                   ARRAYSIZE(L"Windows.UI.Composition.") - 1) == 0 ||
           wcsncmp(className, L"Windows.UI.Input.InputSite.",
                   ARRAYSIZE(L"Windows.UI.Input.InputSite.") - 1) == 0 ||
           _wcsicmp(className, L"InputSiteWindowClass") == 0;
}

// An island host commonly answers WM_NCHITTEST with HTTRANSPARENT so that the
// window hosting it can do its own hit testing, and WindowFromPoint then
// reports that host rather than the island. Walking the children by geometry
// finds it either way.
HWND CompositionHostedWindowFromPoint(POINT pt) {
    HWND hWnd = WindowFromPoint(pt);
    if (!hWnd) {
        return nullptr;
    }

    if (IsCompositionHostedWindow(hWnd)) {
        return hWnd;
    }

    for (int depth = 0; depth < 8; depth++) {
        POINT clientPt = pt;
        if (!ScreenToClient(hWnd, &clientPt)) {
            break;
        }

        HWND hChildWnd = ChildWindowFromPointEx(
            hWnd, clientPt, CWP_SKIPINVISIBLE | CWP_SKIPDISABLED);
        if (!hChildWnd || hChildWnd == hWnd) {
            break;
        }

        hWnd = hChildWnd;
        if (IsCompositionHostedWindow(hWnd)) {
            return hWnd;
        }
    }

    return nullptr;
}

// Alt+click on the caption buttons keeps its regular meaning.
bool IsCaptionButtonHitTest(int hitTest) {
    switch (hitTest) {
        case HTCLOSE:
        case HTMINBUTTON:
        case HTMAXBUTTON:
        case HTHELP:
        case HTSYSMENU:
            return true;
    }

    return false;
}

bool IsPointerMessage(UINT message) {
    return message >= WM_NCPOINTERUPDATE && message <= WM_POINTERROUTEDRELEASED;
}

bool IsPrimaryButtonDownMessage(const MSG* msg) {
    switch (msg->message) {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
            return true;

        case WM_NCLBUTTONDOWN:
        case WM_NCLBUTTONDBLCLK:
            return !IsCaptionButtonHitTest((int)msg->wParam);
    }

    return false;
}

bool IsPrimaryPointerDownMessage(UINT message, WPARAM wParam) {
    switch (message) {
        case WM_POINTERDOWN:
            return IS_POINTER_FIRSTBUTTON_WPARAM(wParam);

        case WM_NCPOINTERDOWN:
            return IS_POINTER_FIRSTBUTTON_WPARAM(wParam) &&
                   !IsCaptionButtonHitTest(HIWORD(wParam));
    }

    return false;
}

bool IsPointerContactEndMessage(UINT message) {
    switch (message) {
        case WM_POINTERUP:
        case WM_NCPOINTERUP:
        case WM_POINTERCAPTURECHANGED:
            return true;
    }

    return false;
}

// Whether the system move loop is running on this thread.
bool IsInMoveLoop() {
    GUITHREADINFO gti{.cbSize = sizeof(gti)};
    return GetGUIThreadInfo(GetCurrentThreadId(), &gti) &&
           (gti.flags & GUI_INMOVESIZE);
}

// Turns a move request retrieved by the root window's thread into the system
// command. The move loop only starts if this thread's synchronized button state
// is down, which it isn't when the press was retrieved by another thread.
void OnMoveRequestRemoved(MSG* msg) {
    if (GetAsyncKeyState(VK_LBUTTON) >= 0) {
        // Released already, a move loop would stick to the cursor.
        Wh_Log(L"Move request for %08X, button already released",
               (DWORD)(ULONG_PTR)msg->hwnd);
        msg->message = WM_NULL;
        msg->wParam = 0;
        msg->lParam = 0;
        return;
    }

    if (GetKeyState(VK_LBUTTON) >= 0) {
        BYTE keyState[256];
        if (GetKeyboardState(keyState)) {
            keyState[VK_LBUTTON] |= 0x80;
            SetKeyboardState(keyState);
        }
    }

    Wh_Log(L"Move request for %08X, starting the move loop",
           (DWORD)(ULONG_PTR)msg->hwnd);

    msg->message = WM_SYSCOMMAND;
    msg->wParam = SC_MOVE | HTCAPTION;
}

// Runs for every message removed from the queue of the current thread, before
// the program sees it.
void OnMessageRemoved(MSG* msg) {
    if (!msg->hwnd) {
        return;
    }

    if (msg->message == g_moveRequestMessage) {
        OnMoveRequestRemoved(msg);
        return;
    }

    // Track Alt so the low level mouse hook exists only while it's held.
    if (msg->message == WM_SYSKEYDOWN || msg->message == WM_KEYDOWN) {
        if (msg->wParam == VK_MENU) {
            // Bit 30 of lParam is set for the auto repeats which arrive while
            // the key is held, including throughout a drag.
            constexpr LPARAM kPreviousKeyStateDown = 1 << 30;
            if (!(msg->lParam & kPreviousKeyStateDown)) {
                g_swallowedPress = false;
            }

            InstallLowLevelMouseHookIfNeeded();
        }
        return;
    }

    if (msg->message == WM_SYSKEYUP || msg->message == WM_KEYUP) {
        if (msg->wParam == VK_MENU) {
            RemoveLowLevelMouseHookIfIdle();

            if (g_swallowedPress.exchange(false)) {
                Wh_Log(L"Swallowing the Alt release of a drag");
                msg->message = WM_NULL;
                msg->wParam = 0;
                msg->lParam = 0;
            }
        }
        return;
    }

    if (msg->message != WM_MOUSEMOVE && !IsPrimaryButtonDownMessage(msg)) {
        return;
    }

    // GetKeyState reflects the state at the time of the retrieved message.
    if (GetKeyState(VK_MENU) >= 0) {
        return;
    }

    if (msg->message == WM_MOUSEMOVE) {
        // The move loop retrieves the moves itself, without dispatching them,
        // holds an internal capture for which no WM_SETCURSOR is sent, and
        // shows the class cursor.
        if (IsInMoveLoop()) {
            SetCursor(LoadCursor(nullptr, IDC_SIZEALL));
        }
        return;
    }

    HWND hRootWnd = GetAncestor(msg->hwnd, GA_ROOT);
    if (!hRootWnd || IsExcludedRootWindow(hRootWnd)) {
        return;
    }

    Wh_Log(L"Message %04X for %08X, requesting a move of root window %08X",
           msg->message, (DWORD)(ULONG_PTR)msg->hwnd,
           (DWORD)(ULONG_PTR)hRootWnd);

    // Posted messages are retrieved before input, so a button release that's
    // already queued is seen by the move loop rather than by the program.
    if (!PostMessage(hRootWnd, g_moveRequestMessage, 0,
                     MAKELPARAM(msg->pt.x, msg->pt.y))) {
        Wh_Log(L"PostMessage error: %u", GetLastError());
        return;
    }

    msg->message = WM_NULL;
    msg->wParam = 0;
    msg->lParam = 0;
}

LRESULT CALLBACK GetMsgProc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto hookScope = HookRefCountScope();

    if (nCode == HC_ACTION && (wParam & PM_REMOVE) && lParam &&
        !g_uninitializing) {
        OnMessageRemoved((MSG*)lParam);
    }

    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

void UnsubclassWindow(HWND hWnd);

LRESULT CALLBACK SubclassProc(HWND hWnd,
                              UINT uMsg,
                              WPARAM wParam,
                              LPARAM lParam,
                              UINT_PTR uIdSubclass,
                              DWORD_PTR dwRefData) {
    auto hookScope = HookRefCountScope();

    if (uMsg == WM_NCDESTROY) {
        if (hWnd == g_contactWnd) {
            g_contactWnd = nullptr;
        }

        UnsubclassWindow(hWnd);
    } else if (uMsg == g_unsubclassRegisteredMessage) {
        UnsubclassWindow(hWnd);
        return 0;
    } else if (IsPointerMessage(uMsg)) {
        UINT pointerId = GET_POINTERID_WPARAM(wParam);
        bool tracked = hWnd == g_contactWnd && pointerId == g_contactPointerId;

        if (IsPrimaryPointerDownMessage(uMsg, wParam)) {
            // A new contact, the mouse reuses its pointer id for every one.
            if (tracked) {
                g_contactWnd = nullptr;
            }

            if (GetAsyncKeyState(VK_MENU) < 0 && !g_uninitializing) {
                Wh_Log(
                    L"Message %04X for %08X, routing pointer %u to "
                    L"DefWindowProc",
                    uMsg, (DWORD)(ULONG_PTR)hWnd, pointerId);
                g_contactWnd = hWnd;
                g_contactPointerId = pointerId;
                return DefWindowProc(hWnd, uMsg, wParam, lParam);
            }
        } else if (tracked) {
            if (IsPointerContactEndMessage(uMsg)) {
                g_contactWnd = nullptr;
            }

            return DefWindowProc(hWnd, uMsg, wParam, lParam);
        }
    }

    LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);

    switch (uMsg) {
        case WM_MOUSEMOVE:
        case WM_NCMOUSEMOVE:
        case WM_POINTERUPDATE:
        case WM_NCPOINTERUPDATE:
            // Overrides the cursor the program chose for the move.
            if (GetAsyncKeyState(VK_MENU) < 0 && !g_uninitializing) {
                SetCursor(LoadCursor(nullptr, IDC_SIZEALL));
            }
            break;
    }

    return result;
}

void UnsubclassWindow(HWND hWnd) {
    RemoveWindowSubclass(hWnd, SubclassProc, 0);

    if (hWnd == g_lastSubclassedWnd) {
        g_lastSubclassedWnd = nullptr;
    }

    std::lock_guard<std::mutex> guard(g_subclassedWindowsMutex);
    g_subclassedWindows.erase(hWnd);
}

void SubclassWindowIfNeeded(HWND hWnd) {
    if (hWnd == g_lastSubclassedWnd) {
        return;
    }

    HWND hRootWnd = GetAncestor(hWnd, GA_ROOT);
    if (!hRootWnd || IsExcludedRootWindow(hRootWnd)) {
        return;
    }

    std::lock_guard<std::mutex> guard(g_subclassedWindowsMutex);
    if (g_uninitializing) {
        return;
    }

    if (!g_subclassedWindows.contains(hWnd)) {
        if (!SetWindowSubclass(hWnd, SubclassProc, 0, 0)) {
            Wh_Log(L"SetWindowSubclass error for %08X", (DWORD)(ULONG_PTR)hWnd);
            return;
        }

        g_subclassedWindows.insert(hWnd);
    }

    g_lastSubclassedWnd = hWnd;
}

LRESULT CALLBACK CallWndProc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto hookScope = HookRefCountScope();

    if (nCode == HC_ACTION && lParam) {
        const CWPSTRUCT* cwp = (const CWPSTRUCT*)lParam;
        // The thread's synchronized key state is stale if the window isn't
        // active yet, e.g. an Alt+click on a background window.
        if (cwp->message == WM_NCHITTEST && GetAsyncKeyState(VK_MENU) < 0) {
            SubclassWindowIfNeeded(cwp->hwnd);
        }
    }

    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

// Where the window is held relative to its origin. A maximized window is
// restored first, like a title bar drag does, keeping the grab proportional.
POINT CalcDragGrab(HWND hRootWnd, POINT ptDown) {
    RECT rc;
    GetWindowRect(hRootWnd, &rc);

    POINT grab;
    if (IsZoomed(hRootWnd)) {
        int width = rc.right - rc.left;
        int height = rc.bottom - rc.top;
        double fractionX =
            width > 0 ? (double)(ptDown.x - rc.left) / width : 0.5;
        double fractionY =
            height > 0 ? (double)(ptDown.y - rc.top) / height : 0.0;

        ShowWindow(hRootWnd, SW_RESTORE);

        RECT restored;
        GetWindowRect(hRootWnd, &restored);
        grab.x = (LONG)(fractionX * (restored.right - restored.left));
        grab.y = (LONG)(fractionY * (restored.bottom - restored.top));
    } else {
        grab.x = ptDown.x - rc.left;
        grab.y = ptDown.y - rc.top;
    }

    return grab;
}

void MoveDraggedWindow(HWND hRootWnd, POINT grab, POINT pt) {
    // SWP_ASYNCWINDOWPOS posts the request when the window belongs to another
    // thread, which keeps a busy owner from blocking the caller.
    SetWindowPos(
        hRootWnd, nullptr, pt.x - grab.x, pt.y - grab.y, 0, 0,
        SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
}

// The class supplies the size cursor, and with DefWindowProc as the window
// procedure no mod code is on the window's call path.
HWND CreateDragOverlay() {
    WNDCLASS wc{
        .lpfnWndProc = DefWindowProc,
        .hInstance = GetModuleHandle(nullptr),
        .hCursor = LoadCursor(nullptr, IDC_SIZEALL),
        .lpszClassName = kDragOverlayClassName,
    };
    if (!RegisterClass(&wc) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        Wh_Log(L"RegisterClass error: %u", GetLastError());
        return nullptr;
    }

    HWND hWnd = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED,
        kDragOverlayClassName, nullptr, WS_POPUP,
        GetSystemMetrics(SM_XVIRTUALSCREEN),
        GetSystemMetrics(SM_YVIRTUALSCREEN),
        GetSystemMetrics(SM_CXVIRTUALSCREEN),
        GetSystemMetrics(SM_CYVIRTUALSCREEN), nullptr, nullptr, wc.hInstance,
        nullptr);
    if (!hWnd) {
        Wh_Log(L"CreateWindowEx error: %u", GetLastError());
        return nullptr;
    }

    // As good as invisible, while an alpha of zero would let the mouse through.
    SetLayeredWindowAttributes(hWnd, 0, 1, LWA_ALPHA);
    ShowWindow(hWnd, SW_SHOWNOACTIVATE);
    return hWnd;
}

void DestroyDragOverlay() {
    if (g_llDragOverlayWnd) {
        DestroyWindow(g_llDragOverlayWnd);
        g_llDragOverlayWnd = nullptr;
    }
}

// Runs for mouse input before it's routed anywhere, which is the only point at
// which a press can be taken away from composition hosted content. Moves are
// never swallowed: that would stop the cursor.
LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    auto hookScope = HookRefCountScope();

    if (nCode != HC_ACTION || g_uninitializing) {
        return CallNextHookEx(nullptr, nCode, wParam, lParam);
    }

    const MSLLHOOKSTRUCT* ms = (const MSLLHOOKSTRUCT*)lParam;

    // Focus can move away while Alt is held, e.g. Alt+Tab, and the Alt release
    // then goes to another process. Drop the hook here instead of keeping it
    // for the rest of the process's life.
    RemoveLowLevelMouseHookIfIdle();

    switch (wParam) {
        case WM_LBUTTONDOWN: {
            g_llCandidate = false;
            g_llDragRoot = nullptr;

            if (GetAsyncKeyState(VK_MENU) >= 0) {
                break;
            }

            HWND hWnd = CompositionHostedWindowFromPoint(ms->pt);
            if (!hWnd) {
                // Delivered as a message, so the paths above handle it and keep
                // the system move loop.
                break;
            }

            HWND hRootWnd = GetAncestor(hWnd, GA_ROOT);
            if (!hRootWnd || IsExcludedRootWindow(hRootWnd)) {
                break;
            }

            Wh_Log(L"Swallowed the press over %08X, root window %08X",
                   (DWORD)(ULONG_PTR)hWnd, (DWORD)(ULONG_PTR)hRootWnd);

            g_llCandidate = true;
            g_llCandidateRoot = hRootWnd;
            g_llDownPt = ms->pt;
            g_swallowedPress = true;
            return 1;
        }

        case WM_LBUTTONUP:
            if (g_llCandidate || g_llDragRoot) {
                g_llCandidate = false;
                g_llDragRoot = nullptr;
                DestroyDragOverlay();
                return 1;
            }
            break;

        case WM_MOUSEMOVE:
            if (g_llDragRoot) {
                MoveDraggedWindow(g_llDragRoot, g_llDragGrab, ms->pt);
            } else if (g_llCandidate && (abs(ms->pt.x - g_llDownPt.x) >=
                                             GetSystemMetrics(SM_CXDRAG) ||
                                         abs(ms->pt.y - g_llDownPt.y) >=
                                             GetSystemMetrics(SM_CYDRAG))) {
                Wh_Log(L"Moving root window %08X",
                       (DWORD)(ULONG_PTR)g_llCandidateRoot);
                g_llCandidate = false;
                g_llDragGrab = CalcDragGrab(g_llCandidateRoot, g_llDownPt);
                g_llDragRoot = g_llCandidateRoot;

                // The press was swallowed, so the window wasn't brought to the
                // front the way a click on it would have been. Allowed because
                // the process holding the hook is the foreground one.
                if (!SetForegroundWindow(g_llDragRoot)) {
                    Wh_Log(L"SetForegroundWindow error: %u", GetLastError());
                }

                g_llDragOverlayWnd = CreateDragOverlay();
                MoveDraggedWindow(g_llDragRoot, g_llDragGrab, ms->pt);
            }
            break;
    }

    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

void InstallLowLevelMouseHookIfNeeded() {
    if (g_lowLevelMouseHook || g_uninitializing) {
        return;
    }

    g_lowLevelMouseHook = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc,
                                           GetModuleHandle(nullptr), 0);
    if (g_lowLevelMouseHook) {
        g_lowLevelMouseHookThreadId = GetCurrentThreadId();
    } else {
        Wh_Log(L"SetWindowsHookEx(WH_MOUSE_LL) error: %u", GetLastError());
    }
}

void RemoveLowLevelMouseHookIfIdle() {
    // Keep it while a drag is in progress or Alt is still held.
    if (!g_lowLevelMouseHook || g_llDragRoot || GetAsyncKeyState(VK_MENU) < 0) {
        return;
    }

    UnhookWindowsHookEx(g_lowLevelMouseHook);
    g_lowLevelMouseHook = nullptr;
    g_llCandidate = false;
}

void SetThreadHooksIfNeeded() {
    if (g_threadHooksAttempted) {
        return;
    }

    g_threadHooksAttempted = true;

    std::lock_guard<std::mutex> guard(g_allThreadHooksMutex);
    if (g_uninitializing) {
        return;
    }

    DWORD dwThreadId = GetCurrentThreadId();

    g_getMessageHook =
        SetWindowsHookEx(WH_GETMESSAGE, GetMsgProc, nullptr, dwThreadId);
    if (g_getMessageHook) {
        g_allThreadHooks.insert(g_getMessageHook);
    } else {
        Wh_Log(L"SetWindowsHookEx(WH_GETMESSAGE) error for thread %u: %u",
               dwThreadId, GetLastError());
    }

    g_callWndProcHook =
        SetWindowsHookEx(WH_CALLWNDPROC, CallWndProc, nullptr, dwThreadId);
    if (g_callWndProcHook) {
        g_allThreadHooks.insert(g_callWndProcHook);
    } else {
        Wh_Log(L"SetWindowsHookEx(WH_CALLWNDPROC) error for thread %u: %u",
               dwThreadId, GetLastError());
    }

    if (g_getMessageHook && g_callWndProcHook) {
        Wh_Log(L"SetWindowsHookEx succeeded for thread %u", dwThreadId);
    }
}

using NtUserGetMessage_t = BOOL(WINAPI*)(MSG* pMsg,
                                         HWND hWnd,
                                         UINT wMsgFilterMin,
                                         UINT wMsgFilterMax);
NtUserGetMessage_t NtUserGetMessage_Original;
BOOL WINAPI NtUserGetMessage_Hook(MSG* pMsg,
                                  HWND hWnd,
                                  UINT wMsgFilterMin,
                                  UINT wMsgFilterMax) {
    SetThreadHooksIfNeeded();

    [[clang::musttail]] return NtUserGetMessage_Original(
        pMsg, hWnd, wMsgFilterMin, wMsgFilterMax);
}

// The last parameter is a flags value which user32 sets from the calling
// PeekMessage variant.
using NtUserPeekMessage_t = BOOL(WINAPI*)(MSG* pMsg,
                                          HWND hWnd,
                                          UINT wMsgFilterMin,
                                          UINT wMsgFilterMax,
                                          UINT wRemoveMsg,
                                          UINT flags);
NtUserPeekMessage_t NtUserPeekMessage_Original;
BOOL WINAPI NtUserPeekMessage_Hook(MSG* pMsg,
                                   HWND hWnd,
                                   UINT wMsgFilterMin,
                                   UINT wMsgFilterMax,
                                   UINT wRemoveMsg,
                                   UINT flags) {
    SetThreadHooksIfNeeded();

    [[clang::musttail]] return NtUserPeekMessage_Original(
        pMsg, hWnd, wMsgFilterMin, wMsgFilterMax, wRemoveMsg, flags);
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    switch (fdwReason) {
        case DLL_PROCESS_ATTACH:
            break;

        case DLL_THREAD_ATTACH:
            break;

        case DLL_THREAD_DETACH:
            if (g_getMessageHook || g_callWndProcHook) {
                std::lock_guard<std::mutex> guard(g_allThreadHooksMutex);

                for (HHOOK hook : {g_getMessageHook, g_callWndProcHook}) {
                    auto it = g_allThreadHooks.find(hook);
                    if (it != g_allThreadHooks.end()) {
                        UnhookWindowsHookEx(hook);
                        g_allThreadHooks.erase(it);
                    }
                }
            }

            // The low level hook and its overlay go away with their thread.
            if (g_lowLevelMouseHook &&
                GetCurrentThreadId() == g_lowLevelMouseHookThreadId) {
                g_lowLevelMouseHook = nullptr;
                g_llCandidate = false;
                g_llDragRoot = nullptr;
                g_llDragOverlayWnd = nullptr;
            }
            break;

        case DLL_PROCESS_DETACH:
            break;
    }

    return TRUE;
}

void LoadSettings() {
    g_settings.dragWindowsWithoutTitleBar =
        Wh_GetIntSetting(L"dragWindowsWithoutTitleBar");
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    HMODULE win32uModule = GetModuleHandle(L"win32u.dll");
    if (!win32uModule) {
        Wh_Log(L"win32u.dll isn't loaded");
        return FALSE;
    }

    void* pNtUserGetMessage =
        (void*)GetProcAddress(win32uModule, "NtUserGetMessage");
    void* pNtUserPeekMessage =
        (void*)GetProcAddress(win32uModule, "NtUserPeekMessage");
    if (!pNtUserGetMessage || !pNtUserPeekMessage) {
        Wh_Log(L"NtUserGetMessage or NtUserPeekMessage not found");
        return FALSE;
    }

    Wh_SetFunctionHook(pNtUserGetMessage, (void*)NtUserGetMessage_Hook,
                       (void**)&NtUserGetMessage_Original);
    Wh_SetFunctionHook(pNtUserPeekMessage, (void*)NtUserPeekMessage_Hook,
                       (void**)&NtUserPeekMessage_Original);

    // Lets a lower integrity process, e.g. a UWP app, request a move of a root
    // window in this process.
    ChangeWindowMessageFilter(g_moveRequestMessage, MSGFLT_ADD);

    return TRUE;
}

void Wh_ModUninit() {
    Wh_Log(L">");

    g_uninitializing = true;

    std::unordered_set<HWND> subclassedWindows;
    {
        std::lock_guard<std::mutex> guard(g_subclassedWindowsMutex);
        subclassedWindows = std::move(g_subclassedWindows);
        g_subclassedWindows.clear();
    }

    for (HWND hWnd : subclassedWindows) {
        SendMessage(hWnd, g_unsubclassRegisteredMessage, 0, 0);
    }

    {
        std::lock_guard<std::mutex> guard(g_allThreadHooksMutex);

        for (HHOOK hook : g_allThreadHooks) {
            UnhookWindowsHookEx(hook);
        }

        g_allThreadHooks.clear();
    }

    ChangeWindowMessageFilter(g_moveRequestMessage, MSGFLT_REMOVE);

    if (g_lowLevelMouseHook) {
        UnhookWindowsHookEx(g_lowLevelMouseHook);
        g_lowLevelMouseHook = nullptr;
    }

    // Destroyed on its thread by DefWindowProc. The class may outlive it, and
    // is then found in place next time.
    if (HWND hOverlayWnd = g_llDragOverlayWnd) {
        PostMessage(hOverlayWnd, WM_CLOSE, 0, 0);
    }

    UnregisterClass(kDragOverlayClassName, GetModuleHandle(nullptr));

    while (g_hookRefCount > 0) {
        Sleep(200);
    }
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}
