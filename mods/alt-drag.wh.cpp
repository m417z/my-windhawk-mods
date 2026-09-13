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
// at all, so neither hook nor subclass sees the press. As a last resort a raw
// input sink observes the Alt press directly and the window is moved by
// SetWindowPos, without the system move loop. To avoid a system-wide stream of
// WM_INPUT in every hooked process, the raw device is registered only while Alt
// is held, and the fallback yields to the message-based paths for any window
// which does deliver the press as a message.
//
// Such a press still reaches the island, which turns the release into a click.
// Only input which hasn't been routed yet can be taken away from it, so while
// Alt is held a low level mouse hook swallows the press over composition hosted
// content of this process and moves the window itself. Windows which deliver
// the press as a message are left to the paths above and keep the system move
// loop. The move loop is of no use here: it retrieves no mouse input while the
// sink owns the contact, so it starts and then tracks nothing.
//
// A swallowed press never reaches the input queue, so as far as the system is
// concerned Alt was tapped on its own, and DefWindowProc turns the release into
// SC_KEYMENU, activating the menu bar. The Alt release which ends such a drag
// is therefore taken as well.

#include <commctrl.h>
#include <windowsx.h>

#include <atomic>
#include <cstdlib>
#include <memory>
#include <mutex>
#include <unordered_set>
#include <vector>

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

// The raw input fallback for composition-hosted content. The sink window and
// its raw device registration are process-wide; the drag state is touched only
// on the sink window's thread.
std::mutex g_rawSinkMutex;
HWND g_rawSinkWnd;
DWORD g_rawSinkThreadId;
bool g_rawSinkCreateFailed;
std::atomic<bool> g_rawInputRegistered;

bool g_rawCandidate;
HWND g_rawCandidateRoot;
POINT g_rawDownPt;
HWND g_rawDragRoot;
POINT g_rawDragGrab;

// The location and time of the last press claimed by the message-based paths,
// used to tell whether a raw press was already handled by them.
std::mutex g_lastHandledPressMutex;
POINT g_lastHandledPressPt;
DWORD g_lastHandledPressTick;

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

void RegisterRawInputIfNeeded();
void UnregisterRawInputIfIdle();
void InstallLowLevelMouseHookIfNeeded();
void RemoveLowLevelMouseHookIfIdle();

auto HookRefCountScope() {
    g_hookRefCount++;
    return std::unique_ptr<decltype(g_hookRefCount),
                           void (*)(decltype(g_hookRefCount)*)>{
        &g_hookRefCount, [](auto hookRefCount) { (*hookRefCount)--; }};
}

void NoteHandledPress(POINT pt) {
    std::lock_guard<std::mutex> guard(g_lastHandledPressMutex);
    g_lastHandledPressPt = pt;
    g_lastHandledPressTick = GetTickCount();
}

bool WasPressAlreadyHandled(POINT pt) {
    std::lock_guard<std::mutex> guard(g_lastHandledPressMutex);
    if (!g_lastHandledPressTick) {
        return false;
    }

    int tolerance = GetSystemMetrics(SM_CXDRAG) + 4;
    return GetTickCount() - g_lastHandledPressTick < 2000 &&
           abs(pt.x - g_lastHandledPressPt.x) <= tolerance &&
           abs(pt.y - g_lastHandledPressPt.y) <= tolerance;
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

    // Track Alt so the raw input sink is registered only while it's held.
    if (msg->message == WM_SYSKEYDOWN || msg->message == WM_KEYDOWN) {
        if (msg->wParam == VK_MENU) {
            // Bit 30 of lParam is set for the auto repeats which arrive while
            // the key is held, including throughout a drag.
            constexpr LPARAM kPreviousKeyStateDown = 1 << 30;
            if (!(msg->lParam & kPreviousKeyStateDown)) {
                g_swallowedPress = false;
            }

            RegisterRawInputIfNeeded();
            InstallLowLevelMouseHookIfNeeded();
        }
        return;
    }

    if (msg->message == WM_SYSKEYUP || msg->message == WM_KEYUP) {
        if (msg->wParam == VK_MENU) {
            UnregisterRawInputIfIdle();
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

    if (!IsPrimaryButtonDownMessage(msg)) {
        return;
    }

    // GetKeyState reflects the state at the time of the retrieved message.
    if (GetKeyState(VK_MENU) >= 0) {
        return;
    }

    HWND hRootWnd = GetAncestor(msg->hwnd, GA_ROOT);
    if (!hRootWnd || IsExcludedRootWindow(hRootWnd)) {
        return;
    }

    Wh_Log(L"Message %04X for %08X, requesting a move of root window %08X",
           msg->message, (DWORD)(ULONG_PTR)msg->hwnd,
           (DWORD)(ULONG_PTR)hRootWnd);

    NoteHandledPress(msg->pt);

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
                NoteHandledPress(
                    POINT{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)});
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
    SetWindowPos(hRootWnd, nullptr, pt.x - grab.x, pt.y - grab.y, 0, 0,
                 SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

void OnRawMouseInput(const RAWMOUSE& mouse) {
    USHORT buttonFlags = mouse.usButtonFlags;

    if (buttonFlags & RI_MOUSE_LEFT_BUTTON_DOWN) {
        g_rawCandidate = false;
        g_rawDragRoot = nullptr;

        POINT pt;
        if (g_uninitializing || GetAsyncKeyState(VK_MENU) >= 0 ||
            !GetCursorPos(&pt)) {
            return;
        }

        HWND hWnd = WindowFromPoint(pt);
        if (!hWnd) {
            return;
        }

        // Only windows in this process can be moved from here, and the
        // message-based paths handle same-process windows which deliver the
        // press as a message.
        DWORD dwProcessId = 0;
        GetWindowThreadProcessId(hWnd, &dwProcessId);
        if (dwProcessId != GetCurrentProcessId()) {
            return;
        }

        HWND hRootWnd = GetAncestor(hWnd, GA_ROOT);
        if (!hRootWnd || IsExcludedRootWindow(hRootWnd)) {
            return;
        }

        g_rawCandidate = true;
        g_rawCandidateRoot = hRootWnd;
        g_rawDownPt = pt;
        return;
    }

    if (buttonFlags & RI_MOUSE_LEFT_BUTTON_UP) {
        g_rawCandidate = false;
        g_rawDragRoot = nullptr;
        RemoveLowLevelMouseHookIfIdle();
        UnregisterRawInputIfIdle();
        return;
    }

    if (g_rawDragRoot) {
        POINT pt;
        if (GetCursorPos(&pt)) {
            MoveDraggedWindow(g_rawDragRoot, g_rawDragGrab, pt);
        }
        return;
    }

    if (!g_rawCandidate || g_uninitializing) {
        return;
    }

    // A real drag keeps the button down. Moves buffered while a system move
    // loop ran arrive after the button was released.
    if (GetAsyncKeyState(VK_LBUTTON) >= 0) {
        g_rawCandidate = false;
        return;
    }

    if (WasPressAlreadyHandled(g_rawDownPt)) {
        g_rawCandidate = false;
        return;
    }

    POINT pt;
    if (!GetCursorPos(&pt)) {
        return;
    }

    if (abs(pt.x - g_rawDownPt.x) < GetSystemMetrics(SM_CXDRAG) &&
        abs(pt.y - g_rawDownPt.y) < GetSystemMetrics(SM_CYDRAG)) {
        return;
    }

    Wh_Log(L"Moving root window %08X by raw input",
           (DWORD)(ULONG_PTR)g_rawCandidateRoot);

    g_rawCandidate = false;
    g_rawDragGrab = CalcDragGrab(g_rawCandidateRoot, g_rawDownPt);
    g_rawDragRoot = g_rawCandidateRoot;
    MoveDraggedWindow(g_rawDragRoot, g_rawDragGrab, pt);
}

constexpr UINT kRawSinkDestroyMessage = WM_APP;

LRESULT CALLBACK RawSinkWndProc(HWND hWnd,
                                UINT uMsg,
                                WPARAM wParam,
                                LPARAM lParam) {
    auto hookScope = HookRefCountScope();

    switch (uMsg) {
        case WM_INPUT: {
            RAWINPUT raw;
            UINT size = sizeof(raw);
            if (GetRawInputData((HRAWINPUT)lParam, RID_INPUT, &raw, &size,
                                sizeof(RAWINPUTHEADER)) != (UINT)-1 &&
                raw.header.dwType == RIM_TYPEMOUSE) {
                OnRawMouseInput(raw.data.mouse);
            }
            break;
        }

        case kRawSinkDestroyMessage: {
            RAWINPUTDEVICE rid = {};
            rid.usUsagePage = 0x01;
            rid.usUsage = 0x02;
            rid.dwFlags = RIDEV_REMOVE;
            RegisterRawInputDevices(&rid, 1, sizeof(rid));
            g_rawInputRegistered = false;
            DestroyWindow(hWnd);
            return 0;
        }
    }

    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

bool ProcessAlreadyUsesRawMouse() {
    UINT count = 0;
    if (GetRegisteredRawInputDevices(nullptr, &count, sizeof(RAWINPUTDEVICE)) !=
            0 ||
        count == 0) {
        return false;
    }

    std::vector<RAWINPUTDEVICE> devices(count);
    UINT written = GetRegisteredRawInputDevices(devices.data(), &count,
                                                sizeof(RAWINPUTDEVICE));
    if (written == (UINT)-1) {
        return false;
    }

    for (UINT i = 0; i < written; i++) {
        if (devices[i].usUsagePage == 0x01 && devices[i].usUsage == 0x02) {
            return true;
        }
    }

    return false;
}

void RegisterRawInputIfNeeded() {
    if (g_rawInputRegistered) {
        return;
    }

    std::lock_guard<std::mutex> guard(g_rawSinkMutex);
    if (g_uninitializing || g_rawInputRegistered) {
        return;
    }

    if (!g_rawSinkWnd) {
        if (g_rawSinkCreateFailed) {
            return;
        }

        // Don't override an app which uses raw mouse input itself.
        if (ProcessAlreadyUsesRawMouse()) {
            g_rawSinkCreateFailed = true;
            return;
        }

        WNDCLASS wndClass = {};
        wndClass.lpfnWndProc = RawSinkWndProc;
        wndClass.hInstance = GetModuleHandle(nullptr);
        wndClass.lpszClassName = L"Windhawk_AltDragRawSink_" WH_MOD_ID;
        RegisterClass(&wndClass);

        g_rawSinkWnd =
            CreateWindowEx(0, wndClass.lpszClassName, L"", 0, 0, 0, 0, 0,
                           HWND_MESSAGE, nullptr, wndClass.hInstance, nullptr);
        if (!g_rawSinkWnd) {
            Wh_Log(L"Raw sink window creation failed: %u", GetLastError());
            g_rawSinkCreateFailed = true;
            return;
        }

        g_rawSinkThreadId = GetCurrentThreadId();
    }

    RAWINPUTDEVICE rid = {};
    rid.usUsagePage = 0x01;
    rid.usUsage = 0x02;
    rid.dwFlags = RIDEV_INPUTSINK;
    rid.hwndTarget = g_rawSinkWnd;
    if (RegisterRawInputDevices(&rid, 1, sizeof(rid))) {
        g_rawInputRegistered = true;
    } else {
        Wh_Log(L"RegisterRawInputDevices error: %u", GetLastError());
    }
}

void UnregisterRawInputIfIdle() {
    if (!g_rawInputRegistered) {
        return;
    }

    std::lock_guard<std::mutex> guard(g_rawSinkMutex);
    // Keep it while a drag is in progress or Alt is still held.
    if (!g_rawInputRegistered || g_rawDragRoot ||
        GetAsyncKeyState(VK_MENU) < 0) {
        return;
    }

    RAWINPUTDEVICE rid = {};
    rid.usUsagePage = 0x01;
    rid.usUsage = 0x02;
    rid.dwFlags = RIDEV_REMOVE;
    RegisterRawInputDevices(&rid, 1, sizeof(rid));
    g_rawInputRegistered = false;
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

            DWORD dwProcessId = 0;
            GetWindowThreadProcessId(hWnd, &dwProcessId);
            if (dwProcessId != GetCurrentProcessId()) {
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

            // Keeps the raw input fallback from acting on the same press.
            NoteHandledPress(ms->pt);
            return 1;
        }

        case WM_LBUTTONUP:
            if (g_llCandidate || g_llDragRoot) {
                g_llCandidate = false;
                g_llDragRoot = nullptr;
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

            // The sink window and the low level hook go away with their
            // thread.
            if (g_rawSinkWnd && GetCurrentThreadId() == g_rawSinkThreadId) {
                g_rawSinkWnd = nullptr;
                g_rawInputRegistered = false;
            }

            if (g_lowLevelMouseHook &&
                GetCurrentThreadId() == g_lowLevelMouseHookThreadId) {
                g_lowLevelMouseHook = nullptr;
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

    HWND rawSinkWnd;
    {
        std::lock_guard<std::mutex> guard(g_rawSinkMutex);
        rawSinkWnd = g_rawSinkWnd;
        g_rawSinkWnd = nullptr;
    }

    // The destroy handler removes the raw device and destroys the window on its
    // own thread. Done without the lock so it can't deadlock against a raw
    // callback waiting for it.
    if (rawSinkWnd && IsWindow(rawSinkWnd)) {
        SendMessage(rawSinkWnd, kRawSinkDestroyMessage, 0, 0);
    }

    UnregisterClass(L"Windhawk_AltDragRawSink_" WH_MOD_ID,
                    GetModuleHandle(nullptr));

    while (g_hookRefCount > 0) {
        Sleep(200);
    }
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    LoadSettings();
}
