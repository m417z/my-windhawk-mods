// ==WindhawkMod==
// @id              virtual-desktop-taskbar-order
// @name            Virtual Desktop Preserve Taskbar Order
// @description     The order on the taskbar isn't preserved between virtual desktop switches, this mod fixes it
// @version         1.0.5
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lcomctl32 -lversion
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
# Virtual Desktop Preserve Taskbar Order

The order on the taskbar isn't preserved between virtual desktop switches, this
mod fixes it.

Only Windows 11 x64 is currently supported. ARM64 is unsupported. For older
Windows versions check out [7+ Taskbar
Tweaker](https://tweaker.ramensoftware.com/).

Experimental ExplorerPatcher support is available for the
[v1.0.4.0 x64 GE taskbar build](https://github.com/ExplorerPatcher/ep_taskbar_releases/releases/tag/v1.0.4.0).
Enable order restoration, move a button on the primary taskbar, then switch
away from that virtual desktop and back. The mod keeps the most recent manual
order for the current Explorer session. Restoration requires a stable pinned
prefix, unpinned groups with one recognized window each, and Windows set to show
only windows on the current desktop.

![Demonstration](https://i.imgur.com/ie8Q9cl.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- explorerPatcherRestoreOrder: false
  $name: Restore order on ExplorerPatcher (experimental)
  $description: Restore the most recent manual order on the primary taskbar for the supported build.
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <array>
#include <atomic>
#include <climits>
#include <cstring>
#include <mutex>
#include <regex>
#include <string>
#include <string_view>

#include <commctrl.h>
#include <psapi.h>

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
};

WinVersion g_winVersion;
DWORD g_taskbarThreadId;
HWND g_hTaskbarWnd;
bool g_nativeTaskbarInitialized;
bool g_explorerPatcherInitialized;
bool g_requireNativeTaskbandOwner;

#pragma region offsets

void* CTaskBand__EnumExistingImmersiveApps;
void* CApplicationViewManager__GetViewInFocus;

size_t OffsetFromAssembly(void* func,
                          size_t defValue,
                          std::string opcode = "mov",
                          int limit = 30) {
    // Example: mov rax, [rcx+0xE0]
    std::regex regex(
        opcode +
        R"( r(?:[a-z]{2}|\d{1,2}), \[r(?:[a-z]{2}|\d{1,2})\+(0x[0-9A-F]+)\])");

    BYTE* p = (BYTE*)func;
    for (int i = 0; i < limit; i++) {
        WH_DISASM_RESULT result;
        if (!Wh_Disasm(p, &result)) {
            break;
        }

        p += result.length;

        std::string_view s = result.text;
        if (s == "ret") {
            break;
        }

        std::match_results<std::string_view::const_iterator> match;
        if (std::regex_match(s.begin(), s.end(), match, regex)) {
            // Wh_Log(L"%S", result.text);
            return std::stoull(match[1], nullptr, 16);
        }
    }

    Wh_Log(L"Failed for %p", func);
    return defValue;
}

LONG_PTR* EV_TASK_SW_APP_VIEW_MGR(LONG_PTR lp) {
    static size_t offset =
        OffsetFromAssembly(CTaskBand__EnumExistingImmersiveApps, 0x220);

    return (LONG_PTR*)(lp + offset);
}

size_t EV_APP_VIEW_MGR_APP_ARRAY_LOCK_OFFSET() {
    static size_t offset = OffsetFromAssembly(
        CApplicationViewManager__GetViewInFocus, 0x210, "lea");

    return offset;
}

SRWLOCK* EV_APP_VIEW_MGR_APP_ARRAY_LOCK(LONG_PTR lp) {
    return (SRWLOCK*)(lp + EV_APP_VIEW_MGR_APP_ARRAY_LOCK_OFFSET());
}

LONG_PTR** EV_APP_VIEW_MGR_APP_ARRAY(LONG_PTR lp) {
    return (LONG_PTR**)(lp + EV_APP_VIEW_MGR_APP_ARRAY_LOCK_OFFSET() + 0x08);
}

size_t* EV_APP_VIEW_MGR_APP_ARRAY_SIZE(LONG_PTR lp) {
    return (size_t*)(lp + EV_APP_VIEW_MGR_APP_ARRAY_LOCK_OFFSET() + 0x10);
}

#pragma endregion  // offsets

void* CTaskBtnGroup_ITaskBtnGroup_vftable;

using CTaskBtnGroup_GetGroup_t = void*(WINAPI*)(void* pThis);
CTaskBtnGroup_GetGroup_t CTaskBtnGroup_GetGroup;

using CTaskBtnGroup_GetNumItems_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetNumItems_t CTaskBtnGroup_GetNumItems;

using CTaskBtnGroup_GetTaskItem_t = void*(WINAPI*)(void* pThis, int index);
CTaskBtnGroup_GetTaskItem_t CTaskBtnGroup_GetTaskItem;

using CTaskGroup_DoesWindowMatch_t =
    HRESULT(WINAPI*)(LPVOID pThis,
                     HWND hCompareWnd,
                     LPVOID pCompareItemIdList,  // ITEMIDLIST*
                     WCHAR* pCompareAppId,
                     int* pnMatch,  // WINDOWMATCHCONFIDENCE*
                     LONG_PTR** p_task_item);
CTaskGroup_DoesWindowMatch_t CTaskGroup_DoesWindowMatch_Original;

using CTaskBand_ViewVirtualDesktopChanged_t = HRESULT(WINAPI*)(
    LPVOID pThis,
    LPVOID applicationView  // IApplicationView*
);
CTaskBand_ViewVirtualDesktopChanged_t
    CTaskBand_ViewVirtualDesktopChanged_Original;

using CTaskListWnd_TryMoveGroup_t =
    bool(WINAPI*)(LPVOID pThis,
                  LPVOID taskGroup,  // ITaskGroup *
                  DWORD index);
CTaskListWnd_TryMoveGroup_t CTaskListWnd_TryMoveGroup_Original;

int g_doesWindowMatchCalls;

HRESULT WINAPI
CTaskGroup_DoesWindowMatch_Hook(LPVOID pThis,
                                HWND hCompareWnd,
                                LPVOID pCompareItemIdList,  // ITEMIDLIST*
                                WCHAR* pCompareAppId,
                                int* pnMatch,  // WINDOWMATCHCONFIDENCE*
                                LONG_PTR** p_task_item) {
    g_doesWindowMatchCalls++;

    HRESULT ret = CTaskGroup_DoesWindowMatch_Original(
        pThis, hCompareWnd, pCompareItemIdList, pCompareAppId, pnMatch,
        p_task_item);

    g_doesWindowMatchCalls--;

    return ret;
}

LONG_PTR g_tryMoveGroup_taskListLongPtr;
LONG_PTR* g_tryMoveGroup_taskGroup;

bool WINAPI CTaskListWnd_TryMoveGroup_Hook(LPVOID pThis,
                                           LPVOID taskGroup,  // ITaskGroup *
                                           DWORD index) {
    g_tryMoveGroup_taskListLongPtr = (LONG_PTR)pThis - 0x28;
    g_tryMoveGroup_taskGroup = (LONG_PTR*)taskGroup;

    bool ret = CTaskListWnd_TryMoveGroup_Original(pThis, taskGroup, index);

    g_tryMoveGroup_taskListLongPtr = 0;
    g_tryMoveGroup_taskGroup = nullptr;

    return ret;
}

#pragma region pointer_redirection

#ifdef _WIN64
#define POINTER_REDIRECTION_ASM_COMMAND "\xFF\x25\xF2\xFF\xFF\xFF"
#else
#define POINTER_REDIRECTION_ASM_COMMAND "\xE8\x00\x00\x00\x00\x58\xFF\x60\xF7"
#endif

#define POINTER_REDIRECTION_SIGNATURE "ptr_redr"

struct POINTER_REDIRECTION {
    void* pOriginalAddress = nullptr;
    void* pRedirectionAddress = nullptr;
    BYTE bAsmCommand[sizeof(POINTER_REDIRECTION_ASM_COMMAND)] =
        POINTER_REDIRECTION_ASM_COMMAND;
    BYTE bSignature[sizeof(POINTER_REDIRECTION_SIGNATURE)] =
        POINTER_REDIRECTION_SIGNATURE;
};

#define POINTER_REDIRECTION_VAR \
    __attribute__((section(".text"))) constinit POINTER_REDIRECTION

void PatchPtr(void** ppAddress, void* pPtr) {
    DWORD dwOldProtect, dwOtherProtect;

    VirtualProtect(ppAddress, sizeof(void*), PAGE_EXECUTE_READWRITE,
                   &dwOldProtect);
    *ppAddress = pPtr;
    VirtualProtect(ppAddress, sizeof(void*), dwOldProtect, &dwOtherProtect);
}

void PointerRedirectionAdd(void** pp, void* pNew, POINTER_REDIRECTION* ppr) {
    PatchPtr(&ppr->pOriginalAddress, *pp);
    PatchPtr(&ppr->pRedirectionAddress, pNew);

    PatchPtr(pp, &ppr->bAsmCommand);
}

void PointerRedirectionRemove(void** pp, POINTER_REDIRECTION* ppr) {
    POINTER_REDIRECTION* pprTemp;

    if (*pp != ppr->bAsmCommand) {
        pprTemp =
            (POINTER_REDIRECTION*)((BYTE*)*pp -
                                   offsetof(POINTER_REDIRECTION, bAsmCommand));
        while (pprTemp->pOriginalAddress != ppr->bAsmCommand) {
            pprTemp = (POINTER_REDIRECTION*)((BYTE*)pprTemp->pOriginalAddress -
                                             offsetof(POINTER_REDIRECTION,
                                                      bAsmCommand));
        }

        PatchPtr(&pprTemp->pOriginalAddress, ppr->pOriginalAddress);
    } else {
        PatchPtr(pp, ppr->pOriginalAddress);
    }
}

#pragma endregion  // pointer_redirection

POINTER_REDIRECTION_VAR prTaskGroupRelease;
POINTER_REDIRECTION_VAR prTaskItemRelease;

LONG_PTR* g_taskGroupVirtualDesktopReleased;
LONG_PTR* g_taskItemVirtualDesktopReleased;

ULONG WINAPI TaskGroupReleaseHook(LONG_PTR this_ptr) {
    ULONG ulRet;

    ulRet = ((ULONG(WINAPI*)(LONG_PTR))prTaskGroupRelease.pOriginalAddress)(
        this_ptr);
    if (ulRet > 0 && g_doesWindowMatchCalls == 0) {
        g_taskGroupVirtualDesktopReleased = (LONG_PTR*)this_ptr;
    }

    return ulRet;
}

ULONG WINAPI TaskItemReleaseHook(LONG_PTR this_ptr) {
    ULONG ulRet;

    ulRet = ((ULONG(WINAPI*)(LONG_PTR))prTaskItemRelease.pOriginalAddress)(
        this_ptr);
    if (ulRet > 0 && g_doesWindowMatchCalls == 0) {
        g_taskItemVirtualDesktopReleased = (LONG_PTR*)this_ptr;
    }

    return ulRet;
}

void OnButtonGroupInserted(LONG_PTR lpTaskSwLongPtr,
                           HDPA hButtonGroupsDpa,
                           int nButtonGroupIndex) {
    Wh_Log(L"> lpTaskSwLongPtr=%p, hButtonGroupsDpa=%p, nButtonGroupIndex=%d",
           (void*)lpTaskSwLongPtr, (void*)hButtonGroupsDpa, nButtonGroupIndex);

    LONG_PTR* plp = (LONG_PTR*)hButtonGroupsDpa;
    int button_groups_count = (int)plp[0];
    LONG_PTR** button_groups = (LONG_PTR**)plp[1];
    LONG_PTR* button_group = button_groups[nButtonGroupIndex];

    int buttons_count = CTaskBtnGroup_GetNumItems(button_group);
    if (buttons_count == 0) {
        return;
    }

    LONG_PTR* task_group = (LONG_PTR*)CTaskBtnGroup_GetGroup(button_group);
    if (!task_group) {
        return;
    }

    LONG_PTR* first_task_item =
        (LONG_PTR*)CTaskBtnGroup_GetTaskItem(button_group, 0);
    if (!first_task_item) {
        return;
    }

    plp = *(LONG_PTR**)task_group;
    void** ppTaskGroupRelease = (void**)&plp[2];
    PointerRedirectionAdd(ppTaskGroupRelease, (void*)TaskGroupReleaseHook,
                          &prTaskGroupRelease);

    plp = *(LONG_PTR**)first_task_item;
    void** ppTaskItemRelease = (void**)&plp[2];
    PointerRedirectionAdd(ppTaskItemRelease, (void*)TaskItemReleaseHook,
                          &prTaskItemRelease);

    LONG_PTR lpAppViewMgr = *EV_TASK_SW_APP_VIEW_MGR(lpTaskSwLongPtr);
    SRWLOCK* pArrayLock = EV_APP_VIEW_MGR_APP_ARRAY_LOCK(lpAppViewMgr);

    AcquireSRWLockExclusive(pArrayLock);

    LONG_PTR* lpArray = *EV_APP_VIEW_MGR_APP_ARRAY(lpAppViewMgr);
    size_t nArraySize = *EV_APP_VIEW_MGR_APP_ARRAY_SIZE(lpAppViewMgr);

    int nMatchCount = 0;
    size_t nRightNeighbourItemIndex = nArraySize;

    // Stage one: move all items in lpArray matching the items
    // in the newly inserted group to the beginning of the array.
    // Their amount is maintained in nMatchCount.
    // Also, find nRightNeighbourItemIndex, which is the index of
    // the item in lpArray which will be the first one before the
    // found matching items.

    for (size_t i = 0; i < nArraySize; i++) {
        g_taskGroupVirtualDesktopReleased = NULL;
        g_taskItemVirtualDesktopReleased = NULL;

        LONG_PTR this_ptr = (LONG_PTR)(lpTaskSwLongPtr + 0x70);
        plp = *(LONG_PTR**)this_ptr;

        ReleaseSRWLockExclusive(pArrayLock);

        CTaskBand_ViewVirtualDesktopChanged_Original((LPVOID)this_ptr,
                                                     (LPVOID)lpArray[i]);

        AcquireSRWLockExclusive(pArrayLock);

        if (lpArray != *EV_APP_VIEW_MGR_APP_ARRAY(lpAppViewMgr) ||
            nArraySize != *EV_APP_VIEW_MGR_APP_ARRAY_SIZE(lpAppViewMgr)) {
            // Something went wrong, abort.
            nMatchCount = 0;
            break;
        }

        if (!g_taskGroupVirtualDesktopReleased) {
            continue;
        }

        if (g_taskGroupVirtualDesktopReleased != task_group) {
            if (nRightNeighbourItemIndex == nArraySize) {
                for (int j = nButtonGroupIndex + 1; j < button_groups_count;
                     j++) {
                    LONG_PTR* check_button_group = button_groups[j];
                    LONG_PTR* check_task_group =
                        (LONG_PTR*)CTaskBtnGroup_GetGroup(check_button_group);
                    if (g_taskGroupVirtualDesktopReleased == check_task_group) {
                        // The current item in lpArray is from the same group
                        // of at least one of the items in button_groups to the
                        // right of the newly added item.
                        nRightNeighbourItemIndex = i - nMatchCount;
                        break;
                    }
                }
            }

            continue;
        }

        if (!g_taskItemVirtualDesktopReleased) {
            continue;
        }

        for (int j = 0; j < buttons_count; j++) {
            LONG_PTR* task_item =
                (LONG_PTR*)CTaskBtnGroup_GetTaskItem(button_group, j);

            if (g_taskItemVirtualDesktopReleased == task_item) {
                // The current item in lpArray matches one of the
                // buttons in the newly added item.
                if (i > (size_t)nMatchCount) {
                    LONG_PTR lpTemp = lpArray[i];
                    memmove(&lpArray[nMatchCount + 1], &lpArray[nMatchCount],
                            (i - nMatchCount) * sizeof(LONG_PTR));
                    lpArray[nMatchCount] = lpTemp;
                }

                nMatchCount++;
                break;
            }
        }
    }

    PointerRedirectionRemove(ppTaskGroupRelease, &prTaskGroupRelease);
    PointerRedirectionRemove(ppTaskItemRelease, &prTaskItemRelease);

    // Stage two: move the found items before the item in
    // nRightNeighbourItemIndex.

    if (nRightNeighbourItemIndex == nArraySize) {
        // By default, move to the right end.
        nRightNeighbourItemIndex = nArraySize - nMatchCount;
    }

    if (nMatchCount > 0 && nRightNeighbourItemIndex > 0) {
        LONG_PTR* lpBuffer = (LONG_PTR*)HeapAlloc(
            GetProcessHeap(), 0, nMatchCount * sizeof(LONG_PTR));
        if (lpBuffer) {
            memcpy(lpBuffer, lpArray, nMatchCount * sizeof(LONG_PTR));
            memmove(&lpArray[0], &lpArray[nMatchCount],
                    nRightNeighbourItemIndex * sizeof(LONG_PTR));
            memcpy(&lpArray[nRightNeighbourItemIndex], lpBuffer,
                   nMatchCount * sizeof(LONG_PTR));

            HeapFree(GetProcessHeap(), 0, lpBuffer);
        }
    }

    ReleaseSRWLockExclusive(pArrayLock);
}

enum class TaskbandOwner { Unknown, Native, ExplorerPatcher };

TaskbandOwner GetTaskbandOwner(LONG_PTR taskband) {
    // Window data can become stale while Explorer is creating or replacing
    // the taskbar. ReadProcessMemory keeps an unreadable object a failed check.
    void* vtable = nullptr;
    HMODULE module = nullptr;
    if (!taskband ||
        !ReadProcessMemory(GetCurrentProcess(), reinterpret_cast<void*>(taskband),
                           &vtable, sizeof(vtable), nullptr) ||
        !vtable ||
        !GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                            reinterpret_cast<PCWSTR>(vtable), &module)) {
        return TaskbandOwner::Unknown;
    }
    if (module == GetModuleHandleW(L"taskbar.dll")) {
        return TaskbandOwner::Native;
    }
    WCHAR path[MAX_PATH];
    DWORD length = GetModuleFileNameW(module, path, ARRAYSIZE(path));
    PCWSTR filename = length && length < ARRAYSIZE(path) ? wcsrchr(path, L'\\') : nullptr;
    if (filename && _wcsnicmp(filename + 1, L"ep_taskbar.", 10) == 0) {
        return TaskbandOwner::ExplorerPatcher;
    }
    return TaskbandOwner::Unknown;
}

void ComFuncVirtualDesktopFixAfterDPA_InsertPtr(HDPA pdpa, int index, void* p) {
    Wh_Log(L"> index=%d, p=%p", index, p);

    if (index == INT_MAX) {
        return;
    }

    if (!g_tryMoveGroup_taskListLongPtr) {
        Wh_Log(L"Not in TryMoveGroup, skipping");
        return;
    }

    if (!p || *(void**)p != CTaskBtnGroup_ITaskBtnGroup_vftable) {
        Wh_Log(L"Invalid pointer or vftable mismatch, skipping");
        return;
    }

    LONG_PTR* button_group = (LONG_PTR*)p;
    LONG_PTR* task_group = (LONG_PTR*)CTaskBtnGroup_GetGroup(button_group);
    if (!task_group || task_group != g_tryMoveGroup_taskGroup) {
        return;
    }

    HWND hTaskSwWnd = (HWND)GetProp(g_hTaskbarWnd, L"TaskbandHWND");
    if (!hTaskSwWnd) {
        return;
    }

    LONG_PTR lpTaskSwLongPtr = GetWindowLongPtr(hTaskSwWnd, 0);
    if (!lpTaskSwLongPtr ||
        (g_requireNativeTaskbandOwner &&
         GetTaskbandOwner(lpTaskSwLongPtr) != TaskbandOwner::Native)) {
        return;
    }

    OnButtonGroupInserted(lpTaskSwLongPtr, pdpa, index);
}

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

void InitializeTaskbarVariables(HWND hTaskbarWnd) {
    Wh_Log(L"Initialized for taskbar %08X", (DWORD)(ULONG_PTR)hTaskbarWnd);
    g_taskbarThreadId = GetWindowThreadProcessId(hTaskbarWnd, nullptr);
    g_hTaskbarWnd = hTaskbarWnd;
}

using DPA_InsertPtr_t = decltype(&DPA_InsertPtr);
DPA_InsertPtr_t DPA_InsertPtr_Original;
auto WINAPI DPA_InsertPtr_Hook(HDPA hdpa, int i, void* p) {
    auto ret = DPA_InsertPtr_Original(hdpa, i, p);

    if (GetCurrentThreadId() == g_taskbarThreadId) {
        ComFuncVirtualDesktopFixAfterDPA_InsertPtr(hdpa, i, p);
    }

    return ret;
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
    if (!hWnd) {
        return hWnd;
    }

    BOOL bTextualClassName = ((ULONG_PTR)lpClassName & ~(ULONG_PTR)0xffff) != 0;

    if (bTextualClassName && wcsicmp(lpClassName, L"Shell_TrayWnd") == 0) {
        Wh_Log(L"Taskbar window created: %08X", (DWORD)(ULONG_PTR)hWnd);
        InitializeTaskbarVariables(hWnd);
    }

    return hWnd;
}

VS_FIXEDFILEINFO* GetModuleVersionInfo(HMODULE hModule, UINT* puPtrLen) {
    void* pFixedFileInfo = nullptr;
    UINT uPtrLen = 0;

    HRSRC hResource =
        FindResource(hModule, MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION);
    if (hResource) {
        HGLOBAL hGlobal = LoadResource(hModule, hResource);
        if (hGlobal) {
            void* pData = LockResource(hGlobal);
            if (pData) {
                if (!VerQueryValue(pData, L"\\", &pFixedFileInfo, &uPtrLen) ||
                    uPtrLen == 0) {
                    pFixedFileInfo = nullptr;
                    uPtrLen = 0;
                }
            }
        }
    }

    if (puPtrLen) {
        *puPtrLen = uPtrLen;
    }

    return (VS_FIXEDFILEINFO*)pFixedFileInfo;
}

WinVersion GetExplorerVersion() {
    VS_FIXEDFILEINFO* fixedFileInfo = GetModuleVersionInfo(nullptr, nullptr);
    if (!fixedFileInfo) {
        return WinVersion::Unsupported;
    }

    WORD major = HIWORD(fixedFileInfo->dwFileVersionMS);
    WORD minor = LOWORD(fixedFileInfo->dwFileVersionMS);
    WORD build = HIWORD(fixedFileInfo->dwFileVersionLS);
    WORD qfe = LOWORD(fixedFileInfo->dwFileVersionLS);

    Wh_Log(L"Version: %u.%u.%u.%u", major, minor, build, qfe);

    switch (major) {
        case 10:
            if (build < 22000) {
                return WinVersion::Win10;
            } else {
                return WinVersion::Win11;
            }
            break;
    }

    return WinVersion::Unsupported;
}

namespace explorer_patcher {

bool restoreOrder;
std::atomic<bool> canApplyHooks;
std::atomic<bool> installAttempted;
std::atomic<bool> ready;
std::atomic<DWORD> taskbarThread;
std::mutex installMutex;
HMODULE hookedModule;

thread_local bool restoringDesktop;

using Update_t = HRESULT(WINAPI*)(void*);
using Move_t = HRESULT(WINAPI*)(void*, int, int);
using CreateGroup_t = void*(WINAPI*)(void*, void*, int);
using TaskCreated_t = HRESULT(WINAPI*)(void*, void*, void*, int);
using GetNumItems_t = int(WINAPI*)(void*);
using GetFlags_t = DWORD(WINAPI*)(void*);
using GetWindow_t = HWND(WINAPI*)(void*);

Update_t updateOriginal;
Move_t moveOriginal;
CreateGroup_t createGroupOriginal;
TaskCreated_t taskCreatedOriginal;
GetNumItems_t getNumItems;
CTaskBtnGroup_GetGroup_t getGroup;
CTaskBtnGroup_GetTaskItem_t getTaskItem;
GetFlags_t getGroupFlags;
GetWindow_t getItemWindow;
void* buttonGroupVtable;
void* taskGroupVtable;
void* windowItemVtable;
decltype(&LoadLibraryExW) loadLibraryOriginal;

bool IsTaskbarThread() {
    if (!ready || !canApplyHooks) {
        return false;
    }
    if (!taskbarThread) {
        HWND tray = FindCurrentProcessTaskbarWnd();
        if (tray) {
            taskbarThread = GetWindowThreadProcessId(tray, nullptr);
        }
    }
    return GetCurrentThreadId() == taskbarThread;
}

HDPA GetButtonGroups(void* taskList) {
    // Verified in _CreateTBGroup and _MoveTaskBtnGroup on the guarded image.
    return *reinterpret_cast<HDPA*>(static_cast<BYTE*>(taskList) + 0xD8);
}

constexpr int kMaxGroups = 64;

struct WindowIdentity {
    HWND hwnd;
    DWORD process;
    DWORD thread;

    bool operator==(const WindowIdentity&) const = default;
};

struct ButtonGroup {
    void* button;
    bool pinned;
};

struct TaskbarOrder {
    void* taskList;
    int pinnedCount;
    int count;
    std::array<WindowIdentity, kMaxGroups> windows;
};

struct TaskCreatedContext {
    void* taskList;
    void* group;
    void* item;
};

// Only the primary taskbar thread captures or reads the saved order.
TaskbarOrder savedOrder;
thread_local TaskCreatedContext taskCreatedContext;

bool HasVtable(void* object, void* vtable) {
    return object && *static_cast<void**>(object) == vtable;
}

bool IsPrimaryTaskList(void* taskList) {
    // In the guarded build, the primary task-list window stores the complete
    // CTaskListWnd object at index 0. Check again after callbacks that can
    // synchronously rebuild the taskbar.
    HWND tray = FindWindowW(L"Shell_TrayWnd", nullptr);
    if (!taskList || !tray ||
        GetWindowThreadProcessId(tray, nullptr) != GetCurrentThreadId()) {
        return false;
    }
    HWND band = static_cast<HWND>(GetPropW(tray, L"TaskbandHWND"));
    HWND list = band ? FindWindowExW(band, nullptr, L"MSTaskListWClass", nullptr)
                     : nullptr;
    return list && GetWindowThreadProcessId(list, nullptr) == GetCurrentThreadId() &&
           reinterpret_cast<void*>(GetWindowLongPtrW(list, 0)) == taskList;
}

bool IsCurrentDesktopOnly() {
    DWORD filter = 0;
    DWORD size = sizeof(filter);
    return RegGetValueW(HKEY_CURRENT_USER,
                        L"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced",
                        L"VirtualDesktopTaskbarFilter", RRF_RT_REG_DWORD,
                        nullptr, &filter, &size) == ERROR_SUCCESS &&
           filter == 1;
}

WindowIdentity ReadWindowIdentity(void* item) {
    if (!HasVtable(item, windowItemVtable)) {
        return {};
    }
    WindowIdentity key{};
    key.hwnd = getItemWindow(item);
    key.thread = key.hwnd ? GetWindowThreadProcessId(key.hwnd, &key.process) : 0;
    if (!key.thread || !key.process) {
        return {};
    }
    return key;
}

ButtonGroup ReadButtonGroup(HDPA groups, int index) {
    void* button = DPA_GetPtr(groups, index);
    if (!HasVtable(button, buttonGroupVtable)) {
        return {};
    }
    void* group = getGroup(button);
    if (!HasVtable(group, taskGroupVtable)) {
        return {};
    }
    return {button, (getGroupFlags(group) & 1) != 0};
}

WindowIdentity ReadSingleWindow(void* button) {
    return getNumItems(button) == 1 ? ReadWindowIdentity(getTaskItem(button, 0))
                                    : WindowIdentity{};
}

int FindRank(const TaskbarOrder& order, const WindowIdentity& window) {
    for (int i = 0; window.hwnd && i < order.count; ++i) {
        if (order.windows[i] == window) {
            return i;
        }
    }
    return -1;
}

TaskbarOrder CaptureOrder(void* taskList) {
    HDPA groups = GetButtonGroups(taskList);
    int count = groups ? DPA_GetPtrCount(groups) : 0;
    if (count <= 0 || count > kMaxGroups) {
        return {};
    }

    TaskbarOrder order{};
    order.taskList = taskList;
    for (int i = 0; i < count; ++i) {
        const auto group = ReadButtonGroup(groups, i);
        if (!group.button || (group.pinned && order.count)) {
            return {};
        }
        if (group.pinned) {
            ++order.pinnedCount;
            continue;
        }
        const auto window = ReadSingleWindow(group.button);
        if (!window.hwnd || FindRank(order, window) >= 0) {
            return {};
        }
        order.windows[order.count++] = window;
    }

    // A window also represented by a pin has ambiguous placement.
    for (int i = 0; i < order.pinnedCount; ++i) {
        void* button = DPA_GetPtr(groups, i);
        int count = getNumItems(button);
        if (count < 0 || count > kMaxGroups) {
            return {};
        }
        for (int j = 0; j < count; ++j) {
            if (FindRank(order, ReadWindowIdentity(getTaskItem(button, j))) >= 0) {
                return {};
            }
        }
    }
    return order;
}

int GetInsertionIndex(void* taskList, void* taskGroup) {
    if (taskList != savedOrder.taskList || taskList != taskCreatedContext.taskList ||
        taskGroup != taskCreatedContext.group ||
        !HasVtable(taskGroup, taskGroupVtable) || (getGroupFlags(taskGroup) & 1)) {
        return INT_MAX;
    }
    int rank = FindRank(savedOrder, ReadWindowIdentity(taskCreatedContext.item));
    if (rank < 0) {
        return INT_MAX;
    }

    HDPA groups = GetButtonGroups(taskList);
    int count = groups ? DPA_GetPtrCount(groups) : 0;
    if (count < savedOrder.pinnedCount || count > kMaxGroups) {
        return INT_MAX;
    }

    int candidate = INT_MAX;
    int previousRank = -1;
    for (int i = 0; i < count; ++i) {
        const auto group = ReadButtonGroup(groups, i);
        if (!group.button || group.pinned != (i < savedOrder.pinnedCount)) {
            return INT_MAX;
        }
        if (group.pinned) {
            continue;
        }
        int liveRank = FindRank(savedOrder, ReadSingleWindow(group.button));
        if (liveRank <= previousRank || liveRank == rank) {
            return INT_MAX;
        }
        if (liveRank > rank && candidate == INT_MAX) {
            candidate = i;
        }
        previousRank = liveRank;
    }
    return candidate;
}

HRESULT WINAPI UpdateHook(void* taskBand) {
    if (!IsTaskbarThread() || restoringDesktop || !IsCurrentDesktopOnly()) {
        return updateOriginal(taskBand);
    }
    restoringDesktop = true;
    HRESULT result = updateOriginal(taskBand);
    restoringDesktop = false;
    return result;
}

HRESULT WINAPI MoveHook(void* taskList, int from, int to) {
    HRESULT result = moveOriginal(taskList, from, to);
    if (result == S_OK && IsTaskbarThread() && !restoringDesktop &&
        IsCurrentDesktopOnly() && IsPrimaryTaskList(taskList)) {
        savedOrder = CaptureOrder(taskList);
    }
    return result;
}

void* WINAPI CreateGroupHook(void* taskList, void* taskGroup, int index) {
    if (restoringDesktop && index == INT_MAX && IsTaskbarThread() &&
        IsPrimaryTaskList(taskList)) {
        index = GetInsertionIndex(taskList, taskGroup);
    }
    return createGroupOriginal(taskList, taskGroup, index);
}

HRESULT WINAPI TaskCreatedHook(void* taskList, void* group, void* item, int flags) {
    if (!restoringDesktop || !IsTaskbarThread() || taskList != savedOrder.taskList) {
        return taskCreatedOriginal(taskList, group, item, flags);
    }
    TaskCreatedContext previous = taskCreatedContext;
    taskCreatedContext = {taskList, group, item};
    HRESULT result = taskCreatedOriginal(taskList, group, item, flags);
    taskCreatedContext = previous;
    return result;
}

bool IsSupportedImage(HMODULE module) {
    MODULEINFO info{};
    if (!GetModuleInformation(GetCurrentProcess(), module, &info, sizeof(info)) ||
        info.SizeOfImage != 0x299000) {
        return false;
    }
    const auto* base = reinterpret_cast<const BYTE*>(module);
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(base);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE ||
        dos->e_lfanew < static_cast<LONG>(sizeof(IMAGE_DOS_HEADER)) ||
        static_cast<DWORD>(dos->e_lfanew) > info.SizeOfImage - sizeof(IMAGE_NT_HEADERS64)) {
        return false;
    }
    const auto* pe = reinterpret_cast<const IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
    if (pe->Signature != IMAGE_NT_SIGNATURE ||
        pe->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 ||
        pe->FileHeader.TimeDateStamp != 0x6AE29682 ||
        pe->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC ||
        pe->OptionalHeader.NumberOfRvaAndSizes <= IMAGE_DIRECTORY_ENTRY_DEBUG) {
        return false;
    }

    const auto& debug = pe->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DEBUG];
    if (!debug.VirtualAddress || debug.VirtualAddress >= info.SizeOfImage ||
        debug.Size > info.SizeOfImage - debug.VirtualAddress) {
        return false;
    }
    // RSDS, ep_taskbar.ge.pdb GUID 8bdf7673-5e1d-4218-b2c6-93af45c4d4f9, age 1.
    constexpr BYTE codeView[] = {
        'R', 'S', 'D', 'S', 0x73, 0x76, 0xDF, 0x8B, 0x1D, 0x5E, 0x18, 0x42,
        0xB2, 0xC6, 0x93, 0xAF, 0x45, 0xC4, 0xD4, 0xF9, 1, 0, 0, 0,
    };
    const auto* entries =
        reinterpret_cast<const IMAGE_DEBUG_DIRECTORY*>(base + debug.VirtualAddress);
    for (size_t i = 0; i < debug.Size / sizeof(*entries); ++i) {
        const auto& entry = entries[i];
        if (entry.Type == IMAGE_DEBUG_TYPE_CODEVIEW &&
            entry.SizeOfData >= sizeof(codeView) && entry.AddressOfRawData &&
            entry.AddressOfRawData <= info.SizeOfImage - sizeof(codeView) &&
            std::memcmp(base + entry.AddressOfRawData, codeView, sizeof(codeView)) == 0) {
            return true;
        }
    }
    return false;
}

void TryInstall(HMODULE module) {
    if (!canApplyHooks || installAttempted ||
        !module || (reinterpret_cast<ULONG_PTR>(module) & 3)) {
        return;
    }

    WCHAR path[MAX_PATH];
    DWORD length = GetModuleFileName(module, path, ARRAYSIZE(path));
    PCWSTR filename = length && length < ARRAYSIZE(path) ? wcsrchr(path, L'\\') : nullptr;
    if (!filename || _wcsicmp(filename + 1, L"ep_taskbar.ge.dll") != 0 ||
        installAttempted.exchange(true)) {
        return;
    }

    // The load hook can run on any thread, including during mod unload.
    std::lock_guard lock(installMutex);
    if (!canApplyHooks) {
        return;
    }

    // Keep the image mapped until Windhawk has removed the hooks.
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                            reinterpret_cast<PCWSTR>(module), &hookedModule)) {
        Wh_Log(L"ExplorerPatcher taskbar could not be retained");
        return;
    }
    if (!IsSupportedImage(module)) {
        Wh_Log(L"Unsupported ExplorerPatcher taskbar build: ep_taskbar.ge.dll");
        FreeLibrary(hookedModule);
        hookedModule = nullptr;
        return;
    }
    struct Hook {
        const char* name;
        void* hook;
        void** original;
        void* address = nullptr;
    };
    Hook hooks[] = {
        {"?_UpdateVirtualDesktopInclusion@CTaskBand@@IEAAJXZ",
         reinterpret_cast<void*>(UpdateHook), reinterpret_cast<void**>(&updateOriginal)},
        {"?_MoveTaskBtnGroup@CTaskListWnd@@IEAAJHH@Z",
         reinterpret_cast<void*>(MoveHook), reinterpret_cast<void**>(&moveOriginal)},
        {"?_CreateTBGroup@CTaskListWnd@@IEAAPEAUITaskBtnGroup@@PEAUITaskGroup@@H@Z",
         reinterpret_cast<void*>(CreateGroupHook), reinterpret_cast<void**>(&createGroupOriginal)},
        {"?_TaskCreated@CTaskListWnd@@IEAAJPEAUITaskGroup@@PEAUITaskItem@@H@Z",
         reinterpret_cast<void*>(TaskCreatedHook), reinterpret_cast<void**>(&taskCreatedOriginal)},
    };
    auto resolve = [module]<typename T>(const char* name, T& target) {
        target = reinterpret_cast<T>(GetProcAddress(module, name));
        if (!target) {
            Wh_Log(L"ExplorerPatcher taskbar export missing: %S", name);
        }
        return target != nullptr;
    };
    if (!resolve("??_7CTaskBtnGroup@@6BITaskBtnGroup@@@", buttonGroupVtable) ||
        !resolve("??_7CTaskGroup@@6BITaskGroup@@@", taskGroupVtable) ||
        !resolve("??_7CWindowTaskItem@@6BITaskItem@@@", windowItemVtable) ||
        !resolve("?GetNumItems@CTaskBtnGroup@@UEAAHXZ", getNumItems) ||
        !resolve("?GetGroup@CTaskBtnGroup@@UEAAPEAUITaskGroup@@XZ", getGroup) ||
        !resolve("?GetTaskItem@CTaskBtnGroup@@UEAAPEAUITaskItem@@H@Z", getTaskItem) ||
        !resolve("?GetFlags@CTaskGroup@@UEBAKXZ", getGroupFlags) ||
        !resolve("?GetWindow@CWindowTaskItem@@UEAAPEAUHWND__@@XZ", getItemWindow)) {
        return;
    }
    for (auto& hook : hooks) {
        hook.address = reinterpret_cast<void*>(GetProcAddress(module, hook.name));
        if (!hook.address) {
            Wh_Log(L"ExplorerPatcher taskbar export missing: %S", hook.name);
            return;
        }
    }
    size_t queued = 0;
    for (; queued < ARRAYSIZE(hooks); ++queued) {
        const auto& hook = hooks[queued];
        if (!Wh_SetFunctionHook(hook.address, hook.hook, hook.original)) {
            Wh_Log(L"ExplorerPatcher taskbar hook could not be queued: %S", hook.name);
            break;
        }
    }
    if (queued == ARRAYSIZE(hooks) && Wh_ApplyHookOperations()) {
        ready = true;
    } else if (queued) {
        for (size_t i = 0; i < queued; ++i) {
            Wh_RemoveFunctionHook(hooks[i].address);
        }
        Wh_ApplyHookOperations();
    }
    Wh_Log(L"ExplorerPatcher taskbar hooks installed: %d", ready.load());
}

HMODULE WINAPI LoadLibraryHook(LPCWSTR filename, HANDLE file, DWORD flags) {
    HMODULE module = loadLibraryOriginal(filename, file, flags);
    if (canApplyHooks) {
        TryInstall(module);
    }
    return module;
}

bool Init() {
    auto loadLibrary = GetProcAddress(GetModuleHandle(L"kernelbase.dll"), "LoadLibraryExW");
    return loadLibrary && Wh_SetFunctionHook(
        reinterpret_cast<void*>(loadLibrary), reinterpret_cast<void*>(LoadLibraryHook),
        reinterpret_cast<void**>(&loadLibraryOriginal));
}

void AfterInit() {
    HWND tray = FindCurrentProcessTaskbarWnd();
    if (tray) {
        taskbarThread = GetWindowThreadProcessId(tray, nullptr);
    }
    canApplyHooks = true;
    HMODULE module = GetModuleHandle(L"ep_taskbar.ge.dll");
    if (module) {
        TryInstall(module);
    } else {
        Wh_Log(L"Waiting for the supported ExplorerPatcher taskbar module ep_taskbar.ge.dll");
    }
}

void BeforeUninit() {
    canApplyHooks = false;
    std::lock_guard lock(installMutex);
    ready = false;
}

void Uninit() {
    if (hookedModule) {
        FreeLibrary(hookedModule);
        hookedModule = nullptr;
    }
}

}  // namespace explorer_patcher

bool InitNativeTaskbar() {
    WindhawkUtils::SYMBOL_HOOK taskbarDllHooks[] = {
        {
            {LR"(const CTaskBtnGroup::`vftable'{for `ITaskBtnGroup'})"},
            &CTaskBtnGroup_ITaskBtnGroup_vftable,
        },
        {
            {LR"(public: virtual struct ITaskGroup * __cdecl CTaskBtnGroup::GetGroup(void))"},
            &CTaskBtnGroup_GetGroup,
        },
        {
            {LR"(public: virtual int __cdecl CTaskBtnGroup::GetNumItems(void))"},
            &CTaskBtnGroup_GetNumItems,
        },
        {
            {LR"(public: virtual struct ITaskItem * __cdecl CTaskBtnGroup::GetTaskItem(int))"},
            &CTaskBtnGroup_GetTaskItem,
        },
        {
            {LR"(public: virtual long __cdecl CTaskGroup::DoesWindowMatch(struct HWND__ *,struct _ITEMIDLIST_ABSOLUTE const *,unsigned short const *,enum WINDOWMATCHCONFIDENCE *,struct ITaskItem * *))"},
            &CTaskGroup_DoesWindowMatch_Original,
            CTaskGroup_DoesWindowMatch_Hook,
        },
        {
            {LR"(public: virtual bool __cdecl CTaskListWnd::TryMoveGroup(struct ITaskGroup *,unsigned int))"},
            &CTaskListWnd_TryMoveGroup_Original,
            CTaskListWnd_TryMoveGroup_Hook,
        },
        {
            {LR"(public: virtual long __cdecl CTaskBand::ViewVirtualDesktopChanged(struct IApplicationView *))"},
            &CTaskBand_ViewVirtualDesktopChanged_Original,
        },
        // For offsets:
        {
            {LR"(protected: void __cdecl CTaskBand::_EnumExistingImmersiveApps(void))"},
            &CTaskBand__EnumExistingImmersiveApps,
        },
    };

    HMODULE taskbarModule =
        LoadLibraryEx(L"taskbar.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!taskbarModule) {
        Wh_Log(L"Couldn't load taskbar.dll");
        return false;
    }

    if (!HookSymbols(taskbarModule, taskbarDllHooks,
                     ARRAYSIZE(taskbarDllHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    // twinui.pcshell.dll
    WindhawkUtils::SYMBOL_HOOK twinuiPcshellSymbolHooks[] = {
        // For offsets:
        {
            {LR"(public: virtual long __cdecl CApplicationViewManager::GetViewInFocus(struct IApplicationView * *))"},
            &CApplicationViewManager__GetViewInFocus,
        },
    };

    HMODULE twinuiPcshellModule = LoadLibraryEx(L"twinui.pcshell.dll", nullptr,
                                                LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!twinuiPcshellModule) {
        Wh_Log(L"Couldn't load twinui.pcshell.dll");
        return false;
    }

    if (!HookSymbols(twinuiPcshellModule, twinuiPcshellSymbolHooks,
                     ARRAYSIZE(twinuiPcshellSymbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    WindhawkUtils::SetFunctionHook(DPA_InsertPtr, DPA_InsertPtr_Hook,
                                   &DPA_InsertPtr_Original);

    WindhawkUtils::SetFunctionHook(CreateWindowExW, CreateWindowExW_Hook,
                                   &CreateWindowExW_Original);

    return true;
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    g_winVersion = GetExplorerVersion();
    if (g_winVersion < WinVersion::Win11) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
    }

#ifdef _M_ARM64
    Wh_Log(L"Unsupported architecture");
    return FALSE;
#endif

    explorer_patcher::restoreOrder = Wh_GetIntSetting(L"explorerPatcherRestoreOrder");
    if (!explorer_patcher::restoreOrder) {
        g_nativeTaskbarInitialized = InitNativeTaskbar();
        return g_nativeTaskbarInitialized;
    }

    HWND tray = FindCurrentProcessTaskbarWnd();
    HWND taskband = tray ? static_cast<HWND>(GetPropW(tray, L"TaskbandHWND")) : nullptr;
    TaskbandOwner owner = GetTaskbandOwner(taskband ? GetWindowLongPtrW(taskband, 0) : 0);
    if (owner != TaskbandOwner::ExplorerPatcher) {
        // When Explorer is still starting, either module may own the future
        // taskbar. Require native ownership before using the native layout.
        g_requireNativeTaskbandOwner = owner == TaskbandOwner::Unknown;
        g_nativeTaskbarInitialized = InitNativeTaskbar();
    }
    if (owner != TaskbandOwner::Native) {
        g_explorerPatcherInitialized = explorer_patcher::Init();
    }
    Wh_Log(L"Taskbar hooks initialized: native=%d, ExplorerPatcher=%d",
           g_nativeTaskbarInitialized, g_explorerPatcherInitialized);
    return g_nativeTaskbarInitialized || g_explorerPatcherInitialized;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (g_explorerPatcherInitialized) {
        explorer_patcher::AfterInit();
    }

    if (g_nativeTaskbarInitialized) {
        HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
        if (hTaskbarWnd) {
            InitializeTaskbarVariables(hTaskbarWnd);
        }
    }
}

void Wh_ModBeforeUninit() {
    if (g_explorerPatcherInitialized) {
        explorer_patcher::BeforeUninit();
    }
}

void Wh_ModUninit() {
    if (g_explorerPatcherInitialized) {
        explorer_patcher::Uninit();
    }
}

BOOL Wh_ModSettingsChanged(BOOL* reload) {
    *reload = explorer_patcher::restoreOrder !=
              (Wh_GetIntSetting(L"explorerPatcherRestoreOrder") != 0);
    return TRUE;
}
