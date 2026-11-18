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

Experimental ExplorerPatcher support is available for the specific taskbar
build recognized by this mod. Enable order restoration, move a button on the
primary taskbar, then switch away from that virtual desktop and back. The mod
keeps the most recent manual order for the current Explorer session. It can
restore recognized single-window groups when the pinned prefix is unchanged
and Windows shows only windows on the current desktop.

Enable ExplorerPatcher diagnostics and mod logging to record the observed
order and insertion decisions. Diagnostics alone perform a dry run.

![Demonstration](https://i.imgur.com/ie8Q9cl.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- explorerPatcherDiagnostics: false
  $name: ExplorerPatcher diagnostics
  $description: Trace order decisions on the supported ExplorerPatcher taskbar build.
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
    if (!lpTaskSwLongPtr) {
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

bool diagnostics;
bool restoreOrder;
std::atomic<bool> canApplyHooks;
std::atomic<bool> installing;
std::atomic<bool> ready;
std::atomic<DWORD> taskbarThread;

// Only accessed on the primary taskbar thread.
unsigned int sequence;
thread_local unsigned int desktopUpdateDepth;
thread_local bool restoreDuringUpdate;

using Update_t = HRESULT(WINAPI*)(void*);
using Move_t = HRESULT(WINAPI*)(void*, int, int);
using CreateGroup_t = void*(WINAPI*)(void*, void*, int);
using TaskCreated_t = HRESULT(WINAPI*)(void*, void*, void*, int);
using TaskDestroyed_t = HRESULT(WINAPI*)(void*, void*, void*, unsigned int);
using GetNumItems_t = int(WINAPI*)(void*);
using GetFlags_t = DWORD(WINAPI*)(void*);
using GetWindow_t = HWND(WINAPI*)(void*);
using IsOnPrimaryTaskband_t = BOOL(WINAPI*)(void*);

Update_t updateOriginal;
Move_t moveOriginal;
CreateGroup_t createGroupOriginal;
TaskCreated_t taskCreatedOriginal;
TaskDestroyed_t taskDestroyedOriginal;
GetNumItems_t getNumItems;
CTaskBtnGroup_GetGroup_t getGroup;
CTaskBtnGroup_GetTaskItem_t getTaskItem;
GetFlags_t getGroupFlags;
GetWindow_t getItemWindow;
IsOnPrimaryTaskband_t isOnPrimaryTaskband;
void* buttonGroupVtable;
void* taskGroupVtable;
void* windowItemVtable;
void* taskListUiVtable;
decltype(&LoadLibraryExW) loadLibraryOriginal;

bool IsTaskbarThread() {
    if (!ready) {
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

bool CanLog() {
    return diagnostics && sequence < 256;
}

HDPA GetButtonGroups(void* taskList) {
    // Verified in _CreateTBGroup and _MoveTaskBtnGroup on the guarded image.
    return *reinterpret_cast<HDPA*>(static_cast<BYTE*>(taskList) + 0xD8);
}

int GetGroupCount(void* taskList) {
    HDPA groups = GetButtonGroups(taskList);
    return groups ? DPA_GetPtrCount(groups) : 0;
}

constexpr int kMaxGroups = 64;
constexpr int kMaxItemsPerGroup = 8;

struct WindowIdentity {
    HWND hwnd;
    DWORD process;
    DWORD thread;

    bool operator==(const WindowIdentity&) const = default;
};

struct GroupSnapshot {
    bool recognized;
    bool pinned;
    int itemCount;
    std::array<unsigned int, kMaxItemsPerGroup> items;
};

struct Snapshot {
    int count;
    std::array<GroupSnapshot, kMaxGroups> groups;
};

struct TaskCreatedContext {
    void* taskList;
    void* group;
    void* item;
};

std::array<WindowIdentity, 256> windowIdentities;
size_t windowIdentityCount;
void* savedTaskList;
Snapshot savedOrder;
Snapshot lastObservedOrder;
bool observedDuringUpdate;
thread_local TaskCreatedContext taskCreatedContext;

bool HasVtable(void* object, void* vtable) {
    return object && *static_cast<void**>(object) == vtable;
}

bool IsPrimaryTaskList(void* taskList) {
    // The guarded CTaskListWnd methods receive the complete object. The
    // ITaskListUI subobject at +0x28 owns IsOnPrimaryTaskband.
    if (!taskList) {
        return false;
    }
    void* taskListUi = static_cast<BYTE*>(taskList) + 0x28;
    return HasVtable(taskListUi, taskListUiVtable) && isOnPrimaryTaskband(taskListUi);
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

unsigned int WindowId(void* item) {
    if (!HasVtable(item, windowItemVtable)) {
        return 0;
    }
    WindowIdentity key{};
    key.hwnd = getItemWindow(item);
    key.thread = key.hwnd ? GetWindowThreadProcessId(key.hwnd, &key.process) : 0;
    if (!key.thread || !key.process) {
        return 0;
    }
    for (size_t i = 0; i < windowIdentityCount; ++i) {
        if (windowIdentities[i] == key) {
            return static_cast<unsigned int>(i + 1);
        }
    }
    if (windowIdentityCount == windowIdentities.size()) {
        return 0;
    }
    windowIdentities[windowIdentityCount++] = key;
    return static_cast<unsigned int>(windowIdentityCount);
}

Snapshot ReadOrder(void* taskList) {
    Snapshot snapshot{};
    HDPA groups = GetButtonGroups(taskList);
    snapshot.count = groups ? DPA_GetPtrCount(groups) : 0;
    for (int i = 0; i < snapshot.count && i < kMaxGroups; ++i) {
        void* buttonGroup = DPA_GetPtr(groups, i);
        if (!HasVtable(buttonGroup, buttonGroupVtable)) {
            continue;
        }
        void* group = getGroup(buttonGroup);
        if (!HasVtable(group, taskGroupVtable)) {
            continue;
        }
        auto& entry = snapshot.groups[i];
        entry.recognized = true;
        entry.pinned = (getGroupFlags(group) & 1) != 0;
        entry.itemCount = getNumItems(buttonGroup);
        for (int j = 0; j < entry.itemCount && j < kMaxItemsPerGroup; ++j) {
            entry.items[j] = WindowId(getTaskItem(buttonGroup, j));
        }
    }
    return snapshot;
}

void LogOrder(PCWSTR label, const Snapshot& snapshot) {
    std::wstring order;
    for (int i = 0; i < snapshot.count && i < kMaxGroups; ++i) {
        const auto& entry = snapshot.groups[i];
        order += entry.recognized ? (entry.pinned ? L"P[" : L"[") : L"?[";
        for (int j = 0; j < entry.itemCount && j < kMaxItemsPerGroup; ++j) {
            if (j) {
                order += L",";
            }
            order += std::to_wstring(entry.items[j]);
        }
        if (entry.itemCount > kMaxItemsPerGroup) {
            order += L",...";
        }
        order += L"] ";
    }
    if (snapshot.count > kMaxGroups) {
        order += L"...";
    }
    Wh_Log(L"EP #%u %s groups=%d order=%s", ++sequence, label,
           snapshot.count, order.c_str());
}

int SavedPosition(unsigned int id) {
    int position = -1;
    for (int i = 0; id && i < savedOrder.count && i < kMaxGroups; ++i) {
        const auto& entry = savedOrder.groups[i];
        for (int j = 0; entry.recognized && j < entry.itemCount &&
                        j < kMaxItemsPerGroup; ++j) {
            if (entry.items[j] == id) {
                if (position >= 0 || entry.pinned || entry.itemCount != 1) {
                    return -1;
                }
                position = i;
            }
        }
    }
    return position;
}

int PinnedPrefix(const Snapshot& snapshot) {
    if (snapshot.count < 0 || snapshot.count > kMaxGroups) {
        return -1;
    }
    int prefix = 0;
    for (int i = 0; i < snapshot.count; ++i) {
        const auto& entry = snapshot.groups[i];
        if (!entry.recognized || (entry.pinned && i != prefix)) {
            return -1;
        }
        prefix += entry.pinned;
    }
    return prefix;
}

int GetInsertionIndex(void* taskList, void* taskGroup) {
    if (taskList != savedTaskList || taskList != taskCreatedContext.taskList ||
        taskGroup != taskCreatedContext.group ||
        !HasVtable(taskGroup, taskGroupVtable)) {
        if (CanLog()) {
            Wh_Log(L"EP #%u insertion skipped: no matching task-created context or drag snapshot",
                   ++sequence);
        }
        return INT_MAX;
    }
    unsigned int id = WindowId(taskCreatedContext.item);
    int rank = SavedPosition(id);
    if ((getGroupFlags(taskGroup) & 1) || rank < 0) {
        if (CanLog()) {
            Wh_Log(L"EP #%u insertion skipped: id=%u pinned or no unique single-window match",
                   ++sequence, id);
        }
        return INT_MAX;
    }

    Snapshot live = ReadOrder(taskList);
    int firstUnpinned = PinnedPrefix(live);
    int savedFirstUnpinned = PinnedPrefix(savedOrder);
    if (firstUnpinned < 0 || firstUnpinned != savedFirstUnpinned) {
        if (CanLog()) {
            Wh_Log(L"EP #%u insertion skipped: id=%u saved index=%d pin boundary saved=%d current=%d",
                   ++sequence, id, rank, savedFirstUnpinned, firstUnpinned);
        }
        return INT_MAX;
    }
    int candidate = live.count;
    std::array<bool, kMaxGroups> seenRanks{};
    seenRanks[rank] = true;
    int previousRank = -1;
    for (int i = firstUnpinned; i < live.count; ++i) {
        const auto& entry = live.groups[i];
        int otherRank = entry.itemCount == 1 ? SavedPosition(entry.items[0]) : -1;
        if (otherRank < 0 || seenRanks[otherRank]) {
            if (CanLog()) {
                Wh_Log(L"EP #%u insertion skipped: id=%u unmatched or ambiguous group at index=%d",
                       ++sequence, id, i);
            }
            return INT_MAX;
        }
        seenRanks[otherRank] = true;
        if (restoreDuringUpdate && otherRank < previousRank) {
            if (CanLog()) {
                Wh_Log(L"EP #%u insertion skipped: id=%u existing groups are out of saved order",
                       ++sequence, id);
            }
            return INT_MAX;
        }
        previousRank = otherRank;
        if (otherRank > rank && candidate == live.count) {
            candidate = i;
        }
    }
    const bool apply = restoreDuringUpdate && candidate < live.count &&
                       IsPrimaryTaskList(taskList);
    if (CanLog()) {
        Wh_Log(L"EP #%u insertion id=%u saved index=%d candidate=%d native append=%d apply=%d",
               ++sequence, id, rank, candidate, live.count, apply);
    }
    return apply ? candidate : INT_MAX;
}

HRESULT WINAPI UpdateHook(void* taskBand) {
    if (!IsTaskbarThread() || (!restoreOrder && !CanLog())) {
        return updateOriginal(taskBand);
    }

    if (++desktopUpdateDepth == 1) {
        observedDuringUpdate = false;
        restoreDuringUpdate = restoreOrder && IsCurrentDesktopOnly();
    }
    if (CanLog()) {
        Wh_Log(L"EP #%u desktop inclusion begin (depth %u, restore=%d)", ++sequence,
               desktopUpdateDepth, restoreDuringUpdate);
    }
    HRESULT result = updateOriginal(taskBand);
    if (--desktopUpdateDepth == 0) {
        restoreDuringUpdate = false;
    }
    if (CanLog()) {
        Wh_Log(L"EP #%u desktop inclusion end (result 0x%08X)", ++sequence,
               static_cast<unsigned int>(result));
        if (!desktopUpdateDepth && observedDuringUpdate && CanLog()) {
            LogOrder(L"last creation during update", lastObservedOrder);
        }
    }
    return result;
}

HRESULT WINAPI MoveHook(void* taskList, int from, int to) {
    HRESULT result = moveOriginal(taskList, from, to);
    if (!IsTaskbarThread()) {
        return result;
    }
    if (CanLog()) {
        Wh_Log(L"EP #%u move %d -> %d, groups=%d, result=0x%08X", ++sequence,
               from, to, GetGroupCount(taskList),
               static_cast<unsigned int>(result));
    }
    if (result == S_OK && !desktopUpdateDepth &&
        (restoreOrder ? IsCurrentDesktopOnly() : CanLog()) &&
        IsPrimaryTaskList(taskList)) {
        savedTaskList = taskList;
        savedOrder = ReadOrder(taskList);
        if (CanLog()) {
            LogOrder(L"order after drag move", savedOrder);
        }
    }
    return result;
}

void* WINAPI CreateGroupHook(void* taskList, void* taskGroup, int index) {
    if (!IsTaskbarThread()) {
        return createGroupOriginal(taskList, taskGroup, index);
    }
    const int before = CanLog() ? GetGroupCount(taskList) : -1;
    int requested = index;
    if (desktopUpdateDepth && index == INT_MAX && (restoreDuringUpdate || CanLog())) {
        requested = GetInsertionIndex(taskList, taskGroup);
    }
    void* result = createGroupOriginal(taskList, taskGroup, requested);
    if (before >= 0 && CanLog()) {
        HDPA groups = GetButtonGroups(taskList);
        int inserted = groups && result ? DPA_GetPtrIndex(groups, result) : -1;
        Wh_Log(L"EP #%u create group requested=%d, passed=%d, inserted=%d, groups=%d -> %d, desktop depth=%u",
               ++sequence, index, requested, inserted, before, GetGroupCount(taskList),
               desktopUpdateDepth);
    }
    return result;
}

HRESULT WINAPI TaskCreatedHook(void* taskList, void* group, void* item, int flags) {
    if (!IsTaskbarThread() || (!restoreDuringUpdate && !CanLog())) {
        return taskCreatedOriginal(taskList, group, item, flags);
    }
    TaskCreatedContext previous = taskCreatedContext;
    taskCreatedContext = {taskList, group, item};
    unsigned int id = CanLog() ? WindowId(item) : 0;
    HRESULT result = taskCreatedOriginal(taskList, group, item, flags);
    taskCreatedContext = previous;
    if (CanLog()) {
        HDPA groups = GetButtonGroups(taskList);
        int count = groups ? DPA_GetPtrCount(groups) : 0;
        int populatedGroups = 0;
        int itemCount = 0;
        for (int index = 0; index < count && index < 64; ++index) {
            void* buttonGroup = DPA_GetPtr(groups, index);
            if (buttonGroup && *static_cast<void**>(buttonGroup) == buttonGroupVtable) {
                itemCount += getNumItems(buttonGroup);
                ++populatedGroups;
            }
        }
        Wh_Log(L"EP #%u task created id=%u, groups=%d, scanned=%d, items=%d, desktop depth=%u, result=0x%08X",
               ++sequence, id, count, populatedGroups, itemCount, desktopUpdateDepth,
               static_cast<unsigned int>(result));
        if (taskList == savedTaskList && CanLog()) {
            if (desktopUpdateDepth) {
                lastObservedOrder = ReadOrder(taskList);
                observedDuringUpdate = true;
            } else {
                LogOrder(L"order after creation outside update", ReadOrder(taskList));
            }
        }
    }
    return result;
}

HRESULT WINAPI TaskDestroyedHook(void* taskList, void* group, void* item,
                                 unsigned int flags) {
    HRESULT result = taskDestroyedOriginal(taskList, group, item, flags);
    if (IsTaskbarThread() && CanLog()) {
        Wh_Log(L"EP #%u task destroyed, desktop depth=%u, result=0x%08X",
               ++sequence, desktopUpdateDepth,
               static_cast<unsigned int>(result));
    }
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
    if (ready || !module || (reinterpret_cast<ULONG_PTR>(module) & 3)) {
        return;
    }

    WCHAR path[MAX_PATH];
    DWORD length = GetModuleFileName(module, path, ARRAYSIZE(path));
    PCWSTR filename = length && length < ARRAYSIZE(path) ? wcsrchr(path, L'\\') : nullptr;
    if (!filename || _wcsicmp(filename + 1, L"ep_taskbar.ge.dll") != 0 ||
        installing.exchange(true)) {
        return;
    }

    if (!IsSupportedImage(module)) {
        Wh_Log(L"EP diagnostics skipped: ep_taskbar.ge.dll has an unrecognized image identity");
        return;
    }
    Wh_Log(L"EP diagnostics: ep_taskbar.ge.dll matched the guarded CodeView build");

    struct Hook {
        const char* name;
        void* hook;
        void** original;
        void* address;
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
        {"?TaskDestroyed@CTaskListWnd@@UEAAJPEAUITaskGroup@@PEAUITaskItem@@W4TaskDestroyedFlags@@@Z",
         reinterpret_cast<void*>(TaskDestroyedHook), reinterpret_cast<void**>(&taskDestroyedOriginal)},
    };
    auto resolve = [module]<typename T>(const char* name, T& target) {
        target = reinterpret_cast<T>(GetProcAddress(module, name));
        if (!target) {
            Wh_Log(L"EP diagnostics skipped: missing export %S", name);
        }
        return target != nullptr;
    };
    if (!resolve("??_7CTaskBtnGroup@@6BITaskBtnGroup@@@", buttonGroupVtable) ||
        !resolve("??_7CTaskGroup@@6BITaskGroup@@@", taskGroupVtable) ||
        !resolve("??_7CWindowTaskItem@@6BITaskItem@@@", windowItemVtable) ||
        !resolve("??_7CTaskListWnd@@6BITaskListUI@@@", taskListUiVtable) ||
        !resolve("?IsOnPrimaryTaskband@CTaskListWnd@@UEAAHXZ", isOnPrimaryTaskband) ||
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
            Wh_Log(L"EP diagnostics skipped: missing export %S", hook.name);
            return;
        }
    }
    size_t queued = 0;
    for (const auto& hook : hooks) {
        if (!Wh_SetFunctionHook(hook.address, hook.hook, hook.original)) {
            Wh_Log(L"EP diagnostics skipped: could not queue %S", hook.name);
            for (size_t i = 0; i < queued; ++i) {
                Wh_RemoveFunctionHook(hooks[i].address);
            }
            Wh_ApplyHookOperations();
            return;
        }
        ++queued;
    }
    ready = Wh_ApplyHookOperations();
    Wh_Log(L"EP diagnostics: hook apply=%d, restore=%d, taskbar thread=%u",
           ready.load(), restoreOrder, taskbarThread.load());
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
    } else {
        Wh_Log(L"EP diagnostics: waiting for the primary taskbar window");
    }
    canApplyHooks = true;
    HMODULE module = GetModuleHandle(L"ep_taskbar.ge.dll");
    if (module) {
        TryInstall(module);
    } else {
        Wh_Log(L"EP diagnostics: waiting for ep_taskbar.ge.dll");
    }
}

}  // namespace explorer_patcher

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

    explorer_patcher::diagnostics = Wh_GetIntSetting(L"explorerPatcherDiagnostics");
    explorer_patcher::restoreOrder = Wh_GetIntSetting(L"explorerPatcherRestoreOrder");
    if (explorer_patcher::diagnostics || explorer_patcher::restoreOrder) {
        return explorer_patcher::Init();
    }

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
        return FALSE;
    }

    if (!HookSymbols(taskbarModule, taskbarDllHooks,
                     ARRAYSIZE(taskbarDllHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
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
        return FALSE;
    }

    if (!HookSymbols(twinuiPcshellModule, twinuiPcshellSymbolHooks,
                     ARRAYSIZE(twinuiPcshellSymbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
    }

    WindhawkUtils::SetFunctionHook(DPA_InsertPtr, DPA_InsertPtr_Hook,
                                   &DPA_InsertPtr_Original);

    WindhawkUtils::SetFunctionHook(CreateWindowExW, CreateWindowExW_Hook,
                                   &CreateWindowExW_Original);

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (explorer_patcher::diagnostics || explorer_patcher::restoreOrder) {
        explorer_patcher::AfterInit();
        return;
    }

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        InitializeTaskbarVariables(hTaskbarWnd);
    }
}

BOOL Wh_ModSettingsChanged(BOOL* reload) {
    *reload =
        explorer_patcher::diagnostics !=
            (Wh_GetIntSetting(L"explorerPatcherDiagnostics") != 0) ||
        explorer_patcher::restoreOrder !=
            (Wh_GetIntSetting(L"explorerPatcherRestoreOrder") != 0);
    return TRUE;
}
