// ==WindhawkMod==
// @id              taskbar-thumbnail-reorder
// @name            Taskbar Thumbnail Reorder
// @description     Reorder taskbar thumbnails with the left mouse button
// @version         1.0.8
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
# Taskbar Thumbnail Reorder

Reorder taskbar thumbnails with the left mouse button.

Only Windows 10 64-bit and Windows 11 are supported. For other Windows versions
check out [7+ Taskbar Tweaker](https://tweaker.ramensoftware.com/).

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.

![demonstration](https://i.imgur.com/wGGe2RS.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool).
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <commctrl.h>
#include <psapi.h>

#include <atomic>

struct {
    bool oldTaskbarOnWin11;
} g_settings;

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_24H2,
};

WinVersion g_winVersion;

std::atomic<bool> g_initialized;
std::atomic<bool> g_explorerPatcherInitialized;

int g_thumbDraggedIndex = -1;
bool g_thumbDragDone;
bool g_taskItemFilterDisallowAll;
DWORD g_taskInclusionChangeLastTickCount;

std::atomic<DWORD> g_getPtr_captureForThreadId;
HDPA g_getPtr_lastHdpa;

// https://www.geoffchappell.com/studies/windows/shell/comctl32/api/da/dpa/dpa.htm
typedef struct _DPA {
    int cpItems;
    PVOID* pArray;
    HANDLE hHeap;
    int cpCapacity;
    int cpGrow;
} DPA, *HDPA;

void* QueryViaVtable(void* object, void* vtable) {
    void* ptr = object;
    while (*(void**)ptr != vtable) {
        ptr = (void**)ptr + 1;
    }
    return ptr;
}

void* QueryViaVtableBackwards(void* object, void* vtable) {
    void* ptr = object;
    while (*(void**)ptr != vtable) {
        ptr = (void**)ptr - 1;
    }
    return ptr;
}

void* CTaskListThumbnailWnd_vftable_1;
void* CTaskListWnd_vftable_ITaskListUI;

using CTaskListThumbnailWnd_GetHoverIndex_t = int(WINAPI*)(void* pThis);
CTaskListThumbnailWnd_GetHoverIndex_t CTaskListThumbnailWnd_GetHoverIndex;

using CTaskListThumbnailWnd__GetTaskItem_t = void*(WINAPI*)(void* pThis,
                                                            int index);
CTaskListThumbnailWnd__GetTaskItem_t CTaskListThumbnailWnd__GetTaskItem;

using CTaskListThumbnailWnd_GetTaskGroup_t = void*(WINAPI*)(void* pThis);
CTaskListThumbnailWnd_GetTaskGroup_t CTaskListThumbnailWnd_GetTaskGroup;

using CTaskListThumbnailWnd_TaskReordered_t = void(WINAPI*)(void* pThis,
                                                            void* taskItem);
CTaskListThumbnailWnd_TaskReordered_t CTaskListThumbnailWnd_TaskReordered;

using CTaskGroup_GetNumItems_t = int(WINAPI*)(PVOID pThis);
CTaskGroup_GetNumItems_t CTaskGroup_GetNumItems;

using CTaskListWnd__GetTBGroupFromGroup_t =
    void*(WINAPI*)(void* pThis, void* pTaskGroup, int* pTaskBtnGroupIndex);
CTaskListWnd__GetTBGroupFromGroup_t CTaskListWnd__GetTBGroupFromGroup;

using CTaskBtnGroup_GetGroupType_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetGroupType_t CTaskBtnGroup_GetGroupType;

using CTaskBtnGroup_IndexOfTaskItem_t = int(WINAPI*)(void* pThis,
                                                     void* taskItem);
CTaskBtnGroup_IndexOfTaskItem_t CTaskBtnGroup_IndexOfTaskItem;

using CTaskListWnd_TaskInclusionChanged_t = HRESULT(WINAPI*)(void* pThis,
                                                             void* pTaskGroup,
                                                             void* pTaskItem);
CTaskListWnd_TaskInclusionChanged_t CTaskListWnd_TaskInclusionChanged;

using TaskItemFilter_IsTaskAllowed_t = bool(WINAPI*)(void* pThis,
                                                     void* pTaskItem);
TaskItemFilter_IsTaskAllowed_t TaskItemFilter_IsTaskAllowed_Original;
bool WINAPI TaskItemFilter_IsTaskAllowed_Hook(void* pThis, void* pTaskItem) {
    Wh_Log(L">");

    if (g_taskItemFilterDisallowAll) {
        return false;
    }

    return TaskItemFilter_IsTaskAllowed_Original(pThis, pTaskItem);
}

bool MoveTaskInTaskList(HWND hMMTaskListWnd,
                        void* lpMMTaskListLongPtr,
                        void* taskGroup,
                        void* taskItemFrom,
                        void* taskItemTo) {
    void* taskBtnGroup = CTaskListWnd__GetTBGroupFromGroup(lpMMTaskListLongPtr,
                                                           taskGroup, nullptr);
    if (!taskBtnGroup) {
        return false;
    }

    int taskBtnGroupType = CTaskBtnGroup_GetGroupType(taskBtnGroup);
    if (taskBtnGroupType != 1 && taskBtnGroupType != 3) {
        return false;
    }

    g_getPtr_lastHdpa = nullptr;

    g_getPtr_captureForThreadId = GetCurrentThreadId();
    int indexFrom = CTaskBtnGroup_IndexOfTaskItem(taskBtnGroup, taskItemFrom);
    g_getPtr_captureForThreadId = 0;

    if (indexFrom == -1) {
        return false;
    }

    HDPA buttonsArray = g_getPtr_lastHdpa;
    if (!buttonsArray) {
        return false;
    }

    g_getPtr_lastHdpa = nullptr;

    g_getPtr_captureForThreadId = GetCurrentThreadId();
    int indexTo = CTaskBtnGroup_IndexOfTaskItem(taskBtnGroup, taskItemTo);
    g_getPtr_captureForThreadId = 0;

    if (indexTo == -1) {
        return false;
    }

    if (g_getPtr_lastHdpa != buttonsArray) {
        return false;
    }

    void* button = (void*)DPA_DeletePtr(buttonsArray, indexFrom);
    if (!button) {
        return false;
    }

    DPA_InsertPtr(buttonsArray, indexTo, button);

    if (g_winVersion <= WinVersion::Win10) {
        InvalidateRect(hMMTaskListWnd, nullptr, FALSE);
    } else {
        g_taskInclusionChangeLastTickCount = GetTickCount();

        HWND hThumbnailWnd = GetCapture();
        if (hThumbnailWnd) {
            WCHAR szClassName[32];
            if (GetClassName(hThumbnailWnd, szClassName,
                             ARRAYSIZE(szClassName)) == 0 ||
                _wcsicmp(szClassName, L"TaskListThumbnailWnd") != 0) {
                hThumbnailWnd = nullptr;
            }
        }

        if (hThumbnailWnd) {
            SendMessage(hThumbnailWnd, WM_SETREDRAW, FALSE, 0);
        }

        g_taskItemFilterDisallowAll = true;

        void* pThis_ITaskListUI = QueryViaVtable(
            lpMMTaskListLongPtr, CTaskListWnd_vftable_ITaskListUI);

        CTaskListWnd_TaskInclusionChanged(pThis_ITaskListUI, taskGroup,
                                          taskItemFrom);

        g_taskItemFilterDisallowAll = false;

        CTaskListWnd_TaskInclusionChanged(pThis_ITaskListUI, taskGroup,
                                          taskItemFrom);

        if (hThumbnailWnd) {
            SendMessage(hThumbnailWnd, WM_SETREDRAW, TRUE, 0);
            RedrawWindow(
                hThumbnailWnd, nullptr, nullptr,
                RDW_ERASE | RDW_FRAME | RDW_INVALIDATE | RDW_ALLCHILDREN);
        }
    }

    return true;
}

HDPA GetTaskItemsArray(void* taskGroup) {
    // This is a horrible hack, but it's the best way I found to get the array
    // of task items from a task group. It relies on the implementation of
    // CTaskGroup::GetNumItems being just this:
    //
    // return DPA_GetPtrCount(this->taskItemsArray);
    //
    // Or in other words:
    //
    // return *(int*)this[taskItemsArrayOffset];
    //
    // Instead of calling it with a real taskGroup object, we call it with an
    // array of pointers to ints. The returned int value is actually the offset
    // to the array member.

    static size_t offset = []() {
        constexpr int kIntArraySize = 256;
        int arrayOfInts[kIntArraySize];
        int* arrayOfIntPtrs[kIntArraySize];
        for (int i = 0; i < kIntArraySize; i++) {
            arrayOfInts[i] = i;
            arrayOfIntPtrs[i] = &arrayOfInts[i];
        }

        return CTaskGroup_GetNumItems(arrayOfIntPtrs);
    }();

    return (HDPA)((void**)taskGroup)[offset];
}

bool MoveTaskInGroup(void* taskGroup, void* taskItemFrom, void* taskItemTo) {
    HDPA taskItemsArray = GetTaskItemsArray(taskGroup);
    if (!taskItemsArray) {
        return false;
    }

    int taskItemsCount = taskItemsArray->cpItems;
    void** taskItems = (void**)taskItemsArray->pArray;

    int indexFrom = -1;
    for (int i = 0; i < taskItemsCount; i++) {
        if (taskItems[i] == taskItemFrom) {
            indexFrom = i;
            break;
        }
    }

    if (indexFrom == -1) {
        return false;
    }

    int indexTo = -1;
    for (int i = 0; i < taskItemsCount; i++) {
        if (taskItems[i] == taskItemTo) {
            indexTo = i;
            break;
        }
    }

    if (indexTo == -1) {
        return false;
    }

    void* taskItemTemp = taskItems[indexFrom];
    if (indexFrom < indexTo) {
        memmove(&taskItems[indexFrom], &taskItems[indexFrom + 1],
                (indexTo - indexFrom) * sizeof(void*));
    } else {
        memmove(&taskItems[indexTo + 1], &taskItems[indexTo],
                (indexFrom - indexTo) * sizeof(void*));
    }
    taskItems[indexTo] = taskItemTemp;

    auto taskbarEnumProc = [taskGroup, taskItemFrom, taskItemTo](
                               HWND hMMTaskbarWnd, bool secondary) {
        HWND hMMTaskSwWnd;
        if (!secondary) {
            hMMTaskSwWnd = (HWND)GetProp(hMMTaskbarWnd, L"TaskbandHWND");
        } else {
            hMMTaskSwWnd =
                FindWindowEx(hMMTaskbarWnd, nullptr, L"WorkerW", nullptr);
        }

        if (!hMMTaskSwWnd) {
            return;
        }

        HWND hMMTaskListWnd =
            FindWindowEx(hMMTaskSwWnd, nullptr, L"MSTaskListWClass", nullptr);
        if (!hMMTaskListWnd) {
            return;
        }

        void* lpMMTaskListLongPtr = (void*)GetWindowLongPtr(hMMTaskListWnd, 0);

        MoveTaskInTaskList(hMMTaskListWnd, lpMMTaskListLongPtr, taskGroup,
                           taskItemFrom, taskItemTo);
    };

    EnumThreadWindows(
        GetCurrentThreadId(),
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            WCHAR szClassName[32];
            if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
                return TRUE;
            }

            bool secondary;
            if (_wcsicmp(szClassName, L"Shell_TrayWnd") == 0) {
                secondary = false;
            } else if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
                secondary = true;
            } else {
                return TRUE;
            }

            auto& proc = *reinterpret_cast<decltype(taskbarEnumProc)*>(lParam);
            proc(hWnd, secondary);
            return TRUE;
        },
        reinterpret_cast<LPARAM>(&taskbarEnumProc));

    return true;
}

bool MoveItemsFromThumbnail(void* lpMMThumbnailLongPtr,
                            int indexFrom,
                            int indexTo) {
    Wh_Log(L">");

    void* taskItemFrom =
        CTaskListThumbnailWnd__GetTaskItem(lpMMThumbnailLongPtr, indexFrom);
    if (!taskItemFrom) {
        return false;
    }

    void* taskItemTo =
        CTaskListThumbnailWnd__GetTaskItem(lpMMThumbnailLongPtr, indexTo);
    if (!taskItemTo) {
        return false;
    }

    if (!MoveTaskInGroup(
            CTaskListThumbnailWnd_GetTaskGroup(lpMMThumbnailLongPtr),
            taskItemFrom, taskItemTo)) {
        return false;
    }

    CTaskListThumbnailWnd_TaskReordered(lpMMThumbnailLongPtr, taskItemFrom);

    return true;
}

using CTaskListThumbnailWnd_v_WndProc_t = LRESULT(
    WINAPI*)(void* pThis, HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
CTaskListThumbnailWnd_v_WndProc_t CTaskListThumbnailWnd_v_WndProc_Original;
LRESULT WINAPI CTaskListThumbnailWnd_v_WndProc_Hook(void* pThis,
                                                    HWND hWnd,
                                                    UINT Msg,
                                                    WPARAM wParam,
                                                    LPARAM lParam) {
    LRESULT result = 0;
    bool processed = false;

    auto OriginalProc = [pThis](HWND hWnd, UINT Msg, WPARAM wParam,
                                LPARAM lParam) {
        return CTaskListThumbnailWnd_v_WndProc_Original(pThis, hWnd, Msg,
                                                        wParam, lParam);
    };

    switch (Msg) {
        case WM_LBUTTONDOWN:
            result = OriginalProc(hWnd, Msg, wParam, lParam);

            {
                void* lpMMThumbnailLongPtr = (void*)GetWindowLongPtr(hWnd, 0);
                void* lpMMThumbnailLongPtr_vftable_1 = QueryViaVtable(
                    lpMMThumbnailLongPtr, CTaskListThumbnailWnd_vftable_1);
                int trackedIndex = CTaskListThumbnailWnd_GetHoverIndex(
                    lpMMThumbnailLongPtr_vftable_1);
                if (trackedIndex >= 0) {
                    g_thumbDragDone = false;
                    g_thumbDraggedIndex = trackedIndex;
                    SetCapture(hWnd);
                }
            }
            break;

        case WM_MOUSEMOVE:
            result = OriginalProc(hWnd, Msg, wParam, lParam);

            if (GetCapture() == hWnd &&
                GetTickCount() - g_taskInclusionChangeLastTickCount > 200) {
                void* lpMMThumbnailLongPtr = (void*)GetWindowLongPtr(hWnd, 0);
                void* lpMMThumbnailLongPtr_vftable_1 = QueryViaVtable(
                    lpMMThumbnailLongPtr, CTaskListThumbnailWnd_vftable_1);
                int trackedIndex = CTaskListThumbnailWnd_GetHoverIndex(
                    lpMMThumbnailLongPtr_vftable_1);
                if (trackedIndex != g_thumbDraggedIndex &&
                    MoveItemsFromThumbnail(lpMMThumbnailLongPtr,
                                           g_thumbDraggedIndex, trackedIndex)) {
                    g_thumbDraggedIndex = trackedIndex;
                    g_thumbDragDone = true;
                }
            }
            break;

        case WM_LBUTTONUP:
            if (GetCapture() == hWnd) {
                ReleaseCapture();
                if (g_thumbDragDone) {
                    result = DefWindowProc(hWnd, Msg, wParam, lParam);
                    processed = true;
                }
            }

            if (!processed) {
                result = OriginalProc(hWnd, Msg, wParam, lParam);
            }
            break;

        case WM_TIMER:
            // Aero peek window of hovered thumbnail.
            if (wParam == 2006) {
                if (GetCapture() == hWnd) {
                    KillTimer(hWnd, wParam);
                    result = DefWindowProc(hWnd, Msg, wParam, lParam);
                    processed = true;
                }
            }

            if (!processed) {
                result = OriginalProc(hWnd, Msg, wParam, lParam);
            }
            break;

        default:
            result = OriginalProc(hWnd, Msg, wParam, lParam);
            break;
    }

    return result;
}

using DPA_GetPtr_t = decltype(&DPA_GetPtr);
DPA_GetPtr_t DPA_GetPtr_Original;
PVOID WINAPI DPA_GetPtr_Hook(HDPA hdpa, INT_PTR i) {
    if (g_getPtr_captureForThreadId == GetCurrentThreadId()) {
        g_getPtr_lastHdpa = hdpa;
    }

    return DPA_GetPtr_Original(hdpa, i);
}

bool HookTaskbarSymbols() {
    HMODULE module;
    if (g_winVersion <= WinVersion::Win10) {
        module = GetModuleHandle(nullptr);
    } else {
        module = LoadLibrary(L"taskbar.dll");
        if (!module) {
            Wh_Log(L"Couldn't load taskbar.dll");
            return false;
        }
    }

    // Taskbar.dll, explorer.exe
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {
                // Windows 11.
                LR"(const CTaskListThumbnailWnd::`vftable'{for `Microsoft::WRL::Details::ImplementsHelper<struct Microsoft::WRL::RuntimeClassFlags<2>,1,struct IExtendedUISwitcher,struct IDropTarget,struct ITaskThumbnailHost>'})",

                // Windows 10.
                LR"(const CTaskListThumbnailWnd::`vftable'{for `Microsoft::WRL::Details::ImplementsHelper<struct Microsoft::WRL::RuntimeClassFlags<2>,1,struct IExtendedUISwitcher,struct IDropTarget,struct ITaskThumbnailHost,struct IUIAnimationStoryboardEventHandler,struct IIconLoaderNotifications2>'})",
            },
            &CTaskListThumbnailWnd_vftable_1,
        },
        {
            {LR"(const CTaskListWnd::`vftable'{for `ITaskListUI'})"},
            &CTaskListWnd_vftable_ITaskListUI,
        },
        {
            {LR"(public: virtual int __cdecl CTaskListThumbnailWnd::GetHoverIndex(void)const )"},
            &CTaskListThumbnailWnd_GetHoverIndex,
        },
        {
            {
                // Windows 11.
                LR"(private: struct ITaskItem * __cdecl CTaskListThumbnailWnd::_GetTaskItem(int)const )",

                // Windows 10.
                LR"(private: struct ITaskItem * __cdecl CTaskListThumbnailWnd::_GetTaskItem(int))",
            },
            &CTaskListThumbnailWnd__GetTaskItem,
        },
        {
            {LR"(public: virtual struct ITaskGroup * __cdecl CTaskListThumbnailWnd::GetTaskGroup(void)const )"},
            &CTaskListThumbnailWnd_GetTaskGroup,
        },
        {
            {LR"(public: virtual void __cdecl CTaskListThumbnailWnd::TaskReordered(struct ITaskItem *))"},
            &CTaskListThumbnailWnd_TaskReordered,
        },
        {
            {LR"(public: virtual int __cdecl CTaskGroup::GetNumItems(void))"},
            &CTaskGroup_GetNumItems,
        },
        {
            {LR"(protected: struct ITaskBtnGroup * __cdecl CTaskListWnd::_GetTBGroupFromGroup(struct ITaskGroup *,int *))"},
            &CTaskListWnd__GetTBGroupFromGroup,
        },
        {
            {LR"(public: virtual enum eTBGROUPTYPE __cdecl CTaskBtnGroup::GetGroupType(void))"},
            &CTaskBtnGroup_GetGroupType,
        },
        {
            {LR"(public: virtual int __cdecl CTaskBtnGroup::IndexOfTaskItem(struct ITaskItem *))"},
            &CTaskBtnGroup_IndexOfTaskItem,
        },
        {
            {LR"(public: virtual long __cdecl CTaskListWnd::TaskInclusionChanged(struct ITaskGroup *,struct ITaskItem *))"},
            &CTaskListWnd_TaskInclusionChanged,
        },
        {
            {LR"(public: virtual bool __cdecl TaskItemFilter::IsTaskAllowed(struct ITaskItem *))"},
            &TaskItemFilter_IsTaskAllowed_Original,
            TaskItemFilter_IsTaskAllowed_Hook,
        },
        {
            {LR"(private: virtual __int64 __cdecl CTaskListThumbnailWnd::v_WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64))"},
            &CTaskListThumbnailWnd_v_WndProc_Original,
            CTaskListThumbnailWnd_v_WndProc_Hook,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
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
            } else if (build < 26100) {
                return WinVersion::Win11;
            } else {
                return WinVersion::Win11_24H2;
            }
            break;
    }

    return WinVersion::Unsupported;
}

struct EXPLORER_PATCHER_HOOK {
    PCSTR symbol;
    void** pOriginalFunction;
    void* hookFunction = nullptr;
    bool optional = false;

    template <typename Prototype>
    EXPLORER_PATCHER_HOOK(
        PCSTR symbol,
        Prototype** originalFunction,
        std::type_identity_t<Prototype*> hookFunction = nullptr,
        bool optional = false)
        : symbol(symbol),
          pOriginalFunction(reinterpret_cast<void**>(originalFunction)),
          hookFunction(reinterpret_cast<void*>(hookFunction)),
          optional(optional) {}
};

bool HookExplorerPatcherSymbols(HMODULE explorerPatcherModule) {
    if (g_explorerPatcherInitialized.exchange(true)) {
        return true;
    }

    if (g_winVersion >= WinVersion::Win11) {
        g_winVersion = WinVersion::Win10;
    }

    EXPLORER_PATCHER_HOOK hooks[] = {
        {R"(??_7CTaskListThumbnailWnd@@6B?$ImplementsHelper@U?$RuntimeClassFlags@$01@WRL@Microsoft@@$00UIDropTarget@@UITaskThumbnailHost@@UIUIAnimationStoryboardEventHandler2@@@Details@WRL@Microsoft@@@)",
         &CTaskListThumbnailWnd_vftable_1},
        {R"(??_7CTaskListWnd@@6BITaskListUI@@@)",
         &CTaskListWnd_vftable_ITaskListUI},
        {R"(?GetHoverIndex@CTaskListThumbnailWnd@@UEBAHXZ)",
         &CTaskListThumbnailWnd_GetHoverIndex},
        {R"(?_GetTaskItem@CTaskListThumbnailWnd@@AEAAPEAUITaskItem@@H@Z)",
         &CTaskListThumbnailWnd__GetTaskItem},
        {R"(?GetTaskGroup@CTaskListThumbnailWnd@@UEBAPEAUITaskGroup@@XZ)",
         &CTaskListThumbnailWnd_GetTaskGroup},
        {R"(?TaskReordered@CTaskListThumbnailWnd@@UEAAXPEAUITaskItem@@@Z)",
         &CTaskListThumbnailWnd_TaskReordered},
        {R"(?GetNumItems@CTaskGroup@@UEAAHXZ)", &CTaskGroup_GetNumItems},
        {R"(?_GetTBGroupFromGroup@CTaskListWnd@@IEAAPEAUITaskBtnGroup@@PEAUITaskGroup@@PEAH@Z)",
         &CTaskListWnd__GetTBGroupFromGroup},
        {R"(?GetGroupType@CTaskBtnGroup@@UEAA?AW4eTBGROUPTYPE@@XZ)",
         &CTaskBtnGroup_GetGroupType},
        {R"(?IndexOfTaskItem@CTaskBtnGroup@@UEAAHPEAUITaskItem@@@Z)",
         &CTaskBtnGroup_IndexOfTaskItem},
        {R"(?TaskInclusionChanged@CTaskListWnd@@UEAAJPEAUITaskGroup@@PEAUITaskItem@@@Z)",
         &CTaskListWnd_TaskInclusionChanged},
        {R"(?IsTaskAllowed@TaskItemFilter@@UEAA_NPEAUITaskItem@@@Z)",
         &TaskItemFilter_IsTaskAllowed_Original,
         TaskItemFilter_IsTaskAllowed_Hook},
        {R"(?v_WndProc@CTaskListThumbnailWnd@@EEAA_JPEAUHWND__@@I_K_J@Z)",
         &CTaskListThumbnailWnd_v_WndProc_Original,
         CTaskListThumbnailWnd_v_WndProc_Hook},
    };

    bool succeeded = true;

    for (const auto& hook : hooks) {
        void* ptr = (void*)GetProcAddress(explorerPatcherModule, hook.symbol);
        if (!ptr) {
            Wh_Log(L"ExplorerPatcher symbol%s doesn't exist: %S",
                   hook.optional ? L" (optional)" : L"", hook.symbol);
            if (!hook.optional) {
                succeeded = false;
            }
            continue;
        }

        if (hook.hookFunction) {
            Wh_SetFunctionHook(ptr, hook.hookFunction, hook.pOriginalFunction);
        } else {
            *hook.pOriginalFunction = ptr;
        }
    }

    if (!succeeded) {
        Wh_Log(L"HookExplorerPatcherSymbols failed");
    } else if (g_initialized) {
        Wh_ApplyHookOperations();
    }

    return succeeded;
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

bool HandleLoadedExplorerPatcher() {
    HMODULE hMods[1024];
    DWORD cbNeeded;
    if (EnumProcessModules(GetCurrentProcess(), hMods, sizeof(hMods),
                           &cbNeeded)) {
        for (size_t i = 0; i < cbNeeded / sizeof(HMODULE); i++) {
            if (IsExplorerPatcherModule(hMods[i])) {
                return HookExplorerPatcherSymbols(hMods[i]);
            }
        }
    }

    return true;
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module && !((ULONG_PTR)module & 3) && !g_explorerPatcherInitialized) {
        if (IsExplorerPatcherModule(module)) {
            HookExplorerPatcherSymbols(module);
        }
    }

    return module;
}

void LoadSettings() {
    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    g_winVersion = GetExplorerVersion();
    if (g_winVersion == WinVersion::Unsupported) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
    }

    if (g_settings.oldTaskbarOnWin11) {
        bool hasWin10Taskbar = g_winVersion < WinVersion::Win11_24H2;

        if (g_winVersion >= WinVersion::Win11) {
            g_winVersion = WinVersion::Win10;
        }

        if (hasWin10Taskbar && !HookTaskbarSymbols()) {
            return FALSE;
        }
    } else if (g_winVersion >= WinVersion::Win11) {
        // if (!HookTaskbarViewDllSymbols()) {
        //     return FALSE;
        // }

        if (!HookTaskbarSymbols()) {
            return FALSE;
        }
    } else {
        if (!HookTaskbarSymbols()) {
            return FALSE;
        }
    }

    if (!HandleLoadedExplorerPatcher()) {
        Wh_Log(L"HandleLoadedExplorerPatcher failed");
        return FALSE;
    }

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(
        kernelBaseModule, "LoadLibraryExW");
    WindhawkUtils::Wh_SetFunctionHookT(pKernelBaseLoadLibraryExW,
                                       LoadLibraryExW_Hook,
                                       &LoadLibraryExW_Original);

    WindhawkUtils::Wh_SetFunctionHookT(DPA_GetPtr, DPA_GetPtr_Hook,
                                       &DPA_GetPtr_Original);

    g_initialized = true;

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    // Try again in case there's a race between the previous attempt and the
    // LoadLibraryExW hook.
    if (!g_explorerPatcherInitialized) {
        HandleLoadedExplorerPatcher();
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;

    LoadSettings();

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11;

    return TRUE;
}
