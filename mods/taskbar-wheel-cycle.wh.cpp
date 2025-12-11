// ==WindhawkMod==
// @id              taskbar-wheel-cycle
// @name            Cycle taskbar buttons with mouse wheel
// @description     Use the mouse wheel and/or keyboard shortcuts to cycle between taskbar buttons
// @version         1.2.0
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lcomctl32 -lole32 -loleaut32 -lruntimeobject -lversion
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
# Cycle taskbar buttons with mouse wheel

Use the mouse wheel while hovering over the taskbar to cycle between taskbar
buttons.

In addition, keyboard shortcuts can be used. The default shortcuts are `Alt+[`
and `Alt+]`, but they can be changed in the mod settings.

## Enhanced Features

**Same-Process Cycling**: Use `Ctrl+Alt+[` and `Ctrl+Alt+]` to cycle only
between windows of the same application (e.g., switch between multiple Chrome
windows or multiple Cursor windows).

**Cross-Process Navigation**: The standard `Alt+[` and `Alt+]` shortcuts now
intelligently skip windows from the same process, making it easier to switch
between different applications.

**Taskbar Button Reordering**: Use `Ctrl+Shift+;` and `Ctrl+Shift+'` to move
the active window's button group left or right in the taskbar, allowing you to
organize your taskbar buttons.

**Multiple Shortcut Sets**: Configure a second set of keyboard shortcuts for all
cycling operations, providing more flexibility in your workflow.

Only Windows 10 64-bit and Windows 11 are supported. For older Windows versions
check out [7+ Taskbar Tweaker](https://tweaker.ramensoftware.com/).

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.

![Demonstration](https://i.imgur.com/FtpUjt1.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- wrapAround: true
  $name: Wrap around
  $description: Wrap around when reaching the end of the taskbar
- reverseScrollingDirection: false
  $name: Reverse scrolling direction
  $description: Reverse the scrolling direction of the mouse wheel
- enableMouseWheelCycling: true
  $name: Enable mouse wheel cycling
  $description: Disable to only use keyboard shortcuts for cycling between taskbar buttons
- cycleLeftKeyboardShortcut: Alt+VK_OEM_4
  $name: Cycle left keyboard shortcut (cross-process)
  $description: >-
    Cycle between windows, skipping windows from the same process.
    Possible modifier keys: Alt, Ctrl, Shift, Win. For possible shortcut keys,
    refer to the following page:
    https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes

    Set to an empty string to disable.
- cycleRightKeyboardShortcut: Alt+VK_OEM_6
  $name: Cycle right keyboard shortcut (cross-process)
  $description: Cycle between windows, skipping windows from the same process. Set to an empty string to disable.
- cycleSameProcessLeftKeyboardShortcut: Ctrl+Alt+VK_OEM_4
  $name: Cycle left keyboard shortcut (same process)
  $description: >-
    Cycle only between windows of the same process.
    For example: switch between 3 Cursor windows.

    Set to an empty string to disable.
- cycleSameProcessRightKeyboardShortcut: Ctrl+Alt+VK_OEM_6
  $name: Cycle right keyboard shortcut (same process)
  $description: >-
    Cycle only between windows of the same process.
    For example: switch between 3 Cursor windows.

    Set to an empty string to disable.
- cycleLeftKeyboardShortcut2: ""
  $name: Cycle left keyboard shortcut 2 (cross-process)
  $description: Second set of shortcuts, same functionality as "Cycle left keyboard shortcut". Set to an empty string to disable.
- cycleRightKeyboardShortcut2: ""
  $name: Cycle right keyboard shortcut 2 (cross-process)
  $description: Second set of shortcuts, same functionality as "Cycle right keyboard shortcut". Set to an empty string to disable.
- cycleSameProcessLeftKeyboardShortcut2: ""
  $name: Cycle left keyboard shortcut 2 (same process)
  $description: Second set of shortcuts, same functionality as "Cycle left keyboard shortcut (same process)". Set to an empty string to disable.
- cycleSameProcessRightKeyboardShortcut2: ""
  $name: Cycle right keyboard shortcut 2 (same process)
  $description: Second set of shortcuts, same functionality as "Cycle right keyboard shortcut (same process)". Set to an empty string to disable.
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: Enable this option to customize the old taskbar on Windows 11 (if using ExplorerPatcher or a similar tool).
- moveButtonGroupLeftShortcut: Ctrl+Shift+VK_OEM_1
  $name: Move button group left shortcut
  $description: >-
    Move the active window's button group one position to the left in the taskbar.
    For example: Ctrl+Shift+; moves Cursor to the left of Wezterm.

    Set to an empty string to disable.
- moveButtonGroupRightShortcut: Ctrl+Shift+VK_OEM_7
  $name: Move button group right shortcut
  $description: >-
    Move the active window's button group one position to the right in the taskbar.
    For example: Ctrl+Shift+' moves Cursor to the right.

    Set to an empty string to disable.
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <commctrl.h>
#include <psapi.h>
#include <windowsx.h>

#undef GetCurrentTime

#include <winrt/Windows.UI.Input.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Input.h>

#include <algorithm>
#include <atomic>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

using namespace winrt::Windows::UI::Xaml;

struct {
    bool wrapAround;
    bool reverseScrollingDirection;
    bool enableMouseWheelCycling;
    WindhawkUtils::StringSetting cycleLeftKeyboardShortcut;
    WindhawkUtils::StringSetting cycleRightKeyboardShortcut;
    WindhawkUtils::StringSetting cycleSameProcessLeftKeyboardShortcut;
    WindhawkUtils::StringSetting cycleSameProcessRightKeyboardShortcut;
    WindhawkUtils::StringSetting cycleLeftKeyboardShortcut2;
    WindhawkUtils::StringSetting cycleRightKeyboardShortcut2;
    WindhawkUtils::StringSetting cycleSameProcessLeftKeyboardShortcut2;
    WindhawkUtils::StringSetting cycleSameProcessRightKeyboardShortcut2;
    bool oldTaskbarOnWin11;
    WindhawkUtils::StringSetting moveButtonGroupLeftShortcut;
    WindhawkUtils::StringSetting moveButtonGroupRightShortcut;
} g_settings;

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_24H2,
};

WinVersion g_winVersion;

std::atomic<bool> g_taskbarViewDllLoaded;
std::atomic<bool> g_initialized;
std::atomic<bool> g_explorerPatcherInitialized;

// DPA structure for button group reordering
// https://www.geoffchappell.com/studies/windows/shell/comctl32/api/da/dpa/dpa.htm
typedef struct _DPA {
    int cpItems;
    PVOID* pArray;
    HANDLE hHeap;
    int cpCapacity;
    int cpGrow;
} DPA, *HDPA;

std::atomic<DWORD> g_getPtr_captureForThreadId;
HDPA g_getPtr_lastHdpa;

struct TaskBtnGroupButtonInfo {
    void* taskBtnGroup;
    int buttonIndex;
};

std::unordered_map<void*, TaskBtnGroupButtonInfo> g_lastTaskListActiveItem;

HWND g_lastScrollTarget = nullptr;
DWORD g_lastScrollTime;
short g_lastScrollDeltaRemainder;

bool g_hotkeyLeftRegistered = false;
bool g_hotkeyRightRegistered = false;
bool g_hotkeySameProcessLeftRegistered = false;
bool g_hotkeySameProcessRightRegistered = false;
bool g_hotkeyLeft2Registered = false;
bool g_hotkeyRight2Registered = false;
bool g_hotkeySameProcessLeft2Registered = false;
bool g_hotkeySameProcessRight2Registered = false;
bool g_hotkeyMoveGroupLeftRegistered = false;
bool g_hotkeyMoveGroupRightRegistered = false;

enum {
    kHotkeyIdLeft = 1682530408,  // From epochconverter.com
    kHotkeyIdRight,
    kHotkeyIdSameProcessLeft,
    kHotkeyIdSameProcessRight,
    kHotkeyIdLeft2,
    kHotkeyIdRight2,
    kHotkeyIdSameProcessLeft2,
    kHotkeyIdSameProcessRight2,
    kHotkeyIdMoveGroupLeft,
    kHotkeyIdMoveGroupRight,
};

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

HWND GetTaskBandWnd() {
    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        return (HWND)GetProp(hTaskbarWnd, L"TaskbandHWND");
    }

    return nullptr;
}

// Forward declarations
HWND TaskListFromMMTaskbarWnd(HWND hMMTaskbarWnd);
HWND GetTaskbarForMonitor(HWND hTaskbarWnd, HMONITOR monitor);

void* CTaskListWnd_vftable_ITaskListUI;
void* CTaskListWnd_vftable_ITaskListSite;
void* CTaskListWnd_vftable_ITaskListAcc;
void* CImmersiveTaskItem_vftable;

using CTaskListWnd_GetButtonGroupCount_t = int(WINAPI*)(void* pThis);
CTaskListWnd_GetButtonGroupCount_t CTaskListWnd_GetButtonGroupCount;

using CTaskListWnd__GetTBGroupFromGroup_t = void*(WINAPI*)(void* pThis,
                                                           void* taskGroup,
                                                           int* index);
CTaskListWnd__GetTBGroupFromGroup_t CTaskListWnd__GetTBGroupFromGroup;

using CTaskBtnGroup_GetGroupType_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetGroupType_t CTaskBtnGroup_GetGroupType;

using CTaskBtnGroup_GetNumItems_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetNumItems_t CTaskBtnGroup_GetNumItems;

using CTaskBtnGroup_GetTaskItem_t = void*(WINAPI*)(void* pThis, int index);
CTaskBtnGroup_GetTaskItem_t CTaskBtnGroup_GetTaskItem;

using CWindowTaskItem_GetWindow_t = HWND(WINAPI*)(PVOID pThis);
CWindowTaskItem_GetWindow_t CWindowTaskItem_GetWindow_Original;

using CImmersiveTaskItem_GetWindow_t = HWND(WINAPI*)(PVOID pThis);
CImmersiveTaskItem_GetWindow_t CImmersiveTaskItem_GetWindow_Original;

using CTaskListWnd_SwitchToItem_t = void(WINAPI*)(void* pThis, void* taskItem);
CTaskListWnd_SwitchToItem_t CTaskListWnd_SwitchToItem_Original;

using CTaskBtnGroup_IndexOfTaskItem_t = int(WINAPI*)(void* pThis, void* taskItem);
CTaskBtnGroup_IndexOfTaskItem_t CTaskBtnGroup_IndexOfTaskItem;

using CTaskListWnd_TaskInclusionChanged_t = HRESULT(WINAPI*)(void* pThis,
                                                             void* pTaskGroup,
                                                             void* pTaskItem);
CTaskListWnd_TaskInclusionChanged_t CTaskListWnd_TaskInclusionChanged;

using DPA_GetPtr_t = decltype(&DPA_GetPtr);
DPA_GetPtr_t DPA_GetPtr_Original;
PVOID WINAPI DPA_GetPtr_Hook(HDPA hdpa, INT_PTR i) {
    if (g_getPtr_captureForThreadId == GetCurrentThreadId()) {
        g_getPtr_lastHdpa = hdpa;
    }
    return DPA_GetPtr_Original(hdpa, i);
}

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

#pragma region scroll

void SwitchToTaskItem(LONG_PTR lpMMTaskListLongPtr, void* taskItem) {
    void* pThis_ITaskListSite = QueryViaVtable(
        (void*)lpMMTaskListLongPtr, CTaskListWnd_vftable_ITaskListSite);

    CTaskListWnd_SwitchToItem_Original(pThis_ITaskListSite, taskItem);
}

HWND GetTaskItemWnd(PVOID taskItem) {
    if (*(void**)taskItem == CImmersiveTaskItem_vftable) {
        return CImmersiveTaskItem_GetWindow_Original(taskItem);
    } else {
        return CWindowTaskItem_GetWindow_Original(taskItem);
    }
}

BOOL IsMinimizedTaskItem(LONG_PTR* task_item) {
    return IsIconic(GetTaskItemWnd(task_item));
}

// Get process name from window handle
std::wstring GetProcessNameFromWindow(HWND hWnd) {
    if (!hWnd) {
        return L"";
    }

    DWORD processId = 0;
    GetWindowThreadProcessId(hWnd, &processId);
    if (!processId) {
        return L"";
    }

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, processId);
    if (!hProcess) {
        return L"";
    }

    WCHAR processPath[MAX_PATH] = {0};
    DWORD size = MAX_PATH;
    if (QueryFullProcessImageName(hProcess, 0, processPath, &size)) {
        CloseHandle(hProcess);
        // Extract just the filename from the full path
        WCHAR* fileName = wcsrchr(processPath, L'\\');
        if (fileName) {
            return std::wstring(fileName + 1);
        }
        return std::wstring(processPath);
    }

    CloseHandle(hProcess);
    return L"";
}

// Compare two windows by Z-order (returns true if hwnd1 is above hwnd2)
bool IsWindowAboveInZOrder(HWND hwnd1, HWND hwnd2) {
    if (!hwnd1 || !hwnd2 || hwnd1 == hwnd2) {
        return false;
    }

    // Enumerate windows in Z-order from top to bottom
    HWND hWnd = GetTopWindow(nullptr);
    while (hWnd) {
        if (hWnd == hwnd1) {
            return true;  // hwnd1 is above hwnd2
        }
        if (hWnd == hwnd2) {
            return false; // hwnd2 is above hwnd1
        }
        hWnd = GetWindow(hWnd, GW_HWNDNEXT);
    }

    return false;
}

BOOL TaskbarScrollRight(int button_groups_count,
                        LONG_PTR** button_groups,
                        int* p_button_group_index,
                        int* p_button_index) {
    int button_group_index = *p_button_group_index;
    int button_index = *p_button_index;
    int button_group_type;

    int buttons_count =
        button_group_index == -1
            ? 0
            : CTaskBtnGroup_GetNumItems(button_groups[button_group_index]);
    if (++button_index >= buttons_count) {
        do {
            button_group_index++;
            if (button_group_index >= button_groups_count) {
                return FALSE;
            }

            button_group_type =
                CTaskBtnGroup_GetGroupType(button_groups[button_group_index]);
        } while (button_group_type != 1 && button_group_type != 3);

        button_index = 0;
    }

    *p_button_group_index = button_group_index;
    *p_button_index = button_index;

    return TRUE;
}

BOOL TaskbarScrollLeft(int button_groups_count,
                       LONG_PTR** button_groups,
                       int* p_button_group_index,
                       int* p_button_index) {
    int button_group_index = *p_button_group_index;
    int button_index = *p_button_index;
    int button_group_type;

    if (button_group_index == -1 || --button_index < 0) {
        if (button_group_index == -1) {
            button_group_index = button_groups_count;
        }

        do {
            button_group_index--;
            if (button_group_index < 0) {
                return FALSE;
            }

            button_group_type =
                CTaskBtnGroup_GetGroupType(button_groups[button_group_index]);
        } while (button_group_type != 1 && button_group_type != 3);

        int buttons_count =
            CTaskBtnGroup_GetNumItems(button_groups[button_group_index]);
        button_index = buttons_count - 1;
    }

    *p_button_group_index = button_group_index;
    *p_button_index = button_index;

    return TRUE;
}

LONG_PTR* TaskbarScrollHelper(int button_groups_count,
                              LONG_PTR** button_groups,
                              int button_group_index_active,
                              int button_index_active,
                              int nRotates,
                              BOOL bWarpAround) {
    int button_group_index, button_index;
    BOOL bRotateRight;
    int prev_button_group_index, prev_button_index;
    BOOL bScrollSucceeded;

    button_group_index = button_group_index_active;
    button_index = button_index_active;

    bRotateRight = TRUE;
    if (nRotates < 0) {
        bRotateRight = FALSE;
        nRotates = -nRotates;
    }

    // Get current process name to skip same processes
    std::wstring currentProcessName;
    if (button_group_index_active >= 0 && button_index_active >= 0) {
        LONG_PTR* currentTaskItem = (LONG_PTR*)CTaskBtnGroup_GetTaskItem(
            button_groups[button_group_index_active], button_index_active);
        if (currentTaskItem) {
            HWND currentWnd = GetTaskItemWnd(currentTaskItem);
            currentProcessName = GetProcessNameFromWindow(currentWnd);
            Wh_Log(L"Current process: %s", currentProcessName.c_str());
        }
    }

    prev_button_group_index = button_group_index;
    prev_button_index = button_index;

    while (nRotates--) {
        // Collect all candidate windows with a different process name
        std::vector<std::pair<int, int>> candidatePositions; // pair<group_index, button_index>
        std::wstring targetProcessName;
        int search_group_index = button_group_index;
        int search_button_index = button_index;
        bool foundDifferent = false;

        // Search in the rotation direction
        for (int attempts = 0; attempts < button_groups_count * 10; attempts++) {
            if (bRotateRight) {
                bScrollSucceeded =
                    TaskbarScrollRight(button_groups_count, button_groups,
                                       &search_group_index, &search_button_index);
            } else {
                bScrollSucceeded =
                    TaskbarScrollLeft(button_groups_count, button_groups,
                                      &search_group_index, &search_button_index);
            }

            if (!bScrollSucceeded) {
                if (bWarpAround && prev_button_group_index != -1) {
                    // Wrap around
                    search_group_index = -1;
                    search_button_index = -1;
                    continue;
                } else {
                    break;
                }
            }

            // Check if we've come full circle
            if (search_group_index == button_group_index_active &&
                search_button_index == button_index_active) {
                break;
            }

            LONG_PTR* taskItem = (LONG_PTR*)CTaskBtnGroup_GetTaskItem(
                button_groups[search_group_index], search_button_index);

            if (!taskItem) {
                continue;
            }

            HWND taskWnd = GetTaskItemWnd(taskItem);
            std::wstring processName = GetProcessNameFromWindow(taskWnd);

            // Skip if same as current process
            if (!currentProcessName.empty() &&
                _wcsicmp(processName.c_str(), currentProcessName.c_str()) == 0) {
                continue;
            }

            // Found a different process
            if (!foundDifferent) {
                foundDifferent = true;
                targetProcessName = processName;
                candidatePositions.push_back({search_group_index, search_button_index});
                Wh_Log(L"Found different process: %s at group=%d, index=%d",
                       processName.c_str(), search_group_index, search_button_index);
            } else if (_wcsicmp(processName.c_str(), targetProcessName.c_str()) == 0) {
                // Same as target process, add to candidates
                candidatePositions.push_back({search_group_index, search_button_index});
                Wh_Log(L"Added candidate: %s at group=%d, index=%d",
                       processName.c_str(), search_group_index, search_button_index);
            } else {
                // Different process found, stop collecting
                break;
            }
        }

        // If no different process found, fall back to original behavior
        if (!foundDifferent) {
            Wh_Log(L"No different process found, using original behavior");
            if (bRotateRight) {
                bScrollSucceeded =
                    TaskbarScrollRight(button_groups_count, button_groups,
                                       &button_group_index, &button_index);
            } else {
                bScrollSucceeded =
                    TaskbarScrollLeft(button_groups_count, button_groups,
                                      &button_group_index, &button_index);
            }

            if (!bScrollSucceeded) {
                if (prev_button_group_index == -1) {
                    return nullptr;
                }

                if (bWarpAround) {
                    button_group_index = -1;
                    button_index = -1;
                    nRotates++;
                } else {
                    button_group_index = prev_button_group_index;
                    button_index = prev_button_index;
                    break;
                }
            }
        } else {
            // Select the best candidate based on Z-order (most recently active)
            HWND bestWindow = nullptr;
            int bestGroupIndex = -1;
            int bestButtonIndex = -1;

            for (const auto& pos : candidatePositions) {
                LONG_PTR* taskItem = (LONG_PTR*)CTaskBtnGroup_GetTaskItem(
                    button_groups[pos.first], pos.second);
                if (!taskItem) {
                    continue;
                }

                HWND wnd = GetTaskItemWnd(taskItem);
                if (!wnd) {
                    continue;
                }

                if (!bestWindow) {
                    bestWindow = wnd;
                    bestGroupIndex = pos.first;
                    bestButtonIndex = pos.second;
                } else if (IsWindowAboveInZOrder(wnd, bestWindow)) {
                    bestWindow = wnd;
                    bestGroupIndex = pos.first;
                    bestButtonIndex = pos.second;
                }
            }

            if (bestGroupIndex >= 0 && bestButtonIndex >= 0) {
                button_group_index = bestGroupIndex;
                button_index = bestButtonIndex;
                Wh_Log(L"Selected best window at group=%d, index=%d",
                       button_group_index, button_index);
            }
        }

        prev_button_group_index = button_group_index;
        prev_button_index = button_index;
    }

    if (button_group_index == button_group_index_active &&
        button_index == button_index_active) {
        return nullptr;
    }

    return (LONG_PTR*)CTaskBtnGroup_GetTaskItem(
        button_groups[button_group_index], button_index);
}

HDPA GetTaskBtnGroupsArray(void* taskList_ITaskListUI) {
    // This is a horrible hack, but it's the best way I found to get the array
    // of task button groups from a task list. It relies on the implementation
    // of CTaskListWnd::GetButtonGroupCount being just this:
    //
    // return DPA_GetPtrCount(this->buttonGroupsArray);
    //
    // Or in other words:
    //
    // return *(int*)this[buttonGroupsArrayOffset];
    //
    // Instead of calling it with a real taskList object, we call it with an
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

        return CTaskListWnd_GetButtonGroupCount(arrayOfIntPtrs);
    }();

    return (HDPA)((void**)taskList_ITaskListUI)[offset];
}

LONG_PTR* TaskbarScroll(LONG_PTR lpMMTaskListLongPtr,
                        int nRotates,
                        BOOL bWarpAround,
                        LONG_PTR* src_task_item) {
    if (nRotates == 0) {
        return nullptr;
    }

    void* taskList_ITaskListUI = QueryViaVtable(
        (void*)lpMMTaskListLongPtr, CTaskListWnd_vftable_ITaskListUI);

    LONG_PTR* plp = (LONG_PTR*)GetTaskBtnGroupsArray(taskList_ITaskListUI);
    if (!plp) {
        return nullptr;
    }

    int button_groups_count = (int)plp[0];
    LONG_PTR** button_groups = (LONG_PTR**)plp[1];

    int button_group_index_active = -1;
    int button_index_active = -1;

    if (src_task_item) {
        for (int i = 0; i < button_groups_count; i++) {
            int button_group_type =
                CTaskBtnGroup_GetGroupType(button_groups[i]);
            if (button_group_type == 1 || button_group_type == 3) {
                int buttons_count = CTaskBtnGroup_GetNumItems(button_groups[i]);
                for (int j = 0; j < buttons_count; j++) {
                    if ((LONG_PTR*)CTaskBtnGroup_GetTaskItem(
                            button_groups[i], j) == src_task_item) {
                        button_group_index_active = i;
                        button_index_active = j;
                        break;
                    }
                }

                if (button_group_index_active != -1) {
                    break;
                }
            }
        }
    } else if (auto it =
                   g_lastTaskListActiveItem.find((void*)lpMMTaskListLongPtr);
               it != g_lastTaskListActiveItem.end()) {
        LONG_PTR* last_button_group_active = (LONG_PTR*)it->second.taskBtnGroup;
        int last_button_index_active = it->second.buttonIndex;
        if (last_button_group_active && last_button_index_active >= 0) {
            for (int i = 0; i < button_groups_count; i++) {
                if (button_groups[i] == last_button_group_active) {
                    int buttons_count =
                        CTaskBtnGroup_GetNumItems(button_groups[i]);
                    if (buttons_count > 0) {
                        button_group_index_active = i;
                        button_index_active = std::min(last_button_index_active,
                                                       buttons_count - 1);
                    }
                    break;
                }
            }
        }
    }

    return TaskbarScrollHelper(button_groups_count, button_groups,
                               button_group_index_active, button_index_active,
                               nRotates, bWarpAround);
}

#pragma endregion  // scroll

#pragma region button_group_reorder

// Find the button group index for the currently active foreground window
int FindActiveButtonGroupIndex(LONG_PTR lpMMTaskListLongPtr) {
    HWND hForeground = GetForegroundWindow();
    if (!hForeground) {
        Wh_Log(L"No foreground window");
        return -1;
    }

    void* taskList_ITaskListUI = QueryViaVtable(
        (void*)lpMMTaskListLongPtr, CTaskListWnd_vftable_ITaskListUI);

    LONG_PTR* plp = (LONG_PTR*)GetTaskBtnGroupsArray(taskList_ITaskListUI);
    if (!plp) {
        Wh_Log(L"Failed to get button groups array");
        return -1;
    }

    int button_groups_count = (int)plp[0];
    LONG_PTR** button_groups = (LONG_PTR**)plp[1];

    // Search through all button groups
    for (int i = 0; i < button_groups_count; i++) {
        int button_group_type = CTaskBtnGroup_GetGroupType(button_groups[i]);
        if (button_group_type != 1 && button_group_type != 3) {
            continue;  // Skip non-app groups
        }

        int buttons_count = CTaskBtnGroup_GetNumItems(button_groups[i]);
        for (int j = 0; j < buttons_count; j++) {
            LONG_PTR* taskItem = (LONG_PTR*)CTaskBtnGroup_GetTaskItem(
                button_groups[i], j);
            if (!taskItem) continue;

            HWND wnd = GetTaskItemWnd(taskItem);
            if (wnd == hForeground) {
                Wh_Log(L"Found active window at group index %d", i);
                return i;
            }
        }
    }

    Wh_Log(L"Active window not found in taskbar");
    return -1;
}

// Swap adjacent button groups in the taskbar
bool SwapAdjacentButtonGroups(HWND hMMTaskListWnd,
                              LONG_PTR lpMMTaskListLongPtr,
                              int currentGroupIndex,
                              bool moveLeft) {
    Wh_Log(L"SwapAdjacentButtonGroups: index=%d, moveLeft=%d",
           currentGroupIndex, moveLeft);

    void* taskList_ITaskListUI = QueryViaVtable(
        (void*)lpMMTaskListLongPtr, CTaskListWnd_vftable_ITaskListUI);

    LONG_PTR* plp = (LONG_PTR*)GetTaskBtnGroupsArray(taskList_ITaskListUI);
    if (!plp) {
        Wh_Log(L"Failed to get button groups array");
        return false;
    }

    int groupCount = (int)plp[0];
    LONG_PTR** buttonGroups = (LONG_PTR**)plp[1];

    // Calculate target index
    int targetIndex = moveLeft ? (currentGroupIndex - 1) : (currentGroupIndex + 1);

    // Check bounds
    if (targetIndex < 0 || targetIndex >= groupCount) {
        Wh_Log(L"Target index %d out of bounds [0, %d)", targetIndex, groupCount);
        return false;
    }

    // Skip if target is not an app group
    int targetGroupType = CTaskBtnGroup_GetGroupType(buttonGroups[targetIndex]);
    if (targetGroupType != 1 && targetGroupType != 3) {
        Wh_Log(L"Target group type %d is not an app group", targetGroupType);
        return false;
    }

    // Save window handles BEFORE swapping for refresh later
    HWND hwndCurrent = NULL;
    HWND hwndTarget = NULL;

    void* taskItemCurrent = CTaskBtnGroup_GetTaskItem(buttonGroups[currentGroupIndex], 0);
    if (taskItemCurrent) {
        hwndCurrent = GetTaskItemWnd((LONG_PTR*)taskItemCurrent);
    }

    void* taskItemTarget = CTaskBtnGroup_GetTaskItem(buttonGroups[targetIndex], 0);
    if (taskItemTarget) {
        hwndTarget = GetTaskItemWnd((LONG_PTR*)taskItemTarget);
    }

    Wh_Log(L"Saved window handles: current=%p, target=%p", hwndCurrent, hwndTarget);

    // Method 1: Try using DPA API (more reliable)
    // Capture the HDPA handle by triggering DPA_GetPtr
    g_getPtr_lastHdpa = nullptr;
    g_getPtr_captureForThreadId = GetCurrentThreadId();

    // Get a task item to trigger DPA_GetPtr
    void* dummyItem = CTaskBtnGroup_GetTaskItem(buttonGroups[currentGroupIndex], 0);
    if (dummyItem) {
        // This should have triggered DPA_GetPtr_Hook
        g_getPtr_captureForThreadId = GetCurrentThreadId();
        CTaskBtnGroup_IndexOfTaskItem(buttonGroups[currentGroupIndex], dummyItem);
    }

    g_getPtr_captureForThreadId = 0;
    HDPA groupsArray = g_getPtr_lastHdpa;

    if (groupsArray && groupsArray->cpItems == groupCount) {
        Wh_Log(L"Using DPA API to reorder");

        // Use DPA API to reorder
        void* group = (void*)DPA_DeletePtr(groupsArray, currentGroupIndex);
        if (!group) {
            Wh_Log(L"DPA_DeletePtr failed");
            return false;
        }

        int insertResult = DPA_InsertPtr(groupsArray, targetIndex, group);
        if (insertResult == -1) {
            Wh_Log(L"DPA_InsertPtr failed");
            // Try to restore
            DPA_InsertPtr(groupsArray, currentGroupIndex, group);
            return false;
        }

        Wh_Log(L"DPA reorder successful");
    } else {
        Wh_Log(L"Using direct memory swap (fallback)");

        // Method 2: Direct memory swap (fallback)
        LONG_PTR* temp = buttonGroups[currentGroupIndex];
        buttonGroups[currentGroupIndex] = buttonGroups[targetIndex];
        buttonGroups[targetIndex] = temp;
    }

    // Refresh the taskbar - use aggressive method to force UI update
    Wh_Log(L"Refreshing taskbar UI");

    // Method 1: Standard window refresh
    InvalidateRect(hMMTaskListWnd, nullptr, TRUE);
    UpdateWindow(hMMTaskListWnd);
    RedrawWindow(hMMTaskListWnd, nullptr, nullptr,
                 RDW_ERASE | RDW_FRAME | RDW_INVALIDATE | RDW_ALLCHILDREN);

    // Method 2: Trigger window flash to force taskbar to update icons
    // This is a hack but it works - briefly flash the windows to force taskbar refresh
    if (hwndCurrent && IsWindow(hwndCurrent)) {
        FLASHWINFO fwi = {sizeof(FLASHWINFO)};
        fwi.hwnd = hwndCurrent;
        fwi.dwFlags = FLASHW_TRAY | FLASHW_TIMERNOFG;
        fwi.uCount = 1;
        fwi.dwTimeout = 0;
        FlashWindowEx(&fwi);
        Wh_Log(L"Flashed current window");
    }

    if (hwndTarget && IsWindow(hwndTarget)) {
        FLASHWINFO fwi = {sizeof(FLASHWINFO)};
        fwi.hwnd = hwndTarget;
        fwi.dwFlags = FLASHW_TRAY | FLASHW_TIMERNOFG;
        fwi.uCount = 1;
        fwi.dwTimeout = 0;
        FlashWindowEx(&fwi);
        Wh_Log(L"Flashed target window");
    }

     // Method 3: Force taskbar to recalculate by sending settings change
     SendNotifyMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0,(LPARAM)L"TraySettings");
    Wh_Log(L"Button group swap completed");
    return true;
}

// Hotkey handler for moving button groups
bool OnMoveButtonGroup(HWND hWnd, bool moveLeft) {
    Wh_Log(L"OnMoveButtonGroup: moveLeft=%d", moveLeft);

    DWORD messagePos = GetMessagePos();
    POINT pt{
        GET_X_LPARAM(messagePos),
        GET_Y_LPARAM(messagePos),
    };

    HMONITOR monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    HWND hTaskbarForMonitor = GetTaskbarForMonitor(hWnd, monitor);
    HWND hMMTaskListWnd = TaskListFromMMTaskbarWnd(
        hTaskbarForMonitor ? hTaskbarForMonitor : hWnd);

    if (!hMMTaskListWnd) {
        Wh_Log(L"Failed to get task list window");
        return false;
    }

    LONG_PTR lpTaskListLongPtr = GetWindowLongPtr(hMMTaskListWnd, 0);

    // Find the button group for the current foreground window
    int currentGroupIndex = FindActiveButtonGroupIndex(lpTaskListLongPtr);
    if (currentGroupIndex == -1) {
        Wh_Log(L"Could not find active button group");
        return false;
    }

    // Swap with adjacent group
    return SwapAdjacentButtonGroups(hMMTaskListWnd, lpTaskListLongPtr,
                                    currentGroupIndex, moveLeft);
}

#pragma endregion  // button_group_reorder

void OnTaskListScroll(HWND hMMTaskListWnd, short delta) {
    if (g_lastScrollTarget == hMMTaskListWnd &&
        GetTickCount() - g_lastScrollTime < 1000 * 5) {
        delta += g_lastScrollDeltaRemainder;
    }

    int clicks = -delta / WHEEL_DELTA;
    Wh_Log(L"%d clicks (delta=%d)", clicks, delta);

    if (clicks != 0) {
        if (g_settings.reverseScrollingDirection) {
            clicks = -clicks;
        }

        LONG_PTR lpMMTaskListLongPtr = GetWindowLongPtr(hMMTaskListWnd, 0);
        PVOID targetTaskItem = TaskbarScroll(lpMMTaskListLongPtr, clicks,
                                             g_settings.wrapAround, nullptr);
        if (targetTaskItem) {
            SwitchToTaskItem(lpMMTaskListLongPtr, targetTaskItem);
        }
    }

    g_lastScrollTarget = hMMTaskListWnd;
    g_lastScrollTime = GetTickCount();
    g_lastScrollDeltaRemainder = delta % WHEEL_DELTA;
}

HWND TaskListFromTaskbarWnd(HWND hTaskbarWnd) {
    HWND hReBarWindow32 =
        FindWindowEx(hTaskbarWnd, nullptr, L"ReBarWindow32", nullptr);
    if (!hReBarWindow32) {
        return nullptr;
    }

    HWND hMSTaskSwWClass =
        FindWindowEx(hReBarWindow32, nullptr, L"MSTaskSwWClass", nullptr);
    if (!hMSTaskSwWClass) {
        return nullptr;
    }

    return FindWindowEx(hMSTaskSwWClass, nullptr, L"MSTaskListWClass", nullptr);
}

HWND TaskListFromSecondaryTaskbarWnd(HWND hSecondaryTaskbarWnd) {
    HWND hWorkerWWnd =
        FindWindowEx(hSecondaryTaskbarWnd, nullptr, L"WorkerW", nullptr);
    if (!hWorkerWWnd) {
        return nullptr;
    }

    return FindWindowEx(hWorkerWWnd, nullptr, L"MSTaskListWClass", nullptr);
}

HWND TaskListFromMMTaskbarWnd(HWND hMMTaskbarWnd) {
    WCHAR szClassName[32];
    if (!GetClassName(hMMTaskbarWnd, szClassName, ARRAYSIZE(szClassName))) {
        return nullptr;
    }

    if (_wcsicmp(szClassName, L"Shell_TrayWnd") == 0) {
        return TaskListFromTaskbarWnd(hMMTaskbarWnd);
    }

    if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
        return TaskListFromSecondaryTaskbarWnd(hMMTaskbarWnd);
    }

    return nullptr;
}

HWND TaskListFromPoint(POINT pt) {
    HWND hPointWnd = WindowFromPoint(pt);
    if (!hPointWnd) {
        return nullptr;
    }

    HWND hRootWnd = GetAncestor(hPointWnd, GA_ROOT);
    if (!hRootWnd) {
        return nullptr;
    }

    return TaskListFromMMTaskbarWnd(hRootWnd);
}

HWND GetTaskbarForMonitor(HWND hTaskbarWnd, HMONITOR monitor) {
    DWORD taskbarThreadId = 0;
    DWORD taskbarProcessId = 0;
    if (!(taskbarThreadId =
              GetWindowThreadProcessId(hTaskbarWnd, &taskbarProcessId)) ||
        taskbarProcessId != GetCurrentProcessId()) {
        return nullptr;
    }

    if (MonitorFromWindow(hTaskbarWnd, MONITOR_DEFAULTTONEAREST) == monitor) {
        return hTaskbarWnd;
    }

    HWND hResultWnd = nullptr;

    auto enumWindowsProc = [monitor, &hResultWnd](HWND hWnd) -> BOOL {
        WCHAR szClassName[32];
        if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
            return TRUE;
        }

        if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") != 0) {
            return TRUE;
        }

        if (MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST) != monitor) {
            return TRUE;
        }

        hResultWnd = hWnd;
        return FALSE;
    };

    EnumThreadWindows(
        taskbarThreadId,
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(enumWindowsProc)*>(lParam);
            return proc(hWnd);
        },
        reinterpret_cast<LPARAM>(&enumWindowsProc));

    return hResultWnd;
}

bool FromStringHotKey(std::wstring_view hotkeyString,
                      UINT* modifiersOut,
                      UINT* vkOut) {
    static const std::unordered_map<std::wstring_view, UINT> modifiersMap = {
        {L"ALT", MOD_ALT},           {L"CTRL", MOD_CONTROL},
        {L"NOREPEAT", MOD_NOREPEAT}, {L"SHIFT", MOD_SHIFT},
        {L"WIN", MOD_WIN},
    };

    static const std::unordered_map<std::wstring_view, UINT> vkMap = {
        {L"VK_LBUTTON", 0x01},
        {L"VK_RBUTTON", 0x02},
        {L"VK_CANCEL", 0x03},
        {L"VK_MBUTTON", 0x04},
        {L"VK_XBUTTON1", 0x05},
        {L"VK_XBUTTON2", 0x06},
        {L"VK_BACK", 0x08},
        {L"VK_TAB", 0x09},
        {L"VK_CLEAR", 0x0C},
        {L"VK_RETURN", 0x0D},
        {L"VK_SHIFT", 0x10},
        {L"VK_CONTROL", 0x11},
        {L"VK_MENU", 0x12},
        {L"VK_PAUSE", 0x13},
        {L"VK_CAPITAL", 0x14},
        {L"VK_KANA", 0x15},
        {L"VK_HANGUEL", 0x15},
        {L"VK_HANGUL", 0x15},
        {L"VK_IME_ON", 0x16},
        {L"VK_JUNJA", 0x17},
        {L"VK_FINAL", 0x18},
        {L"VK_HANJA", 0x19},
        {L"VK_KANJI", 0x19},
        {L"VK_IME_OFF", 0x1A},
        {L"VK_ESCAPE", 0x1B},
        {L"VK_CONVERT", 0x1C},
        {L"VK_NONCONVERT", 0x1D},
        {L"VK_ACCEPT", 0x1E},
        {L"VK_MODECHANGE", 0x1F},
        {L"VK_SPACE", 0x20},
        {L"VK_PRIOR", 0x21},
        {L"VK_NEXT", 0x22},
        {L"VK_END", 0x23},
        {L"VK_HOME", 0x24},
        {L"VK_LEFT", 0x25},
        {L"VK_UP", 0x26},
        {L"VK_RIGHT", 0x27},
        {L"VK_DOWN", 0x28},
        {L"VK_SELECT", 0x29},
        {L"VK_PRINT", 0x2A},
        {L"VK_EXECUTE", 0x2B},
        {L"VK_SNAPSHOT", 0x2C},
        {L"VK_INSERT", 0x2D},
        {L"VK_DELETE", 0x2E},
        {L"VK_HELP", 0x2F},
        {L"0", 0x30},
        {L"1", 0x31},
        {L"2", 0x32},
        {L"3", 0x33},
        {L"4", 0x34},
        {L"5", 0x35},
        {L"6", 0x36},
        {L"7", 0x37},
        {L"8", 0x38},
        {L"9", 0x39},
        {L"A", 0x41},
        {L"B", 0x42},
        {L"C", 0x43},
        {L"D", 0x44},
        {L"E", 0x45},
        {L"F", 0x46},
        {L"G", 0x47},
        {L"H", 0x48},
        {L"I", 0x49},
        {L"J", 0x4A},
        {L"K", 0x4B},
        {L"L", 0x4C},
        {L"M", 0x4D},
        {L"N", 0x4E},
        {L"O", 0x4F},
        {L"P", 0x50},
        {L"Q", 0x51},
        {L"R", 0x52},
        {L"S", 0x53},
        {L"T", 0x54},
        {L"U", 0x55},
        {L"V", 0x56},
        {L"W", 0x57},
        {L"X", 0x58},
        {L"Y", 0x59},
        {L"Z", 0x5A},
        {L"VK_LWIN", 0x5B},
        {L"VK_RWIN", 0x5C},
        {L"VK_APPS", 0x5D},
        {L"VK_SLEEP", 0x5F},
        {L"VK_NUMPAD0", 0x60},
        {L"VK_NUMPAD1", 0x61},
        {L"VK_NUMPAD2", 0x62},
        {L"VK_NUMPAD3", 0x63},
        {L"VK_NUMPAD4", 0x64},
        {L"VK_NUMPAD5", 0x65},
        {L"VK_NUMPAD6", 0x66},
        {L"VK_NUMPAD7", 0x67},
        {L"VK_NUMPAD8", 0x68},
        {L"VK_NUMPAD9", 0x69},
        {L"VK_MULTIPLY", 0x6A},
        {L"VK_ADD", 0x6B},
        {L"VK_SEPARATOR", 0x6C},
        {L"VK_SUBTRACT", 0x6D},
        {L"VK_DECIMAL", 0x6E},
        {L"VK_DIVIDE", 0x6F},
        {L"VK_F1", 0x70},
        {L"VK_F2", 0x71},
        {L"VK_F3", 0x72},
        {L"VK_F4", 0x73},
        {L"VK_F5", 0x74},
        {L"VK_F6", 0x75},
        {L"VK_F7", 0x76},
        {L"VK_F8", 0x77},
        {L"VK_F9", 0x78},
        {L"VK_F10", 0x79},
        {L"VK_F11", 0x7A},
        {L"VK_F12", 0x7B},
        {L"VK_F13", 0x7C},
        {L"VK_F14", 0x7D},
        {L"VK_F15", 0x7E},
        {L"VK_F16", 0x7F},
        {L"VK_F17", 0x80},
        {L"VK_F18", 0x81},
        {L"VK_F19", 0x82},
        {L"VK_F20", 0x83},
        {L"VK_F21", 0x84},
        {L"VK_F22", 0x85},
        {L"VK_F23", 0x86},
        {L"VK_F24", 0x87},
        {L"VK_NUMLOCK", 0x90},
        {L"VK_SCROLL", 0x91},
        {L"VK_LSHIFT", 0xA0},
        {L"VK_RSHIFT", 0xA1},
        {L"VK_LCONTROL", 0xA2},
        {L"VK_RCONTROL", 0xA3},
        {L"VK_LMENU", 0xA4},
        {L"VK_RMENU", 0xA5},
        {L"VK_BROWSER_BACK", 0xA6},
        {L"VK_BROWSER_FORWARD", 0xA7},
        {L"VK_BROWSER_REFRESH", 0xA8},
        {L"VK_BROWSER_STOP", 0xA9},
        {L"VK_BROWSER_SEARCH", 0xAA},
        {L"VK_BROWSER_FAVORITES", 0xAB},
        {L"VK_BROWSER_HOME", 0xAC},
        {L"VK_VOLUME_MUTE", 0xAD},
        {L"VK_VOLUME_DOWN", 0xAE},
        {L"VK_VOLUME_UP", 0xAF},
        {L"VK_MEDIA_NEXT_TRACK", 0xB0},
        {L"VK_MEDIA_PREV_TRACK", 0xB1},
        {L"VK_MEDIA_STOP", 0xB2},
        {L"VK_MEDIA_PLAY_PAUSE", 0xB3},
        {L"VK_LAUNCH_MAIL", 0xB4},
        {L"VK_LAUNCH_MEDIA_SELECT", 0xB5},
        {L"VK_LAUNCH_APP1", 0xB6},
        {L"VK_LAUNCH_APP2", 0xB7},
        {L"VK_OEM_1", 0xBA},
        {L"VK_OEM_PLUS", 0xBB},
        {L"VK_OEM_COMMA", 0xBC},
        {L"VK_OEM_MINUS", 0xBD},
        {L"VK_OEM_PERIOD", 0xBE},
        {L"VK_OEM_2", 0xBF},
        {L"VK_OEM_3", 0xC0},
        {L"VK_OEM_4", 0xDB},
        {L"VK_OEM_5", 0xDC},
        {L"VK_OEM_6", 0xDD},
        {L"VK_OEM_7", 0xDE},
        {L"VK_OEM_8", 0xDF},
        {L"VK_OEM_102", 0xE2},
        {L"VK_PROCESSKEY", 0xE5},
        {L"VK_PACKET", 0xE7},
        {L"VK_ATTN", 0xF6},
        {L"VK_CRSEL", 0xF7},
        {L"VK_EXSEL", 0xF8},
        {L"VK_EREOF", 0xF9},
        {L"VK_PLAY", 0xFA},
        {L"VK_ZOOM", 0xFB},
        {L"VK_NONAME", 0xFC},
        {L"VK_PA1", 0xFD},
        {L"VK_OEM_CLEAR", 0xFE},
    };

    // https://stackoverflow.com/a/46931770
    auto splitStringView = [](std::wstring_view s, WCHAR delimiter) {
        size_t pos_start = 0, pos_end;
        std::wstring_view token;
        std::vector<std::wstring_view> res;

        while ((pos_end = s.find(delimiter, pos_start)) !=
               std::wstring_view::npos) {
            token = s.substr(pos_start, pos_end - pos_start);
            pos_start = pos_end + 1;
            res.push_back(token);
        }

        res.push_back(s.substr(pos_start));
        return res;
    };

    // https://stackoverflow.com/a/54364173
    auto trimStringView = [](std::wstring_view s) {
        s.remove_prefix(std::min(s.find_first_not_of(L" \t\r\v\n"), s.size()));
        s.remove_suffix(std::min(
            s.size() - s.find_last_not_of(L" \t\r\v\n") - 1, s.size()));
        return s;
    };

    UINT modifiers = 0;
    UINT vk = 0;

    auto hotkeyParts = splitStringView(hotkeyString, '+');
    for (auto hotkeyPart : hotkeyParts) {
        hotkeyPart = trimStringView(hotkeyPart);
        std::wstring hotkeyPartUpper{hotkeyPart};
        std::transform(hotkeyPartUpper.begin(), hotkeyPartUpper.end(),
                       hotkeyPartUpper.begin(), ::toupper);

        if (auto it = modifiersMap.find(hotkeyPartUpper);
            it != modifiersMap.end()) {
            modifiers |= it->second;
            continue;
        }

        if (vk) {
            // Only one is allowed.
            return false;
        }

        if (auto it = vkMap.find(hotkeyPartUpper); it != vkMap.end()) {
            vk = it->second;
            continue;
        }

        size_t pos;
        try {
            vk = std::stoi(hotkeyPartUpper, &pos, 0);
            if (hotkeyPartUpper[pos] != L'\0' || !vk) {
                return false;
            }
        } catch (const std::exception&) {
            return false;
        }
    }

    if (!vk) {
        return false;
    }

    *modifiersOut = modifiers;
    *vkOut = vk;
    return true;
}

void UnregisterHotkeys(HWND hWnd) {
    if (g_hotkeyLeftRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdLeft);
        g_hotkeyLeftRegistered = false;
    }

    if (g_hotkeyRightRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdRight);
        g_hotkeyRightRegistered = false;
    }

    if (g_hotkeySameProcessLeftRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdSameProcessLeft);
        g_hotkeySameProcessLeftRegistered = false;
    }

    if (g_hotkeySameProcessRightRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdSameProcessRight);
        g_hotkeySameProcessRightRegistered = false;
    }

    if (g_hotkeyLeft2Registered) {
        UnregisterHotKey(hWnd, kHotkeyIdLeft2);
        g_hotkeyLeft2Registered = false;
    }

    if (g_hotkeyRight2Registered) {
        UnregisterHotKey(hWnd, kHotkeyIdRight2);
        g_hotkeyRight2Registered = false;
    }

    if (g_hotkeySameProcessLeft2Registered) {
        UnregisterHotKey(hWnd, kHotkeyIdSameProcessLeft2);
        g_hotkeySameProcessLeft2Registered = false;
    }

    if (g_hotkeySameProcessRight2Registered) {
        UnregisterHotKey(hWnd, kHotkeyIdSameProcessRight2);
        g_hotkeySameProcessRight2Registered = false;
    }

    if (g_hotkeyMoveGroupLeftRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdMoveGroupLeft);
        g_hotkeyMoveGroupLeftRegistered = false;
    }

    if (g_hotkeyMoveGroupRightRegistered) {
        UnregisterHotKey(hWnd, kHotkeyIdMoveGroupRight);
        g_hotkeyMoveGroupRightRegistered = false;
    }
}

void RegisterHotkeys(HWND hWnd) {
    if (!*g_settings.cycleLeftKeyboardShortcut &&
        !*g_settings.cycleRightKeyboardShortcut &&
        !*g_settings.cycleSameProcessLeftKeyboardShortcut &&
        !*g_settings.cycleSameProcessRightKeyboardShortcut) {
        return;
    }

    UINT modifiers;
    UINT vk;

    if (FromStringHotKey(g_settings.cycleLeftKeyboardShortcut.get(), &modifiers,
                         &vk)) {
        g_hotkeyLeftRegistered =
            RegisterHotKey(hWnd, kHotkeyIdLeft, modifiers, vk);
        if (!g_hotkeyLeftRegistered) {
            Wh_Log(L"Couldn't register hotkey: %s",
                   g_settings.cycleLeftKeyboardShortcut.get());
        }
    } else {
        Wh_Log(L"Couldn't parse hotkey: %s",
               g_settings.cycleLeftKeyboardShortcut.get());
    }

    if (FromStringHotKey(g_settings.cycleRightKeyboardShortcut.get(),
                         &modifiers, &vk)) {
        g_hotkeyRightRegistered =
            RegisterHotKey(hWnd, kHotkeyIdRight, modifiers, vk);
        if (!g_hotkeyRightRegistered) {
            Wh_Log(L"Couldn't register hotkey: %s",
                   g_settings.cycleRightKeyboardShortcut.get());
        }
    } else {
        Wh_Log(L"Couldn't parse hotkey: %s",
               g_settings.cycleRightKeyboardShortcut.get());
    }

    // Use default value if same process left shortcut is empty
    const WCHAR* sameProcessLeftShortcut = g_settings.cycleSameProcessLeftKeyboardShortcut.get();
    if (!sameProcessLeftShortcut || wcslen(sameProcessLeftShortcut) == 0) {
        sameProcessLeftShortcut = L"Ctrl+Alt+VK_OEM_4";
        Wh_Log(L"Using default same process left hotkey: %s", sameProcessLeftShortcut);
    }

    if (FromStringHotKey(sameProcessLeftShortcut, &modifiers, &vk)) {
        g_hotkeySameProcessLeftRegistered =
            RegisterHotKey(hWnd, kHotkeyIdSameProcessLeft, modifiers, vk);
        if (!g_hotkeySameProcessLeftRegistered) {
            Wh_Log(L"Couldn't register same process left hotkey: %s",
                   sameProcessLeftShortcut);
        } else {
            Wh_Log(L"Registered same process left hotkey: %s", sameProcessLeftShortcut);
        }
    } else {
        Wh_Log(L"Couldn't parse same process left hotkey: %s",
               sameProcessLeftShortcut);
    }

    // Use default value if same process right shortcut is empty
    const WCHAR* sameProcessRightShortcut = g_settings.cycleSameProcessRightKeyboardShortcut.get();
    if (!sameProcessRightShortcut || wcslen(sameProcessRightShortcut) == 0) {
        sameProcessRightShortcut = L"Ctrl+Alt+VK_OEM_6";
        Wh_Log(L"Using default same process right hotkey: %s", sameProcessRightShortcut);
    }

    if (FromStringHotKey(sameProcessRightShortcut, &modifiers, &vk)) {
        g_hotkeySameProcessRightRegistered =
            RegisterHotKey(hWnd, kHotkeyIdSameProcessRight, modifiers, vk);
        if (!g_hotkeySameProcessRightRegistered) {
            Wh_Log(L"Couldn't register same process right hotkey: %s",
                   sameProcessRightShortcut);
        } else {
            Wh_Log(L"Registered same process right hotkey: %s", sameProcessRightShortcut);
        }
    } else {
        Wh_Log(L"Couldn't parse same process right hotkey: %s",
               sameProcessRightShortcut);
    }

    // Register second set of hotkeys (only if user configured them)
    const WCHAR* leftShortcut2 = g_settings.cycleLeftKeyboardShortcut2.get();
    if (leftShortcut2 && wcslen(leftShortcut2) > 0) {
        if (FromStringHotKey(leftShortcut2, &modifiers, &vk)) {
            g_hotkeyLeft2Registered = RegisterHotKey(hWnd, kHotkeyIdLeft2, modifiers, vk);
            if (g_hotkeyLeft2Registered) {
                Wh_Log(L"Registered left hotkey 2: %s", leftShortcut2);
            }
        }
    }

    const WCHAR* rightShortcut2 = g_settings.cycleRightKeyboardShortcut2.get();
    if (rightShortcut2 && wcslen(rightShortcut2) > 0) {
        if (FromStringHotKey(rightShortcut2, &modifiers, &vk)) {
            g_hotkeyRight2Registered = RegisterHotKey(hWnd, kHotkeyIdRight2, modifiers, vk);
            if (g_hotkeyRight2Registered) {
                Wh_Log(L"Registered right hotkey 2: %s", rightShortcut2);
            }
        }
    }

    const WCHAR* sameProcessLeftShortcut2 = g_settings.cycleSameProcessLeftKeyboardShortcut2.get();
    if (sameProcessLeftShortcut2 && wcslen(sameProcessLeftShortcut2) > 0) {
        if (FromStringHotKey(sameProcessLeftShortcut2, &modifiers, &vk)) {
            g_hotkeySameProcessLeft2Registered = RegisterHotKey(hWnd, kHotkeyIdSameProcessLeft2, modifiers, vk);
            if (g_hotkeySameProcessLeft2Registered) {
                Wh_Log(L"Registered same process left hotkey 2: %s", sameProcessLeftShortcut2);
            }
        }
    }

    const WCHAR* sameProcessRightShortcut2 = g_settings.cycleSameProcessRightKeyboardShortcut2.get();
    if (sameProcessRightShortcut2 && wcslen(sameProcessRightShortcut2) > 0) {
        if (FromStringHotKey(sameProcessRightShortcut2, &modifiers, &vk)) {
            g_hotkeySameProcessRight2Registered = RegisterHotKey(hWnd, kHotkeyIdSameProcessRight2, modifiers, vk);
            if (g_hotkeySameProcessRight2Registered) {
                Wh_Log(L"Registered same process right hotkey 2: %s", sameProcessRightShortcut2);
            }
        }
    }

    // Register button group move hotkeys
    const WCHAR* moveLeftShortcut = g_settings.moveButtonGroupLeftShortcut.get();
    if (moveLeftShortcut && wcslen(moveLeftShortcut) > 0) {
        if (FromStringHotKey(moveLeftShortcut, &modifiers, &vk)) {
            g_hotkeyMoveGroupLeftRegistered = RegisterHotKey(hWnd, kHotkeyIdMoveGroupLeft, modifiers, vk);
            if (g_hotkeyMoveGroupLeftRegistered) {
                Wh_Log(L"Registered move group left hotkey: %s", moveLeftShortcut);
            } else {
                Wh_Log(L"Couldn't register move group left hotkey: %s", moveLeftShortcut);
            }
        }
    }

    const WCHAR* moveRightShortcut = g_settings.moveButtonGroupRightShortcut.get();
    if (moveRightShortcut && wcslen(moveRightShortcut) > 0) {
        if (FromStringHotKey(moveRightShortcut, &modifiers, &vk)) {
            g_hotkeyMoveGroupRightRegistered = RegisterHotKey(hWnd, kHotkeyIdMoveGroupRight, modifiers, vk);
            if (g_hotkeyMoveGroupRightRegistered) {
                Wh_Log(L"Registered move group right hotkey: %s", moveRightShortcut);
            } else {
                Wh_Log(L"Couldn't register move group right hotkey: %s", moveRightShortcut);
            }
        }
    }
}

// Switch to another window of the same process
bool OnTaskbarSameProcessHotkey(HWND hWnd, int hotkeyId) {
    Wh_Log(L"Same process hotkey triggered");

    DWORD messagePos = GetMessagePos();
    POINT pt{
        GET_X_LPARAM(messagePos),
        GET_Y_LPARAM(messagePos),
    };

    HMONITOR monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    HWND hTaskbarForMonitor = GetTaskbarForMonitor(hWnd, monitor);
    HWND hMMTaskListWnd = TaskListFromMMTaskbarWnd(
        hTaskbarForMonitor ? hTaskbarForMonitor : hWnd);
    if (!hMMTaskListWnd) {
        return false;
    }

    LONG_PTR lpTaskListLongPtr = GetWindowLongPtr(hMMTaskListWnd, 0);
    void* taskList_ITaskListUI = QueryViaVtable(
        (void*)lpTaskListLongPtr, CTaskListWnd_vftable_ITaskListUI);

    LONG_PTR* plp = (LONG_PTR*)GetTaskBtnGroupsArray(taskList_ITaskListUI);
    if (!plp) {
        return false;
    }

    int button_groups_count = (int)plp[0];
    LONG_PTR** button_groups = (LONG_PTR**)plp[1];

    // Get current foreground window process name
    HWND hForeground = GetForegroundWindow();
    std::wstring currentProcessName = GetProcessNameFromWindow(hForeground);
    if (currentProcessName.empty()) {
        Wh_Log(L"Could not get current process name");
        return false;
    }

    Wh_Log(L"Current foreground process: %s", currentProcessName.c_str());

    // Collect all windows of the same process
    struct WindowInfo {
        int groupIndex;
        int buttonIndex;
        HWND hwnd;
        LONG_PTR* taskItem;
    };
    std::vector<WindowInfo> sameProcessWindows;

    for (int i = 0; i < button_groups_count; i++) {
        int button_group_type = CTaskBtnGroup_GetGroupType(button_groups[i]);
        if (button_group_type == 1 || button_group_type == 3) {
            int buttons_count = CTaskBtnGroup_GetNumItems(button_groups[i]);
            for (int j = 0; j < buttons_count; j++) {
                LONG_PTR* taskItem = (LONG_PTR*)CTaskBtnGroup_GetTaskItem(button_groups[i], j);
                if (!taskItem) {
                    continue;
                }

                HWND wnd = GetTaskItemWnd(taskItem);
                std::wstring processName = GetProcessNameFromWindow(wnd);

                if (_wcsicmp(processName.c_str(), currentProcessName.c_str()) == 0) {
                    sameProcessWindows.push_back({i, j, wnd, taskItem});
                    Wh_Log(L"Found same process window: %s at group=%d, index=%d",
                           processName.c_str(), i, j);
                }
            }
        }
    }

    // If only one window or no windows found, do nothing
    if (sameProcessWindows.size() <= 1) {
        Wh_Log(L"Only one or no windows of this process, nothing to switch");
        return false;
    }

    // Windows are kept in taskbar order (no sorting)
    // This ensures consistent cycling behavior
    Wh_Log(L"Total %d same process windows in taskbar order",
           (int)sameProcessWindows.size());

    // Find current window in the list
    int currentIndex = -1;
    for (size_t i = 0; i < sameProcessWindows.size(); i++) {
        if (sameProcessWindows[i].hwnd == hForeground) {
            currentIndex = static_cast<int>(i);
            Wh_Log(L"Current window found at taskbar index %d (group=%d, button=%d)",
                   currentIndex,
                   sameProcessWindows[i].groupIndex,
                   sameProcessWindows[i].buttonIndex);
            break;
        }
    }

    // If current window not found, use the first one in taskbar
    if (currentIndex == -1) {
        currentIndex = 0;
        Wh_Log(L"Current window not found in list, starting from first window");
    }

    // Calculate next index based on direction
    int nextIndex;
    bool rotateRight = (hotkeyId == kHotkeyIdSameProcessRight || hotkeyId == kHotkeyIdSameProcessRight2);

    if (rotateRight) {
        nextIndex = (currentIndex + 1) % sameProcessWindows.size();
    } else {
        nextIndex = (currentIndex - 1 + sameProcessWindows.size()) % sameProcessWindows.size();
    }

    Wh_Log(L"Switching from taskbar index %d to %d (group=%d, button=%d)",
           currentIndex, nextIndex,
           sameProcessWindows[nextIndex].groupIndex,
           sameProcessWindows[nextIndex].buttonIndex);

    // Switch to the target window
    SwitchToTaskItem(lpTaskListLongPtr, sameProcessWindows[nextIndex].taskItem);

    return true;
}

bool OnTaskbarHotkey(HWND hWnd, int hotkeyId) {
    Wh_Log(L">");

    DWORD messagePos = GetMessagePos();
    POINT pt{
        GET_X_LPARAM(messagePos),
        GET_Y_LPARAM(messagePos),
    };

    HMONITOR monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);

    HWND hTaskbarForMonitor = GetTaskbarForMonitor(hWnd, monitor);

    HWND hMMTaskListWnd = TaskListFromMMTaskbarWnd(
        hTaskbarForMonitor ? hTaskbarForMonitor : hWnd);
    if (!hMMTaskListWnd) {
        return false;
    }

    int clicks = (hotkeyId == kHotkeyIdLeft || hotkeyId == kHotkeyIdLeft2) ? -1 : 1;

    LONG_PTR lpTaskListLongPtr = GetWindowLongPtr(hMMTaskListWnd, 0);
    PVOID targetTaskItem = TaskbarScroll(lpTaskListLongPtr, clicks,
                                         g_settings.wrapAround, nullptr);
    if (targetTaskItem) {
        SwitchToTaskItem(lpTaskListLongPtr, targetTaskItem);
    }

    return true;
}

UINT g_hotkeyRegisteredMsg =
    RegisterWindowMessage(L"Windhawk_hotkey_" WH_MOD_ID);

enum {
    HOTKEY_REGISTER,
    HOTKEY_UNREGISTER,
    HOTKEY_UPDATE,
};

using CTaskListWnd__SetActiveItem_t = void(WINAPI*)(void* pThis,
                                                    void* taskBtnGroup,
                                                    int buttonIndex);
CTaskListWnd__SetActiveItem_t CTaskListWnd__SetActiveItem_Original;
void WINAPI CTaskListWnd__SetActiveItem_Hook(void* pThis,
                                             void* taskBtnGroup,
                                             int buttonIndex) {
    Wh_Log(L">");

    g_lastTaskListActiveItem[pThis] = {
        .taskBtnGroup = taskBtnGroup,
        .buttonIndex = buttonIndex,
    };

    CTaskListWnd__SetActiveItem_Original(pThis, taskBtnGroup, buttonIndex);
}

using CTaskBand_v_WndProc_t = LRESULT(
    WINAPI*)(void* pThis, HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
CTaskBand_v_WndProc_t CTaskBand_v_WndProc_Original;
LRESULT WINAPI CTaskBand_v_WndProc_Hook(void* pThis,
                                        HWND hWnd,
                                        UINT Msg,
                                        WPARAM wParam,
                                        LPARAM lParam) {
    LRESULT result = 0;

    auto originalProc = [pThis](HWND hWnd, UINT Msg, WPARAM wParam,
                                LPARAM lParam) {
        return CTaskBand_v_WndProc_Original(pThis, hWnd, Msg, wParam, lParam);
    };

    switch (Msg) {
        case WM_HOTKEY:
            switch (wParam) {
                case kHotkeyIdLeft:
                case kHotkeyIdRight:
                    OnTaskbarHotkey(GetAncestor(hWnd, GA_ROOT),
                                    static_cast<int>(wParam));
                    break;

                case kHotkeyIdSameProcessLeft:
                case kHotkeyIdSameProcessRight:
                    OnTaskbarSameProcessHotkey(GetAncestor(hWnd, GA_ROOT),
                                               static_cast<int>(wParam));
                    break;

                case kHotkeyIdLeft2:
                case kHotkeyIdRight2:
                    // Second set of cross-process hotkeys, same behavior
                    OnTaskbarHotkey(GetAncestor(hWnd, GA_ROOT),
                                    static_cast<int>(wParam));
                    break;

                case kHotkeyIdSameProcessLeft2:
                case kHotkeyIdSameProcessRight2:
                    // Second set of same-process hotkeys, same behavior
                    OnTaskbarSameProcessHotkey(GetAncestor(hWnd, GA_ROOT),
                                               static_cast<int>(wParam));
                    break;

                case kHotkeyIdMoveGroupLeft:
                    OnMoveButtonGroup(GetAncestor(hWnd, GA_ROOT), true);
                    break;

                case kHotkeyIdMoveGroupRight:
                    OnMoveButtonGroup(GetAncestor(hWnd, GA_ROOT), false);
                    break;

                default:
                    result = originalProc(hWnd, Msg, wParam, lParam);
                    break;
            }
            break;

        case WM_CREATE:
            result = originalProc(hWnd, Msg, wParam, lParam);
            RegisterHotkeys(hWnd);
            break;

        case WM_DESTROY:
            UnregisterHotkeys(hWnd);
            result = originalProc(hWnd, Msg, wParam, lParam);
            break;

        default:
            if (Msg == g_hotkeyRegisteredMsg) {
                switch (wParam) {
                    case HOTKEY_REGISTER:
                        RegisterHotkeys(hWnd);
                        break;

                    case HOTKEY_UNREGISTER:
                        UnregisterHotkeys(hWnd);
                        break;

                    case HOTKEY_UPDATE:
                        UnregisterHotkeys(hWnd);
                        RegisterHotkeys(hWnd);
                        break;
                }
            } else {
                result = originalProc(hWnd, Msg, wParam, lParam);
            }
            break;
    }

    return result;
}

using TrayUI_WndProc_t = LRESULT(WINAPI*)(void* pThis,
                                          HWND hWnd,
                                          UINT Msg,
                                          WPARAM wParam,
                                          LPARAM lParam,
                                          bool* flag);
TrayUI_WndProc_t TrayUI_WndProc_Original;
LRESULT WINAPI TrayUI_WndProc_Hook(void* pThis,
                                   HWND hWnd,
                                   UINT Msg,
                                   WPARAM wParam,
                                   LPARAM lParam,
                                   bool* flag) {
    if (Msg == WM_MOUSEWHEEL && g_settings.enableMouseWheelCycling) {
        HWND hTaskListWnd = TaskListFromTaskbarWnd(hWnd);

        RECT rc{};
        GetWindowRect(hTaskListWnd, &rc);

        POINT pt{
            .x = GET_X_LPARAM(lParam),
            .y = GET_Y_LPARAM(lParam),
        };

        if (PtInRect(&rc, pt)) {
            short delta = GET_WHEEL_DELTA_WPARAM(wParam);

            // Allows to steal focus.
            INPUT input{};
            SendInput(1, &input, sizeof(INPUT));

            OnTaskListScroll(hTaskListWnd, delta);

            *flag = false;
            return 0;
        }
    }

    LRESULT ret =
        TrayUI_WndProc_Original(pThis, hWnd, Msg, wParam, lParam, flag);

    return ret;
}

using CSecondaryTray_v_WndProc_t = LRESULT(
    WINAPI*)(void* pThis, HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
CSecondaryTray_v_WndProc_t CSecondaryTray_v_WndProc_Original;
LRESULT WINAPI CSecondaryTray_v_WndProc_Hook(void* pThis,
                                             HWND hWnd,
                                             UINT Msg,
                                             WPARAM wParam,
                                             LPARAM lParam) {
    if (Msg == WM_MOUSEWHEEL && g_settings.enableMouseWheelCycling) {
        HWND hSecondaryTaskListWnd = TaskListFromSecondaryTaskbarWnd(hWnd);

        RECT rc{};
        GetWindowRect(hSecondaryTaskListWnd, &rc);

        POINT pt{
            .x = GET_X_LPARAM(lParam),
            .y = GET_Y_LPARAM(lParam),
        };

        if (PtInRect(&rc, pt)) {
            short delta = GET_WHEEL_DELTA_WPARAM(wParam);

            // Allows to steal focus.
            INPUT input{};
            SendInput(1, &input, sizeof(INPUT));

            OnTaskListScroll(hSecondaryTaskListWnd, delta);

            return 0;
        }
    }

    LRESULT ret =
        CSecondaryTray_v_WndProc_Original(pThis, hWnd, Msg, wParam, lParam);

    return ret;
}

using TaskbarFrame_OnPointerWheelChanged_t = int(WINAPI*)(PVOID pThis,
                                                          PVOID pArgs);
TaskbarFrame_OnPointerWheelChanged_t
    TaskbarFrame_OnPointerWheelChanged_Original;
int TaskbarFrame_OnPointerWheelChanged_Hook(PVOID pThis, PVOID pArgs) {
    Wh_Log(L">");

    auto original = [=]() {
        return TaskbarFrame_OnPointerWheelChanged_Original(pThis, pArgs);
    };

    if (!g_settings.enableMouseWheelCycling) {
        return original();
    }

    winrt::Windows::Foundation::IInspectable taskbarFrame = nullptr;
    ((IUnknown*)pThis)
        ->QueryInterface(
            winrt::guid_of<winrt::Windows::Foundation::IInspectable>(),
            winrt::put_abi(taskbarFrame));

    if (!taskbarFrame) {
        return original();
    }

    auto className = winrt::get_class_name(taskbarFrame);
    Wh_Log(L"%s", className.c_str());

    if (className != L"Taskbar.TaskbarFrame") {
        return original();
    }

    auto taskbarFrameElement = taskbarFrame.as<UIElement>();

    Input::PointerRoutedEventArgs args = nullptr;
    ((IUnknown*)pArgs)
        ->QueryInterface(winrt::guid_of<Input::PointerRoutedEventArgs>(),
                         winrt::put_abi(args));
    if (!args) {
        return original();
    }

    DWORD messagePos = GetMessagePos();
    POINT pt = {GET_X_LPARAM(messagePos), GET_Y_LPARAM(messagePos)};
    HWND hMMTaskListWnd = TaskListFromPoint(pt);
    if (!hMMTaskListWnd) {
        return original();
    }

    auto currentPoint = args.GetCurrentPoint(taskbarFrameElement);
    double delta = currentPoint.Properties().MouseWheelDelta();
    if (!delta) {
        return original();
    }

    // Allows to steal focus.
    INPUT input;
    ZeroMemory(&input, sizeof(INPUT));
    SendInput(1, &input, sizeof(INPUT));

    OnTaskListScroll(hMMTaskListWnd, static_cast<short>(delta));

    args.Handled(true);
    return 0;
}

void LoadSettings() {
    g_settings.wrapAround = Wh_GetIntSetting(L"wrapAround");
    g_settings.reverseScrollingDirection =
        Wh_GetIntSetting(L"reverseScrollingDirection");
    g_settings.enableMouseWheelCycling =
        Wh_GetIntSetting(L"enableMouseWheelCycling");
    g_settings.cycleLeftKeyboardShortcut =
        WindhawkUtils::StringSetting::make(L"cycleLeftKeyboardShortcut");
    g_settings.cycleRightKeyboardShortcut =
        WindhawkUtils::StringSetting::make(L"cycleRightKeyboardShortcut");
    g_settings.cycleSameProcessLeftKeyboardShortcut =
        WindhawkUtils::StringSetting::make(L"cycleSameProcessLeftKeyboardShortcut");
    g_settings.cycleSameProcessRightKeyboardShortcut =
        WindhawkUtils::StringSetting::make(L"cycleSameProcessRightKeyboardShortcut");
    g_settings.cycleLeftKeyboardShortcut2 =
        WindhawkUtils::StringSetting::make(L"cycleLeftKeyboardShortcut2");
    g_settings.cycleRightKeyboardShortcut2 =
        WindhawkUtils::StringSetting::make(L"cycleRightKeyboardShortcut2");
    g_settings.cycleSameProcessLeftKeyboardShortcut2 =
        WindhawkUtils::StringSetting::make(L"cycleSameProcessLeftKeyboardShortcut2");
    g_settings.cycleSameProcessRightKeyboardShortcut2 =
        WindhawkUtils::StringSetting::make(L"cycleSameProcessRightKeyboardShortcut2");
    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
    g_settings.moveButtonGroupLeftShortcut =
        WindhawkUtils::StringSetting::make(L"moveButtonGroupLeftShortcut");
    g_settings.moveButtonGroupRightShortcut =
        WindhawkUtils::StringSetting::make(L"moveButtonGroupRightShortcut");
}

bool HookTaskbarViewDllSymbols(HMODULE module) {
    // Taskbar.View.dll, ExplorerExtensions.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: virtual int __cdecl winrt::impl::produce<struct winrt::Taskbar::implementation::TaskbarFrame,struct winrt::Windows::UI::Xaml::Controls::IControlOverrides>::OnPointerWheelChanged(void *))"},
            &TaskbarFrame_OnPointerWheelChanged_Original,
            TaskbarFrame_OnPointerWheelChanged_Hook,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
}

HMODULE GetTaskbarViewModuleHandle() {
    HMODULE module = GetModuleHandle(L"Taskbar.View.dll");
    if (!module) {
        module = GetModuleHandle(L"ExplorerExtensions.dll");
    }

    return module;
}

void HandleLoadedModuleIfTaskbarView(HMODULE module, LPCWSTR lpLibFileName) {
    if (g_winVersion >= WinVersion::Win11 && !g_taskbarViewDllLoaded &&
        GetTaskbarViewModuleHandle() == module &&
        !g_taskbarViewDllLoaded.exchange(true)) {
        Wh_Log(L"Loaded %s", lpLibFileName);

        if (HookTaskbarViewDllSymbols(module)) {
            Wh_ApplyHookOperations();
        }
    }
}

bool HookTaskbarSymbols() {
    HMODULE module;
    if (g_winVersion <= WinVersion::Win10) {
        module = GetModuleHandle(nullptr);
    } else {
        module = LoadLibraryEx(L"taskbar.dll", nullptr,
                               LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (!module) {
            Wh_Log(L"Couldn't load taskbar.dll");
            return false;
        }
    }

    // Taskbar.dll, explorer.exe
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(const CTaskListWnd::`vftable'{for `ITaskListUI'})"},
            &CTaskListWnd_vftable_ITaskListUI,
        },
        {
            {LR"(const CTaskListWnd::`vftable'{for `ITaskListSite'})"},
            &CTaskListWnd_vftable_ITaskListSite,
        },
        {
            {LR"(const CTaskListWnd::`vftable'{for `ITaskListAcc'})"},
            &CTaskListWnd_vftable_ITaskListAcc,
        },
        {
            {LR"(const CImmersiveTaskItem::`vftable'{for `ITaskItem'})"},
            &CImmersiveTaskItem_vftable,
        },
        {
            {LR"(public: virtual int __cdecl CTaskListWnd::GetButtonGroupCount(void))"},
            &CTaskListWnd_GetButtonGroupCount,
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
            {LR"(public: virtual int __cdecl CTaskBtnGroup::GetNumItems(void))"},
            &CTaskBtnGroup_GetNumItems,
        },
        {
            {LR"(public: virtual struct ITaskItem * __cdecl CTaskBtnGroup::GetTaskItem(int))"},
            &CTaskBtnGroup_GetTaskItem,
        },
        {
            {LR"(public: virtual struct HWND__ * __cdecl CWindowTaskItem::GetWindow(void))"},
            &CWindowTaskItem_GetWindow_Original,
        },
        {
            {LR"(public: virtual struct HWND__ * __cdecl CImmersiveTaskItem::GetWindow(void))"},
            &CImmersiveTaskItem_GetWindow_Original,
        },
        {
            {LR"(public: virtual void __cdecl CTaskListWnd::SwitchToItem(struct ITaskItem *))"},
            &CTaskListWnd_SwitchToItem_Original,
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
            {LR"(protected: void __cdecl CTaskListWnd::_SetActiveItem(struct ITaskBtnGroup *,int))"},
            &CTaskListWnd__SetActiveItem_Original,
            CTaskListWnd__SetActiveItem_Hook,
        },
        {
            {LR"(protected: virtual __int64 __cdecl CTaskBand::v_WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64))"},
            &CTaskBand_v_WndProc_Original,
            CTaskBand_v_WndProc_Hook,
        },
        {
            {LR"(public: virtual __int64 __cdecl TrayUI::WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64,bool *))"},
            &TrayUI_WndProc_Original,
            TrayUI_WndProc_Hook,
        },
        {
            {LR"(private: virtual __int64 __cdecl CSecondaryTray::v_WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64))"},
            &CSecondaryTray_v_WndProc_Original,
            CSecondaryTray_v_WndProc_Hook,
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
        {R"(??_7CTaskListWnd@@6BITaskListUI@@@)",
         &CTaskListWnd_vftable_ITaskListUI},
        {R"(??_7CTaskListWnd@@6BITaskListSite@@@)",
         &CTaskListWnd_vftable_ITaskListSite},
        {R"(??_7CTaskListWnd@@6BITaskListAcc@@@)",
         &CTaskListWnd_vftable_ITaskListAcc},
        {R"(??_7CImmersiveTaskItem@@6BITaskItem@@@)",
         &CImmersiveTaskItem_vftable},
        {R"(?GetButtonGroupCount@CTaskListWnd@@UEAAHXZ)",
         &CTaskListWnd_GetButtonGroupCount},
        {R"(?_GetTBGroupFromGroup@CTaskListWnd@@IEAAPEAUITaskBtnGroup@@PEAUITaskGroup@@PEAH@Z)",
         &CTaskListWnd__GetTBGroupFromGroup},
        {R"(?GetGroupType@CTaskBtnGroup@@UEAA?AW4eTBGROUPTYPE@@XZ)",
         &CTaskBtnGroup_GetGroupType},
        {R"(?GetNumItems@CTaskBtnGroup@@UEAAHXZ)", &CTaskBtnGroup_GetNumItems},
        {R"(?GetTaskItem@CTaskBtnGroup@@UEAAPEAUITaskItem@@H@Z)",
         &CTaskBtnGroup_GetTaskItem},
        {R"(?GetWindow@CWindowTaskItem@@UEAAPEAUHWND__@@XZ)",
         &CWindowTaskItem_GetWindow_Original},
        {R"(?GetWindow@CImmersiveTaskItem@@UEAAPEAUHWND__@@XZ)",
         &CImmersiveTaskItem_GetWindow_Original},
        {R"(?SwitchToItem@CTaskListWnd@@UEAAXPEAUITaskItem@@@Z)",
         &CTaskListWnd_SwitchToItem_Original},
        {R"(?_SetActiveItem@CTaskListWnd@@IEAAXPEAUITaskBtnGroup@@H@Z)",
         &CTaskListWnd__SetActiveItem_Original,
         CTaskListWnd__SetActiveItem_Hook},
        {R"(?v_WndProc@CTaskBand@@MEAA_JPEAUHWND__@@I_K_J@Z)",
         &CTaskBand_v_WndProc_Original, CTaskBand_v_WndProc_Hook},
        {R"(?WndProc@TrayUI@@UEAA_JPEAUHWND__@@I_K_JPEA_N@Z)",
         &TrayUI_WndProc_Original, TrayUI_WndProc_Hook},
        // Exported after 67.1:
        {R"(?v_WndProc@CSecondaryTray@@EEAA_JPEAUHWND__@@I_K_J@Z)",
         &CSecondaryTray_v_WndProc_Original, CSecondaryTray_v_WndProc_Hook,
         true},
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

void HandleLoadedModuleIfExplorerPatcher(HMODULE module) {
    if (module && !((ULONG_PTR)module & 3) && !g_explorerPatcherInitialized) {
        if (IsExplorerPatcherModule(module)) {
            HookExplorerPatcherSymbols(module);
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
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            g_taskbarViewDllLoaded = true;
            if (!HookTaskbarViewDllSymbols(taskbarViewModule)) {
                return FALSE;
            }
        } else {
            Wh_Log(L"Taskbar view module not loaded yet");
        }

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
    WindhawkUtils::SetFunctionHook(pKernelBaseLoadLibraryExW,
                                   LoadLibraryExW_Hook,
                                   &LoadLibraryExW_Original);

    // Hook DPA_GetPtr for button group reordering
    WindhawkUtils::SetFunctionHook(DPA_GetPtr, DPA_GetPtr_Hook,
                                   &DPA_GetPtr_Original);

    g_initialized = true;

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    if (g_winVersion >= WinVersion::Win11 && !g_taskbarViewDllLoaded) {
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
    if (!g_explorerPatcherInitialized) {
        HandleLoadedExplorerPatcher();
    }

    if (HWND hTaskBandWnd = GetTaskBandWnd()) {
        SendMessage(hTaskBandWnd, g_hotkeyRegisteredMsg, HOTKEY_REGISTER, 0);
    }
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");

    if (HWND hTaskBandWnd = GetTaskBandWnd()) {
        SendMessage(hTaskBandWnd, g_hotkeyRegisteredMsg, HOTKEY_UNREGISTER, 0);
    }
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;

    LoadSettings();

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11;
    if (*bReload) {
        return TRUE;
    }

    if (HWND hTaskBandWnd = GetTaskBandWnd()) {
        SendMessage(hTaskBandWnd, g_hotkeyRegisteredMsg, HOTKEY_UPDATE, 0);
    }

    return TRUE;
}
