// ==WindhawkMod==
// @id              taskbar-button-click
// @name            Middle click to close on the taskbar - Fork
// @description     Close programs with a middle click on the taskbar instead of creating a new instance. Allows disabling override with modifier keys.
// @version         1.0.9 
// @author          m417z (original), Maxim Fomin
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lversion -lwininet
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
# Middle click to close on the taskbar

Close programs with the middle click on the taskbar instead of creating a new
instance.

Holding Ctrl while middle clicking will end the running task. The key
combination can be configured or disabled in the mod settings.

Holding a configured key (Ctrl or Alt) while middle clicking can disable the
override, reverting to the default Windows behavior (launching a new instance).
This can be configured in the mod settings.

**Priority Warning:** If you configure the *same* modifier key (e.g., Ctrl) for both "End Task"
and "Disable Override", pressing that key during a middle click will prioritize
**disabling the override** (launching a new instance) over ending the task.

Only Windows 10 64-bit and Windows 11 are supported. For other Windows versions
check out [7+ Taskbar Tweaker](https://tweaker.ramensoftware.com/).

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.

![Demonstration](https://i.imgur.com/qeO9tLG.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- multipleItemsBehavior: closeAll
  $name: Multiple items behavior
  $description: >-
    You can choose the desired behavior for middle clicking on a group of
    windows
  $options:
  - closeAll: Close all windows
  - closeForeground: Close foreground window
  - none: Do nothing
- keysToEndTask:
  - Ctrl: true
  - Alt: false
  $name: Keys to end task
  $description: >-
    A combination of keys that can be pressed while middle clicking to
    forcefully end the running task.

    Note: This option won't have effect on a group of taskbar items.
    Warning: If the same key is selected here and in "Keys to disable override",
    disabling the override will take priority.
- keysToDisableOverride:
  - Ctrl: false
  - Alt: false
  $name: Keys to disable override
  $description: >-
    A combination of keys that can be pressed while middle clicking to
    revert to the default behavior (launching a new instance) instead of
    closing the window.
    Warning: If the same key is selected here and in "Keys to end task",
    this action (disabling the override) will take priority.
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool). Note: For Windhawk versions older than
    1.3, you have to disable and re-enable the mod to apply this option.
*/
// ==/WindhawkModSettings==

#include <psapi.h>
#include <wininet.h>

#include <algorithm>
#include <atomic>
#include <string>
#include <string_view>
#include <vector>
#include <optional> // Required for std::optional

// Enum for multiple items behavior setting
enum {
    MULTIPLE_ITEMS_BEHAVIOR_NONE,
    MULTIPLE_ITEMS_BEHAVIOR_CLOSE_ALL,
    MULTIPLE_ITEMS_BEHAVIOR_CLOSE_FOREGROUND,
};

// Structure to hold mod settings
struct {
    int multipleItemsBehavior;
    bool keysToEndTaskCtrl;
    bool keysToEndTaskAlt;
    bool keysToDisableOverrideCtrl; // New setting
    bool keysToDisableOverrideAlt;  // New setting
    bool oldTaskbarOnWin11;
} g_settings;

// Enum for Windows version detection
enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_24H2,
};

WinVersion g_winVersion; // Global variable for detected Windows version

// Atomic flags for initialization state
std::atomic<bool> g_initialized;
std::atomic<bool> g_explorerPatcherInitialized;

// Typedefs for original function pointers
using CTaskListWnd_HandleClick_t = long(WINAPI*)(
    LPVOID pThis,
    LPVOID,  // ITaskGroup *
    LPVOID,  // ITaskItem *
    LPVOID   // winrt::Windows::System::LauncherOptions const &
);
CTaskListWnd_HandleClick_t CTaskListWnd_HandleClick_Original;

using CTaskListWnd__HandleClick_t =
    void(WINAPI*)(LPVOID pThis,
                  LPVOID,  // ITaskBtnGroup *
                  int,
                  int,  // enum CTaskListWnd::eCLICKACTION
                  int,
                  int);
CTaskListWnd__HandleClick_t CTaskListWnd__HandleClick_Original;

using CTaskBand_Launch_t = long(WINAPI*)(LPVOID pThis,
                                         LPVOID,  // ITaskGroup *
                                         LPVOID,  // tagPOINT const &
                                         int  // enum LaunchFromTaskbarOptions
);
CTaskBand_Launch_t CTaskBand_Launch_Original;

using CTaskListWnd_GetActiveBtn_t = HRESULT(WINAPI*)(LPVOID pThis,
                                                     LPVOID*,  // ITaskGroup **
                                                     int*);
CTaskListWnd_GetActiveBtn_t CTaskListWnd_GetActiveBtn_Original;

using CTaskListWnd_ProcessJumpViewCloseWindow_t =
    void(WINAPI*)(LPVOID pThis,
                  HWND,
                  LPVOID,  // struct ITaskGroup *
                  HMONITOR);
CTaskListWnd_ProcessJumpViewCloseWindow_t
    CTaskListWnd_ProcessJumpViewCloseWindow_Original;

using CTaskBand__EndTask_t = void(WINAPI*)(LPVOID pThis,
                                           HWND hWnd,
                                           BOOL bForce);
CTaskBand__EndTask_t CTaskBand__EndTask_Original;

using CTaskBtnGroup_GetGroupType_t = int(WINAPI*)(LPVOID pThis);
CTaskBtnGroup_GetGroupType_t CTaskBtnGroup_GetGroupType_Original;

using CTaskBtnGroup_GetGroup_t = LPVOID(WINAPI*)(LPVOID pThis);
CTaskBtnGroup_GetGroup_t CTaskBtnGroup_GetGroup_Original;

using CTaskBtnGroup_GetTaskItem_t = void*(WINAPI*)(LPVOID pThis, int);
CTaskBtnGroup_GetTaskItem_t CTaskBtnGroup_GetTaskItem_Original;

using CWindowTaskItem_GetWindow_t = HWND(WINAPI*)(LPVOID pThis);
CWindowTaskItem_GetWindow_t CWindowTaskItem_GetWindow_Original;

using CImmersiveTaskItem_GetWindow_t = HWND(WINAPI*)(LPVOID pThis);
CImmersiveTaskItem_GetWindow_t CImmersiveTaskItem_GetWindow_Original;

// Pointer to the vftable of CImmersiveTaskItem (used for type checking)
void* CImmersiveTaskItem_vftable;

// Global variables to store context during click handling
LPVOID g_pCTaskListWndHandlingClick; // Pointer to CTaskListWnd instance handling the click
LPVOID g_pCTaskListWndTaskBtnGroup;  // Pointer to the ITaskBtnGroup being clicked
int g_CTaskListWndTaskItemIndex = -1; // Index of the specific item clicked within a group
int g_CTaskListWndClickAction = -1;   // Type of click action (e.g., middle click)

// Hook for CTaskListWnd::HandleClick (Win11 specific)
long WINAPI CTaskListWnd_HandleClick_Hook(LPVOID pThis,
                                          LPVOID param1,
                                          LPVOID param2,
                                          LPVOID param3) {
    Wh_Log(L"> CTaskListWnd::HandleClick");

    // Store the 'this' pointer to provide context for later hooks
    g_pCTaskListWndHandlingClick = pThis;

    // Call the original function
    long ret = CTaskListWnd_HandleClick_Original(pThis, param1, param2, param3);

    // Clear the context pointer
    g_pCTaskListWndHandlingClick = nullptr;

    return ret;
}

// Hook for CTaskListWnd::_HandleClick (Win10 and Win11)
void WINAPI CTaskListWnd__HandleClick_Hook(LPVOID pThis,
                                           LPVOID taskBtnGroup,
                                           int taskItemIndex,
                                           int clickAction,
                                           int param4,
                                           int param5) {
    Wh_Log(L"> CTaskListWnd::_HandleClick Action: %d", clickAction);

    // If the Win11 specific hook isn't available, estimate the CTaskListWnd pointer
    // This is a workaround for Win10 where HandleClick isn't directly hooked.
    if (!CTaskListWnd_HandleClick_Original) {
        // A magic offset determined by reverse engineering (may be fragile)
        g_pCTaskListWndHandlingClick = (BYTE*)pThis + 0x28;
    }

    // Store context information about the click
    g_pCTaskListWndTaskBtnGroup = taskBtnGroup;
    g_CTaskListWndTaskItemIndex = taskItemIndex;
    g_CTaskListWndClickAction = clickAction;

    // Call the original function
    CTaskListWnd__HandleClick_Original(pThis, taskBtnGroup, taskItemIndex,
                                       clickAction, param4, param5);

    // Clear context information if we estimated the pointer (Win10 case)
    if (!CTaskListWnd_HandleClick_Original) {
        g_pCTaskListWndHandlingClick = nullptr;
    }

    // Clear the rest of the context information
    g_pCTaskListWndTaskBtnGroup = nullptr;
    g_CTaskListWndTaskItemIndex = -1;
    g_CTaskListWndClickAction = -1;
}

// Hook for CTaskBand::Launch, where the core logic resides
long WINAPI CTaskBand_Launch_Hook(LPVOID pThis,
                                  LPVOID taskGroup,
                                  LPVOID param2,
                                  int param3) {
    Wh_Log(L"> CTaskBand::Launch");

    // Lambda to easily call the original function
    auto original = [=]() {
        return CTaskBand_Launch_Original(pThis, taskGroup, param2, param3);
    };

    // Check if we have the necessary context from the _HandleClick hooks
    if (!g_pCTaskListWndHandlingClick || !g_pCTaskListWndTaskBtnGroup) {
        Wh_Log(L"Missing click context, calling original.");
        return original();
    }

    // Get the real task group from taskBtnGroup.
    // This is a workaround for compatibility with other mods (like taskbar-grouping)
    // that might modify the 'taskGroup' parameter passed to this function.
    LPVOID realTaskGroup =
        CTaskBtnGroup_GetGroup_Original(g_pCTaskListWndTaskBtnGroup);
    if (!realTaskGroup) {
        Wh_Log(L"Failed to get real task group, calling original.");
        return original();
    }

    // Check if the action is a middle click (action code 3).
    // Also check if Shift key is NOT pressed, because Shift+LeftClick also triggers
    // the launch action, but we only want to intercept middle clicks.
    if (g_CTaskListWndClickAction != 3 || GetKeyState(VK_SHIFT) < 0) {
        Wh_Log(L"Not a middle click or Shift is pressed, calling original.");
        return original();
    }

    // --- Priority Logic: Check for "Disable Override" first ---
    bool ctrlDown = GetKeyState(VK_CONTROL) < 0;
    bool altDown = GetKeyState(VK_MENU) < 0; // VK_MENU is Alt

    bool disableOverride = (ctrlDown && g_settings.keysToDisableOverrideCtrl) ||
                           (altDown && g_settings.keysToDisableOverrideAlt);

    if (disableOverride) {
        Wh_Log(L"Override disabled by modifier key(s). Calling original function (priority).");
        // Revert to default Windows behavior (launch new instance)
        return original(); // <--- PRIORITY EXIT
    }
    // --- End of Priority Logic ---


    // Determine the type of the taskbar button group.
    // Group types:
    // 1 - Single item or multiple uncombined items
    // 2 - Pinned item (we don't want to close pinned items on middle click)
    // 3 - Multiple combined items (group)
    int groupType =
        CTaskBtnGroup_GetGroupType_Original(g_pCTaskListWndTaskBtnGroup);
    Wh_Log(L"Group type: %d", groupType);

    if (groupType != 1 && groupType != 3) {
        Wh_Log(L"Group type is not 1 or 3, calling original.");
        // Don't interfere with pinned items or unknown types
        return original();
    }

    int taskItemIndex = -1; // Index of the window to close/end

    // Handle behavior for grouped items (type 3)
    if (groupType == 3) {
        switch (g_settings.multipleItemsBehavior) {
            case MULTIPLE_ITEMS_BEHAVIOR_NONE:
                Wh_Log(L"Multiple items behavior: None. Doing nothing.");
                return 0; // Do nothing as configured

            case MULTIPLE_ITEMS_BEHAVIOR_CLOSE_FOREGROUND:
                {
                    Wh_Log(L"Multiple items behavior: Close Foreground.");
                    void* activeTaskGroup = nullptr;
                    // Get the currently active (foreground) button in the task list
                    // Note: The CTaskListWnd pointer might need adjustment (0x18 offset observed)
                    HRESULT hr = CTaskListWnd_GetActiveBtn_Original(
                        (BYTE*)g_pCTaskListWndHandlingClick + 0x18, // Potential offset needed
                        &activeTaskGroup, &taskItemIndex);

                    if (FAILED(hr) || !activeTaskGroup) {
                        Wh_Log(L"Failed to get active button or no active button.");
                        return 0; // Failed to get active button
                    }

                    // Release the obtained interface pointer
                    if (activeTaskGroup) { // Check if pointer is valid before releasing
                       ((IUnknown*)activeTaskGroup)->Release();
                    }


                    // Check if the active button belongs to the group we clicked
                    // and if we got a valid index.
                    if (activeTaskGroup != realTaskGroup || taskItemIndex < 0) {
                        Wh_Log(L"Active button is not in the clicked group or index invalid.");
                        return 0; // Active button is not in this group, do nothing
                    }
                    Wh_Log(L"Targeting foreground item at index: %d", taskItemIndex);
                }
                break;

            case MULTIPLE_ITEMS_BEHAVIOR_CLOSE_ALL:
            default:
                 Wh_Log(L"Multiple items behavior: Close All (or default). Targeting group.");
                // For "Close All", we don't need a specific index.
                // We'll pass nullptr for HWND later, which ProcessJumpViewCloseWindow
                // interprets as closing all windows in the group.
                taskItemIndex = -1; // Indicate closing the whole group
                break;
        }
    } else { // Single item (type 1)
        taskItemIndex = g_CTaskListWndTaskItemIndex; // Use the index from the click context
        Wh_Log(L"Single item clicked, index: %d", taskItemIndex);
    }

    HWND hWnd = nullptr; // Window handle to close or end

    // If we have a valid index (not closing all in a group), get the HWND
    if (taskItemIndex >= 0) {
        // Get the specific task item object from the group
        void* taskItem = CTaskBtnGroup_GetTaskItem_Original(
            g_pCTaskListWndTaskBtnGroup, taskItemIndex);

        if (taskItem) {
            bool isImmersive = false;
            // Check if it's a UWP/Immersive app window
            if (CImmersiveTaskItem_vftable) {
                // Preferred method: Compare vftable pointer
                isImmersive = (*(void**)taskItem == CImmersiveTaskItem_vftable);
            } else {
                // Fallback for ExplorerPatcher where vftables might not be exported:
                // Call the IsImmersive method via the vtable index (57 observed).
                // This is less robust than direct symbol hooking.
                // Add basic validation for taskItem pointer before dereferencing
                try {
                    using IsImmersive_t = bool(WINAPI*)(PVOID pThis);
                    IsImmersive_t pIsImmersive =
                        (IsImmersive_t)(*(void***)taskItem)[57]; // Magic index 57
                    if (pIsImmersive) { // Check if function pointer seems valid
                       isImmersive = pIsImmersive(taskItem);
                    } else {
                        Wh_Log(L"Warning: Could not get IsImmersive function pointer from vtable.");
                    }
                } catch (...) {
                     Wh_Log(L"Warning: Exception accessing taskItem vtable for IsImmersive check.");
                     // Assume not immersive if check fails
                }
            }


            // Call the appropriate GetWindow method based on the type
            if (isImmersive) {
                hWnd = CImmersiveTaskItem_GetWindow_Original(taskItem);
                Wh_Log(L"Identified as Immersive Task Item.");
            } else {
                hWnd = CWindowTaskItem_GetWindow_Original(taskItem);
                Wh_Log(L"Identified as Window Task Item.");
            }
        } else {
             Wh_Log(L"Failed to get Task Item at index %d.", taskItemIndex);
        }
    }

    // Check if the "End Task" modifier keys are pressed according to settings
    // This check only happens if the "Disable Override" check above was false.
    bool endTask = (hWnd != nullptr || taskItemIndex >= 0) && // Don't end task on group click
                   (ctrlDown || altDown) && // Use the already fetched key states
                   g_settings.keysToEndTaskCtrl == ctrlDown &&
                   g_settings.keysToEndTaskAlt == altDown;

    if (endTask) {
        // Forcefully end the task
        if (hWnd) {
            Wh_Log(L"Ending task for HWND %p", hWnd);
            // Use the internal CTaskBand::_EndTask function
            CTaskBand__EndTask_Original(pThis, hWnd, TRUE); // TRUE for force
        } else {
            // This case should ideally not happen if endTask is true,
            // but log it just in case.
            Wh_Log(L"End task requested, but no valid HWND found.");
        }
    } else {
        // Gracefully close the window(s)
        // This block is reached if:
        // 1. No relevant modifier key was pressed.
        // 2. A modifier key was pressed, but it wasn't configured for "Disable Override" or "End Task".
        // 3. A modifier key was pressed and configured for "End Task", but not for "Disable Override".
        Wh_Log(L"Closing HWND %p (nullptr means close all in group)", hWnd);

        POINT pt;
        GetCursorPos(&pt); // Get cursor position to determine the monitor
        HMONITOR monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);

        // Use the internal CTaskListWnd::ProcessJumpViewCloseWindow function.
        // Passing nullptr for hWnd when taskItemIndex was -1 (Close All)
        // signals it to close all windows in the realTaskGroup.
        CTaskListWnd_ProcessJumpViewCloseWindow_Original(
            g_pCTaskListWndHandlingClick, hWnd, realTaskGroup, monitor);
    }

    // Prevent the original launch action by returning 0 (S_OK usually means success)
    return 0;
}


// Helper function to get module version information
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
            // No need to explicitly FreeResource for RT_VERSION resources
        }
    }

    if (puPtrLen) {
        *puPtrLen = uPtrLen;
    }

    return (VS_FIXEDFILEINFO*)pFixedFileInfo;
}

// Function to determine the Windows version based on explorer.exe version
WinVersion GetExplorerVersion() {
    VS_FIXEDFILEINFO* fixedFileInfo = GetModuleVersionInfo(nullptr, nullptr); // nullptr for current process (explorer.exe)
    if (!fixedFileInfo) {
        Wh_Log(L"Failed to get explorer.exe version info.");
        return WinVersion::Unsupported;
    }

    WORD major = HIWORD(fixedFileInfo->dwFileVersionMS);
    WORD minor = LOWORD(fixedFileInfo->dwFileVersionMS);
    WORD build = HIWORD(fixedFileInfo->dwFileVersionLS);
    WORD qfe = LOWORD(fixedFileInfo->dwFileVersionLS);

    Wh_Log(L"Explorer Version: %u.%u.%u.%u", major, minor, build, qfe);

    // Classify based on build number
    if (major == 10) { // Windows 10/11 use major version 10
        if (build < 22000) {
            return WinVersion::Win10;
        } else if (build < 26100) { // 24H2 starts at 26100
            return WinVersion::Win11;
        } else {
            return WinVersion::Win11_24H2;
        }
    }

    Wh_Log(L"Unsupported major version: %u", major);
    return WinVersion::Unsupported;
}

// Structure to define a symbol hook configuration
struct SYMBOL_HOOK {
    std::vector<std::wstring_view> symbols; // List of possible symbol names (mangled)
    void** pOriginalFunction;               // Pointer to store the original function address
    void* hookFunction = nullptr;           // The hook function address (if any)
    bool optional = false;                  // Whether finding this symbol is optional
};

// Function to hook symbols, using a cache for performance
bool HookSymbols(HMODULE module,
                 const SYMBOL_HOOK* symbolHooks,
                 size_t symbolHooksCount,
                 bool cacheOnly = false) {
    const WCHAR cacheVer = L'1'; // Version marker for cache format
    const WCHAR cacheSep = L'@'; // Separator character in cache string
    constexpr size_t cacheMaxSize = 10240; // Max size for the cache string

    WCHAR moduleFilePath[MAX_PATH];
    if (!GetModuleFileName(module, moduleFilePath, ARRAYSIZE(moduleFilePath))) {
        Wh_Log(L"HookSymbols: GetModuleFileName failed");
        return false;
    }

    PCWSTR moduleFileName = wcsrchr(moduleFilePath, L'\\');
    if (!moduleFileName) {
        moduleFileName = moduleFilePath; // Use the full path if no backslash
    } else {
        moduleFileName++; // Skip the backslash
    }

    // Construct the cache key using the module filename
    WCHAR cacheStrKeyBuffer[MAX_PATH + 30];
    swprintf_s(cacheStrKeyBuffer, ARRAYSIZE(cacheStrKeyBuffer), L"symbol-cache-%s", moduleFileName);
    std::wstring_view cacheStrKey(cacheStrKeyBuffer);


    WCHAR cacheBuffer[cacheMaxSize + 1];
    Wh_GetStringValue(cacheStrKey.data(), cacheBuffer, ARRAYSIZE(cacheBuffer));

    std::wstring_view cacheBufferView(cacheBuffer);

    // Helper lambda to split a string_view
    auto splitStringView = [](std::wstring_view s, WCHAR delimiter) {
        size_t pos_start = 0, pos_end;
        std::wstring_view token;
        std::vector<std::wstring_view> res;

        while ((pos_end = s.find(delimiter, pos_start)) != std::wstring_view::npos) {
            token = s.substr(pos_start, pos_end - pos_start);
            pos_start = pos_end + 1;
            res.push_back(token);
        }

        res.push_back(s.substr(pos_start));
        return res;
    };

    auto cacheParts = splitStringView(cacheBufferView, cacheSep);

    std::vector<bool> symbolResolved(symbolHooksCount, false);
    std::wstring newSystemCacheStr; // String to build the new cache entry

    // Lambda called when a symbol is resolved (either from cache or live lookup)
    auto onSymbolResolved = [&](std::wstring_view symbol, void* address) {
        for (size_t i = 0; i < symbolHooksCount; i++) {
            if (symbolResolved[i]) {
                continue; // Already resolved this hook entry
            }

            // Check if the found symbol matches any of the names for this hook entry
            bool match = false;
            for (auto hookSymbol : symbolHooks[i].symbols) {
                if (hookSymbol == symbol) {
                    match = true;
                    break;
                }
            }

            if (!match) {
                continue; // Symbol doesn't match this hook entry
            }

            // Apply the hook or store the original address
            if (symbolHooks[i].hookFunction) {
                Wh_SetFunctionHook(address, symbolHooks[i].hookFunction,
                                   symbolHooks[i].pOriginalFunction);
                Wh_Log(L"Hooked %p: %.*s", address, (int)symbol.length(), symbol.data());
            } else {
                *symbolHooks[i].pOriginalFunction = address;
                Wh_Log(L"Found %p: %.*s", address, (int)symbol.length(), symbol.data());
            }

            symbolResolved[i] = true; // Mark this hook entry as resolved

            // Add the resolved symbol and its offset to the new cache string
            newSystemCacheStr += cacheSep;
            newSystemCacheStr += symbol;
            newSystemCacheStr += cacheSep;
            newSystemCacheStr += std::to_wstring((ULONG_PTR)address - (ULONG_PTR)module);

            break; // Move to the next hook entry
        }
    };

    // Get module timestamp and size for cache validation
    IMAGE_DOS_HEADER* dosHeader = (IMAGE_DOS_HEADER*)module;
    IMAGE_NT_HEADERS* header = nullptr;
     if (dosHeader && dosHeader->e_magic == IMAGE_DOS_SIGNATURE) {
        header = (IMAGE_NT_HEADERS*)((BYTE*)dosHeader + dosHeader->e_lfanew);
        if (!IsBadReadPtr(header, sizeof(IMAGE_NT_HEADERS)) && header->Signature == IMAGE_NT_SIGNATURE) {
             // Header looks valid
        } else {
             header = nullptr; // Invalid header
        }
    }


    if (!header) {
        Wh_Log(L"HookSymbols: Invalid module header for %s", moduleFileName);
        return false; // Cannot proceed without valid header
    }

    auto timeStamp = std::to_wstring(header->FileHeader.TimeDateStamp);
    auto imageSize = std::to_wstring(header->OptionalHeader.SizeOfImage);

    // Start building the new cache string with validation info
    newSystemCacheStr += cacheVer;
    newSystemCacheStr += cacheSep;
    newSystemCacheStr += timeStamp;
    newSystemCacheStr += cacheSep;
    newSystemCacheStr += imageSize;

    // Try to resolve symbols using the existing cache
    bool cacheValid = false;
    if (cacheParts.size() >= 3 &&
        cacheParts[0] == std::wstring_view(&cacheVer, 1) &&
        cacheParts[1] == timeStamp && cacheParts[2] == imageSize)
    {
        cacheValid = true;
        Wh_Log(L"Cache is valid for %s", moduleFileName);
        for (size_t i = 3; i + 1 < cacheParts.size(); i += 2) {
            auto symbol = cacheParts[i];
            auto addressStr = cacheParts[i + 1];
            if (addressStr.length() == 0) { // Symbol not found in previous lookup
                continue;
            }

            // Calculate the absolute address from the cached offset
            void* addressPtr = nullptr;
            try {
                 addressPtr = (void*)(std::stoull(std::wstring(addressStr)) + (ULONG_PTR)module);
            } catch (const std::exception& e) {
                Wh_Log(L"Error parsing cached address '%ls': %S", addressStr.data(), e.what());
                continue;
            }

            onSymbolResolved(symbol, addressPtr);
        }

        // Check if optional symbols were explicitly marked as not found in the cache
        for (size_t i = 0; i < symbolHooksCount; i++) {
            if (symbolResolved[i] || !symbolHooks[i].optional) {
                continue; // Already resolved or mandatory
            }

            size_t notFoundCount = 0;
            for (const auto& hookSymbol : symbolHooks[i].symbols) {
                 bool foundInCache = false;
                 bool markedAsNotFound = false;
                 for (size_t j = 3; j + 1 < cacheParts.size(); j += 2) {
                     if (cacheParts[j] == hookSymbol) {
                         foundInCache = true;
                         if (cacheParts[j+1].length() == 0) {
                             markedAsNotFound = true;
                         }
                         break;
                     }
                 }
                 if (foundInCache && markedAsNotFound) {
                     notFoundCount++;
                 } else if (!foundInCache) {
                     // If *any* symbol name for this optional hook wasn't even in the cache,
                     // we can't be sure it doesn't exist without a live lookup.
                     notFoundCount = 0; // Force live lookup
                     break;
                 }
            }


            if (notFoundCount == symbolHooks[i].symbols.size()) {
                Wh_Log(L"Optional symbol %d confirmed absent from cache for %s", (int)i, moduleFileName);
                symbolResolved[i] = true; // Mark as resolved (because it's confirmed absent)
                // Add entries to the new cache indicating absence
                for (auto hookSymbol : symbolHooks[i].symbols) {
                    newSystemCacheStr += cacheSep;
                    newSystemCacheStr += hookSymbol;
                    newSystemCacheStr += cacheSep; // Empty address means not found
                }
            }
        }
    } else {
         Wh_Log(L"Cache invalid or missing for %s", moduleFileName);
    }


    // If all symbols were resolved from the cache, we're done
    if (std::all_of(symbolResolved.begin(), symbolResolved.end(), [](bool b){ return b; })) {
        Wh_Log(L"All symbols resolved from cache for %s.", moduleFileName);
        return true;
    }

    // If cacheOnly was requested and we couldn't resolve everything, return failure
    if (cacheOnly) {
        Wh_Log(L"Cache-only lookup failed for %s.", moduleFileName);
        return false;
    }

    // --- Perform Live Symbol Lookup ---
    Wh_Log(L"Performing live symbol lookup for %s...", moduleFileName);

    WH_FIND_SYMBOL findSymbol;
    HANDLE findSymbolHandle = Wh_FindFirstSymbol(module, nullptr, &findSymbol);
    if (!findSymbolHandle || findSymbolHandle == INVALID_HANDLE_VALUE) {
        Wh_Log(L"Wh_FindFirstSymbol failed for %s (Error: %p)", moduleFileName, GetLastError());
        // Don't immediately fail, maybe optional symbols were the only ones missing
    } else {
        do {
            // Pass found symbols to our handler
            onSymbolResolved(findSymbol.symbol, findSymbol.address);
        } while (Wh_FindNextSymbol(findSymbolHandle, &findSymbol));

        Wh_FindCloseSymbol(findSymbolHandle);
    }


    // Check final resolution status
    bool allResolved = true;
    for (size_t i = 0; i < symbolHooksCount; i++) {
        if (symbolResolved[i]) {
            continue; // Already resolved
        }

        // If it's still unresolved after live lookup
        if (!symbolHooks[i].optional) {
            // Mandatory symbol not found - this is an error
            Wh_Log(L"Mandatory symbol %d unresolved for %s. First name: %.*s",
                   (int)i, moduleFileName, (int)symbolHooks[i].symbols[0].length(), symbolHooks[i].symbols[0].data());
            allResolved = false;
            // Don't break, continue checking other symbols for logging purposes
        } else {
            // Optional symbol not found - this is acceptable
            Wh_Log(L"Optional symbol %d confirmed absent after live lookup for %s", (int)i, moduleFileName);
            symbolResolved[i] = true; // Mark as resolved (because it's confirmed absent)
            // Add entries to the new cache indicating absence
            for (auto hookSymbol : symbolHooks[i].symbols) {
                newSystemCacheStr += cacheSep;
                newSystemCacheStr += hookSymbol;
                newSystemCacheStr += cacheSep; // Empty address means not found
            }
        }
    }

    // Save the newly generated cache string if it's not too large
    if (allResolved) {
        if (newSystemCacheStr.length() <= cacheMaxSize) {
            Wh_SetStringValue(cacheStrKey.data(), newSystemCacheStr.c_str());
            Wh_Log(L"Updated symbol cache for %s", moduleFileName);
        } else {
            Wh_Log(L"New cache for %s is too large (%zu bytes), not saving.", moduleFileName, newSystemCacheStr.length());
            // Consider deleting the old invalid cache entry
            Wh_SetStringValue(cacheStrKey.data(), L"");
        }
    } else {
         Wh_Log(L"Failed to resolve all mandatory symbols for %s.", moduleFileName);
         // Delete potentially invalid cache
         Wh_SetStringValue(cacheStrKey.data(), L"");
    }


    return allResolved;
}


// --- Online Cache Fallback (Optional, kept from original for reference) ---
// Helper function to download content from a URL
std::optional<std::wstring> GetUrlContent(PCWSTR lpUrl) {
    Wh_Log(L"Attempting to download URL: %s", lpUrl);
    HINTERNET hOpenHandle = InternetOpen(
        L"WindhawkMod/1.0 (Taskbar Middle Click)", // User agent string
        INTERNET_OPEN_TYPE_PRECONFIG, // Use system proxy settings
        nullptr, nullptr, 0);
    if (!hOpenHandle) {
        Wh_Log(L"InternetOpen failed: %lu", GetLastError());
        return std::nullopt;
    }

    // RAII wrapper for internet handles
    struct InternetHandleCloser {
        HINTERNET handle;
        ~InternetHandleCloser() { if (handle) InternetCloseHandle(handle); }
    };

    InternetHandleCloser openCloser{hOpenHandle};

    HINTERNET hUrlHandle = InternetOpenUrl(
        hOpenHandle, lpUrl, nullptr, 0,
        INTERNET_FLAG_NO_AUTH | INTERNET_FLAG_NO_CACHE_WRITE |
        INTERNET_FLAG_NO_COOKIES | INTERNET_FLAG_NO_UI |
        INTERNET_FLAG_PRAGMA_NOCACHE | INTERNET_FLAG_RELOAD |
        INTERNET_FLAG_SECURE, // Assume HTTPS
        0);

    if (!hUrlHandle) {
        Wh_Log(L"InternetOpenUrl failed for %s: %lu", lpUrl, GetLastError());
        return std::nullopt;
    }
    InternetHandleCloser urlCloser{hUrlHandle};

    DWORD dwStatusCode = 0;
    DWORD dwStatusCodeSize = sizeof(dwStatusCode);
    if (!HttpQueryInfo(hUrlHandle,
                       HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER,
                       &dwStatusCode, &dwStatusCodeSize, nullptr)) {
         Wh_Log(L"HttpQueryInfo (Status Code) failed: %lu", GetLastError());
         return std::nullopt;
    }

    Wh_Log(L"HTTP Status Code: %lu", dwStatusCode);
    if (dwStatusCode != 200) {
        return std::nullopt; // Request failed
    }

    std::vector<BYTE> buffer;
    constexpr DWORD READ_CHUNK_SIZE = 4096;
    BYTE readBuffer[READ_CHUNK_SIZE];
    DWORD dwBytesRead = 0;

    while (InternetReadFile(hUrlHandle, readBuffer, READ_CHUNK_SIZE, &dwBytesRead) && dwBytesRead > 0) {
        buffer.insert(buffer.end(), readBuffer, readBuffer + dwBytesRead);
    }

     if (GetLastError() != ERROR_SUCCESS && GetLastError() != ERROR_INTERNET_CONNECTION_ABORTED) {
         // Check if it failed for a reason other than finishing the read
         Wh_Log(L"InternetReadFile failed mid-stream: %lu", GetLastError());
         // Might still have partial data, but let's consider it a failure
         // return std::nullopt;
     }


    if (buffer.empty()) {
        Wh_Log(L"Downloaded content is empty.");
        return std::nullopt;
    }

    // Assume UTF-8 encoding for the downloaded cache file
    int charsNeeded = MultiByteToWideChar(CP_UTF8, 0, (PCSTR)buffer.data(), (int)buffer.size(), nullptr, 0);
    if (charsNeeded <= 0) {
        Wh_Log(L"MultiByteToWideChar (determining size) failed: %lu", GetLastError());
        return std::nullopt;
    }

    std::wstring unicodeContent(charsNeeded, L'\0');
    int charsWritten = MultiByteToWideChar(CP_UTF8, 0, (PCSTR)buffer.data(), (int)buffer.size(), unicodeContent.data(), (int)unicodeContent.size());

    if (charsWritten <= 0) {
         Wh_Log(L"MultiByteToWideChar (conversion) failed: %lu", GetLastError());
         return std::nullopt;
    }

    Wh_Log(L"Successfully downloaded and decoded %d characters.", charsWritten);
    return unicodeContent;
}

// Function to attempt hooking symbols, falling back to an online cache if local fails
bool HookSymbolsWithOnlineCacheFallback(HMODULE module,
                                        const SYMBOL_HOOK* symbolHooks,
                                        size_t symbolHooksCount) {
    // Use a generic mod ID for the cache URL to potentially share caches
    // between the original and forks if symbols are compatible.
    constexpr WCHAR kModIdForCache[] = L"taskbar-button-click";

    // First, try hooking using only the local cache.
    if (HookSymbols(module, symbolHooks, symbolHooksCount, /*cacheOnly=*/true)) {
        Wh_Log(L"HookSymbols succeeded using local cache.");
        return true;
    }

    Wh_Log(L"HookSymbols() from local cache failed or incomplete, trying online cache fallback...");

    // --- Prepare information for online cache URL ---
    WCHAR moduleFilePath[MAX_PATH];
    DWORD moduleFilePathLen = GetModuleFileName(module, moduleFilePath, ARRAYSIZE(moduleFilePath));
    if (!moduleFilePathLen || moduleFilePathLen >= ARRAYSIZE(moduleFilePath)) {
        Wh_Log(L"Online Cache: GetModuleFileName failed.");
        // Proceed to try live lookup without online cache
        return HookSymbols(module, symbolHooks, symbolHooksCount, /*cacheOnly=*/false);
    }

    PWSTR moduleFileName = wcsrchr(moduleFilePath, L'\\');
    if (!moduleFileName) {
        moduleFileName = moduleFilePath; // Use full path if no backslash
    } else {
        moduleFileName++; // Skip the backslash
    }

    // Normalize filename to lowercase for URL consistency
    // Be careful with buffer size if moduleFileName points into moduleFilePath
    DWORD moduleFileNameLen = (DWORD)wcslen(moduleFileName);
    LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_LOWERCASE, moduleFileName, moduleFileNameLen + 1, // Include null terminator for safety
                  moduleFileName, moduleFileNameLen + 1, nullptr, nullptr, 0);


    // Get module timestamp and size
    IMAGE_DOS_HEADER* dosHeader = (IMAGE_DOS_HEADER*)module;
     IMAGE_NT_HEADERS* header = nullptr;
     if (dosHeader && dosHeader->e_magic == IMAGE_DOS_SIGNATURE) {
        header = (IMAGE_NT_HEADERS*)((BYTE*)dosHeader + dosHeader->e_lfanew);
         if (!IsBadReadPtr(header, sizeof(IMAGE_NT_HEADERS)) && header->Signature == IMAGE_NT_SIGNATURE) {
             // Header looks valid
         } else {
             header = nullptr; // Invalid header
         }
    }
    if (!header) {
        Wh_Log(L"Online Cache: Invalid module header for %s", moduleFileName);
        return HookSymbols(module, symbolHooks, symbolHooksCount, /*cacheOnly=*/false);
    }

    auto timeStamp = std::to_wstring(header->FileHeader.TimeDateStamp);
    auto imageSize = std::to_wstring(header->OptionalHeader.SizeOfImage);

    // Construct the local cache key name (used in the URL path)
    std::wstring cacheStrKey =
#if defined(_M_IX86)
        L"symbol-x86-cache-";
#elif defined(_M_X64)
        L"symbol-cache-"; // Keep original name convention
#else
#error "Unsupported architecture"
#endif
    cacheStrKey += moduleFileName;

    // Construct the full URL for the online cache file
    std::wstring onlineCacheUrl = L"https://ramensoftware.github.io/windhawk-mod-symbol-cache/"; // Base URL
    onlineCacheUrl += kModIdForCache;
    onlineCacheUrl += L'/';
    onlineCacheUrl += cacheStrKey; // e.g., symbol-cache-explorer.exe
    onlineCacheUrl += L'/';
    onlineCacheUrl += timeStamp;
    onlineCacheUrl += L'-';
    onlineCacheUrl += imageSize;
    onlineCacheUrl += L".txt"; // File extension

    Wh_Log(L"Looking for online cache at: %s", onlineCacheUrl.c_str());

    // Attempt to download the cache content
    auto onlineCacheContent = GetUrlContent(onlineCacheUrl.c_str());

    if (onlineCacheContent) {
        Wh_Log(L"Online cache downloaded successfully. Saving locally.");
        // Save the downloaded content as the local cache
        Wh_SetStringValue(cacheStrKey.c_str(), onlineCacheContent->c_str());
    } else {
        Wh_Log(L"Failed to download or use online cache.");
        // If download fails, the local cache (if any) remains unchanged or empty.
    }

    // Finally, attempt hooking again. This time, it will use the potentially
    // updated local cache (if download succeeded) OR perform a live lookup
    // if the cache is still invalid or incomplete.
    Wh_Log(L"Attempting HookSymbols again (with potentially updated cache or live lookup)...");
    return HookSymbols(module, symbolHooks, symbolHooksCount, /*cacheOnly=*/false);
}
// --- End of Online Cache Fallback ---


// Function to hook symbols specifically for ExplorerPatcher's taskbar module
bool HookExplorerPatcherSymbols(HMODULE explorerPatcherModule) {
    // Prevent redundant initialization
    if (g_explorerPatcherInitialized.exchange(true)) {
        Wh_Log(L"ExplorerPatcher symbols already initialized.");
        return true;
    }

    Wh_Log(L"Initializing ExplorerPatcher symbol hooks...");

    // Define the symbols needed from ExplorerPatcher's DLL (ep_taskbar_*.dll)
    struct EXPLORER_PATCHER_HOOK {
        PCSTR symbol;             // Mangled symbol name
        void** pOriginalFunction; // Pointer to store original function address
        void* hookFunction = nullptr; // Hook function address (if any)
        bool optional = false;    // Is finding this symbol optional?
    };

    EXPLORER_PATCHER_HOOK hooks[] = {
        // Core click handler (Win10/Win11 style)
        { "?_HandleClick@CTaskListWnd@@IEAAXPEAUITaskBtnGroup@@HW4eCLICKACTION@1@HH@Z",
          (void**)&CTaskListWnd__HandleClick_Original, (void*)CTaskListWnd__HandleClick_Hook },
        // Launch action handler
        { "?Launch@CTaskBand@@UEAAJPEAUITaskGroup@@AEBUtagPOINT@@W4LaunchFromTaskbarOptions@@@Z",
          (void**)&CTaskBand_Launch_Original, (void*)CTaskBand_Launch_Hook },
        // Get active button (for close foreground)
        { "?GetActiveBtn@CTaskListWnd@@UEAAJPEAPEAUITaskGroup@@PEAH@Z",
          (void**)&CTaskListWnd_GetActiveBtn_Original },
        // Close window handler
        { "?ProcessJumpViewCloseWindow@CTaskListWnd@@UEAAXPEAUHWND__@@PEAUITaskGroup@@PEAUHMONITOR__@@@Z",
          (void**)&CTaskListWnd_ProcessJumpViewCloseWindow_Original },
        // End task handler
        { "?_EndTask@CTaskBand@@IEAAXQEAUHWND__@@H@Z", // Common pattern
          (void**)&CTaskBand__EndTask_Original },
        // Get group type
        { "?GetGroupType@CTaskBtnGroup@@UEAA?AW4eTBGROUPTYPE@@XZ",
          (void**)&CTaskBtnGroup_GetGroupType_Original },
        // Get group interface
        { "?GetGroup@CTaskBtnGroup@@UEAAPEAUITaskGroup@@XZ",
          (void**)&CTaskBtnGroup_GetGroup_Original },
        // Get task item from group
        { "?GetTaskItem@CTaskBtnGroup@@UEAAPEAUITaskItem@@H@Z",
          (void**)&CTaskBtnGroup_GetTaskItem_Original },
        // Get window handle (standard window)
        { "?GetWindow@CWindowTaskItem@@UEAAPEAUHWND__@@XZ",
          (void**)&CWindowTaskItem_GetWindow_Original },
        // Get window handle (UWP/Immersive window)
        { "?GetWindow@CImmersiveTaskItem@@UEAAPEAUHWND__@@XZ",
          (void**)&CImmersiveTaskItem_GetWindow_Original },
        // Get vftable for CImmersiveTaskItem (optional, used for type check)
        { "??_7CImmersiveTaskItem@@6B@ITaskItem@@@", // Example mangled vftable name
          (void**)&CImmersiveTaskItem_vftable, nullptr, /*optional=*/true },
    };

    bool succeeded = true;
    Wh_Log(L"Attempting to resolve %d symbols from ExplorerPatcher module...", (int)ARRAYSIZE(hooks));


    for (const auto& hook : hooks) {
        void* ptr = (void*)GetProcAddress(explorerPatcherModule, hook.symbol);
        if (!ptr) {
            Wh_Log(L"EP Symbol %s: %S", hook.optional ? L"(Optional) NOT FOUND" : L"NOT FOUND", hook.symbol);
            if (!hook.optional) {
                succeeded = false; // Mandatory symbol missing
            }
            // Clear the target pointer if it wasn't found
             if (hook.pOriginalFunction) *hook.pOriginalFunction = nullptr;
        } else {
            Wh_Log(L"EP Symbol FOUND: %S at %p", hook.symbol, ptr);
            if (hook.hookFunction) {
                // Apply the hook immediately if a hook function is provided
                Wh_SetFunctionHook(ptr, hook.hookFunction, hook.pOriginalFunction);
                 Wh_Log(L"  -> Hooked with %p", hook.hookFunction);
            } else if (hook.pOriginalFunction) {
                // Store the found address if no hook function is needed (e.g., for calling original or getting vftable)
                *hook.pOriginalFunction = ptr;
                 Wh_Log(L"  -> Stored address in pOriginalFunction");
            }
        }
    }

    // Apply any hooks set via Wh_SetFunctionHook if the main mod is already initialized
    if (g_initialized) {
        Wh_Log(L"Applying ExplorerPatcher hook operations.");
        Wh_ApplyHookOperations();
    } else {
         Wh_Log(L"Main mod not yet initialized, EP hooks will be applied later.");
    }

    if (!succeeded) {
         Wh_Log(L"Failed to resolve one or more mandatory ExplorerPatcher symbols.");
    } else {
         Wh_Log(L"ExplorerPatcher symbol resolution complete.");
    }

    return succeeded;
}

// Callback function for module loading, used to detect ExplorerPatcher
bool HandleModuleIfExplorerPatcher(HMODULE module) {
    WCHAR moduleFilePath[MAX_PATH];
    DWORD pathLen = GetModuleFileName(module, moduleFilePath, ARRAYSIZE(moduleFilePath));

    // Basic validation of the path
    if (pathLen == 0 || pathLen >= ARRAYSIZE(moduleFilePath)) {
        return true; // Continue processing other modules
    }

    PCWSTR moduleFileName = wcsrchr(moduleFilePath, L'\\');
    if (!moduleFileName) {
        moduleFileName = moduleFilePath; // Use full path if no backslash
    } else {
        moduleFileName++; // Skip the backslash
    }

    // Check if the filename starts with "ep_taskbar." (case-insensitive)
    if (_wcsnicmp(L"ep_taskbar.", moduleFileName, wcslen(L"ep_taskbar.")) == 0) {
        Wh_Log(L"ExplorerPatcher taskbar module detected: %s", moduleFileName);
        // Attempt to hook the symbols within this module
        return HookExplorerPatcherSymbols(module);
    }

    // Not the module we're looking for
    return true; // Continue processing other modules
}

// Function to iterate through already loaded modules and check for ExplorerPatcher
void HandleLoadedExplorerPatcher() {
    Wh_Log(L"Checking already loaded modules for ExplorerPatcher...");
    std::vector<HMODULE> hMods(1024); // Preallocate buffer for module handles
    DWORD cbNeeded;
    // EnumProcessModulesEx might be better for filtering, but this is simpler
    if (EnumProcessModules(GetCurrentProcess(), hMods.data(), (DWORD)(hMods.size() * sizeof(HMODULE)), &cbNeeded)) {
        size_t numModules = cbNeeded / sizeof(HMODULE);
        if (numModules > hMods.size()) {
             Wh_Log(L"Warning: Module list truncated (%zu modules found, buffer size %zu).", numModules, hMods.size());
             numModules = hMods.size(); // Process only the modules we got handles for
        }


        Wh_Log(L"Enumerated %zu loaded modules.", numModules);
        for (size_t i = 0; i < numModules; i++) {
            // If EP is found and hooked, HandleModuleIfExplorerPatcher returns,
            // and g_explorerPatcherInitialized will be set, stopping further checks.
            if (!HandleModuleIfExplorerPatcher(hMods[i])) {
                 Wh_Log(L"Error occurred while handling module %p.", hMods[i]);
                 // Decide if we should stop or continue checking other modules
                 // break; // Example: Stop on first error
            }
            if (g_explorerPatcherInitialized) {
                 Wh_Log(L"ExplorerPatcher handled. Stopping module scan.");
                 break; // Stop scanning once EP is found and handled
            }
        }
    } else {
         Wh_Log(L"EnumProcessModules failed: %lu", GetLastError());
    }
     if (!g_explorerPatcherInitialized) {
         Wh_Log(L"ExplorerPatcher module not found among loaded modules.");
     }
}

// Hook for LoadLibraryExW to detect ExplorerPatcher loading *after* the mod starts
using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    // Call the original function first
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);

    // Check if a module was loaded successfully and if we haven't already handled EP
    // The check `!((ULONG_PTR)module & 3)` is a heuristic to filter out resource-only loads.
    if (module && !((ULONG_PTR)module & 3) && !g_explorerPatcherInitialized) {
        // Check if the newly loaded module is the EP taskbar DLL
        HandleModuleIfExplorerPatcher(module);
    }

    return module;
}

// Function to hook the primary taskbar symbols (either in explorer.exe or taskbar.dll)
bool HookTaskbarSymbols() {
    Wh_Log(L"Initializing primary taskbar symbol hooks...");

    // Define the symbols needed from the main taskbar implementation
    SYMBOL_HOOK symbolHooks[] = {
        // --- Win11 Only Symbols ---
        // CTaskListWnd::HandleClick (used for context)
        { { LR"(public: virtual long __cdecl CTaskListWnd::HandleClick(struct ITaskGroup *,struct ITaskItem *,struct winrt::Windows::System::LauncherOptions const &))",
            LR"(public: virtual long __cdecl CTaskListWnd::HandleClick(struct ITaskGroup * __ptr64,struct ITaskItem * __ptr64,struct winrt::Windows::System::LauncherOptions const & __ptr64) __ptr64)" }, // 64-bit variant
          (void**)&CTaskListWnd_HandleClick_Original, (void*)CTaskListWnd_HandleClick_Hook, /*optional=*/ (g_winVersion <= WinVersion::Win10) }, // Optional on Win10

        // --- Win10 and Win11 Symbols ---
        // CTaskListWnd::_HandleClick (main click handler)
        { { LR"(protected: void __cdecl CTaskListWnd::_HandleClick(struct ITaskBtnGroup *,int,enum CTaskListWnd::eCLICKACTION,int,int))",
            LR"(protected: void __cdecl CTaskListWnd::_HandleClick(struct ITaskBtnGroup * __ptr64,int,enum CTaskListWnd::eCLICKACTION,int,int) __ptr64)" }, // 64-bit variant
          (void**)&CTaskListWnd__HandleClick_Original, (void*)CTaskListWnd__HandleClick_Hook },
        // CTaskBand::Launch (where middle click action is decided)
        { { LR"(public: virtual long __cdecl CTaskBand::Launch(struct ITaskGroup *,struct tagPOINT const &,enum LaunchFromTaskbarOptions))",
            LR"(public: virtual long __cdecl CTaskBand::Launch(struct ITaskGroup * __ptr64,struct tagPOINT const & __ptr64,enum LaunchFromTaskbarOptions) __ptr64)" }, // 64-bit variant
          (void**)&CTaskBand_Launch_Original, (void*)CTaskBand_Launch_Hook },
        // CTaskListWnd::GetActiveBtn (for close foreground)
        { { LR"(public: virtual long __cdecl CTaskListWnd::GetActiveBtn(struct ITaskGroup * *,int *))",
            LR"(public: virtual long __cdecl CTaskListWnd::GetActiveBtn(struct ITaskGroup * __ptr64 * __ptr64,int * __ptr64) __ptr64)" }, // 64-bit variant
          (void**)&CTaskListWnd_GetActiveBtn_Original },
        // CTaskListWnd::ProcessJumpViewCloseWindow (to close windows)
        { { LR"(public: virtual void __cdecl CTaskListWnd::ProcessJumpViewCloseWindow(struct HWND__ *,struct ITaskGroup *,struct HMONITOR__ *))",
            LR"(public: virtual void __cdecl CTaskListWnd::ProcessJumpViewCloseWindow(struct HWND__ * __ptr64,struct ITaskGroup * __ptr64,struct HMONITOR__ * __ptr64) __ptr64)" }, // 64-bit variant
          (void**)&CTaskListWnd_ProcessJumpViewCloseWindow_Original },
        // CTaskBand::_EndTask (to end tasks) - Names vary!
        { { // Win11 / newer __cdecl:
            LR"(protected: void __cdecl CTaskBand::_EndTask(struct HWND__ * const,int))",
            LR"(protected: void __cdecl CTaskBand::_EndTask(struct HWND__ * __ptr64 const,int) __ptr64)",
            // Win10 / older __thiscall:
            LR"(protected: void __thiscall CTaskBand::_EndTask(struct HWND__ * const,int))",
            LR"(protected: void __thiscall CTaskBand::_EndTask(struct HWND__ * __ptr64 const,int) __ptr64)" },
          (void**)&CTaskBand__EndTask_Original },
        // CTaskBtnGroup::GetGroupType
        { { LR"(public: virtual enum eTBGROUPTYPE __cdecl CTaskBtnGroup::GetGroupType(void))",
            LR"(public: virtual enum eTBGROUPTYPE __cdecl CTaskBtnGroup::GetGroupType(void) __ptr64)" }, // 64-bit variant
          (void**)&CTaskBtnGroup_GetGroupType_Original },
        // CTaskBtnGroup::GetGroup
        { { LR"(public: virtual struct ITaskGroup * __cdecl CTaskBtnGroup::GetGroup(void))",
            LR"(public: virtual struct ITaskGroup * __ptr64 __cdecl CTaskBtnGroup::GetGroup(void) __ptr64)" }, // 64-bit variant
          (void**)&CTaskBtnGroup_GetGroup_Original },
        // CTaskBtnGroup::GetTaskItem
        { { LR"(public: virtual struct ITaskItem * __cdecl CTaskBtnGroup::GetTaskItem(int))",
            LR"(public: virtual struct ITaskItem * __ptr64 __cdecl CTaskBtnGroup::GetTaskItem(int) __ptr64)" }, // 64-bit variant
          (void**)&CTaskBtnGroup_GetTaskItem_Original },
        // CWindowTaskItem::GetWindow
        { { LR"(public: virtual struct HWND__ * __cdecl CWindowTaskItem::GetWindow(void))",
            LR"(public: virtual struct HWND__ * __ptr64 __cdecl CWindowTaskItem::GetWindow(void) __ptr64)" }, // 64-bit variant
          (void**)&CWindowTaskItem_GetWindow_Original },
        // CImmersiveTaskItem::GetWindow
        { { LR"(public: virtual struct HWND__ * __cdecl CImmersiveTaskItem::GetWindow(void))",
            LR"(public: virtual struct HWND__ * __ptr64 __cdecl CImmersiveTaskItem::GetWindow(void) __ptr64)" }, // 64-bit variant
          (void**)&CImmersiveTaskItem_GetWindow_Original },
        // CImmersiveTaskItem vftable (optional, for type check)
        { { LR"(const CImmersiveTaskItem::`vftable'{for `ITaskItem'})", // Example name
            LR"(??_7CImmersiveTaskItem@@6B@ITaskItem@@@)", // Another possible mangled name
            LR"(const CImmersiveTaskItem::`vftable'{for `ITaskItem'} __ptr64)" }, // 64-bit variant
          (void**)&CImmersiveTaskItem_vftable, nullptr, /*optional=*/true },
    };

    // Determine the target module based on Windows version
    HMODULE targetModule = nullptr;
    if (g_winVersion <= WinVersion::Win10) {
        // On Win10 (or when EP Win10 taskbar is used), symbols are in explorer.exe
        targetModule = GetModuleHandle(nullptr); // Get handle to current process (explorer.exe)
        Wh_Log(L"Targeting explorer.exe for symbols (Win10/EP mode).");
         return HookSymbolsWithOnlineCacheFallback(targetModule, symbolHooks, ARRAYSIZE(symbolHooks));

    } else {
        // On Win11 (native taskbar), symbols are primarily in taskbar.dll
        targetModule = LoadLibrary(L"taskbar.dll");
        if (!targetModule) {
            Wh_Log(L"Failed to load taskbar.dll: %lu", GetLastError());
            return false;
        }
        Wh_Log(L"Targeting taskbar.dll for symbols (Win11 native mode).");
        return HookSymbolsWithOnlineCacheFallback(targetModule, symbolHooks, ARRAYSIZE(symbolHooks));
        // Note: We don't unload taskbar.dll as explorer likely keeps it loaded anyway.
    }
}

// Function to load settings from Windhawk
void LoadSettings() {
    Wh_Log(L"Loading settings...");

    // Multiple Items Behavior
    PCWSTR multipleItemsBehaviorStr = Wh_GetStringSetting(L"multipleItemsBehavior");
    if (multipleItemsBehaviorStr) {
        if (wcscmp(multipleItemsBehaviorStr, L"closeForeground") == 0) {
            g_settings.multipleItemsBehavior = MULTIPLE_ITEMS_BEHAVIOR_CLOSE_FOREGROUND;
            Wh_Log(L" Setting: multipleItemsBehavior = closeForeground");
        } else if (wcscmp(multipleItemsBehaviorStr, L"none") == 0) {
            g_settings.multipleItemsBehavior = MULTIPLE_ITEMS_BEHAVIOR_NONE;
             Wh_Log(L" Setting: multipleItemsBehavior = none");
        } else { // Default to closeAll
            g_settings.multipleItemsBehavior = MULTIPLE_ITEMS_BEHAVIOR_CLOSE_ALL;
             Wh_Log(L" Setting: multipleItemsBehavior = closeAll (default)");
        }
        Wh_FreeStringSetting(multipleItemsBehaviorStr);
    } else {
         g_settings.multipleItemsBehavior = MULTIPLE_ITEMS_BEHAVIOR_CLOSE_ALL; // Default
         Wh_Log(L" Setting: multipleItemsBehavior = closeAll (default, failed to read)");
    }


    // Keys to End Task
    g_settings.keysToEndTaskCtrl = Wh_GetIntSetting(L"keysToEndTask.Ctrl");
    g_settings.keysToEndTaskAlt = Wh_GetIntSetting(L"keysToEndTask.Alt");
    Wh_Log(L" Setting: keysToEndTask.Ctrl = %d", g_settings.keysToEndTaskCtrl);
    Wh_Log(L" Setting: keysToEndTask.Alt = %d", g_settings.keysToEndTaskAlt);

    // Keys to Disable Override (New Setting)
    g_settings.keysToDisableOverrideCtrl = Wh_GetIntSetting(L"keysToDisableOverride.Ctrl");
    g_settings.keysToDisableOverrideAlt = Wh_GetIntSetting(L"keysToDisableOverride.Alt");
    Wh_Log(L" Setting: keysToDisableOverride.Ctrl = %d", g_settings.keysToDisableOverrideCtrl);
    Wh_Log(L" Setting: keysToDisableOverride.Alt = %d", g_settings.keysToDisableOverrideAlt);


    // Old Taskbar on Win11 (ExplorerPatcher mode)
    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
     Wh_Log(L" Setting: oldTaskbarOnWin11 = %d", g_settings.oldTaskbarOnWin11);

     // Add a log warning if there's a configuration conflict
     if ((g_settings.keysToEndTaskCtrl && g_settings.keysToDisableOverrideCtrl) ||
         (g_settings.keysToEndTaskAlt && g_settings.keysToDisableOverrideAlt)) {
         Wh_Log(L"Warning: Same modifier key configured for 'End Task' and 'Disable Override'. 'Disable Override' will take priority.");
     }


     Wh_Log(L"Settings loaded.");
}

// Main initialization function called by Windhawk
BOOL Wh_ModInit() {
    Wh_Log(L"> Wh_ModInit");

    // Load settings first
    LoadSettings();

    // Determine Windows version
    g_winVersion = GetExplorerVersion();
    if (g_winVersion == WinVersion::Unsupported) {
        Wh_Log(L"Initialization failed: Unsupported Windows version.");
        return FALSE; // Stop initialization
    }
     Wh_Log(L"Detected Windows version class: %d", (int)g_winVersion);

    // Handle the "Old Taskbar on Win11" setting (ExplorerPatcher mode)
    if (g_settings.oldTaskbarOnWin11 && g_winVersion >= WinVersion::Win11) {
        Wh_Log(L"ExplorerPatcher mode enabled. Adjusting hooks for Win10-style taskbar.");
        // Treat as Win10 for symbol hooking purposes
        WinVersion originalVersion = g_winVersion; // Store original for potential future use
        g_winVersion = WinVersion::Win10; // Override detected version for hooking logic

        // Hook primary taskbar symbols (expecting them in explorer.exe now)
        if (!HookTaskbarSymbols()) {
             Wh_Log(L"Initialization failed: Failed to hook primary symbols in ExplorerPatcher mode.");
             g_winVersion = originalVersion; // Restore original version before failing
            return FALSE;
        }

        // Check if ExplorerPatcher's DLL is already loaded
        HandleLoadedExplorerPatcher();

        // Hook LoadLibraryExW to catch EP loading later
        HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
        if (kernelBaseModule) {
            FARPROC pKernelBaseLoadLibraryExW = GetProcAddress(kernelBaseModule, "LoadLibraryExW");
            if (pKernelBaseLoadLibraryExW) {
                 Wh_Log(L"Hooking kernelbase!LoadLibraryExW to detect late EP loading.");
                Wh_SetFunctionHook((void*)pKernelBaseLoadLibraryExW,
                                   (void*)LoadLibraryExW_Hook,
                                   (void**)&LoadLibraryExW_Original);
            } else {
                 Wh_Log(L"Warning: Could not get address of kernelbase!LoadLibraryExW.");
            }
        } else {
             Wh_Log(L"Warning: Could not get handle to kernelbase.dll.");
        }

    } else {
        // Standard mode (Native Win10 or Win11 taskbar)
         Wh_Log(L"Standard mode enabled (Native Win10/Win11 taskbar).");
        if (!HookTaskbarSymbols()) {
             Wh_Log(L"Initialization failed: Failed to hook primary taskbar symbols.");
            return FALSE;
        }
    }

    // Apply all queued hook operations
    Wh_ApplyHookOperations();

    // Mark initialization as complete
    g_initialized = true;
    Wh_Log(L"Initialization successful.");
    return TRUE;
}

// Called by Windhawk after initialization (optional)
void Wh_ModAfterInit() {
    Wh_Log(L"> Wh_ModAfterInit");

    // In ExplorerPatcher mode, double-check if EP loaded between Wh_ModInit
    // and now, in case the LoadLibrary hook missed it or wasn't set up in time.
    if (g_settings.oldTaskbarOnWin11 && !g_explorerPatcherInitialized) {
        Wh_Log(L"Re-checking for loaded ExplorerPatcher module in Wh_ModAfterInit...");
        HandleLoadedExplorerPatcher();
    }
     Wh_Log(L"Wh_ModAfterInit complete.");
}

// Called by Windhawk when settings are changed in the UI
BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L"> Wh_ModSettingsChanged");

    // Store the previous state of the setting that requires a reload
    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;

    // Reload all settings from Windhawk
    LoadSettings();

    // Determine if a full mod reload is necessary.
    // A reload is needed if the ExplorerPatcher mode setting changed,
    // as this affects which module we hook and which symbols are needed.
    *bReload = (g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11);

    Wh_Log(L"Settings changed. Reload required: %s", *bReload ? L"Yes" : L"No");

    // Indicate success
    return TRUE;
}

// Compatibility function for Windhawk versions older than 1.3
void Wh_ModSettingsChanged() {
    Wh_Log(L"> Wh_ModSettingsChanged (Compatibility)");
    BOOL bReload = FALSE; // Assume no reload needed by default for old versions
    Wh_ModSettingsChanged(&bReload);
}

// Optional: Called by Windhawk before the mod is unloaded
void Wh_ModUninit() {
     Wh_Log(L"> Wh_ModUninit");
     // Hooks are automatically removed by Windhawk.
     g_initialized = false;
     g_explorerPatcherInitialized = false; // Reset EP state on unload
     Wh_Log(L"Uninitialization complete.");
}
