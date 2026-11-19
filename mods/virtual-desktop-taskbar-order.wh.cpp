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

Only Windows 11 is currently supported. For older Windows versions check out [7+
Taskbar Tweaker](https://tweaker.ramensoftware.com/).

![Demonstration](https://i.imgur.com/ie8Q9cl.gif)
*/
// ==/WindhawkModReadme==

#include <windhawk_utils.h>

#include <algorithm>
#include <climits>
#include <map>
#include <utility>
#include <vector>

#include <commctrl.h>

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
};

WinVersion g_winVersion;
DWORD g_taskbarThreadId;

void* CTaskBtnGroup_ITaskBtnGroup_vftable;

using CTaskBtnGroup_GetGroup_t = void*(WINAPI*)(void* pThis);
CTaskBtnGroup_GetGroup_t CTaskBtnGroup_GetGroup;

bool IsButtonGroupOf(void* p, void* taskGroup) {
    return taskGroup && p &&
           *(void**)p == CTaskBtnGroup_ITaskBtnGroup_vftable &&
           CTaskBtnGroup_GetGroup(p) == taskGroup;
}

// The task groups of each taskbar's buttons, in order, per virtual desktop. A
// taskbar is identified by the DPA of its button groups, and a desktop by its
// IVirtualDesktop pointer. The keys and the task groups are only compared.
std::map<std::pair<HDPA, void*>, std::vector<void*>> g_desktopArrangements;

// The desktop which the taskbar shows, null until the first switch since the
// mod was loaded.
void* g_currentDesktop;

struct DesktopSwitch {
    void* oldDesktop;
    void* newDesktop;
    std::vector<HDPA> savedDpas;
};

// Set while the taskbar updates its buttons for a virtual desktop switch.
thread_local DesktopSwitch* g_desktopSwitch;

std::vector<void*> GetButtonGroupsOrder(HDPA hdpa) {
    std::vector<void*> taskGroups;
    for (int i = 0; i < DPA_GetPtrCount(hdpa); i++) {
        taskGroups.push_back(CTaskBtnGroup_GetGroup(DPA_GetPtr(hdpa, i)));
    }

    return taskGroups;
}

// Saves the arrangement of the old desktop the first time the switch changes
// the taskbar's buttons.
void SaveArrangementBeforeSwitchChange(HDPA hdpa) {
    std::vector<HDPA>& savedDpas = g_desktopSwitch->savedDpas;
    if (std::ranges::find(savedDpas, hdpa) != savedDpas.end()) {
        return;
    }

    savedDpas.push_back(hdpa);

    g_desktopArrangements[{hdpa, g_desktopSwitch->oldDesktop}] =
        GetButtonGroupsOrder(hdpa);
}

// Returns the index at which the task group's new button group gets the
// position it had in the arrangement of the new desktop, or INT_MAX to append
// it.
int GetDesktopArrangementIndex(HDPA hdpa, void* taskGroup) {
    auto it = g_desktopArrangements.find({hdpa, g_desktopSwitch->newDesktop});
    if (it == g_desktopArrangements.end()) {
        return INT_MAX;
    }

    const std::vector<void*>& arrangement = it->second;
    auto position = std::ranges::find(arrangement, taskGroup);
    if (position == arrangement.end()) {
        return INT_MAX;
    }

    // Insert before the first button group which follows it in the
    // arrangement.
    for (int i = 0; i < DPA_GetPtrCount(hdpa); i++) {
        void* buttonTaskGroup = CTaskBtnGroup_GetGroup(DPA_GetPtr(hdpa, i));
        if (std::find(position + 1, arrangement.end(), buttonTaskGroup) !=
            arrangement.end()) {
            return i;
        }
    }

    return INT_MAX;
}

template <typename Function>
HRESULT HandleDesktopSwitch(void* oldDesktop,
                            void* newDesktop,
                            Function original) {
    if (GetCurrentThreadId() != g_taskbarThreadId || g_desktopSwitch) {
        return original();
    }

    Wh_Log(L"> oldDesktop=%p, newDesktop=%p", oldDesktop, newDesktop);

    DesktopSwitch desktopSwitch{oldDesktop, newDesktop, {}};
    g_desktopSwitch = &desktopSwitch;
    HRESULT hr = original();
    g_desktopSwitch = nullptr;

    if (!g_currentDesktop) {
        // Drags before the first switch saved arrangements without the
        // desktop. Keep them unless the switch saved newer ones.
        std::vector<HDPA> pendingDpas;
        for (const auto& [key, arrangement] : g_desktopArrangements) {
            if (!key.second) {
                pendingDpas.push_back(key.first);
            }
        }

        for (HDPA hdpa : pendingDpas) {
            auto node = g_desktopArrangements.extract({hdpa, nullptr});
            node.key().second = oldDesktop;
            g_desktopArrangements.insert(std::move(node));
        }
    }

    g_currentDesktop = newDesktop;

    return hr;
}

using CTaskBand_CurrentVirtualDesktopChanged_t =
    HRESULT(WINAPI*)(void* pThis,
                     void* oldDesktop,   // IVirtualDesktop*
                     void* newDesktop);  // IVirtualDesktop*
CTaskBand_CurrentVirtualDesktopChanged_t
    CTaskBand_CurrentVirtualDesktopChanged_Original;
HRESULT WINAPI CTaskBand_CurrentVirtualDesktopChanged_Hook(void* pThis,
                                                           void* oldDesktop,
                                                           void* newDesktop) {
    return HandleDesktopSwitch(oldDesktop, newDesktop, [&] {
        return CTaskBand_CurrentVirtualDesktopChanged_Original(
            pThis, oldDesktop, newDesktop);
    });
}

using CTaskBand_CurrentVirtualDesktopChangedPre24H2_t =
    HRESULT(WINAPI*)(void* pThis,
                     void* objectArray,  // IObjectArray*
                     void* oldDesktop,   // IVirtualDesktop*
                     void* newDesktop);  // IVirtualDesktop*
CTaskBand_CurrentVirtualDesktopChangedPre24H2_t
    CTaskBand_CurrentVirtualDesktopChangedPre24H2_Original;
HRESULT WINAPI
CTaskBand_CurrentVirtualDesktopChangedPre24H2_Hook(void* pThis,
                                                   void* objectArray,
                                                   void* oldDesktop,
                                                   void* newDesktop) {
    return HandleDesktopSwitch(oldDesktop, newDesktop, [&] {
        return CTaskBand_CurrentVirtualDesktopChangedPre24H2_Original(
            pThis, objectArray, oldDesktop, newDesktop);
    });
}

CTaskBand_CurrentVirtualDesktopChanged_t
    CTaskBand_CurrentVirtualDesktopChangedAnimated_Original;
HRESULT WINAPI
CTaskBand_CurrentVirtualDesktopChangedAnimated_Hook(void* pThis,
                                                    void* oldDesktop,
                                                    void* newDesktop) {
    return HandleDesktopSwitch(oldDesktop, newDesktop, [&] {
        return CTaskBand_CurrentVirtualDesktopChangedAnimated_Original(
            pThis, oldDesktop, newDesktop);
    });
}

thread_local void* g_createTBGroup_taskGroup;

using CTaskListWnd__CreateTBGroup_t =
    void*(WINAPI*)(void* pThis,
                   void* taskGroup,  // ITaskGroup*
                   int index);
CTaskListWnd__CreateTBGroup_t CTaskListWnd__CreateTBGroup_Original;
void* WINAPI CTaskListWnd__CreateTBGroup_Hook(void* pThis,
                                              void* taskGroup,
                                              int index) {
    g_createTBGroup_taskGroup = taskGroup;

    void* ret = CTaskListWnd__CreateTBGroup_Original(pThis, taskGroup, index);

    g_createTBGroup_taskGroup = nullptr;

    return ret;
}

thread_local void* g_taskDestroyedFinal_buttonGroup;

using CTaskListWnd__TaskDestroyedFinal_t =
    void(WINAPI*)(void* pThis,
                  void* buttonGroup,  // ITaskBtnGroup*
                  int param);
CTaskListWnd__TaskDestroyedFinal_t CTaskListWnd__TaskDestroyedFinal_Original;
void WINAPI CTaskListWnd__TaskDestroyedFinal_Hook(void* pThis,
                                                  void* buttonGroup,
                                                  int param) {
    g_taskDestroyedFinal_buttonGroup = buttonGroup;

    CTaskListWnd__TaskDestroyedFinal_Original(pThis, buttonGroup, param);

    g_taskDestroyedFinal_buttonGroup = nullptr;
}

thread_local void* g_tryMoveGroup_taskGroup;

using CTaskListWnd_TryMoveGroup_t =
    bool(WINAPI*)(void* pThis,
                  void* taskGroup,  // ITaskGroup*
                  DWORD index);
CTaskListWnd_TryMoveGroup_t CTaskListWnd_TryMoveGroup_Original;
bool WINAPI CTaskListWnd_TryMoveGroup_Hook(void* pThis,
                                           void* taskGroup,
                                           DWORD index) {
    g_tryMoveGroup_taskGroup = taskGroup;

    bool ret = CTaskListWnd_TryMoveGroup_Original(pThis, taskGroup, index);

    g_tryMoveGroup_taskGroup = nullptr;

    return ret;
}

using DPA_InsertPtr_t = decltype(&DPA_InsertPtr);
DPA_InsertPtr_t DPA_InsertPtr_Original;
int WINAPI DPA_InsertPtr_Hook(HDPA hdpa, int i, void* p) {
    if (g_desktopSwitch && IsButtonGroupOf(p, g_createTBGroup_taskGroup)) {
        SaveArrangementBeforeSwitchChange(hdpa);
        if (i == INT_MAX) {
            i = GetDesktopArrangementIndex(hdpa, g_createTBGroup_taskGroup);
            Wh_Log(L"Inserting button group %p at %d", p, i);
        }
    }

    int ret = DPA_InsertPtr_Original(hdpa, i, p);

    if (IsButtonGroupOf(p, g_tryMoveGroup_taskGroup) &&
        GetCurrentThreadId() == g_taskbarThreadId) {
        Wh_Log(L"Button group %p moved to %d", p, ret);
        g_desktopArrangements[{hdpa, g_currentDesktop}] =
            GetButtonGroupsOrder(hdpa);
    }

    return ret;
}

using DPA_DeletePtr_t = decltype(&DPA_DeletePtr);
DPA_DeletePtr_t DPA_DeletePtr_Original;
void* WINAPI DPA_DeletePtr_Hook(HDPA hdpa, int i) {
    if (g_desktopSwitch && g_taskDestroyedFinal_buttonGroup &&
        DPA_GetPtr(hdpa, i) == g_taskDestroyedFinal_buttonGroup) {
        SaveArrangementBeforeSwitchChange(hdpa);
    }

    return DPA_DeletePtr_Original(hdpa, i);
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

BOOL Wh_ModInit() {
    Wh_Log(L">");

    g_winVersion = GetExplorerVersion();
    if (g_winVersion < WinVersion::Win11) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
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
            {LR"(public: virtual bool __cdecl CTaskListWnd::TryMoveGroup(struct ITaskGroup *,unsigned int))"},
            &CTaskListWnd_TryMoveGroup_Original,
            CTaskListWnd_TryMoveGroup_Hook,
        },
        {
            {LR"(protected: struct ITaskBtnGroup * __cdecl CTaskListWnd::_CreateTBGroup(struct ITaskGroup *,int))"},
            &CTaskListWnd__CreateTBGroup_Original,
            CTaskListWnd__CreateTBGroup_Hook,
        },
        {
            {LR"(protected: void __cdecl CTaskListWnd::_TaskDestroyedFinal(struct ITaskBtnGroup *,int))"},
            &CTaskListWnd__TaskDestroyedFinal_Original,
            CTaskListWnd__TaskDestroyedFinal_Hook,
        },
        {
            // Windows 11 version 24H2 and newer.
            {LR"(public: virtual long __cdecl CTaskBand::CurrentVirtualDesktopChanged(struct IVirtualDesktop *,struct IVirtualDesktop *))"},
            &CTaskBand_CurrentVirtualDesktopChanged_Original,
            CTaskBand_CurrentVirtualDesktopChanged_Hook,
            true,
        },
        {
            // Before Windows 11 version 24H2.
            {LR"(public: virtual long __cdecl CTaskBand::CurrentVirtualDesktopChanged(struct IObjectArray *,struct IVirtualDesktop *,struct IVirtualDesktop *))"},
            &CTaskBand_CurrentVirtualDesktopChangedPre24H2_Original,
            CTaskBand_CurrentVirtualDesktopChangedPre24H2_Hook,
            true,
        },
        {
            {LR"(public: virtual long __cdecl CTaskBand::CurrentVirtualDesktopChangedAnimated(struct IVirtualDesktop *,struct IVirtualDesktop *))"},
            &CTaskBand_CurrentVirtualDesktopChangedAnimated_Original,
            CTaskBand_CurrentVirtualDesktopChangedAnimated_Hook,
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

    WindhawkUtils::SetFunctionHook(DPA_InsertPtr, DPA_InsertPtr_Hook,
                                   &DPA_InsertPtr_Original);

    WindhawkUtils::SetFunctionHook(DPA_DeletePtr, DPA_DeletePtr_Hook,
                                   &DPA_DeletePtr_Original);

    WindhawkUtils::SetFunctionHook(CreateWindowExW, CreateWindowExW_Hook,
                                   &CreateWindowExW_Original);

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (hTaskbarWnd) {
        InitializeTaskbarVariables(hTaskbarWnd);
    }
}
