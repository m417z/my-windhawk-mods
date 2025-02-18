// ==WindhawkMod==
// @id              pinned-items-double-click
// @name            Open pinned items with double click
// @description     Only open pinned items when double clicking on them to avoid accidental clicks
// @version         1.0.1
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -lversion
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
# Open pinned items with double click

Only open pinned items when double clicking on them to avoid accidental clicks.

![Demonstration](https://i.imgur.com/Si3siPm.gif)

Only Windows 10 64-bit and Windows 11 are supported. For older Windows versions
check out [7+ Taskbar Tweaker](https://tweaker.ramensoftware.com/).

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.
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

#include <windowsx.h>

struct {
    bool oldTaskbarOnWin11;
} g_settings;

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
};

WinVersion g_winVersion;

bool IsDoubleClickDistance(DWORD pos1, DWORD pos2) {
    return abs(GET_X_LPARAM(pos1) - GET_X_LPARAM(pos2)) <=
               GetSystemMetrics(SM_CXDOUBLECLK) &&
           abs(GET_Y_LPARAM(pos1) - GET_Y_LPARAM(pos2)) <=
               GetSystemMetrics(SM_CYDOUBLECLK);
}

using CTaskBtnGroup_GetGroupType_t = int(WINAPI*)(PVOID pThis);
CTaskBtnGroup_GetGroupType_t CTaskBtnGroup_GetGroupType_Original;

using CTaskListWnd__HandleClick_t = void(WINAPI*)(PVOID pThis,
                                                  PVOID taskBtnGroup,
                                                  int taskItemIndex,
                                                  int clickAction,
                                                  int param4,
                                                  int param5);
CTaskListWnd__HandleClick_t CTaskListWnd__HandleClick_Original;
void WINAPI CTaskListWnd__HandleClick_Hook(PVOID pThis,
                                           PVOID taskBtnGroup,
                                           int taskItemIndex,
                                           int clickAction,
                                           int param4,
                                           int param5) {
    Wh_Log(L"> %d", clickAction);

    auto original = [&]() {
        CTaskListWnd__HandleClick_Original(pThis, taskBtnGroup, taskItemIndex,
                                           clickAction, param4, param5);
    };

    if (clickAction != 0) {
        return original();
    }

    // Group types:
    // 1 - Single item or multiple uncombined items
    // 2 - Pinned item
    // 3 - Multiple combined items
    int groupType = CTaskBtnGroup_GetGroupType_Original(taskBtnGroup);
    if (groupType != 2) {
        return original();
    }

    static ULONGLONG firstClickTickCount = 0;
    static DWORD firstClickMessagePos;
    static PVOID firstClickTaskBtnGroup;

    ULONGLONG tickCount = GetTickCount64();
    DWORD messagePos = GetMessagePos();

    if (firstClickTickCount &&
        IsDoubleClickDistance(firstClickMessagePos, messagePos) &&
        firstClickTaskBtnGroup == taskBtnGroup &&
        tickCount - firstClickTickCount <= GetDoubleClickTime()) {
        // Double click detected, proceed.
        firstClickTickCount = 0;
        return original();
    }

    firstClickTickCount = tickCount;
    firstClickMessagePos = messagePos;
    firstClickTaskBtnGroup = taskBtnGroup;
}

using CTaskListWnd__HandleMouseButtonDown_t = void(WINAPI*)(PVOID pThis,
                                                            ULONGLONG param1,
                                                            const POINT* point,
                                                            bool isDoubleClick);
CTaskListWnd__HandleMouseButtonDown_t
    CTaskListWnd__HandleMouseButtonDown_Original;
void WINAPI CTaskListWnd__HandleMouseButtonDown_Hook(PVOID pThis,
                                                     ULONGLONG param1,
                                                     const POINT* point,
                                                     bool isDoubleClick) {
    Wh_Log(L">");

    // In Windows 10, double clicks on pinned items are ignored. Make all clicks
    // seem like single clicks.
    isDoubleClick = false;

    CTaskListWnd__HandleMouseButtonDown_Original(pThis, param1, point,
                                                 isDoubleClick);
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

    if (g_winVersion >= WinVersion::Win11 && g_settings.oldTaskbarOnWin11) {
        g_winVersion = WinVersion::Win10;
    }

    // Taskbar.dll, explorer.exe
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: virtual enum eTBGROUPTYPE __cdecl CTaskBtnGroup::GetGroupType(void))"},
            (void**)&CTaskBtnGroup_GetGroupType_Original,
        },
        {
            {LR"(protected: void __cdecl CTaskListWnd::_HandleClick(struct ITaskBtnGroup *,int,enum CTaskListWnd::eCLICKACTION,int,int))"},
            (void**)&CTaskListWnd__HandleClick_Original,
            (void*)CTaskListWnd__HandleClick_Hook,
        },
        {
            // Windows 10 only.
            {LR"(protected: void __cdecl CTaskListWnd::_HandleMouseButtonDown(unsigned __int64,struct tagPOINT const &,bool))"},
            (void**)&CTaskListWnd__HandleMouseButtonDown_Original,
            (void*)CTaskListWnd__HandleMouseButtonDown_Hook,
            true,
        },
    };

    HMODULE module;
    if (g_winVersion <= WinVersion::Win10) {
        module = GetModuleHandle(nullptr);
    } else {
        module = LoadLibrary(L"taskbar.dll");
        if (!module) {
            Wh_Log(L"Couldn't load taskbar.dll");
            return FALSE;
        }
    }

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return FALSE;
    }

    return TRUE;
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;

    LoadSettings();

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11;

    return TRUE;
}
