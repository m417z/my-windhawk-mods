// ==WindhawkMod==
// @id              taskbar-labels
// @name            Taskbar Labels for Windows 11
// @description     Show text labels for running programs on the taskbar (Windows 11 only)
// @version         1.0.3
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -loleaut32 -lole32 -lruntimeobject
// ==/WindhawkMod==

// Source code is published under The GNU General Public License v3.0.

// ==WindhawkModReadme==
/*
# Taskbar Labels for Windows 11
Show text labels for running programs on the taskbar (Windows 11 only).

By default, the Windows 11 taskbar only shows icons for taskbar items, without
any text labels. This mod adds text labels, similarly to the way it was possible
to configure in older Windows versions.

Before:

![Before screenshot](https://i.imgur.com/SjHSF7g.png)

After:

![After screenshot](https://i.imgur.com/qpc4iFh.png)

## Known limitations

* When there are too many items, the rightmost items become unreachable.
* Switching between virtual desktops might result in missing labels.
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- taskbarItemWidth: 160
  $name: Taskbar item width
  $description: >-
    The width of taskbar items for which text labels are added
*/
// ==/WindhawkModSettings==

#undef GetCurrentTime

#include <initguid.h>  // must come before knownfolders.h

#include <inspectable.h>
#include <knownfolders.h>
#include <shlobj.h>

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Markup.h>
#include <winrt/Windows.UI.Xaml.Media.h>

#include <memory>
#include <regex>
#include <type_traits>

// #define EXTRA_DBG_LOG

struct {
    int taskbarItemWidth;
} g_settings;

bool g_applyingSettings = false;
bool g_unloading = false;

#ifndef SPI_SETLOGICALDPIOVERRIDE
#define SPI_SETLOGICALDPIOVERRIDE 0x009F
#endif

winrt::Windows::UI::Xaml::FrameworkElement FindChildByClassName(
    winrt::Windows::UI::Xaml::FrameworkElement element,
    PCWSTR className) {
    int childrenCount =
        winrt::Windows::UI::Xaml::Media::VisualTreeHelper::GetChildrenCount(
            element);

    for (int i = 0; i < childrenCount; i++) {
        auto child =
            winrt::Windows::UI::Xaml::Media::VisualTreeHelper::GetChild(element,
                                                                        i)
                .try_as<winrt::Windows::UI::Xaml::FrameworkElement>();
        if (child && child.Name() == className) {
            return child;
        }
    }

    return nullptr;
}

#if defined(EXTRA_DBG_LOG)

void LogElement(Windows_UI_Xaml_IDependencyObject* pChild, int depth) {
    HRESULT hr;

    {
        HString hsChild(nullptr, &WindowsDeleteString);
        {
            HSTRING hs = nullptr;
            hr = pChild->GetRuntimeClassName(&hs);
            if (SUCCEEDED(hr))
                hsChild.reset(hs);
        }

        PCWSTR pwszName =
            hsChild ? WindowsGetStringRawBuffer(hsChild.get(), 0) : L"-";

        WCHAR padding[MAX_PATH]{};
        for (int i = 0; i < depth * 4; i++) {
            padding[i] = L' ';
        }
        Wh_Log(L"%sClass: %s", padding, pwszName);
    }

    ComPtr<Windows_UI_Xaml_IFrameworkElement> pFrameworkElement;
    hr = pChild->QueryInterface(IID_Windows_UI_Xaml_IFrameworkElement,
                                (void**)&pFrameworkElement);
    if (SUCCEEDED(hr)) {
        HString hsChild(nullptr, &WindowsDeleteString);
        {
            HSTRING hs = nullptr;
            hr = pFrameworkElement->get_Name(&hs);
            if (SUCCEEDED(hr))
                hsChild.reset(hs);
        }

        PCWSTR pwszName =
            hsChild ? WindowsGetStringRawBuffer(hsChild.get(), 0) : L"-";

        WCHAR padding[MAX_PATH]{};
        for (int i = 0; i < depth * 4; i++) {
            padding[i] = L' ';
        }
        Wh_Log(L"%sName: %s", padding, pwszName);
    }
}

void LogElementTreeAux(
    Windows_UI_Xaml_IDependencyObject* pRootDependencyObject,
    Windows_UI_Xaml_IVisualTreeHelperStatics* pVisualTreeHelperStatics,
    int depth) {
    if (!pRootDependencyObject) {
        return;
    }

    LogElement(pRootDependencyObject, depth);

    HRESULT hr = S_OK;
    INT32 Count = -1;
    hr = pVisualTreeHelperStatics->GetChildrenCount(pRootDependencyObject,
                                                    &Count);
    if (SUCCEEDED(hr)) {
        for (INT32 Index = 0; Index < Count; ++Index) {
            ComPtr<Windows_UI_Xaml_IDependencyObject> pChild;
            hr = pVisualTreeHelperStatics->GetChild(pRootDependencyObject,
                                                    Index, &pChild);
            if (SUCCEEDED(hr)) {
                LogElementTreeAux(pChild.Get(), pVisualTreeHelperStatics,
                                  depth + 1);
            }
        }
    }
}

void LogElementTree(IInspectable* pElement) {
    HRESULT hr;

    HStringReference hsVisualTreeHelperStatics(
        L"Windows.UI.Xaml.Media.VisualTreeHelper");
    ComPtr<Windows_UI_Xaml_IVisualTreeHelperStatics> pVisualTreeHelperStatics;
    hr = RoGetActivationFactory(hsVisualTreeHelperStatics.Get(),
                                IID_Windows_UI_Xaml_IVisualTreeHelperStatics,
                                (void**)&pVisualTreeHelperStatics);
    if (SUCCEEDED(hr)) {
        ComPtr<Windows_UI_Xaml_IDependencyObject> pRootDependencyObject;
        hr = pElement->QueryInterface(IID_Windows_UI_Xaml_IDependencyObject,
                                      (void**)&pRootDependencyObject);
        if (SUCCEEDED(hr)) {
            LogElementTreeAux(pRootDependencyObject.Get(),
                              pVisualTreeHelperStatics.Get(), 0);
        }
    }
}

#endif  // defined(EXTRA_DBG_LOG)

////////////////////////////////////////////////////////////////////////////////

using CTaskListWnd__GetTBGroupFromGroup_t = void*(WINAPI*)(void* pThis,
                                                           void* taskGroup,
                                                           int* index);
CTaskListWnd__GetTBGroupFromGroup_t CTaskListWnd__GetTBGroupFromGroup_Original;

using CTaskBtnGroup_GetNumItems_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetNumItems_t CTaskBtnGroup_GetNumItems_Original;

using CTaskBtnGroup_GetTaskItem_t = void*(WINAPI*)(void* pThis, int index);
CTaskBtnGroup_GetTaskItem_t CTaskBtnGroup_GetTaskItem_Original;

using CTaskGroup_GetTitleText_t = LONG_PTR(WINAPI*)(void* pThis,
                                                    void* taskItem,
                                                    WCHAR* text,
                                                    int bufferSize);
CTaskGroup_GetTitleText_t CTaskGroup_GetTitleText_Original;

using IconContainer_IsStorageRecreationRequired_t = bool(WINAPI*)(void* pThis,
                                                                  void* param1,
                                                                  int flags);
IconContainer_IsStorageRecreationRequired_t
    IconContainer_IsStorageRecreationRequired_Original;
bool WINAPI IconContainer_IsStorageRecreationRequired_Hook(void* pThis,
                                                           void* param1,
                                                           int flags) {
    if (g_applyingSettings) {
        return true;
    }

    return IconContainer_IsStorageRecreationRequired_Original(pThis, param1,
                                                              flags);
}

bool g_inGroupChanged;
WCHAR g_taskBtnGroupTitleInGroupChanged[256];

using CTaskListWnd_GroupChanged_t = LONG_PTR(WINAPI*)(void* pThis,
                                                      void* taskGroup,
                                                      int taskGroupProperty);
CTaskListWnd_GroupChanged_t CTaskListWnd_GroupChanged_Original;
LONG_PTR WINAPI CTaskListWnd_GroupChanged_Hook(void* pThis,
                                               void* taskGroup,
                                               int taskGroupProperty) {
    Wh_Log(L">");

    int numItems = 0;
    void* taskItem = nullptr;

    void* taskBtnGroup = CTaskListWnd__GetTBGroupFromGroup_Original(
        (BYTE*)pThis - 0x28, taskGroup, nullptr);
    if (taskBtnGroup) {
        numItems = CTaskBtnGroup_GetNumItems_Original(taskBtnGroup);
        if (numItems == 1) {
            taskItem = CTaskBtnGroup_GetTaskItem_Original(taskBtnGroup, 0);
        }
    }

    WCHAR* textBuffer = g_taskBtnGroupTitleInGroupChanged;
    int textBufferSize = ARRAYSIZE(g_taskBtnGroupTitleInGroupChanged);

    if (numItems > 1) {
        int written = wsprintf(textBuffer, L"[%d] ", numItems);
        textBuffer += written;
        textBufferSize -= written;
    }

    CTaskGroup_GetTitleText_Original(taskGroup, taskItem, textBuffer,
                                     textBufferSize);

    g_inGroupChanged = true;
    LONG_PTR ret =
        CTaskListWnd_GroupChanged_Original(pThis, taskGroup, taskGroupProperty);
    g_inGroupChanged = false;

    *g_taskBtnGroupTitleInGroupChanged = L'\0';

    return ret;
}

using CTaskListWnd_TaskDestroyed_t = LONG_PTR(WINAPI*)(void* pThis,
                                                       void* taskGroup,
                                                       void* taskItem,
                                                       int taskDestroyedFlags);
CTaskListWnd_TaskDestroyed_t CTaskListWnd_TaskDestroyed_Original;
LONG_PTR WINAPI CTaskListWnd_TaskDestroyed_Hook(void* pThis,
                                                void* taskGroup,
                                                void* taskItem,
                                                int taskDestroyedFlags) {
    Wh_Log(L">");

    LONG_PTR ret = CTaskListWnd_TaskDestroyed_Original(
        pThis, taskGroup, taskItem, taskDestroyedFlags);

    // Trigger CTaskListWnd::GroupChanged to trigger the title change.
    int taskGroupProperty = 4;  // saw this in the debugger
    CTaskListWnd_GroupChanged_Hook(pThis, taskGroup, taskGroupProperty);

    return ret;
}

////////////////////////////////////////////////////////////////////////////////

double* double_48_value_Original;

using ITaskListButton_get_IsRunning_t = HRESULT(WINAPI*)(void* pThis,
                                                         bool* running);
ITaskListButton_get_IsRunning_t ITaskListButton_get_IsRunning_Original;

void UpdateTaskListButtonCustomizations(void* pTaskListButtonImpl) {
    winrt::Windows::Foundation::IInspectable taskListButtonIInspectable;
    winrt::copy_from_abi(taskListButtonIInspectable,
                         *(IInspectable**)pTaskListButtonImpl);

    auto taskListButtonElement =
        taskListButtonIInspectable
            .as<winrt::Windows::UI::Xaml::FrameworkElement>();

    auto iconPanelElement =
        FindChildByClassName(taskListButtonElement, L"IconPanel");
    if (!iconPanelElement) {
        return;
    }

    auto iconElement = FindChildByClassName(iconPanelElement, L"Icon");
    if (!iconElement) {
        return;
    }

    double taskListButtonWidth = taskListButtonElement.ActualWidth();

    double iconPanelWidth = iconPanelElement.ActualWidth();

    // Check if non-positive or NaN.
    if (!(taskListButtonWidth > 0) || !(iconPanelWidth > 0)) {
        return;
    }

    static double initialWidth = iconPanelWidth;

    void* taskListButtonProducer = (BYTE*)pTaskListButtonImpl + 0x10;

    bool isRunning = false;
    ITaskListButton_get_IsRunning_Original(taskListButtonProducer, &isRunning);

    bool showLabels = isRunning && !g_unloading;

    double widthToSet = showLabels ? g_settings.taskbarItemWidth : initialWidth;

    if (widthToSet != taskListButtonWidth || widthToSet != iconPanelWidth) {
        taskListButtonElement.Width(widthToSet);
        iconPanelElement.Width(widthToSet);

        auto horizontalAlignment =
            showLabels ? winrt::Windows::UI::Xaml::HorizontalAlignment::Left
                       : winrt::Windows::UI::Xaml::HorizontalAlignment::Center;
        iconElement.HorizontalAlignment(horizontalAlignment);

        winrt::Windows::UI::Xaml::Thickness margin{};
        if (showLabels) {
            margin.Left = 10;
        }
        iconElement.Margin(margin);
    }

    auto windhawkTextElement =
        FindChildByClassName(iconPanelElement, L"WindhawkText");

    if (showLabels && !windhawkTextElement) {
        auto iconPanel =
            iconPanelElement.as<winrt::Windows::UI::Xaml::Controls::Panel>();

        auto children = iconPanel.Children();

        PCWSTR xaml =
            LR"(
                <TextBlock
                    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                    mc:Ignorable="d"
                    Name="WindhawkText"
                    VerticalAlignment="Center"
                    FontSize="12"
                    TextTrimming="CharacterEllipsis"
                />
            )";

        auto pUIElement =
            winrt::Windows::UI::Xaml::Markup::XamlReader::Load(xaml)
                .as<winrt::Windows::UI::Xaml::FrameworkElement>();
        if (pUIElement) {
            children.Append(pUIElement);

            winrt::Windows::UI::Xaml::Thickness margin{
                .Left = 10 + 24 + 8,
                .Right = 10,
                .Bottom = 2,
            };
            pUIElement.Margin(margin);
        }
    } else if (!showLabels && windhawkTextElement) {
        // Don't remove, for some reason it causes a bug - the running indicator
        // ends up being behind the semi-transparent rectangle of the active
        // button.
        /*
        ComPtr<Windows_UI_Xaml_Controls_IPanel> pIconPanel;
        hr = pIconPanelElement->QueryInterface(
            IID_Windows_UI_Xaml_Controls_IPanel, (void**)&pIconPanel);
        if (SUCCEEDED(hr)) {
            ComPtr<IVector> children;
            hr = pIconPanel->get_Children(&children);
            if (SUCCEEDED(hr)) {
                unsigned int index = -1;
                boolean found = false;
                hr = children->IndexOf(pWindhawkTextElement.Get(), &index,
                                       &found);
                if (SUCCEEDED(hr) && found) {
                    hr = children->RemoveAt(index);
                }
            }
        }
        */

        // Set empty text instead.
        auto windhawkTextControl =
            windhawkTextElement
                .as<winrt::Windows::UI::Xaml::Controls::TextBlock>();
        windhawkTextControl.Text(L"");
    }
}

using TaskListButton_UpdateVisualStates_t = void(WINAPI*)(void* pThis);
TaskListButton_UpdateVisualStates_t TaskListButton_UpdateVisualStates_Original;
void WINAPI TaskListButton_UpdateVisualStates_Hook(void* pThis) {
    Wh_Log(L">");

    TaskListButton_UpdateVisualStates_Original(pThis);

    void* pTaskListButtonImpl = (BYTE*)pThis + 0x08;

#if defined(EXTRA_DBG_LOG)
    Wh_Log(L"=====");
    IInspectable* pTaskListButton = *(IInspectable**)pTaskListButtonImpl;
    LogElementTree(pTaskListButton);
    Wh_Log(L"=====");
#endif  // defined(EXTRA_DBG_LOG)

    UpdateTaskListButtonCustomizations(pTaskListButtonImpl);
}

using TaskListButton_UpdateButtonPadding_t = void(WINAPI*)(void* pThis);
TaskListButton_UpdateButtonPadding_t
    TaskListButton_UpdateButtonPadding_Original;
void WINAPI TaskListButton_UpdateButtonPadding_Hook(void* pThis) {
    Wh_Log(L">");

    TaskListButton_UpdateButtonPadding_Original(pThis);

    void* pTaskListButtonImpl = (BYTE*)pThis + 0x08;

    UpdateTaskListButtonCustomizations(pTaskListButtonImpl);
}

using TaskListButton_Icon_t = void(WINAPI*)(void* pThis,
                                            LONG_PTR randomAccessStream);
TaskListButton_Icon_t TaskListButton_Icon_Original;
void WINAPI TaskListButton_Icon_Hook(void* pThis, LONG_PTR randomAccessStream) {
    Wh_Log(L">");

    TaskListButton_Icon_Original(pThis, randomAccessStream);

    if (!g_inGroupChanged) {
        return;
    }

    void* pTaskListButtonImpl = (BYTE*)pThis + 0x08;

    winrt::Windows::Foundation::IInspectable taskListButtonIInspectable;
    winrt::copy_from_abi(taskListButtonIInspectable,
                         *(IInspectable**)pTaskListButtonImpl);

    auto taskListButtonElement =
        taskListButtonIInspectable
            .as<winrt::Windows::UI::Xaml::FrameworkElement>();

    auto iconPanelElement =
        FindChildByClassName(taskListButtonElement, L"IconPanel");
    if (iconPanelElement) {
        auto windhawkTextElement =
            FindChildByClassName(iconPanelElement, L"WindhawkText");
        if (windhawkTextElement) {
            auto windhawkTextControl =
                windhawkTextElement
                    .as<winrt::Windows::UI::Xaml::Controls::TextBlock>();
            windhawkTextControl.Text(g_taskBtnGroupTitleInGroupChanged);
        }
    }
}

void LoadSettings() {
    g_settings.taskbarItemWidth = Wh_GetIntSetting(L"taskbarItemWidth");
}

void FreeSettings() {
    // Nothing for now.
}

bool ProtectAndMemcpy(DWORD protect, void* dst, const void* src, size_t size) {
    DWORD oldProtect;
    if (!VirtualProtect(dst, size, protect, &oldProtect)) {
        return false;
    }

    memcpy(dst, src, size);
    VirtualProtect(dst, size, oldProtect, &oldProtect);
    return true;
}

void ApplySettings() {
    HWND hTaskbarWnd = FindWindow(L"Shell_TrayWnd", nullptr);
    DWORD dwTaskbarProcessId = 0;
    if (hTaskbarWnd &&
        GetWindowThreadProcessId(hTaskbarWnd, &dwTaskbarProcessId) &&
        dwTaskbarProcessId != GetCurrentProcessId()) {
        hTaskbarWnd = nullptr;
    }

    if (!hTaskbarWnd) {
        return;
    }

    g_applyingSettings = true;

    double prevTaskbarHeight = *double_48_value_Original;

    RECT taskbarRect{};
    GetWindowRect(hTaskbarWnd, &taskbarRect);

    // Temporarily change the height to force a UI refresh.
    double tempTaskbarHeight = 1;
    ProtectAndMemcpy(PAGE_READWRITE, double_48_value_Original,
                     &tempTaskbarHeight, sizeof(double));

    // Trigger TrayUI::_HandleSettingChange.
    SendMessage(hTaskbarWnd, WM_SETTINGCHANGE, SPI_SETLOGICALDPIOVERRIDE, 0);

    // Wait for the change to apply.
    RECT newTaskbarRect{};
    int counter = 0;
    while (GetWindowRect(hTaskbarWnd, &newTaskbarRect) &&
           newTaskbarRect.top == taskbarRect.top) {
        if (++counter >= 100) {
            break;
        }
        Sleep(100);
    }

    ProtectAndMemcpy(PAGE_READWRITE, double_48_value_Original,
                     &prevTaskbarHeight, sizeof(double));

    if (hTaskbarWnd) {
        // Trigger TrayUI::_HandleSettingChange.
        SendMessage(hTaskbarWnd, WM_SETTINGCHANGE, SPI_SETLOGICALDPIOVERRIDE,
                    0);

        HWND hReBarWindow32 =
            FindWindowEx(hTaskbarWnd, nullptr, L"ReBarWindow32", nullptr);
        if (hReBarWindow32) {
            HWND hMSTaskSwWClass = FindWindowEx(hReBarWindow32, nullptr,
                                                L"MSTaskSwWClass", nullptr);
            if (hMSTaskSwWClass) {
                // Trigger CTaskBand::_HandleSyncDisplayChange.
                SendMessage(hMSTaskSwWClass, 0x452, 3, 0);
            }
        }
    }

    g_applyingSettings = false;
}

struct SYMBOL_HOOK {
    std::wregex symbolRegex;
    void** pOriginalFunction;
    void* hookFunction = nullptr;
    bool optional = false;
};

bool HookSymbols(HMODULE module,
                 SYMBOL_HOOK* symbolHooks,
                 size_t symbolHooksCount) {
    WH_FIND_SYMBOL symbol;
    HANDLE findSymbol = Wh_FindFirstSymbol(module, nullptr, &symbol);
    if (!findSymbol) {
        return false;
    }

    do {
        for (size_t i = 0; i < symbolHooksCount; i++) {
            if (!*symbolHooks[i].pOriginalFunction &&
                std::regex_match(symbol.symbol, symbolHooks[i].symbolRegex)) {
                if (symbolHooks[i].hookFunction) {
                    Wh_SetFunctionHook(symbol.address,
                                       symbolHooks[i].hookFunction,
                                       symbolHooks[i].pOriginalFunction);
                    Wh_Log(L"Hooked %p (%s)", symbol.address, symbol.symbol);
                } else {
                    *symbolHooks[i].pOriginalFunction = symbol.address;
                    Wh_Log(L"Found %p (%s)", symbol.address, symbol.symbol);
                }
                break;
            }
        }
    } while (Wh_FindNextSymbol(findSymbol, &symbol));

    Wh_FindCloseSymbol(findSymbol);

    for (size_t i = 0; i < symbolHooksCount; i++) {
        if (!symbolHooks[i].optional && !*symbolHooks[i].pOriginalFunction) {
            Wh_Log(L"Missing symbol: %d", i);
            return false;
        }
    }

    return true;
}

bool HookTaskbarDllSymbols() {
    HMODULE module = LoadLibrary(L"taskbar.dll");
    if (!module) {
        Wh_Log(L"Failed to load taskbar.dll");
        return false;
    }

    SYMBOL_HOOK symbolHooks[] = {
        {std::wregex(
             LR"(protected: struct ITaskBtnGroup \* __ptr64 __cdecl CTaskListWnd::_GetTBGroupFromGroup\(struct ITaskGroup \* __ptr64,int \* __ptr64\) __ptr64)"),
         (void**)&CTaskListWnd__GetTBGroupFromGroup_Original, nullptr},

        {std::wregex(
             LR"(public: virtual int __cdecl CTaskBtnGroup::GetNumItems\(void\) __ptr64)"),
         (void**)&CTaskBtnGroup_GetNumItems_Original, nullptr},

        {std::wregex(
             LR"(public: virtual struct ITaskItem \* __ptr64 __cdecl CTaskBtnGroup::GetTaskItem\(int\) __ptr64)"),
         (void**)&CTaskBtnGroup_GetTaskItem_Original, nullptr},

        {std::wregex(
             LR"(public: virtual long __cdecl CTaskGroup::GetTitleText\(struct ITaskItem \* __ptr64,unsigned short \* __ptr64,int\) __ptr64)"),
         (void**)&CTaskGroup_GetTitleText_Original, nullptr},

        {std::wregex(
             LR"(public: virtual bool __cdecl IconContainer::IsStorageRecreationRequired\(class CCoSimpleArray<unsigned int,4294967294,class CSimpleArrayStandardCompareHelper<unsigned int> > const & __ptr64,enum IconContainerFlags\) __ptr64)"),
         (void**)&IconContainer_IsStorageRecreationRequired_Original,
         (void*)IconContainer_IsStorageRecreationRequired_Hook},

        {std::wregex(
             LR"(public: virtual void __cdecl CTaskListWnd::GroupChanged\(struct ITaskGroup \* __ptr64,enum winrt::WindowsUdk::UI::Shell::TaskGroupProperty\) __ptr64)"),
         (void**)&CTaskListWnd_GroupChanged_Original,
         (void*)CTaskListWnd_GroupChanged_Hook},

        {std::wregex(
             LR"(public: virtual long __cdecl CTaskListWnd::TaskDestroyed\(struct ITaskGroup \* __ptr64,struct ITaskItem \* __ptr64,enum TaskDestroyedFlags\) __ptr64)"),
         (void**)&CTaskListWnd_TaskDestroyed_Original,
         (void*)CTaskListWnd_TaskDestroyed_Hook},
    };

    return HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks));
}

bool HookTaskbarViewDllSymbols() {
    WCHAR szWindowsDirectory[MAX_PATH];
    if (!GetWindowsDirectory(szWindowsDirectory,
                             ARRAYSIZE(szWindowsDirectory))) {
        Wh_Log(L"GetWindowsDirectory failed");
        return false;
    }

    bool windowsVersionIdentified = false;
    HMODULE module;

    WCHAR szTargetDllPath[MAX_PATH];
    wcscpy_s(szTargetDllPath, szWindowsDirectory);
    wcscat_s(
        szTargetDllPath,
        LR"(\SystemApps\MicrosoftWindows.Client.Core_cw5n1h2txyewy\Taskbar.View.dll)");
    if (GetFileAttributes(szTargetDllPath) != INVALID_FILE_ATTRIBUTES) {
        // Windows 11 version 22H2.
        windowsVersionIdentified = true;

        module = GetModuleHandle(szTargetDllPath);
        if (!module) {
            // Try to load dependency DLLs. At process start, if they're not
            // loaded, loading the taskbar view DLL fails.
            WCHAR szRuntimeDllPath[MAX_PATH];

            wcscpy_s(szRuntimeDllPath, szWindowsDirectory);
            wcscat_s(
                szRuntimeDllPath,
                LR"(\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\vcruntime140_app.dll)");
            LoadLibrary(szRuntimeDllPath);

            wcscpy_s(szRuntimeDllPath, szWindowsDirectory);
            wcscat_s(
                szRuntimeDllPath,
                LR"(\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\vcruntime140_1_app.dll)");
            LoadLibrary(szRuntimeDllPath);

            wcscpy_s(szRuntimeDllPath, szWindowsDirectory);
            wcscat_s(
                szRuntimeDllPath,
                LR"(\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\msvcp140_app.dll)");
            LoadLibrary(szRuntimeDllPath);

            module = LoadLibrary(szTargetDllPath);
        }
    }

    if (!windowsVersionIdentified) {
        wcscpy_s(szTargetDllPath, szWindowsDirectory);
        wcscat_s(
            szTargetDllPath,
            LR"(\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\ExplorerExtensions.dll)");
        if (GetFileAttributes(szTargetDllPath) != INVALID_FILE_ATTRIBUTES) {
            // Windows 11 version 21H2.
            windowsVersionIdentified = true;

            module = GetModuleHandle(szTargetDllPath);
            if (!module) {
                // Try to load dependency DLLs. At process start, if they're not
                // loaded, loading the ExplorerExtensions DLL fails.
                WCHAR szRuntimeDllPath[MAX_PATH];

                PWSTR pProgramFilesDirectory;
                if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_ProgramFiles, 0,
                                                   nullptr,
                                                   &pProgramFilesDirectory))) {
                    wcscpy_s(szRuntimeDllPath, pProgramFilesDirectory);
                    wcscat_s(
                        szRuntimeDllPath,
                        LR"(\WindowsApps\Microsoft.VCLibs.140.00_14.0.29231.0_x64__8wekyb3d8bbwe\vcruntime140_app.dll)");
                    LoadLibrary(szRuntimeDllPath);

                    wcscpy_s(szRuntimeDllPath, pProgramFilesDirectory);
                    wcscat_s(
                        szRuntimeDllPath,
                        LR"(\WindowsApps\Microsoft.VCLibs.140.00_14.0.29231.0_x64__8wekyb3d8bbwe\vcruntime140_1_app.dll)");
                    LoadLibrary(szRuntimeDllPath);

                    wcscpy_s(szRuntimeDllPath, pProgramFilesDirectory);
                    wcscat_s(
                        szRuntimeDllPath,
                        LR"(\WindowsApps\Microsoft.VCLibs.140.00_14.0.29231.0_x64__8wekyb3d8bbwe\msvcp140_app.dll)");
                    LoadLibrary(szRuntimeDllPath);

                    CoTaskMemFree(pProgramFilesDirectory);

                    module = LoadLibrary(szTargetDllPath);
                }
            }
        }
    }

    if (!module) {
        Wh_Log(L"Failed to load module");
        return FALSE;
    }

    SYMBOL_HOOK symbolHooks[] = {
        {std::wregex(LR"(__real@4048000000000000)"),
         (void**)&double_48_value_Original, nullptr},

        {std::wregex(
             LR"(public: virtual int __cdecl winrt::impl::produce<struct winrt::Taskbar::implementation::TaskListButton,struct winrt::Taskbar::ITaskListButton>::get_IsRunning\(bool \* __ptr64\) __ptr64)"),
         (void**)&ITaskListButton_get_IsRunning_Original, nullptr},

        {std::wregex(
             LR"(private: void __cdecl winrt::Taskbar::implementation::TaskListButton::UpdateVisualStates\(void\) __ptr64)"),
         (void**)&TaskListButton_UpdateVisualStates_Original,
         (void*)TaskListButton_UpdateVisualStates_Hook},

        {std::wregex(
             LR"(private: void __cdecl winrt::Taskbar::implementation::TaskListButton::UpdateButtonPadding\(void\) __ptr64)"),
         (void**)&TaskListButton_UpdateButtonPadding_Original,
         (void*)TaskListButton_UpdateButtonPadding_Hook},

        {std::wregex(
             LR"(public: void __cdecl winrt::Taskbar::implementation::TaskListButton::Icon\(struct winrt::Windows::Storage::Streams::IRandomAccessStream\) __ptr64)"),
         (void**)&TaskListButton_Icon_Original,
         (void*)TaskListButton_Icon_Hook},
    };

    return HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks));
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    LoadSettings();

    if (!HookTaskbarDllSymbols()) {
        return FALSE;
    }

    if (!HookTaskbarViewDllSymbols()) {
        return FALSE;
    }

    return TRUE;
}

void Wh_ModAfterInit(void) {
    Wh_Log(L">");

    ApplySettings();
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");

    g_unloading = true;
    ApplySettings();
}

void Wh_ModUninit() {
    Wh_Log(L">");

    FreeSettings();
}

void Wh_ModSettingsChanged() {
    Wh_Log(L">");

    FreeSettings();
    LoadSettings();

    ApplySettings();
}
