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
// @compilerOptions -lcomctl32 -lole32 -lversion
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
#include <atomic>
#include <climits>
#include <map>
#include <mutex>
#include <utility>
#include <vector>

#include <commctrl.h>
#include <objectarray.h>

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
};

WinVersion g_winVersion;
DWORD g_taskbarThreadId;
HWND g_hTaskbarWnd;

void* CTaskBtnGroup_ITaskBtnGroup_vftable;

using CTaskBtnGroup_GetGroup_t = void*(WINAPI*)(void* pThis);
CTaskBtnGroup_GetGroup_t CTaskBtnGroup_GetGroup;

using CTaskBtnGroup_GetNumItems_t = int(WINAPI*)(void* pThis);
CTaskBtnGroup_GetNumItems_t CTaskBtnGroup_GetNumItems;

using CTaskBtnGroup_GetTaskItem_t = void*(WINAPI*)(void* pThis, int index);
CTaskBtnGroup_GetTaskItem_t CTaskBtnGroup_GetTaskItem;

using CTaskBand__MatchApplicationView_t =
    HRESULT(WINAPI*)(LPVOID pThis,
                     IUnknown* applicationView,  // IApplicationView*
                     IUnknown** taskGroup,       // ITaskGroup**
                     IUnknown** taskItem);       // ITaskItem**
CTaskBand__MatchApplicationView_t CTaskBand__MatchApplicationView;

using CTaskListWnd_TryMoveGroup_t =
    bool(WINAPI*)(LPVOID pThis,
                  LPVOID taskGroup,  // ITaskGroup *
                  DWORD index);
CTaskListWnd_TryMoveGroup_t CTaskListWnd_TryMoveGroup_Original;

LPVOID g_tryMoveGroup_taskList;
LONG_PTR* g_tryMoveGroup_taskGroup;

bool WINAPI CTaskListWnd_TryMoveGroup_Hook(LPVOID pThis,
                                           LPVOID taskGroup,  // ITaskGroup *
                                           DWORD index) {
    g_tryMoveGroup_taskList = pThis;
    g_tryMoveGroup_taskGroup = (LONG_PTR*)taskGroup;

    bool ret = CTaskListWnd_TryMoveGroup_Original(pThis, taskGroup, index);

    g_tryMoveGroup_taskList = nullptr;
    g_tryMoveGroup_taskGroup = nullptr;

    return ret;
}

#pragma region view_order

// The order of app views which GetViews returns. The taskbar adds the buttons
// in this order after a virtual desktop switch. Holds no references.
std::mutex g_viewOrderMutex;
std::vector<IUnknown*> g_viewOrder;

// The CApplicationViewManager interface which GetViews is called on.
std::atomic<void*> g_appViewCollection;

struct GetViewsCapture {
    void* collection = nullptr;
    std::vector<IUnknown*> views;
};

thread_local GetViewsCapture* g_getViewsCapture;

using CEnumerableObjectCollection_AddObject_t =
    HRESULT(WINAPI*)(void* pThis, IUnknown* object);
CEnumerableObjectCollection_AddObject_t
    CEnumerableObjectCollection_AddObject_Original;
HRESULT WINAPI CEnumerableObjectCollection_AddObject_Hook(void* pThis,
                                                          IUnknown* object) {
    GetViewsCapture* capture = g_getViewsCapture;
    if (capture && object &&
        (!capture->collection || capture->collection == pThis)) {
        // Keep the view, it's added to the collection after GetViews returns.
        capture->collection = pThis;
        object->AddRef();
        capture->views.push_back(object);
        return S_OK;
    }

    return CEnumerableObjectCollection_AddObject_Original(pThis, object);
}

// Makes g_viewOrder contain exactly the given views, keeping the order of the
// known views and appending the new ones, and returns the views in that order.
std::vector<IUnknown*> UpdateViewOrder(const std::vector<IUnknown*>& views) {
    std::lock_guard<std::mutex> lock(g_viewOrderMutex);

    std::erase_if(g_viewOrder, [&views](IUnknown* view) {
        return std::ranges::find(views, view) == views.end();
    });

    for (IUnknown* view : views) {
        if (std::ranges::find(g_viewOrder, view) == g_viewOrder.end()) {
            g_viewOrder.push_back(view);
        }
    }

    return g_viewOrder;
}

using CApplicationViewManager_GetViews_t =
    HRESULT(WINAPI*)(void* pThis, IObjectArray** views);
CApplicationViewManager_GetViews_t CApplicationViewManager_GetViews_Original;

// Calls GetViews and adds the views to the returned collection in the order of
// g_viewOrder. On success, orderedViews receives the views in the same order,
// with a reference to each of them.
HRESULT GetViewsInOrder(void* pThis,
                        IObjectArray** views,
                        std::vector<IUnknown*>* orderedViews) {
    GetViewsCapture capture;
    GetViewsCapture* prevCapture = std::exchange(g_getViewsCapture, &capture);
    HRESULT hr = CApplicationViewManager_GetViews_Original(pThis, views);
    g_getViewsCapture = prevCapture;

    if (SUCCEEDED(hr)) {
        *orderedViews = UpdateViewOrder(capture.views);
        for (IUnknown* view : *orderedViews) {
            hr = CEnumerableObjectCollection_AddObject_Original(
                capture.collection, view);
            if (FAILED(hr)) {
                (*views)->Release();
                *views = nullptr;
                break;
            }
        }
    }

    if (FAILED(hr)) {
        for (IUnknown* view : capture.views) {
            view->Release();
        }

        orderedViews->clear();
    }

    return hr;
}

HRESULT WINAPI CApplicationViewManager_GetViews_Hook(void* pThis,
                                                     IObjectArray** views) {
    g_appViewCollection = pThis;

    std::vector<IUnknown*> orderedViews;
    HRESULT hr = GetViewsInOrder(pThis, views, &orderedViews);

    for (IUnknown* view : orderedViews) {
        view->Release();
    }

    return hr;
}

using CApplicationViewManager_RemoveApplicationView_t =
    HRESULT(WINAPI*)(void* pThis,
                     IUnknown* applicationView,  // IApplicationView*
                     int options);  // REMOVE_APPLICATION_VIEW_OPTIONS
CApplicationViewManager_RemoveApplicationView_t
    CApplicationViewManager_RemoveApplicationView_Original;
HRESULT WINAPI
CApplicationViewManager_RemoveApplicationView_Hook(void* pThis,
                                                   IUnknown* applicationView,
                                                   int options) {
    g_appViewCollection = pThis;

    HRESULT hr = CApplicationViewManager_RemoveApplicationView_Original(
        pThis, applicationView, options);

    // Forget the view right away, a new view might reuse its address.
    std::lock_guard<std::mutex> lock(g_viewOrderMutex);
    std::erase(g_viewOrder, applicationView);

    return hr;
}

const GUID kClsidImmersiveShell = {
    0xC2F03A33,
    0x21F5,
    0x47FA,
    {0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39}};
const GUID kIidApplicationViewCollection = {
    0x1841C6D7,
    0x4F9D,
    0x42C0,
    {0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5}};

// Only GetViews, the first method, is declared.
struct IApplicationViewCollection : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetViews(IObjectArray** views) = 0;
};

// Calls GetViews of the immersive shell's view collection service, which sets
// g_appViewCollection in the GetViews hook.
void QueryAppViewCollection() {
    HRESULT hrInit = CoInitializeEx(nullptr, COINIT_MULTITHREADED);

    IServiceProvider* serviceProvider = nullptr;
    HRESULT hr =
        CoCreateInstance(kClsidImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER,
                         IID_PPV_ARGS(&serviceProvider));

    IApplicationViewCollection* collection = nullptr;
    if (SUCCEEDED(hr)) {
        hr = serviceProvider->QueryService(kIidApplicationViewCollection,
                                           kIidApplicationViewCollection,
                                           (void**)&collection);
        serviceProvider->Release();
    }

    if (SUCCEEDED(hr)) {
        IObjectArray* views = nullptr;
        hr = collection->GetViews(&views);
        if (SUCCEEDED(hr)) {
            views->Release();
        }

        collection->Release();
    }

    if (FAILED(hr)) {
        Wh_Log(L"Querying the app view collection failed: %08X", hr);
    }

    if (SUCCEEDED(hrInit)) {
        CoUninitialize();
    }
}

#pragma endregion  // view_order

#pragma region desktop_arrangement

// The task groups of each taskbar's buttons, in order, saved per virtual
// desktop when switching away from it. A taskbar is identified by the DPA of
// its button groups. The keys and the task groups are only compared.
std::map<std::pair<HDPA, void*>, std::vector<void*>> g_desktopArrangements;

struct DesktopSwitch {
    void* oldDesktop;
    void* newDesktop;
    std::vector<HDPA> savedDpas;
};

// Set while the taskbar updates its buttons for a virtual desktop switch.
thread_local DesktopSwitch* g_desktopSwitch;

// Saves the order of the taskbar's button groups for the old desktop, the first
// time the switch changes them.
void SaveDesktopArrangement(HDPA hdpa) {
    std::vector<HDPA>& savedDpas = g_desktopSwitch->savedDpas;
    if (std::ranges::find(savedDpas, hdpa) != savedDpas.end()) {
        return;
    }

    savedDpas.push_back(hdpa);

    std::vector<void*> taskGroups;
    for (int i = 0; i < DPA_GetPtrCount(hdpa); i++) {
        taskGroups.push_back(CTaskBtnGroup_GetGroup(DPA_GetPtr(hdpa, i)));
    }

    g_desktopArrangements[{hdpa, g_desktopSwitch->oldDesktop}] =
        std::move(taskGroups);
}

// Returns the index at which the task group's new button group gets the
// position it had in the arrangement saved for the new desktop, or INT_MAX to
// append it.
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

    // An arrangement of the old desktop which the switch didn't replace might
    // be outdated.
    std::erase_if(g_desktopArrangements, [&desktopSwitch](const auto& item) {
        return item.first.second == desktopSwitch.oldDesktop &&
               std::ranges::find(desktopSwitch.savedDpas, item.first.first) ==
                   desktopSwitch.savedDpas.end();
    });

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
                     IObjectArray* objectArray,
                     void* oldDesktop,   // IVirtualDesktop*
                     void* newDesktop);  // IVirtualDesktop*
CTaskBand_CurrentVirtualDesktopChangedPre24H2_t
    CTaskBand_CurrentVirtualDesktopChangedPre24H2_Original;
HRESULT WINAPI
CTaskBand_CurrentVirtualDesktopChangedPre24H2_Hook(void* pThis,
                                                   IObjectArray* objectArray,
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

#pragma endregion  // desktop_arrangement

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

    void* appViewCollection = g_appViewCollection;
    if (!appViewCollection) {
        Wh_Log(L"GetViews wasn't called yet, skipping");
        return;
    }

    IObjectArray* views;
    std::vector<IUnknown*> orderedViews;
    HRESULT hr = GetViewsInOrder(appViewCollection, &views, &orderedViews);
    if (FAILED(hr)) {
        Wh_Log(L"GetViews failed: %08X", hr);
        return;
    }

    views->Release();

    // Find the views of the buttons in the newly inserted group, and the first
    // view which belongs to one of the groups to the right of it.
    std::vector<IUnknown*> matchingViews;
    IUnknown* rightNeighbourView = nullptr;

    for (IUnknown* view : orderedViews) {
        IUnknown* viewTaskGroup = nullptr;
        IUnknown* viewTaskItem = nullptr;
        CTaskBand__MatchApplicationView((LPVOID)lpTaskSwLongPtr, view,
                                        &viewTaskGroup, &viewTaskItem);

        if (viewTaskGroup && viewTaskGroup == (IUnknown*)task_group) {
            for (int j = 0; j < buttons_count; j++) {
                if (viewTaskItem ==
                    CTaskBtnGroup_GetTaskItem(button_group, j)) {
                    matchingViews.push_back(view);
                    break;
                }
            }
        } else if (viewTaskGroup && !rightNeighbourView) {
            for (int j = nButtonGroupIndex + 1; j < button_groups_count; j++) {
                if (viewTaskGroup == CTaskBtnGroup_GetGroup(button_groups[j])) {
                    rightNeighbourView = view;
                    break;
                }
            }
        }

        if (viewTaskGroup) {
            viewTaskGroup->Release();
        }

        if (viewTaskItem) {
            viewTaskItem->Release();
        }
    }

    // Move the matching views before the right neighbour view, or to the end.
    if (!matchingViews.empty()) {
        std::lock_guard<std::mutex> lock(g_viewOrderMutex);

        // Skip views which were removed in the meantime.
        std::erase_if(matchingViews, [](IUnknown* view) {
            return std::ranges::find(g_viewOrder, view) == g_viewOrder.end();
        });

        std::erase_if(g_viewOrder, [&matchingViews](IUnknown* view) {
            return std::ranges::find(matchingViews, view) !=
                   matchingViews.end();
        });

        g_viewOrder.insert(std::ranges::find(g_viewOrder, rightNeighbourView),
                           matchingViews.begin(), matchingViews.end());
    }

    for (IUnknown* view : orderedViews) {
        view->Release();
    }
}

void ComFuncVirtualDesktopFixAfterDPA_InsertPtr(HDPA pdpa, int index, void* p) {
    Wh_Log(L"> index=%d, p=%p", index, p);

    if (index == INT_MAX) {
        return;
    }

    if (!g_tryMoveGroup_taskList) {
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
    if (g_desktopSwitch && g_createTBGroup_taskGroup && p &&
        *(void**)p == CTaskBtnGroup_ITaskBtnGroup_vftable &&
        CTaskBtnGroup_GetGroup(p) == g_createTBGroup_taskGroup) {
        SaveDesktopArrangement(hdpa);
        if (i == INT_MAX) {
            i = GetDesktopArrangementIndex(hdpa, g_createTBGroup_taskGroup);
            Wh_Log(L"Inserting button group %p at %d", p, i);
        }
    }

    auto ret = DPA_InsertPtr_Original(hdpa, i, p);

    if (GetCurrentThreadId() == g_taskbarThreadId) {
        ComFuncVirtualDesktopFixAfterDPA_InsertPtr(hdpa, i, p);
    }

    return ret;
}

using DPA_DeletePtr_t = decltype(&DPA_DeletePtr);
DPA_DeletePtr_t DPA_DeletePtr_Original;
void* WINAPI DPA_DeletePtr_Hook(HDPA hdpa, int i) {
    if (g_desktopSwitch && g_taskDestroyedFinal_buttonGroup &&
        DPA_GetPtr(hdpa, i) == g_taskDestroyedFinal_buttonGroup) {
        SaveDesktopArrangement(hdpa);
    }

    return DPA_DeletePtr_Original(hdpa, i);
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
            {LR"(public: virtual int __cdecl CTaskBtnGroup::GetNumItems(void))"},
            &CTaskBtnGroup_GetNumItems,
        },
        {
            {LR"(public: virtual struct ITaskItem * __cdecl CTaskBtnGroup::GetTaskItem(int))"},
            &CTaskBtnGroup_GetTaskItem,
        },
        {
            {LR"(protected: long __cdecl CTaskBand::_MatchApplicationView(struct IApplicationView *,struct ITaskGroup * *,struct ITaskItem * *))"},
            &CTaskBand__MatchApplicationView,
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

    // twinui.pcshell.dll
    WindhawkUtils::SYMBOL_HOOK twinuiPcshellSymbolHooks[] = {
        {
            {LR"(public: virtual long __cdecl CApplicationViewManager::GetViews(struct IObjectArray * *))"},
            &CApplicationViewManager_GetViews_Original,
            CApplicationViewManager_GetViews_Hook,
        },
        {
            {LR"(public: virtual long __cdecl CApplicationViewManager::RemoveApplicationView(struct IApplicationView *,enum REMOVE_APPLICATION_VIEW_OPTIONS))"},
            &CApplicationViewManager_RemoveApplicationView_Original,
            CApplicationViewManager_RemoveApplicationView_Hook,
        },
        {
            {LR"(public: virtual long __cdecl CEnumerableObjectCollection::AddObject(struct IUnknown *))"},
            &CEnumerableObjectCollection_AddObject_Original,
            CEnumerableObjectCollection_AddObject_Hook,
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

        // The taskbar already started, and won't call GetViews until the next
        // virtual desktop switch.
        QueryAppViewCollection();
    }
}
