// ==WindhawkMod==
// @id              taskbar-auto-hide-when-maximized
// @name            Taskbar auto-hide when maximized
// @description     Makes the taskbar auto-hide only when a window is maximized or intersects the taskbar
// @version         1.2.7
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -ldwmapi -lgdi32 -lole32 -loleaut32 -lversion
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
# Taskbar auto-hide when maximized

Makes the taskbar auto-hide only when a window is maximized or intersects the
taskbar.

The **Dock** mode is compatible with the visual regions published by Windows
11 Taskbar Styler 1.8 and newer, including its Click-through taskbar style. If
the visual region is unavailable, invalid, stale, or unsupported, the complete
native taskbar rectangle is used as a safe fallback. In Dock mode, a taskbar
revealed by Windows is restored above normal maximized applications. A true
fullscreen foreground window on the same monitor still suppresses that
z-order repair, so edge reveal doesn't force the taskbar above fullscreen
content.

![Demonstration](https://i.imgur.com/hEz1lhs.gif)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- mode: intersected
  $name: Mode
  $options:
  - intersected: Auto-hide when a window is maximized or intersects the taskbar
  - dock: Auto-hide when a window is maximized or intersects the visible taskbar dock
  - maximized: Auto-hide only when a window is maximized
  - never: Never auto-hide
- foregroundWindowOnly: false
  $name: Apply only to foreground window
  $description: >-
    Enable this option to apply the auto-hide taskbar feature only to the
    selected window.
- excludedPrograms: [""]
  $name: Excluded programs
  $description: >-
    The taskbar won't auto-hide due to windows of these programs being maximized
    or intersecting the taskbar.

    Entries can be process names, paths or application IDs, for example:

    mspaint.exe

    C:\Windows\System32\notepad.exe

    Microsoft.WindowsCalculator_8wekyb3d8bbwe!App
- primaryMonitorOnly: false
  $name: Primary monitor only
  $description: >-
    Apply the mod's behavior only to the primary monitor taskbar. Secondary
    monitors will use Windows' default auto-hide behavior.
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool).
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <dwmapi.h>
#include <psapi.h>

#include <winrt/base.h>

#include <algorithm>
#include <atomic>
#include <climits>
#include <condition_variable>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <type_traits>
#include <unordered_map>
#include <unordered_set>
#include <vector>

enum class Mode {
    intersected,
    dock,
    maximized,
    fullscreen,
    never,
};

struct Settings {
    Mode mode = Mode::intersected;
    bool foregroundWindowOnly = false;
    std::unordered_set<std::wstring> excludedPrograms;
    bool primaryMonitorOnly = false;
    bool oldTaskbarOnWin11 = false;
};

struct DockCapture {
    HMONITOR monitor = nullptr;
    SIZE taskbarSize{};
    uint64_t generation = 0;
    bool physicallyShown = false;
    bool valid = false;
    std::vector<RECT> bands;
};

struct DockRegionCache {
    HMONITOR monitor = nullptr;
    SIZE taskbarSize{};
    uint64_t generation = 0;
    std::vector<RECT> bands;
    bool initialized = false;
    bool shownDerived = false;
    bool usable = false;
};

bool RectEqual(const RECT& left, const RECT& right) {
    return left.left == right.left && left.top == right.top &&
           left.right == right.right && left.bottom == right.bottom;
}

bool RectVectorsEqual(const std::vector<RECT>& left,
                      const std::vector<RECT>& right) {
    return left.size() == right.size() &&
           std::equal(left.begin(), left.end(), right.begin(), RectEqual);
}

// Fullscreen-aware reveal shape adapted from Taskbar Dock Click-Through by
// primez-x: https://github.com/ramensoftware/windhawk-mods/pull/4974
constexpr bool IsFullscreenWindowShapeForDockReveal(
    LONG_PTR style,
    UINT showCommand,
    const RECT& windowRect,
    const RECT& monitorRect) noexcept {
    const bool coversMonitor =
        windowRect.left <= monitorRect.left + 2 &&
        windowRect.top <= monitorRect.top + 2 &&
        windowRect.right >= monitorRect.right - 2 &&
        windowRect.bottom >= monitorRect.bottom - 2;
    if (!coversMonitor) {
        return false;
    }

    const bool normalMaximizedFrame =
        showCommand == SW_SHOWMAXIMIZED &&
        (style & WS_CAPTION) == WS_CAPTION &&
        (style & WS_THICKFRAME) == WS_THICKFRAME && !(style & WS_POPUP);
    if (normalMaximizedFrame) {
        return false;
    }

    return (style & WS_POPUP) || (style & WS_CAPTION) != WS_CAPTION ||
           (style & WS_THICKFRAME) != WS_THICKFRAME;
}

bool NormalizeDockRegionRect(const RECT& rawRect,
                             SIZE taskbarSize,
                             bool rightToLeft,
                             RECT* normalizedRect) noexcept {
    if (!normalizedRect || taskbarSize.cx <= 0 || taskbarSize.cy <= 0) {
        return false;
    }

    const int64_t width = taskbarSize.cx;
    const int64_t height = taskbarSize.cy;
    int64_t left = rawRect.left;
    int64_t top = rawRect.top;
    int64_t right = rawRect.right;
    int64_t bottom = rawRect.bottom;
    if (left < 0 || top < 0 || right > width || bottom > height ||
        left >= right || top >= bottom) {
        return false;
    }

    if (rightToLeft) {
        const int64_t mirroredLeft = width - right;
        right = width - left;
        left = mirroredLeft;
    }
    if (left < 0 || top < 0 || right > width || bottom > height ||
        left >= right || top >= bottom) {
        return false;
    }

    *normalizedRect = {
        static_cast<LONG>(left), static_cast<LONG>(top),
        static_cast<LONG>(right), static_cast<LONG>(bottom)};
    return true;
}

bool ParseDockRegionData(int regionType,
                         const void* buffer,
                         size_t bufferSize,
                         SIZE taskbarSize,
                         bool rightToLeft,
                         std::vector<RECT>* bands) noexcept {
    if (!bands) {
        return false;
    }
    bands->clear();
    try {
        if ((regionType != SIMPLEREGION && regionType != COMPLEXREGION) ||
            !buffer || bufferSize < sizeof(RGNDATAHEADER) ||
            taskbarSize.cx <= 0 || taskbarSize.cy <= 0) {
            return false;
        }

        const auto* data = static_cast<const RGNDATA*>(buffer);
        const auto& header = data->rdh;
        if (header.dwSize != sizeof(RGNDATAHEADER) ||
            header.iType != RDH_RECTANGLES || header.nCount == 0) {
            return false;
        }

        const uint64_t rectBytes =
            static_cast<uint64_t>(header.nCount) * sizeof(RECT);
        const uint64_t expectedSize =
            static_cast<uint64_t>(header.dwSize) + rectBytes;
        if (rectBytes != header.nRgnSize || expectedSize != bufferSize) {
            return false;
        }

        const auto* rects = reinterpret_cast<const RECT*>(data->Buffer);
        bands->reserve(header.nCount);
        for (DWORD i = 0; i < header.nCount; i++) {
            RECT rect{};
            if (!NormalizeDockRegionRect(rects[i], taskbarSize, rightToLeft,
                                         &rect)) {
                bands->clear();
                return false;
            }
            bands->push_back(rect);
        }

        std::sort(bands->begin(), bands->end(), [](const RECT& left,
                                                   const RECT& right) {
            if (left.top != right.top) {
                return left.top < right.top;
            }
            if (left.bottom != right.bottom) {
                return left.bottom < right.bottom;
            }
            if (left.left != right.left) {
                return left.left < right.left;
            }
            return left.right < right.right;
        });

        std::vector<RECT> normalized;
        normalized.reserve(bands->size());
        for (const RECT& rect : *bands) {
            if (!normalized.empty() && normalized.back().top == rect.top &&
                normalized.back().bottom == rect.bottom &&
                rect.left <= normalized.back().right) {
                normalized.back().right =
                    std::max(normalized.back().right, rect.right);
            } else {
                normalized.push_back(rect);
            }
        }
        *bands = std::move(normalized);
        return !bands->empty();
    } catch (...) {
        bands->clear();
        return false;
    }
}

bool WindowIntersectsDockBands(const RECT& windowRect,
                               const RECT& taskbarRect,
                               const std::vector<RECT>& bands) {
    const int64_t localLeft =
        static_cast<int64_t>(windowRect.left) - taskbarRect.left;
    const int64_t localTop =
        static_cast<int64_t>(windowRect.top) - taskbarRect.top;
    const int64_t localRight =
        static_cast<int64_t>(windowRect.right) - taskbarRect.left;
    const int64_t localBottom =
        static_cast<int64_t>(windowRect.bottom) - taskbarRect.top;
    for (const RECT& band : bands) {
        if (localLeft < band.right && band.left < localRight &&
            localTop < band.bottom && band.top < localBottom) {
            return true;
        }
    }
    return false;
}

DockRegionCache ApplyDockCapture(const DockRegionCache& previous,
                                 const DockCapture& capture) {
    DockRegionCache next;
    next.monitor = capture.monitor;
    next.taskbarSize = capture.taskbarSize;
    next.generation = capture.generation;
    next.initialized = true;

    if (capture.physicallyShown && capture.valid) {
        next.bands = capture.bands;
        next.shownDerived = true;
        next.usable = true;
        return next;
    }

    const bool sameIdentity =
        previous.initialized && previous.monitor == capture.monitor &&
        previous.taskbarSize.cx == capture.taskbarSize.cx &&
        previous.taskbarSize.cy == capture.taskbarSize.cy &&
        previous.generation == capture.generation;
    if (!capture.physicallyShown && capture.valid && sameIdentity &&
        previous.shownDerived &&
        RectVectorsEqual(previous.bands, capture.bands)) {
        next.bands = previous.bands;
        next.shownDerived = true;
        next.usable = true;
    }
    return next;
}

std::vector<RECT> EffectiveDockBands(const DockRegionCache& cache,
                                     SIZE taskbarSize) {
    if (cache.usable && !cache.bands.empty()) {
        return cache.bands;
    }
    if (taskbarSize.cx <= 0 || taskbarSize.cy <= 0) {
        return {};
    }
    return {{0, 0, taskbarSize.cx, taskbarSize.cy}};
}

bool DockPolicySemanticallyEqual(const DockRegionCache& left,
                                 const DockRegionCache& right) {
    if (left.usable != right.usable) {
        return false;
    }
    if (left.initialized && right.initialized &&
        (left.monitor != right.monitor ||
         left.taskbarSize.cx != right.taskbarSize.cx ||
         left.taskbarSize.cy != right.taskbarSize.cy)) {
        return false;
    }
    return !left.usable || RectVectorsEqual(left.bands, right.bands);
}

struct DockRefreshToken {
    uint64_t dockEpoch = 0;
    uint64_t generation = 0;

    bool operator==(const DockRefreshToken&) const = default;
};

struct DockRefreshPending {
    bool pending = false;
    DockRefreshToken token;
};

bool IsDockRefreshTokenCurrent(const DockRefreshToken& token,
                               uint64_t dockEpoch,
                               uint64_t generation) {
    return token.dockEpoch == dockEpoch && token.generation == generation;
}

bool BeginDockRefresh(DockRefreshPending& pending,
                      const DockRefreshToken& token) {
    if (pending.pending) {
        return false;
    }
    pending.pending = true;
    pending.token = token;
    return true;
}

bool CompleteDockRefresh(DockRefreshPending& pending,
                         const DockRefreshToken& token) {
    if (!pending.pending || pending.token != token) {
        return false;
    }
    pending = {};
    return true;
}

bool FailDockRefreshPost(DockRefreshPending& pending,
                         const DockRefreshToken& token) {
    return CompleteDockRefresh(pending, token);
}

uintptr_t EncodeEligibilitySettingsEpoch(uint64_t epoch) {
    uintptr_t encoded = static_cast<uintptr_t>(epoch);
    return encoded ? encoded : 1;
}

bool EligibilitySettingsEpochMatches(uintptr_t encoded, uint64_t epoch) {
    return encoded == EncodeEligibilitySettingsEpoch(epoch);
}

uintptr_t EncodeEligibilityLoadNonce(uintptr_t nonce) {
    return nonce ? nonce : 1;
}

bool EligibilityLoadNonceMatches(uintptr_t encoded, uintptr_t nonce) {
    return encoded == EncodeEligibilityLoadNonce(nonce);
}

uintptr_t GenerateEligibilityLoadNonce() noexcept {
    GUID guid{};
    uintptr_t nonce = 0;
    if (SUCCEEDED(CoCreateGuid(&guid))) {
        nonce = guid.Data1 ^
                (static_cast<uintptr_t>(guid.Data2) << 16) ^ guid.Data3;
        for (BYTE value : guid.Data4) {
            nonce = (nonce << 5) ^ (nonce >> 2) ^ value;
        }
    }
    LARGE_INTEGER performanceCounter{};
    if (QueryPerformanceCounter(&performanceCounter)) {
        nonce ^= static_cast<uintptr_t>(performanceCounter.QuadPart);
    }
    nonce ^= static_cast<uintptr_t>(GetTickCount64());
    nonce ^= static_cast<uintptr_t>(GetCurrentProcessId()) << 17;
    return EncodeEligibilityLoadNonce(nonce);
}

bool IsDockRefreshMessage(UINT registeredMessage, UINT message) {
    return registeredMessage != 0 && message == registeredMessage;
}

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_24H2,
};

enum class TaskbarBackend {
    NativeModern,
    NativeLegacy,
    ExplorerPatcherLegacy,
};

enum class ExplorerPatcherInstallState {
    Unseen,
    Installing,
    RegisteredPendingInitialApply,
    Active,
    FailedTerminal,
};

enum class TaskbarViewInstallState {
    Unseen,
    Installing,
    RegisteredPendingInitialApply,
    Active,
    FailedTerminal,
};

enum class DockRevealPhysicalState {
    Unknown,
    Hidden,
    Shown,
};

struct DockRevealRepairToken {
    uint64_t settingsEpoch = 0;
    uint64_t dockEpoch = 0;
    uint64_t backendEpoch = 0;
    uint64_t generation = 0;
    uint64_t request = 0;

    bool operator==(const DockRevealRepairToken&) const = default;
};

struct DockRevealEnrollment {
    uint64_t settingsEpoch = 0;
    uint64_t dockEpoch = 0;
    uint64_t backendEpoch = 0;
    uint64_t generation = 0;
    HMONITOR monitor = nullptr;
    RECT shownRect{};
    bool secondary = false;
    bool scopeApplicable = false;
    DockRevealPhysicalState physicalState = DockRevealPhysicalState::Unknown;
};

struct DockRevealRepairPending {
    DockRevealRepairToken token;
};

struct DockRevealDeferredDelivery {
    DockRevealRepairToken token;
    DWORD threadId = 0;
};

enum class InitialHookQueueState {
    AwaitingApply,
    Applying,
    Completed,
    FailedTerminal,
};

std::atomic<WinVersion> g_actualWinVersion;
TaskbarBackend g_taskbarBackend = TaskbarBackend::NativeLegacy;

std::atomic<bool> g_initialized;
std::atomic<ExplorerPatcherInstallState> g_explorerPatcherInstallState;
std::atomic<TaskbarViewInstallState> g_taskbarViewInstallState;
std::atomic<InitialHookQueueState> g_initialHookQueueState{
    InitialHookQueueState::AwaitingApply};

bool g_wasAutoHideProcessed;
bool g_wasAutoHideDisabled;
std::mutex g_stateMutex;
std::mutex g_hookOperationsMutex;
std::mutex g_policyCallbackMutex;
std::condition_variable g_policyCallbackCondition;
bool g_acceptPolicyCallbacks;
bool g_policyCallbackTeardownStarted;
size_t g_activePolicyCallbacks;
std::shared_ptr<const Settings> g_settingsSnapshot =
    std::make_shared<const Settings>();
uint64_t g_settingsEpoch = 1;
uintptr_t g_eligibilityLoadNonce;
uint64_t g_dockEpoch = 1;
uint64_t g_taskbarBackendEpoch = 1;
uint64_t g_nextTaskbarGeneration = 1;
uint64_t g_nextDockRevealRepairRequest = 1;
bool g_acceptDockWork;
bool g_fullPolicyReevaluationPending;
uint64_t g_fullPolicyReevaluationToken = 1;
std::unordered_map<HWND, uint64_t> g_taskbarGenerations;
std::unordered_map<HWND, DockRegionCache> g_dockRegionCaches;
std::unordered_map<HWND, DockRefreshPending> g_pendingDockRefreshes;
std::unordered_map<HWND, DockRevealEnrollment> g_dockRevealEnrollments;
std::unordered_map<HWND, DockRevealRepairPending>
    g_pendingDockRevealRepairs;
std::unordered_map<HWND, DockRevealRepairToken> g_dockRevealRepairsInProgress;
std::unordered_map<HWND, DockRevealDeferredDelivery>
    g_deferredDockRevealRepairs;
std::unordered_map<void*, HWND> g_taskbarsKeptShown;
std::unordered_map<HWND, void*> g_taskbarToViewCoordinator;
UINT_PTR g_pendingEventsTimer;
std::atomic<HWND> g_multitaskingViewHwnd;
std::atomic<HWND> g_altTabViewHwnd;

void ClearDockTaskbarRevealLifecycleLocked(HWND hWnd);
void ClearAllDockTaskbarRevealStateLocked();
void FlushDeferredDockRevealRepairsForCurrentThread() noexcept;

// Keep TLS trivial: Explorer's UI thread can outlive the mod image, so a
// thread_local container with a destructor would be unsafe at thread exit.
thread_local unsigned int g_taskbarMutationDeferralDepth;
thread_local bool g_taskbarMutationDeferralHasDeferredDelivery;

class TaskbarMutationDeferralScope {
   public:
    TaskbarMutationDeferralScope() noexcept {
        if (g_taskbarMutationDeferralDepth != UINT_MAX) {
            g_taskbarMutationDeferralDepth++;
            entered_ = true;
        }
    }

    ~TaskbarMutationDeferralScope() {
        if (!entered_) {
            return;
        }
        if (--g_taskbarMutationDeferralDepth == 0 &&
            g_taskbarMutationDeferralHasDeferredDelivery) {
            FlushDeferredDockRevealRepairsForCurrentThread();
        }
    }

    TaskbarMutationDeferralScope(const TaskbarMutationDeferralScope&) =
        delete;
    TaskbarMutationDeferralScope& operator=(
        const TaskbarMutationDeferralScope&) = delete;

   private:
    bool entered_ = false;
};

struct SettingsStateSnapshot {
    std::shared_ptr<const Settings> settings;
    uint64_t settingsEpoch = 0;
    uintptr_t eligibilityLoadNonce = 0;
    TaskbarBackend taskbarBackend = TaskbarBackend::NativeLegacy;
};

SettingsStateSnapshot GetSettingsStateSnapshot() {
    std::lock_guard<std::mutex> lock(g_stateMutex);
    return {g_settingsSnapshot, g_settingsEpoch, g_eligibilityLoadNonce,
            g_taskbarBackend};
}

bool EligibilityPublicationAllowed(uint64_t settingsEpoch,
                                   uintptr_t loadNonce) {
    if (!g_initialized.load(std::memory_order_acquire)) {
        return false;
    }
    std::lock_guard<std::mutex> lock(g_stateMutex);
    return g_initialized.load(std::memory_order_relaxed) &&
           settingsEpoch == g_settingsEpoch &&
           loadNonce == g_eligibilityLoadNonce;
}

#ifdef WH_EDITING
using PolicyCallbackEnteredTestHook = void (*)();
PolicyCallbackEnteredTestHook g_policyCallbackEnteredTestHook;
using BackendActiveCheckCompletedTestHook = void (*)();
BackendActiveCheckCompletedTestHook g_backendActiveCheckCompletedTestHook;
using InitialHookFailureBeforePolicyDrainTestHook = void (*)();
InitialHookFailureBeforePolicyDrainTestHook
    g_initialHookFailureBeforePolicyDrainTestHook;
#endif

void InvokePolicyCallbackEnteredTestHook() noexcept {
#ifdef WH_EDITING
    try {
        if (g_policyCallbackEnteredTestHook) {
            g_policyCallbackEnteredTestHook();
        }
    } catch (...) {
    }
#endif
}

void InvokeBackendActiveCheckCompletedTestHook() noexcept {
#ifdef WH_EDITING
    try {
        if (g_backendActiveCheckCompletedTestHook) {
            g_backendActiveCheckCompletedTestHook();
        }
    } catch (...) {
    }
#endif
}

void InvokeInitialHookFailureBeforePolicyDrainTestHook() noexcept {
#ifdef WH_EDITING
    try {
        if (g_initialHookFailureBeforePolicyDrainTestHook) {
            g_initialHookFailureBeforePolicyDrainTestHook();
        }
    } catch (...) {
    }
#endif
}

class PolicyCallbackScope {
   public:
    PolicyCallbackScope() {
        {
            std::lock_guard<std::mutex> lock(g_policyCallbackMutex);
            if (g_acceptPolicyCallbacks) {
                g_activePolicyCallbacks++;
                entered_ = true;
            }
        }
        if (entered_) {
            InvokePolicyCallbackEnteredTestHook();
        }
    }

    ~PolicyCallbackScope() {
        if (!entered_) {
            return;
        }
        std::lock_guard<std::mutex> lock(g_policyCallbackMutex);
        g_activePolicyCallbacks--;
        if (!g_activePolicyCallbacks) {
            g_policyCallbackCondition.notify_all();
        }
    }

    explicit operator bool() const {
        return entered_;
    }

    PolicyCallbackScope(const PolicyCallbackScope&) = delete;
    PolicyCallbackScope& operator=(const PolicyCallbackScope&) = delete;

   private:
    bool entered_ = false;
};

void SetPolicyCallbacksAccepted(bool accepted) {
    std::lock_guard<std::mutex> lock(g_policyCallbackMutex);
    g_acceptPolicyCallbacks = accepted;
}

void ResetPolicyCallbackTeardown() {
    std::lock_guard<std::mutex> lock(g_policyCallbackMutex);
    g_policyCallbackTeardownStarted = false;
}

void ClosePolicyCallbacksForTeardown() {
    std::lock_guard<std::mutex> lock(g_policyCallbackMutex);
    g_policyCallbackTeardownStarted = true;
    g_acceptPolicyCallbacks = false;
}

void WaitForPolicyCallbacks() {
    std::unique_lock<std::mutex> lock(g_policyCallbackMutex);
    g_policyCallbackCondition.wait(lock,
                                   [] { return !g_activePolicyCallbacks; });
}

constexpr unsigned kHookRescanExplorerPatcher = 1u << 0;
constexpr unsigned kHookRescanTaskbarView = 1u << 1;
constexpr DWORD kHookRescanInitialRetryDelayMs = 50;
constexpr DWORD kHookRescanMaximumRetryDelayMs = 1000;

struct HookRescanWorkerOperations {
    void* context = nullptr;
    void (*process)(void* context, unsigned bits) = nullptr;
    BOOL (*signal)(void* context, HANDLE workEvent) = nullptr;
    void (*processFailed)(void* context) = nullptr;
    BOOL (*signalStop)(void* context, HANDLE stopEvent) = nullptr;
    DWORD (*waitThread)(void* context, HANDLE thread, DWORD timeout) = nullptr;
    DWORD (*waitRetry)(void* context, HANDLE stopEvent, DWORD timeout) =
        nullptr;
};

class HookRescanWorkerLifecycle {
   public:
    HookRescanWorkerLifecycle() = default;
    HookRescanWorkerLifecycle(const HookRescanWorkerLifecycle&) = delete;
    HookRescanWorkerLifecycle& operator=(const HookRescanWorkerLifecycle&) =
        delete;

    ~HookRescanWorkerLifecycle() {
        StopAndDrain();
    }

    bool Start(const HookRescanWorkerOperations& operations) noexcept {
        try {
            std::lock_guard<std::mutex> lock(lifecycleMutex_);
            if (thread_) {
                return accepting_.load(std::memory_order_acquire);
            }
            if (!operations.process || !operations.signal) {
                return false;
            }

            workEvent_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            stopEvent_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            if (!workEvent_ || !stopEvent_) {
                CloseHandlesLocked();
                return false;
            }

            operations_ = operations;
            pendingBits_.store(0, std::memory_order_relaxed);
            activeSchedulers_.store(0, std::memory_order_relaxed);
            thread_ = CreateThread(nullptr, 0, ThreadEntry, this, 0, nullptr);
            if (!thread_) {
                operations_ = {};
                CloseHandlesLocked();
                return false;
            }
            accepting_.store(true, std::memory_order_release);
            return true;
        } catch (...) {
            return false;
        }
    }

    bool Request(unsigned bits) noexcept {
        activeSchedulers_.fetch_add(1, std::memory_order_acq_rel);
        if (!accepting_.load(std::memory_order_acquire)) {
            FinishSchedule();
            return false;
        }

        if (bits) {
            pendingBits_.fetch_or(bits, std::memory_order_acq_rel);
        } else if (!pendingBits_.load(std::memory_order_acquire)) {
            FinishSchedule();
            return true;
        }

        BOOL signaled = FALSE;
        try {
            signaled = operations_.signal(operations_.context, workEvent_);
        } catch (...) {
        }
        FinishSchedule();
        return signaled;
    }

    bool RetryPending() noexcept {
        return Request(0);
    }

    void CloseAdmission() noexcept {
        accepting_.store(false, std::memory_order_release);
    }

    bool StopAndDrain() noexcept {
        CloseAdmission();
        try {
            {
                std::unique_lock<std::mutex> schedulerLock(
                    schedulerDrainMutex_);
                schedulerDrainCondition_.wait(schedulerLock, [&] {
                    return !activeSchedulers_.load(
                        std::memory_order_acquire);
                });
            }

            std::lock_guard<std::mutex> lock(lifecycleMutex_);
            if (!thread_) {
                return true;
            }
            BOOL stopSignaled = operations_.signalStop
                                    ? operations_.signalStop(
                                          operations_.context, stopEvent_)
                                    : SetEvent(stopEvent_);
            if (!stopSignaled) {
                Wh_Log(L"Error: Couldn't signal the hook-rescan worker stop");
                return false;
            }

            auto waitThread = [&](DWORD timeout) {
                return operations_.waitThread
                           ? operations_.waitThread(operations_.context,
                                                    thread_, timeout)
                           : WaitForSingleObject(thread_, timeout);
            };
            DWORD waitResult = waitThread(5000);
            if (waitResult == WAIT_TIMEOUT) {
                Wh_Log(L"Diagnostic: Hook-rescan worker stop exceeded five "
                       L"seconds");
                waitResult = waitThread(INFINITE);
            }
            if (waitResult != WAIT_OBJECT_0) {
                Wh_Log(L"Error: Hook-rescan worker exit wasn't confirmed: "
                       L"result=%u error=%u",
                       waitResult,
                       waitResult == WAIT_FAILED ? GetLastError() : 0);
                return false;
            }
            CloseHandle(thread_);
            thread_ = nullptr;
            operations_ = {};
            pendingBits_.store(0, std::memory_order_release);
            CloseHandlesLocked();
            return true;
        } catch (...) {
            Wh_Log(L"Error: Couldn't drain the hook-rescan worker");
            return false;
        }
    }

   private:
    static DWORD WINAPI ThreadEntry(void* context) noexcept {
        static_cast<HookRescanWorkerLifecycle*>(context)->Run();
        return 0;
    }

    void Run() noexcept {
        HANDLE waits[]{stopEvent_, workEvent_};
        DWORD retryDelay = kHookRescanInitialRetryDelayMs;
        for (;;) {
            DWORD waitResult =
                WaitForMultipleObjects(ARRAYSIZE(waits), waits, FALSE, INFINITE);
            if (waitResult == WAIT_OBJECT_0) {
                return;
            }
            if (waitResult != WAIT_OBJECT_0 + 1) {
                Wh_Log(L"Error: Hook-rescan worker wait failed: %u",
                       GetLastError());
                return;
            }

            for (;;) {
                ResetEvent(workEvent_);
                if (!accepting_.load(std::memory_order_acquire)) {
                    break;
                }
                unsigned bits =
                    pendingBits_.exchange(0, std::memory_order_acq_rel);
                if (!bits) {
                    break;
                }
                try {
                    operations_.process(operations_.context, bits);
                    retryDelay = kHookRescanInitialRetryDelayMs;
                } catch (...) {
                    pendingBits_.fetch_or(bits, std::memory_order_acq_rel);
                    if (operations_.processFailed) {
                        try {
                            operations_.processFailed(operations_.context);
                        } catch (...) {
                        }
                    }
                    Wh_Log(L"Error: Deferred hook rescan failed");
                    DWORD retryWaitResult = WAIT_FAILED;
                    try {
                        retryWaitResult = operations_.waitRetry
                                              ? operations_.waitRetry(
                                                    operations_.context,
                                                    stopEvent_, retryDelay)
                                              : WaitForSingleObject(
                                                    stopEvent_, retryDelay);
                    } catch (...) {
                    }
                    if (retryWaitResult == WAIT_OBJECT_0) {
                        return;
                    }
                    if (retryWaitResult != WAIT_TIMEOUT) {
                        Wh_Log(L"Error: Hook-rescan retry wait failed: %u",
                               retryWaitResult == WAIT_FAILED ? GetLastError()
                                                               : 0);
                        return;
                    }
                    retryDelay =
                        retryDelay < kHookRescanMaximumRetryDelayMs
                            ? std::min<DWORD>(retryDelay * 2,
                                              kHookRescanMaximumRetryDelayMs)
                            : kHookRescanMaximumRetryDelayMs;
                    continue;
                }
                if (!pendingBits_.load(std::memory_order_acquire)) {
                    break;
                }
            }
        }
    }

    void FinishSchedule() noexcept {
        if (activeSchedulers_.fetch_sub(1, std::memory_order_acq_rel) == 1) {
            try {
                std::lock_guard<std::mutex> lock(schedulerDrainMutex_);
                schedulerDrainCondition_.notify_all();
            } catch (...) {
            }
        }
    }

    void CloseHandlesLocked() noexcept {
        if (stopEvent_) {
            CloseHandle(stopEvent_);
            stopEvent_ = nullptr;
        }
        if (workEvent_) {
            CloseHandle(workEvent_);
            workEvent_ = nullptr;
        }
    }

    std::mutex lifecycleMutex_;
    std::atomic<bool> accepting_ = false;
    std::atomic<unsigned> pendingBits_ = 0;
    std::atomic<size_t> activeSchedulers_ = 0;
    std::mutex schedulerDrainMutex_;
    std::condition_variable schedulerDrainCondition_;
    HookRescanWorkerOperations operations_;
    HANDLE workEvent_ = nullptr;
    HANDLE stopEvent_ = nullptr;
    HANDLE thread_ = nullptr;
};

HookRescanWorkerLifecycle g_hookRescanWorker;

void StopHookRescanWorkerConfirmed() noexcept {
    while (!g_hookRescanWorker.StopAndDrain()) {
        Wh_Log(L"Error: Refusing teardown without a confirmed hook-rescan "
               L"worker exit");
        Sleep(100);
    }
}

enum class WinEventThreadState {
    Stopped,
    Starting,
    CancelingStart,
    Running,
    Stopping,
};

bool WaitForConfirmedSignal(HANDLE handle, PCWSTR label) {
    if (!handle) {
        Wh_Log(L"Error: %s has no wait handle", label);
        return false;
    }

    DWORD waitResult = WaitForSingleObject(handle, 5000);
    if (waitResult == WAIT_TIMEOUT) {
        Wh_Log(L"Diagnostic: %s exceeded five seconds", label);
        waitResult = WaitForSingleObject(handle, INFINITE);
    }
    if (waitResult != WAIT_OBJECT_0) {
        DWORD error = waitResult == WAIT_FAILED ? GetLastError() : ERROR_SUCCESS;
        Wh_Log(L"Error: %s wait was not confirmed: result=%u error=%u", label,
               waitResult, error);
        return false;
    }
    return true;
}

class WinEventThreadContext {
   public:
    WinEventThreadContext(uint64_t settingsEpoch, bool foregroundWindowOnly)
        : settingsEpoch_(settingsEpoch),
          foregroundWindowOnly_(foregroundWindowOnly),
          stopEvent_(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          readyEvent_(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          startupDone_(CreateEventW(nullptr, TRUE, FALSE, nullptr)),
          stopDone_(CreateEventW(nullptr, TRUE, FALSE, nullptr)) {}

    ~WinEventThreadContext() {
        if (thread_) {
            CloseHandle(thread_);
        }
        if (stopDone_) {
            CloseHandle(stopDone_);
        }
        if (startupDone_) {
            CloseHandle(startupDone_);
        }
        if (readyEvent_) {
            CloseHandle(readyEvent_);
        }
        if (stopEvent_) {
            CloseHandle(stopEvent_);
        }
    }

    bool IsValid() const {
        return stopEvent_ && readyEvent_ && startupDone_ && stopDone_;
    }

    HANDLE StopEvent() const {
        return stopEvent_;
    }

    void SignalReady() const {
        SetEvent(readyEvent_);
    }

    bool ForegroundWindowOnly() const {
        return foregroundWindowOnly_;
    }

    uint64_t SettingsEpoch() const {
        return settingsEpoch_;
    }

   private:
    friend class WinEventThreadLifecycle;

    uint64_t settingsEpoch_;
    bool foregroundWindowOnly_;
    HANDLE stopEvent_ = nullptr;
    HANDLE readyEvent_ = nullptr;
    HANDLE startupDone_ = nullptr;
    HANDLE stopDone_ = nullptr;
    HANDLE thread_ = nullptr;
};

using WinEventThreadWorker = DWORD(WINAPI*)(WinEventThreadContext* context);

#ifdef WH_EDITING
using WinEventStartWaitEnteredTestHook = void (*)();
WinEventStartWaitEnteredTestHook g_winEventStartWaitEnteredTestHook;
#endif

void InvokeWinEventStartWaitEnteredTestHook() noexcept {
#ifdef WH_EDITING
    try {
        if (g_winEventStartWaitEnteredTestHook) {
            g_winEventStartWaitEnteredTestHook();
        }
    } catch (...) {
    }
#endif
}

class WinEventThreadLifecycle {
   public:
    WinEventThreadLifecycle() = default;
    WinEventThreadLifecycle(const WinEventThreadLifecycle&) = delete;
    WinEventThreadLifecycle& operator=(const WinEventThreadLifecycle&) = delete;

    ~WinEventThreadLifecycle() {
        SetAdmission(false, 0);
        Stop();
    }

    void SetAdmission(bool enabled, uint64_t settingsEpoch) {
        std::lock_guard<std::mutex> lock(mutex_);
        admissionEnabled_ = enabled;
        admittedSettingsEpoch_ = settingsEpoch;
    }

    bool Start(WinEventThreadWorker worker,
               uint64_t settingsEpoch,
               bool foregroundWindowOnly) {
        std::shared_ptr<WinEventThreadContext> context;
        HANDLE phaseEvent = nullptr;
        bool createThread = false;
        bool joiningStartupPhase = false;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            if (!AdmissionMatchesLocked(settingsEpoch)) {
                return false;
            }
            switch (state_) {
                case WinEventThreadState::Stopped:
                    try {
                        context = std::make_shared<WinEventThreadContext>(
                            settingsEpoch, foregroundWindowOnly);
                    } catch (...) {
                        return false;
                    }
                    if (!context->IsValid()) {
                        return false;
                    }
                    context_ = context;
                    state_ = WinEventThreadState::Starting;
                    createThread = true;
                    break;
                case WinEventThreadState::Starting:
                case WinEventThreadState::CancelingStart:
                    context = context_;
                    if (!ContextMatches(context, settingsEpoch,
                                        foregroundWindowOnly)) {
                        return false;
                    }
                    phaseEvent = context->startupDone_;
                    joiningStartupPhase = true;
                    break;
                case WinEventThreadState::Running:
                    return ContextMatches(context_, settingsEpoch,
                                          foregroundWindowOnly);
                case WinEventThreadState::Stopping:
                    context = context_;
                    phaseEvent = context->stopDone_;
                    break;
            }
        }

        if (!createThread) {
            if (joiningStartupPhase) {
                InvokeWinEventStartWaitEnteredTestHook();
            }
            WaitUntilConfirmed(phaseEvent, L"WinEvent lifecycle phase");
            std::lock_guard<std::mutex> lock(mutex_);
            return AdmissionMatchesLocked(settingsEpoch) &&
                   state_ == WinEventThreadState::Running &&
                   ContextMatches(context_, settingsEpoch,
                                  foregroundWindowOnly);
        }

        ThreadStartParameter* parameter = nullptr;
        try {
            parameter = new ThreadStartParameter{context, worker};
        } catch (...) {
        }
        HANDLE thread =
            parameter
                ? CreateThread(nullptr, 0, ThreadEntry, parameter, 0, nullptr)
                : nullptr;
        if (!thread) {
            delete parameter;
            std::lock_guard<std::mutex> lock(mutex_);
            if (context_ == context) {
                state_ = WinEventThreadState::Stopped;
                SetEvent(context->startupDone_);
                context_.reset();
            }
            return false;
        }

        {
            std::lock_guard<std::mutex> lock(mutex_);
            context->thread_ = thread;
            if (state_ == WinEventThreadState::CancelingStart ||
                !AdmissionMatchesLocked(settingsEpoch)) {
                if (state_ == WinEventThreadState::Starting) {
                    state_ = WinEventThreadState::CancelingStart;
                }
                SetEvent(context->stopEvent_);
            }
        }

        HANDLE startupWaits[]{context->readyEvent_, context->thread_};
        DWORD waitResult =
            WaitForMultipleObjects(2, startupWaits, FALSE, 5000);
        if (waitResult == WAIT_TIMEOUT) {
            Wh_Log(L"Diagnostic: WinEvent startup exceeded five seconds");
            waitResult =
                WaitForMultipleObjects(2, startupWaits, FALSE, INFINITE);
        }
        if (waitResult == WAIT_FAILED) {
            Wh_Log(L"Error: WinEvent startup wait failed: error=%u",
                   GetLastError());
        }

        bool running = false;
        bool cleanup = false;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            if (context_ == context) {
                if (state_ == WinEventThreadState::Starting &&
                    waitResult == WAIT_OBJECT_0 &&
                    AdmissionMatchesLocked(settingsEpoch) &&
                    ContextMatches(context, settingsEpoch,
                                   foregroundWindowOnly)) {
                    state_ = WinEventThreadState::Running;
                    SetEvent(context->startupDone_);
                    running = true;
                } else {
                    if (state_ == WinEventThreadState::Starting) {
                        state_ = WinEventThreadState::CancelingStart;
                        SetEvent(context->stopEvent_);
                    }
                    cleanup = true;
                }
            }
        }

        if (running) {
            return true;
        }
        if (cleanup) {
            WaitUntilConfirmed(context->thread_,
                               L"WinEvent canceled startup exit");
            std::lock_guard<std::mutex> lock(mutex_);
            if (context_ == context) {
                state_ = WinEventThreadState::Stopped;
                SetEvent(context->startupDone_);
                context_.reset();
            }
        }
        return false;
    }

    void Stop() {
        for (;;) {
            std::shared_ptr<WinEventThreadContext> context;
            HANDLE phaseEvent = nullptr;
            bool ownsRunningStop = false;
            {
                std::lock_guard<std::mutex> lock(mutex_);
                switch (state_) {
                    case WinEventThreadState::Stopped:
                        return;
                    case WinEventThreadState::Starting:
                        state_ = WinEventThreadState::CancelingStart;
                        context = context_;
                        SetEvent(context->stopEvent_);
                        phaseEvent = context->startupDone_;
                        break;
                    case WinEventThreadState::CancelingStart:
                        context = context_;
                        phaseEvent = context->startupDone_;
                        break;
                    case WinEventThreadState::Running:
                        state_ = WinEventThreadState::Stopping;
                        context = context_;
                        SetEvent(context->stopEvent_);
                        ownsRunningStop = true;
                        break;
                    case WinEventThreadState::Stopping:
                        context = context_;
                        phaseEvent = context->stopDone_;
                        break;
                }
            }

            if (ownsRunningStop) {
                WaitUntilConfirmed(context->thread_, L"WinEvent thread exit");
                std::lock_guard<std::mutex> lock(mutex_);
                if (context_ == context &&
                    state_ == WinEventThreadState::Stopping) {
                    state_ = WinEventThreadState::Stopped;
                    SetEvent(context->stopDone_);
                    context_.reset();
                }
                return;
            }

            WaitUntilConfirmed(phaseEvent, L"WinEvent stop barrier");
        }
    }

    WinEventThreadState State() const {
        std::lock_guard<std::mutex> lock(mutex_);
        return state_;
    }

   private:
    struct ThreadStartParameter {
        std::shared_ptr<WinEventThreadContext> context;
        WinEventThreadWorker worker;
    };

    static DWORD WINAPI ThreadEntry(void* rawParameter) {
        std::unique_ptr<ThreadStartParameter> parameter(
            static_cast<ThreadStartParameter*>(rawParameter));
        return parameter->worker(parameter->context.get());
    }

    static bool ContextMatches(
        const std::shared_ptr<WinEventThreadContext>& context,
        uint64_t settingsEpoch,
        bool foregroundWindowOnly) {
        return context && context->SettingsEpoch() == settingsEpoch &&
               context->ForegroundWindowOnly() == foregroundWindowOnly;
    }

    bool AdmissionMatchesLocked(uint64_t settingsEpoch) const {
        return admissionEnabled_ && admittedSettingsEpoch_ == settingsEpoch;
    }

    static void WaitUntilConfirmed(HANDLE handle, PCWSTR label) {
        while (!WaitForConfirmedSignal(handle, label)) {
            Wh_Log(L"Error: refusing lifecycle transition without confirmed "
                   L"%s completion",
                   label);
            Sleep(100);
        }
    }

    mutable std::mutex mutex_;
    WinEventThreadState state_ = WinEventThreadState::Stopped;
    std::shared_ptr<WinEventThreadContext> context_;
    bool admissionEnabled_ = false;
    uint64_t admittedSettingsEpoch_ = 0;
};

WinEventThreadLifecycle g_winEventThreadLifecycle;

// TrayUI::_HandleTrayPrivateSettingMessage
constexpr UINT kHandleTrayPrivateSettingMessage = WM_USER + 0x1CA;

enum {
    kTrayPrivateSettingAutoHideGet = 3,
    kTrayPrivateSettingAutoHideSet = 4,
};

constexpr WCHAR kCanHideTaskbarEligibilityProp[] =
    L"Windhawk_CanHideTaskbar_" WH_MOD_ID;

constexpr WCHAR kCanHideTaskbarEligibilityEpochProp[] =
    L"Windhawk_CanHideTaskbarSettingsEpoch_" WH_MOD_ID;

constexpr WCHAR kCanHideTaskbarEligibilityNonceProp[] =
    L"Windhawk_CanHideTaskbarLoadNonce_" WH_MOD_ID;

const HANDLE kCanHideTaskbarNotEligible = (HANDLE)1;
const HANDLE kCanHideTaskbarEligible = (HANDLE)2;

constexpr WCHAR kUpdateTaskbarStatePendingTickCount[] =
    L"Windhawk_UpdateTaskbarStatePendingTickCount_" WH_MOD_ID;

static const UINT g_getTaskbarRectRegisteredMsg =
    RegisterWindowMessage(L"Windhawk_GetTaskbarRect_" WH_MOD_ID);

static const UINT g_updateTaskbarStateRegisteredMsg =
    RegisterWindowMessage(L"Windhawk_UpdateTaskbarState_" WH_MOD_ID);

static const UINT g_dockRegionRefreshRegisteredMsg =
    RegisterWindowMessage(L"Windhawk_DockRegionRefresh_" WH_MOD_ID);

static const UINT g_dockRevealRepairRegisteredMsg =
    RegisterWindowMessage(L"Windhawk_DockRevealRepair_" WH_MOD_ID);

enum {
    kTrayUITimerHide = 2,
    kTrayUITimerUnhide = 3,
};

#if __cplusplus < 202302L
// Missing in older MinGW headers.
DECLARE_HANDLE(CO_MTA_USAGE_COOKIE);
WINOLEAPI CoIncrementMTAUsage(CO_MTA_USAGE_COOKIE* pCookie);
WINOLEAPI CoDecrementMTAUsage(CO_MTA_USAGE_COOKIE Cookie);
#endif

// Missing in older MinGW headers.
#ifndef EVENT_OBJECT_CLOAKED
#define EVENT_OBJECT_CLOAKED 0x8017
#endif
#ifndef EVENT_OBJECT_UNCLOAKED
#define EVENT_OBJECT_UNCLOAKED 0x8018
#endif

using IsWindowArranged_t = BOOL(WINAPI*)(HWND hwnd);
IsWindowArranged_t pIsWindowArranged;

// Private API for window band (z-order band).
// https://blog.adeltax.com/window-z-order-in-windows-10/
using GetWindowBand_t = BOOL(WINAPI*)(HWND hWnd, PDWORD pdwBand);
GetWindowBand_t pGetWindowBand;

// https://devblogs.microsoft.com/oldnewthing/20200302-00/?p=103507
bool IsWindowCloaked(HWND hwnd) {
    BOOL isCloaked = FALSE;
    return SUCCEEDED(DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, &isCloaked,
                                           sizeof(isCloaked))) &&
           isCloaked;
}

enum class MultitaskingViewType {
    None,
    WinTab,
    AltTab,
};

// Detects Alt+Tab or Win+Tab and returns the type.
MultitaskingViewType GetMultitaskingViewType(HWND hWnd) {
    // Must be in the current process (explorer.exe).
    DWORD dwProcessId = 0;
    DWORD dwThreadId = GetWindowThreadProcessId(hWnd, &dwProcessId);
    if (!dwThreadId || dwProcessId != GetCurrentProcessId()) {
        return MultitaskingViewType::None;
    }

    WCHAR className[64];
    if (!GetClassName(hWnd, className, ARRAYSIZE(className)) ||
        _wcsicmp(className, L"XamlExplorerHostIslandWindow") != 0) {
        return MultitaskingViewType::None;
    }

    // The Win+Tab window uses band ZBID_IMMERSIVE_APPCHROME.
    constexpr DWORD ZBID_IMMERSIVE_APPCHROME = 5;

    // The Alt+Tab window uses band ZBID_SYSTEM_TOOLS. The virtual desktop
    // switcher (which we don't want) uses band ZBID_IMMERSIVE_EDGY.
    constexpr DWORD ZBID_SYSTEM_TOOLS = 16;

    DWORD band = 0;
    if (!pGetWindowBand || !pGetWindowBand(hWnd, &band)) {
        return MultitaskingViewType::None;
    }

    MultitaskingViewType type;
    if (band == ZBID_IMMERSIVE_APPCHROME) {
        type = MultitaskingViewType::WinTab;
    } else if (band == ZBID_SYSTEM_TOOLS) {
        type = MultitaskingViewType::AltTab;
    } else {
        return MultitaskingViewType::None;
    }

    // Check thread description for "MultitaskingView".
    HANDLE hThread =
        OpenThread(THREAD_QUERY_LIMITED_INFORMATION, FALSE, dwThreadId);
    if (!hThread) {
        return MultitaskingViewType::None;
    }

    bool isMultitaskingView = false;
    PWSTR description = nullptr;
    if (SUCCEEDED(GetThreadDescription(hThread, &description)) && description) {
        isMultitaskingView = wcscmp(description, L"MultitaskingView") == 0;
        LocalFree(description);
    }

    CloseHandle(hThread);
    return isMultitaskingView ? type : MultitaskingViewType::None;
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

bool IsTaskbarWindow(HWND hWnd) {
    WCHAR szClassName[32];
    if (!GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName))) {
        return false;
    }

    return _wcsicmp(szClassName, L"Shell_TrayWnd") == 0 ||
           _wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0;
}

HWND FindTaskbarWindows(std::unordered_set<HWND>* secondaryTaskbarWindows) {
    secondaryTaskbarWindows->clear();

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (!hTaskbarWnd) {
        return nullptr;
    }

    DWORD taskbarThreadId = GetWindowThreadProcessId(hTaskbarWnd, nullptr);
    if (!taskbarThreadId) {
        return nullptr;
    }

    auto enumWindowsProc = [&secondaryTaskbarWindows](HWND hWnd) -> BOOL {
        WCHAR szClassName[32];
        if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
            return TRUE;
        }

        if (_wcsicmp(szClassName, L"Shell_SecondaryTrayWnd") == 0) {
            secondaryTaskbarWindows->insert(hWnd);
        }

        return TRUE;
    };

    EnumThreadWindows(
        taskbarThreadId,
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(enumWindowsProc)*>(lParam);
            return proc(hWnd);
        },
        reinterpret_cast<LPARAM>(&enumWindowsProc));

    return hTaskbarWnd;
}

bool GetTaskbarRectForMonitor(HMONITOR monitor, RECT* rect) {
    SetRectEmpty(rect);

    HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
    if (!hTaskbarWnd) {
        return false;
    }

    SendMessage(hTaskbarWnd, g_getTaskbarRectRegisteredMsg, (WPARAM)monitor,
                (LPARAM)rect);
    return true;
}

// https://gist.github.com/m417z/451dfc2dad88d7ba88ed1814779a26b4
std::wstring GetWindowAppId(HWND hWnd) {
    // {c8900b66-a973-584b-8cae-355b7f55341b}
    constexpr winrt::guid CLSID_StartMenuCacheAndAppResolver{
        0x660b90c8,
        0x73a9,
        0x4b58,
        {0x8c, 0xae, 0x35, 0x5b, 0x7f, 0x55, 0x34, 0x1b}};

    // {de25675a-72de-44b4-9373-05170450c140}
    constexpr winrt::guid IID_IAppResolver_8{
        0xde25675a,
        0x72de,
        0x44b4,
        {0x93, 0x73, 0x05, 0x17, 0x04, 0x50, 0xc1, 0x40}};

    struct IAppResolver_8 : public IUnknown {
       public:
        virtual HRESULT STDMETHODCALLTYPE GetAppIDForShortcut() = 0;
        virtual HRESULT STDMETHODCALLTYPE GetAppIDForShortcutObject() = 0;
        virtual HRESULT STDMETHODCALLTYPE
        GetAppIDForWindow(HWND hWnd,
                          WCHAR** pszAppId,
                          void* pUnknown1,
                          void* pUnknown2,
                          void* pUnknown3) = 0;
        virtual HRESULT STDMETHODCALLTYPE
        GetAppIDForProcess(DWORD dwProcessId,
                           WCHAR** pszAppId,
                           void* pUnknown1,
                           void* pUnknown2,
                           void* pUnknown3) = 0;
    };

    HRESULT hr;
    std::wstring result;

    CO_MTA_USAGE_COOKIE cookie;
    bool mtaUsageIncreased = SUCCEEDED(CoIncrementMTAUsage(&cookie));

    winrt::com_ptr<IAppResolver_8> appResolver;
    hr = CoCreateInstance(CLSID_StartMenuCacheAndAppResolver, nullptr,
                          CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER,
                          IID_IAppResolver_8, appResolver.put_void());
    if (SUCCEEDED(hr)) {
        WCHAR* pszAppId;
        hr = appResolver->GetAppIDForWindow(hWnd, &pszAppId, nullptr, nullptr,
                                            nullptr);
        if (SUCCEEDED(hr)) {
            result = pszAppId;
            CoTaskMemFree(pszAppId);
        }
    }

    appResolver = nullptr;

    if (mtaUsageIncreased) {
        CoDecrementMTAUsage(cookie);
    }

    return result;
}

std::wstring GetProcessFileName(DWORD dwProcessId) {
    HANDLE hProcess =
        OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dwProcessId);
    if (!hProcess) {
        return std::wstring{};
    }

    WCHAR processPath[MAX_PATH];

    DWORD dwSize = ARRAYSIZE(processPath);
    if (!QueryFullProcessImageName(hProcess, 0, processPath, &dwSize)) {
        CloseHandle(hProcess);
        return std::wstring{};
    }

    CloseHandle(hProcess);

    PCWSTR processFileName = wcsrchr(processPath, L'\\');
    if (!processFileName) {
        return std::wstring{};
    }

    processFileName++;
    return processFileName;
}

std::wstring GetWindowLogInfo(HWND hWnd) {
    DWORD dwProcessId = 0;
    GetWindowThreadProcessId(hWnd, &dwProcessId);
    std::wstring processName = GetProcessFileName(dwProcessId);

    WCHAR className[256];
    if (!GetClassName(hWnd, className, ARRAYSIZE(className))) {
        wcscpy_s(className, L"<unknown>");
    }

    WCHAR windowName[256];
    if (!GetWindowText(hWnd, windowName, ARRAYSIZE(windowName))) {
        windowName[0] = L'\0';
    }

    LONG style = GetWindowLong(hWnd, GWL_STYLE);
    LONG exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);

    RECT rect{};
    GetWindowRect(hWnd, &rect);

    WCHAR buffer[1024];
    swprintf_s(buffer,
               L"window %08X: PID=%u, process=%s, class=%s, name=%s, "
               L"style=0x%08X, exStyle=0x%08X, rect={%d,%d,%d,%d}",
               (DWORD)(DWORD_PTR)hWnd, dwProcessId, processName.c_str(),
               className, windowName, style, exStyle, rect.left, rect.top,
               rect.right, rect.bottom);
    return buffer;
}

bool IsWindowExcluded(HWND hWnd, const Settings& settings) {
    if (settings.excludedPrograms.empty()) {
        return false;
    }

    DWORD resolvedWindowProcessPathLen = 0;
    WCHAR resolvedWindowProcessPath[MAX_PATH];
    WCHAR resolvedWindowProcessPathUpper[MAX_PATH];

    DWORD dwProcessId = 0;
    if (GetWindowThreadProcessId(hWnd, &dwProcessId)) {
        HANDLE hProcess =
            OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, dwProcessId);
        if (hProcess) {
            DWORD dwSize = ARRAYSIZE(resolvedWindowProcessPath);
            if (QueryFullProcessImageName(hProcess, 0,
                                          resolvedWindowProcessPath, &dwSize)) {
                resolvedWindowProcessPathLen = dwSize;
            }

            CloseHandle(hProcess);
        }
    }

    if (resolvedWindowProcessPathLen > 0) {
        LCMapStringEx(LOCALE_NAME_USER_DEFAULT, LCMAP_UPPERCASE,
                      resolvedWindowProcessPath,
                      resolvedWindowProcessPathLen + 1,
                      resolvedWindowProcessPathUpper,
                      resolvedWindowProcessPathLen + 1, nullptr, nullptr, 0);
    } else {
        *resolvedWindowProcessPath = L'\0';
        *resolvedWindowProcessPathUpper = L'\0';
    }

    if (resolvedWindowProcessPathLen > 0 &&
        settings.excludedPrograms.contains(resolvedWindowProcessPathUpper)) {
        return true;
    }

    if (PCWSTR programFileNameUpper =
            wcsrchr(resolvedWindowProcessPathUpper, L'\\')) {
        programFileNameUpper++;
        if (*programFileNameUpper &&
            settings.excludedPrograms.contains(programFileNameUpper)) {
            return true;
        }
    }

    std::wstring appId = GetWindowAppId(hWnd);
    LCMapStringEx(LOCALE_NAME_USER_DEFAULT, LCMAP_UPPERCASE, appId.data(),
                  appId.length(), appId.data(), appId.length(), nullptr,
                  nullptr, 0);
    if (settings.excludedPrograms.contains(appId.c_str())) {
        return true;
    }

    return false;
}

bool IsWindowEligibleForHidingTaskbar(HWND hWnd, const Settings& settings) {
    if (!IsWindowVisible(hWnd) || IsWindowCloaked(hWnd) || IsIconic(hWnd) ||
        (GetWindowLong(hWnd, GWL_EXSTYLE) & WS_EX_NOACTIVATE)) {
        return false;
    }

    // Ignore the Alt+Tab overlay so it doesn't affect taskbar state.
    if (hWnd == g_altTabViewHwnd) {
        return false;
    }

    if (hWnd == GetShellWindow() || GetProp(hWnd, L"DesktopWindow")) {
        return false;
    }

    // Exclude menus (#32768).
    if (GetClassWord(hWnd, GCW_ATOM) == 32768) {
        return false;
    }

    // Check this after the other checks, as it's the most expensive one.
    if (IsWindowExcluded(hWnd, settings)) {
        return false;
    }

    return true;
}

bool CanHideTaskbarForWindow(HWND hWnd,
                             HMONITOR monitor,
                             const MONITORINFO* monitorInfo,
                             const RECT* taskbarRect,
                             const Settings& settings,
                             uint64_t settingsEpoch,
                             uintptr_t eligibilityLoadNonce,
                             const std::vector<RECT>& dockBands) {
    HANDLE prop = GetProp(hWnd, kCanHideTaskbarEligibilityProp);
    uintptr_t propEpoch = reinterpret_cast<uintptr_t>(
        GetProp(hWnd, kCanHideTaskbarEligibilityEpochProp));
    uintptr_t propNonce = reinterpret_cast<uintptr_t>(
        GetProp(hWnd, kCanHideTaskbarEligibilityNonceProp));
    if (!prop || !EligibilitySettingsEpochMatches(propEpoch, settingsEpoch) ||
        !EligibilityLoadNonceMatches(propNonce, eligibilityLoadNonce)) {
        prop = IsWindowEligibleForHidingTaskbar(hWnd, settings)
                   ? kCanHideTaskbarEligible
                   : kCanHideTaskbarNotEligible;
        if (EligibilityPublicationAllowed(settingsEpoch,
                                          eligibilityLoadNonce)) {
            SetProp(hWnd, kCanHideTaskbarEligibilityProp, prop);
            SetProp(hWnd, kCanHideTaskbarEligibilityEpochProp,
                    reinterpret_cast<HANDLE>(
                        EncodeEligibilitySettingsEpoch(settingsEpoch)));
            SetProp(hWnd, kCanHideTaskbarEligibilityNonceProp,
                    reinterpret_cast<HANDLE>(EncodeEligibilityLoadNonce(
                        eligibilityLoadNonce)));
        }
    }

    if (prop == kCanHideTaskbarNotEligible) {
        return false;
    }

    WINDOWPLACEMENT wp{
        .length = sizeof(WINDOWPLACEMENT),
    };

    if (GetWindowPlacement(hWnd, &wp) && wp.showCmd == SW_SHOWMAXIMIZED) {
        if (MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST) == monitor) {
            return true;
        }

        return false;
    }

    bool isWindowArranged = pIsWindowArranged && pIsWindowArranged(hWnd);
    if (isWindowArranged &&
        MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST) != monitor) {
        return false;
    }

    RECT windowRect{};
    DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, &windowRect,
                          sizeof(windowRect));

    // Don't keep the taskbar shown for a fullscreen window.
    if (EqualRect(&windowRect, &monitorInfo->rcMonitor)) {
        return true;
    }

    // It makes sense to treat arranged windows (e.g. with Win+left) as
    // maximized, as they occupy the whole monitor height. Still check for
    // intersection, as a window can also just occupy the upper side of the
    // screen (e.g. Win+left, Win+up).
    if (settings.mode == Mode::intersected ||
        (settings.mode == Mode::maximized && isWindowArranged)) {
        RECT intersectRect;
        if (IntersectRect(&intersectRect, &windowRect, taskbarRect)) {
            return true;
        }
    } else if (settings.mode == Mode::dock) {
        if (dockBands.empty()) {
            RECT intersectRect;
            if (IntersectRect(&intersectRect, &windowRect, taskbarRect)) {
                return true;
            }
        } else if (WindowIntersectsDockBands(windowRect, *taskbarRect,
                                            dockBands)) {
            return true;
        }
    }

    return false;
}

bool ShouldKeepTaskbarShown(HWND hTaskbarWnd, HMONITOR monitor) {
    std::shared_ptr<const Settings> settings;
    uint64_t settingsEpoch;
    uintptr_t eligibilityLoadNonce;
    TaskbarBackend taskbarBackend;
    DockRegionCache dockCache;
    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        settings = g_settingsSnapshot;
        settingsEpoch = g_settingsEpoch;
        eligibilityLoadNonce = g_eligibilityLoadNonce;
        taskbarBackend = g_taskbarBackend;
        if (auto it = g_dockRegionCaches.find(hTaskbarWnd);
            it != g_dockRegionCaches.end()) {
            try {
                dockCache = it->second;
            } catch (...) {
                dockCache = {};
            }
        }
    }

    if (settings->primaryMonitorOnly &&
        monitor != MonitorFromPoint({0, 0}, MONITOR_DEFAULTTOPRIMARY)) {
        return false;
    }

    if (settings->mode == Mode::never) {
        return true;
    }

    // Always show taskbar when MultitaskingView (Win+Tab) is active.
    if (g_multitaskingViewHwnd) {
        return true;
    }

    if (settings->mode == Mode::fullscreen) {
        QUERY_USER_NOTIFICATION_STATE state;
        if (SHQueryUserNotificationState(&state) == S_OK) {
            return !(state == QUNS_BUSY ||
                     state == QUNS_RUNNING_D3D_FULL_SCREEN ||
                     state == QUNS_PRESENTATION_MODE);
        }

        return true;
    }

    MONITORINFO monitorInfo{
        .cbSize = sizeof(MONITORINFO),
    };
    GetMonitorInfo(monitor, &monitorInfo);

    RECT taskbarRect{};
    GetTaskbarRectForMonitor(monitor, &taskbarRect);

    SIZE taskbarSize{taskbarRect.right - taskbarRect.left,
                     taskbarRect.bottom - taskbarRect.top};
    const bool modernDockMode =
        settings->mode == Mode::dock &&
        taskbarBackend == TaskbarBackend::NativeModern &&
        !settings->oldTaskbarOnWin11;
    if (!modernDockMode || dockCache.monitor != monitor ||
        dockCache.taskbarSize.cx != taskbarSize.cx ||
        dockCache.taskbarSize.cy != taskbarSize.cy) {
        dockCache = {};
    }
    std::vector<RECT> dockBands;
    try {
        dockBands = EffectiveDockBands(dockCache, taskbarSize);
    } catch (...) {
        dockBands.clear();
    }

    if (settings->foregroundWindowOnly) {
        HWND hForegroundWnd = GetForegroundWindow();
        return !hForegroundWnd ||
               !CanHideTaskbarForWindow(hForegroundWnd, monitor, &monitorInfo,
                                         &taskbarRect, *settings, settingsEpoch,
                                         eligibilityLoadNonce, dockBands);
    }

    bool canHideTaskbar = false;

    DWORD dwTaskbarThreadId = GetCurrentThreadId();

    auto enumWindowsProc = [&](HWND hWnd) -> BOOL {
        if (GetWindowThreadProcessId(hWnd, nullptr) == dwTaskbarThreadId) {
            return TRUE;
        }

        canHideTaskbar =
            CanHideTaskbarForWindow(hWnd, monitor, &monitorInfo, &taskbarRect,
                                    *settings, settingsEpoch,
                                    eligibilityLoadNonce, dockBands);
        if (!canHideTaskbar) {
            return TRUE;
        }

        Wh_Log(L"Can hide taskbar %08X for %s", (DWORD)(DWORD_PTR)hTaskbarWnd,
               GetWindowLogInfo(hWnd).c_str());
        return FALSE;
    };

    EnumWindows(
        [](HWND hWnd, LPARAM lParam) -> BOOL {
            auto& proc = *reinterpret_cast<decltype(enumWindowsProc)*>(lParam);
            return proc(hWnd);
        },
        reinterpret_cast<LPARAM>(&enumWindowsProc));

    return !canHideTaskbar;
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

DWORD WINAPI WinEventHookWorker(WinEventThreadContext* context);

void AdjustTaskbar(HWND hMMTaskbarWnd,
                   bool clearPendingWhenDone = false) noexcept {
    try {
        if (!g_initialized.load(std::memory_order_acquire)) {
            return;
        }
        auto state = GetSettingsStateSnapshot();
        if (state.settings->mode != Mode::never) {
            if (!g_winEventThreadLifecycle.Start(
                    WinEventHookWorker, state.settingsEpoch,
                    state.settings->foregroundWindowOnly)) {
                Wh_Log(L"Error: Couldn't start WinEvent thread");
            }
        }

        PostMessage(hMMTaskbarWnd, g_updateTaskbarStateRegisteredMsg,
                    clearPendingWhenDone, 0);
    } catch (...) {
        Wh_Log(L"Error: Couldn't schedule taskbar policy adjustment");
    }
}

template <typename Callback>
bool ForEachTaskbarWindowNoexcept(Callback&& callback) noexcept {
    try {
        using CallbackType = std::remove_reference_t<Callback>;
        HWND primary = FindCurrentProcessTaskbarWnd();
        if (!primary) {
            return false;
        }

        bool succeeded = false;
        try {
            succeeded = callback(primary);
        } catch (...) {
            succeeded = false;
        }

        DWORD taskbarThreadId = GetWindowThreadProcessId(primary, nullptr);
        if (!taskbarThreadId) {
            return false;
        }

        struct EnumerationContext {
            CallbackType* callback;
            bool succeeded;
        } context{&callback, succeeded};
        BOOL enumerated = EnumThreadWindows(
            taskbarThreadId,
            [](HWND hWnd, LPARAM lParam) -> BOOL {
                WCHAR className[32];
                if (!GetClassNameW(hWnd, className, ARRAYSIZE(className)) ||
                    _wcsicmp(className, L"Shell_SecondaryTrayWnd") != 0) {
                    return TRUE;
                }
                auto& context =
                    *reinterpret_cast<EnumerationContext*>(lParam);
                try {
                    if (!(*context.callback)(hWnd)) {
                        context.succeeded = false;
                    }
                } catch (...) {
                    context.succeeded = false;
                }
                return TRUE;
            },
            reinterpret_cast<LPARAM>(&context));
        return enumerated && context.succeeded;
    } catch (...) {
        return false;
    }
}

void AdjustAllTaskbars() noexcept {
    ForEachTaskbarWindowNoexcept([](HWND hWnd) noexcept {
        AdjustTaskbar(hWnd);
        return true;
    });
}

void RequestFullPolicyReevaluationLocked() noexcept {
    g_fullPolicyReevaluationPending = true;
    g_fullPolicyReevaluationToken++;
    if (!g_fullPolicyReevaluationToken) {
        g_fullPolicyReevaluationToken = 1;
    }
}

template <typename PublishBackend>
bool CommitExplorerPatcherBackendWith(
    PublishBackend&& publishBackend) noexcept {
    bool restorePolicyAdmission = false;
    bool committed = false;
    try {
        {
            std::lock_guard<std::mutex> policyLock(g_policyCallbackMutex);
            restorePolicyAdmission = g_acceptPolicyCallbacks;
            g_acceptPolicyCallbacks = false;
        }
        WaitForPolicyCallbacks();
        bool publishAllowed = false;
        {
            std::lock_guard<std::mutex> policyLock(g_policyCallbackMutex);
            publishAllowed = !g_policyCallbackTeardownStarted;
        }
        if (publishAllowed &&
            g_initialized.load(std::memory_order_acquire)) {
            committed = publishBackend();
        }
    } catch (...) {
        Wh_Log(L"Error: Couldn't commit the ExplorerPatcher backend");
    }
    try {
        std::lock_guard<std::mutex> policyLock(g_policyCallbackMutex);
        if (committed && restorePolicyAdmission &&
            !g_policyCallbackTeardownStarted &&
            g_initialized.load(std::memory_order_acquire)) {
            g_acceptPolicyCallbacks = true;
        }
    } catch (...) {
    }
    return committed;
}

bool CommitExplorerPatcherBackend() noexcept {
    return CommitExplorerPatcherBackendWith([] {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (g_taskbarBackend == TaskbarBackend::ExplorerPatcherLegacy) {
            return false;
        }

        g_taskbarsKeptShown.clear();
        g_taskbarToViewCoordinator.clear();
        g_taskbarBackend = TaskbarBackend::ExplorerPatcherLegacy;
        g_taskbarBackendEpoch++;
        g_dockEpoch++;
        g_dockRegionCaches.clear();
        g_pendingDockRefreshes.clear();
        ClearAllDockTaskbarRevealStateLocked();
        RequestFullPolicyReevaluationLocked();
        return true;
    });
}

template <typename PostMessageCallback>
bool TrySchedulePendingFullPolicyReevaluationWith(
    PostMessageCallback&& postMessageCallback) noexcept {
    try {
        uint64_t reevaluationToken;
        {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            if (!g_fullPolicyReevaluationPending) {
                return true;
            }
            reevaluationToken = g_fullPolicyReevaluationToken;
        }
        if (!g_updateTaskbarStateRegisteredMsg) {
            return false;
        }

        bool scheduled = ForEachTaskbarWindowNoexcept([&](HWND hWnd) {
            return postMessageCallback(hWnd,
                                       g_updateTaskbarStateRegisteredMsg, 0,
                                       0) != FALSE;
        });
        if (!scheduled) {
            return false;
        }

        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (g_fullPolicyReevaluationPending &&
            g_fullPolicyReevaluationToken == reevaluationToken) {
            g_fullPolicyReevaluationPending = false;
            return true;
        }
        return false;
    } catch (...) {
        Wh_Log(L"Error: Full taskbar policy reevaluation remains pending");
        return false;
    }
}

bool TrySchedulePendingFullPolicyReevaluation() noexcept {
    return TrySchedulePendingFullPolicyReevaluationWith(
        [](HWND hWnd, UINT message, WPARAM wParam,
           LPARAM lParam) noexcept -> BOOL {
            return PostMessageW(hWnd, message, wParam, lParam);
        });
}

bool DockModeEffective(const Settings& settings) {
    return settings.mode == Mode::dock &&
           g_taskbarBackend == TaskbarBackend::NativeModern &&
           !settings.oldTaskbarOnWin11 &&
           g_dockRegionRefreshRegisteredMsg != 0;
}

uint64_t EnsureTaskbarGenerationLocked(HWND hWnd) {
    auto [it, inserted] =
        g_taskbarGenerations.try_emplace(hWnd, g_nextTaskbarGeneration);
    if (inserted) {
        g_nextTaskbarGeneration++;
        if (!g_nextTaskbarGeneration) {
            g_nextTaskbarGeneration = 1;
        }
    }
    return it->second;
}

bool QueueDockRegionRefresh(HWND hWnd) noexcept {
    DockRefreshToken token;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (!g_acceptDockWork || !DockModeEffective(*g_settingsSnapshot)) {
            return false;
        }
        auto pending = g_pendingDockRefreshes.try_emplace(hWnd).first;
        if (pending->second.pending) {
            return false;
        }
        token = {g_dockEpoch, EnsureTaskbarGenerationLocked(hWnd)};
        BeginDockRefresh(pending->second, token);
    } catch (...) {
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto pending = g_pendingDockRefreshes.find(hWnd);
            if (pending != g_pendingDockRefreshes.end() &&
                !pending->second.pending) {
                g_pendingDockRefreshes.erase(pending);
            }
        } catch (...) {
        }
        Wh_Log(L"Error: Couldn't queue Dock refresh state");
        return false;
    }

    auto failPending = [&]() noexcept {
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto it = g_pendingDockRefreshes.find(hWnd);
            if (it != g_pendingDockRefreshes.end() &&
                FailDockRefreshPost(it->second, token)) {
                g_pendingDockRefreshes.erase(it);
            }
        } catch (...) {
            Wh_Log(L"Error: Couldn't clear failed Dock refresh state");
        }
    };

    if (!PostMessageW(hWnd, g_dockRegionRefreshRegisteredMsg,
                      static_cast<WPARAM>(token.dockEpoch),
                      static_cast<LPARAM>(token.generation))) {
        failPending();
        return false;
    }
    return true;
}

void QueueAllDockRegionRefreshes() noexcept {
    ForEachTaskbarWindowNoexcept([](HWND hWnd) noexcept {
        QueueDockRegionRefresh(hWnd);
        return true;
    });
}

void ResetDockTaskbarLifecycle(HWND hWnd, bool destroying) noexcept {
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_dockRegionCaches.erase(hWnd);
        g_pendingDockRefreshes.erase(hWnd);
        g_taskbarGenerations.erase(hWnd);
        ClearDockTaskbarRevealLifecycleLocked(hWnd);
    } catch (...) {
        Wh_Log(L"Error: Couldn't reset taskbar Dock lifecycle state");
    }
}

void ResetDockTopology(HWND hWnd) noexcept {
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_dockEpoch++;
        g_dockRegionCaches.clear();
        g_pendingDockRefreshes.clear();
        ClearAllDockTaskbarRevealStateLocked();
        RequestFullPolicyReevaluationLocked();
    } catch (...) {
        Wh_Log(L"Error: Couldn't reset Dock topology state");
        return;
    }
    QueueAllDockRegionRefreshes();
    AdjustAllTaskbars();
}

bool RectContains(const RECT& outer, const RECT& inner) {
    return outer.left <= inner.left && outer.top <= inner.top &&
           outer.right >= inner.right && outer.bottom >= inner.bottom;
}

class ScopedPhysicalCoordinates {
   public:
    ScopedPhysicalCoordinates() {
        DPI_AWARENESS_CONTEXT current = GetThreadDpiAwarenessContext();
        if (!AreDpiAwarenessContextsEqual(
                current, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)) {
            previous_ = SetThreadDpiAwarenessContext(
                DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            valid_ = previous_ != nullptr;
        }
    }

    ~ScopedPhysicalCoordinates() {
        if (previous_) {
            SetThreadDpiAwarenessContext(previous_);
        }
    }

    ScopedPhysicalCoordinates(const ScopedPhysicalCoordinates&) = delete;
    ScopedPhysicalCoordinates& operator=(const ScopedPhysicalCoordinates&) =
        delete;

    explicit operator bool() const {
        return valid_;
    }

   private:
    DPI_AWARENESS_CONTEXT previous_ = nullptr;
    bool valid_ = true;
};

DockCapture CaptureDockRegion(HWND hWnd,
                              HMONITOR monitor,
                              uint64_t generation) noexcept {
    DockCapture capture;
    capture.monitor = monitor;
    capture.generation = generation;

    const ScopedPhysicalCoordinates coordinates;
    if (!coordinates) {
        return capture;
    }

    RECT taskbarRect{};
    RECT windowRect{};
    MONITORINFO monitorInfo{};
    monitorInfo.cbSize = sizeof(monitorInfo);
    if (!monitor || !GetTaskbarRectForMonitor(monitor, &taskbarRect) ||
        !GetMonitorInfoW(monitor, &monitorInfo) ||
        !GetWindowRect(hWnd, &windowRect)) {
        return capture;
    }

    capture.taskbarSize = {taskbarRect.right - taskbarRect.left,
                           taskbarRect.bottom - taskbarRect.top};
    if (capture.taskbarSize.cx <= 0 || capture.taskbarSize.cy <= 0 ||
        windowRect.right - windowRect.left != capture.taskbarSize.cx ||
        windowRect.bottom - windowRect.top != capture.taskbarSize.cy ||
        MonitorFromRect(&taskbarRect, MONITOR_DEFAULTTONEAREST) != monitor) {
        return capture;
    }
    capture.physicallyShown = RectContains(monitorInfo.rcMonitor, windowRect);

    HRGN region = CreateRectRgn(0, 0, 0, 0);
    if (!region) {
        return capture;
    }
    const int regionType = GetWindowRgn(hWnd, region);
    DWORD dataSize = 0;
    if (regionType == SIMPLEREGION || regionType == COMPLEXREGION) {
        dataSize = GetRegionData(region, 0, nullptr);
    }
    std::vector<BYTE> regionData;
    try {
        regionData.resize(dataSize);
    } catch (...) {
        DeleteObject(region);
        return capture;
    }
    bool dataRead = dataSize >= sizeof(RGNDATAHEADER) &&
                    GetRegionData(
                        region, dataSize,
                        reinterpret_cast<RGNDATA*>(regionData.data())) ==
                        dataSize;
    DeleteObject(region);

    if (!dataRead) {
        return capture;
    }
    const bool rightToLeft =
        (GetWindowLongPtrW(hWnd, GWL_EXSTYLE) & WS_EX_LAYOUTRTL) != 0;
    capture.valid = ParseDockRegionData(
        regionType, regionData.data(), regionData.size(), capture.taskbarSize,
        rightToLeft, &capture.bands);
    return capture;
}

using DockCaptureFunction =
    DockCapture (*)(HWND hWnd, HMONITOR monitor, uint64_t generation);

void ProcessDockRegionRefreshWithCapture(
    HWND hWnd,
    HMONITOR monitor,
    const DockRefreshToken& token,
    DockCaptureFunction captureFunction) noexcept {
    bool adjustForFallback = false;
    try {
        bool current = false;
        {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto pending = g_pendingDockRefreshes.find(hWnd);
            auto generation = g_taskbarGenerations.find(hWnd);
            if (pending == g_pendingDockRefreshes.end() ||
                generation == g_taskbarGenerations.end() ||
                !CompleteDockRefresh(pending->second, token)) {
                return;
            }
            g_pendingDockRefreshes.erase(pending);
            current = g_acceptDockWork &&
                      DockModeEffective(*g_settingsSnapshot) &&
                      IsDockRefreshTokenCurrent(token, g_dockEpoch,
                                                generation->second);
        }
        if (!current) {
            return;
        }

        DockCapture capture =
            captureFunction(hWnd, monitor, token.generation);

        bool policyChanged = false;
        {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto generation = g_taskbarGenerations.find(hWnd);
            if (!g_acceptDockWork ||
                !DockModeEffective(*g_settingsSnapshot) ||
                generation == g_taskbarGenerations.end() ||
                !IsDockRefreshTokenCurrent(token, g_dockEpoch,
                                           generation->second)) {
                return;
            }

            DockRegionCache previous;
            if (auto it = g_dockRegionCaches.find(hWnd);
                it != g_dockRegionCaches.end()) {
                previous = it->second;
            }
            DockRegionCache next = ApplyDockCapture(previous, capture);
            policyChanged = !DockPolicySemanticallyEqual(previous, next);
            g_dockRegionCaches[hWnd] = std::move(next);
        }

        if (policyChanged) {
            AdjustTaskbar(hWnd);
        }
        return;
    } catch (...) {
        Wh_Log(L"Error: Dock refresh failed; using the full taskbar fallback");
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            adjustForFallback = g_dockRegionCaches.erase(hWnd) != 0;
        } catch (...) {
            adjustForFallback = true;
        }
    }

    if (adjustForFallback) {
        try {
            AdjustTaskbar(hWnd);
        } catch (...) {
            Wh_Log(L"Error: Couldn't schedule Dock fallback adjustment");
        }
    }
}

void ProcessDockRegionRefresh(HWND hWnd,
                              HMONITOR monitor,
                              const DockRefreshToken& token) noexcept {
    ProcessDockRegionRefreshWithCapture(hWnd, monitor, token,
                                        CaptureDockRegion);
}

bool AdjustAllTaskbarsIfNotPending() {
    if (!g_initialized.load(std::memory_order_acquire)) {
        return true;
    }
    std::unordered_set<HWND> secondaryTaskbarWindows;
    HWND hWnd = FindTaskbarWindows(&secondaryTaskbarWindows);

    DWORD currentTickCount = GetTickCount();
    DWORD pendingTickCount = 0;

    if (hWnd) {
        DWORD tickCount = (DWORD)(DWORD_PTR)GetProp(
            hWnd, kUpdateTaskbarStatePendingTickCount);
        if (tickCount > pendingTickCount) {
            pendingTickCount = tickCount;
        }
    }

    for (HWND hSecondaryWnd : secondaryTaskbarWindows) {
        DWORD tickCount = (DWORD)(DWORD_PTR)GetProp(
            hSecondaryWnd, kUpdateTaskbarStatePendingTickCount);
        if (tickCount > pendingTickCount) {
            pendingTickCount = tickCount;
        }
    }

    // Consider times larger than 10 seconds as expired, to prevent having
    // it stuck in this state.
    if (pendingTickCount && currentTickCount - pendingTickCount < 1000 * 10) {
        return false;
    }

    if (hWnd) {
        SetProp(hWnd, kUpdateTaskbarStatePendingTickCount,
                (HANDLE)(DWORD_PTR)currentTickCount);
        AdjustTaskbar(hWnd, /*clearPendingWhenDone=*/true);
    }

    for (HWND hSecondaryWnd : secondaryTaskbarWindows) {
        SetProp(hSecondaryWnd, kUpdateTaskbarStatePendingTickCount,
                (HANDLE)(DWORD_PTR)currentTickCount);
        AdjustTaskbar(hSecondaryWnd, /*clearPendingWhenDone=*/true);
    }

    return true;
}

using ViewCoordinator_ShouldTaskbarBeExpanded_t =
    bool(WINAPI*)(void* pThis, HWND hMMTaskbarWnd, bool expanded);
ViewCoordinator_ShouldTaskbarBeExpanded_t
    ViewCoordinator_ShouldTaskbarBeExpanded_Original;
bool WINAPI ViewCoordinator_ShouldTaskbarBeExpanded_Hook(void* pThis,
                                                         HWND hMMTaskbarWnd,
                                                         bool expanded) {
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        return ViewCoordinator_ShouldTaskbarBeExpanded_Original(
            pThis, hMMTaskbarWnd, expanded);
    }
    bool nativeModernActive = false;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        nativeModernActive =
            g_taskbarBackend == TaskbarBackend::NativeModern &&
            g_taskbarViewInstallState.load(std::memory_order_acquire) ==
                TaskbarViewInstallState::Active;
    } catch (...) {
    }
    InvokeBackendActiveCheckCompletedTestHook();
    if (!nativeModernActive) {
        return ViewCoordinator_ShouldTaskbarBeExpanded_Original(
            pThis, hMMTaskbarWnd, expanded);
    }
    Wh_Log(L"> hMMTaskbarWnd=%08X, expanded=%d",
           (DWORD)(ULONG_PTR)hMMTaskbarWnd, expanded);

    g_taskbarToViewCoordinator[hMMTaskbarWnd] = pThis;

    // Return true if the taskbar should be kept shown.
    for (const auto& pair : g_taskbarsKeptShown) {
        if (pair.second == hMMTaskbarWnd) {
            Wh_Log(L"Returning true for taskbar kept shown");
            return true;
        }
    }

    return ViewCoordinator_ShouldTaskbarBeExpanded_Original(
        pThis, hMMTaskbarWnd, expanded);
}

using ViewCoordinator_UpdateIsExpanded_t = void(WINAPI*)(void* pThis,
                                                         HWND hMMTaskbarWnd,
                                                         int reason);
ViewCoordinator_UpdateIsExpanded_t ViewCoordinator_UpdateIsExpanded_Original;

void UpdateViewCoordinatorIsExpanded(HWND hWnd) {
    ViewCoordinator_UpdateIsExpanded_t updateIsExpandedOriginal = nullptr;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (g_taskbarBackend != TaskbarBackend::NativeModern ||
            g_taskbarViewInstallState.load(std::memory_order_acquire) !=
                TaskbarViewInstallState::Active) {
            return;
        }
        // The acquire above synchronizes with the release publication of
        // Active. Capture the non-atomic resolved pointer only afterward.
        updateIsExpandedOriginal =
            ViewCoordinator_UpdateIsExpanded_Original;
    } catch (...) {
        return;
    }
    // From ViewCoordinator::HandleIsPointerOverTaskbarFrameChanged.
    constexpr int kReasonIsPointerOverTaskbarFrameChanged = 7;

    if (updateIsExpandedOriginal) {
        auto it = g_taskbarToViewCoordinator.find(hWnd);
        if (it != g_taskbarToViewCoordinator.end()) {
            updateIsExpandedOriginal(
                it->second, hWnd, kReasonIsPointerOverTaskbarFrameChanged);
        }
    }
}

using TrayUI_GetStuckMonitor_t = HMONITOR(WINAPI*)(void* pThis);
using CSecondaryTray_GetMonitor_t = HMONITOR(WINAPI*)(void* pThis);
using TrayUI_GetStuckRectForMonitor_t = bool(WINAPI*)(void* pThis,
                                                      HMONITOR hMonitor,
                                                      RECT* rect);
using TrayUI_GetStuckRectForMonitor_Win10_t = RECT*(WINAPI*)(void* pThis,
                                                             RECT* rect,
                                                             HMONITOR hMonitor);
using TrayUI__Hide_t = void(WINAPI*)(void* pThis);
using CSecondaryTray__AutoHide_t = void(WINAPI*)(void* pThis, bool param1);
using TrayUI_Unhide_t = void(WINAPI*)(void* pThis,
                                      int trayUnhideFlags,
                                      int unhideRequest);
using CSecondaryTray__Unhide_t = void(WINAPI*)(void* pThis,
                                               int trayUnhideFlags,
                                               int unhideRequest);
using TrayUI_WndProc_t = LRESULT(WINAPI*)(void* pThis,
                                          HWND hWnd,
                                          UINT Msg,
                                          WPARAM wParam,
                                          LPARAM lParam,
                                          bool* flag);
using CSecondaryTray_v_WndProc_t = LRESULT(
    WINAPI*)(void* pThis, HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

struct TaskbarDispatch {
    TaskbarBackend backend = TaskbarBackend::NativeLegacy;
    void* trayUIVtableInspectable = nullptr;
    void* trayUIVtableTrayComponentHost = nullptr;
    void* secondaryTrayVtableSecondaryTray = nullptr;
    TrayUI_GetStuckMonitor_t trayUIGetStuckMonitorOriginal = nullptr;
    CSecondaryTray_GetMonitor_t secondaryTrayGetMonitorOriginal = nullptr;
    TrayUI_GetStuckRectForMonitor_t trayUIGetStuckRectForMonitorOriginal =
        nullptr;
    TrayUI_GetStuckRectForMonitor_Win10_t
        trayUIGetStuckRectForMonitorWin10Original = nullptr;
    TrayUI__Hide_t trayUIHideOriginal = nullptr;
    CSecondaryTray__AutoHide_t secondaryTrayAutoHideOriginal = nullptr;
    TrayUI_Unhide_t trayUIUnhideOriginal = nullptr;
    CSecondaryTray__Unhide_t secondaryTrayUnhideOriginal = nullptr;
    TrayUI_WndProc_t trayUIWndProcOriginal = nullptr;
    CSecondaryTray_v_WndProc_t secondaryTrayWndProcOriginal = nullptr;
};

struct TaskbarRevealForegroundSnapshot {
    HMONITOR monitor = nullptr;
    LONG_PTR style = 0;
    UINT showCommand = SW_SHOWNORMAL;
    RECT windowRect{};
    RECT monitorRect{};
};

enum class TaskbarRevealForegroundQueryResult {
    NoneEligible,
    SnapshotReady,
    InspectionFailed,
};

struct TaskbarRevealZOrderOperations {
    void* context = nullptr;
    bool (*validateTaskbar)(void* context, HWND hWnd, bool secondary) =
        nullptr;
    bool (*getShownTaskbarRect)(void* context,
                                HMONITOR monitor,
                                RECT* rect) = nullptr;
    BOOL (*getWindowRect)(void* context, HWND hWnd, RECT* rect) = nullptr;
    HMONITOR (*getWindowMonitor)(void* context, HWND hWnd) = nullptr;
    TaskbarRevealForegroundQueryResult (*queryForeground)(
        void* context,
        HWND taskbar,
        HMONITOR taskbarMonitor,
        TaskbarRevealForegroundSnapshot* snapshot) = nullptr;
    BOOL (*setWindowPos)(void* context,
                         HWND hWnd,
                         HWND insertAfter,
                         int x,
                         int y,
                         int cx,
                         int cy,
                         UINT flags) = nullptr;
    BOOL (*postMessage)(void* context,
                        HWND hWnd,
                        UINT message,
                        WPARAM wParam,
                        LPARAM lParam) = nullptr;
};

#ifdef WH_EDITING
const TaskbarRevealZOrderOperations*
    g_taskbarRevealZOrderOperationsOverride;
#endif

bool IsEligibleApplicationWindowForDockReveal(HWND hWnd) noexcept {
    try {
        if (!hWnd || !IsWindowVisible(hWnd) || hWnd == GetShellWindow() ||
            IsIconic(hWnd)) {
            return false;
        }
        const LONG_PTR exStyle = GetWindowLongPtrW(hWnd, GWL_EXSTYLE);
        if (exStyle & (WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE)) {
            return false;
        }

        WCHAR className[48]{};
        GetClassNameW(hWnd, className, ARRAYSIZE(className));
        if (wcscmp(className, L"Shell_TrayWnd") == 0 ||
            wcscmp(className, L"Shell_SecondaryTrayWnd") == 0 ||
            wcscmp(className, L"Progman") == 0 ||
            wcscmp(className, L"WorkerW") == 0 ||
            wcscmp(className, L"#32768") == 0) {
            return false;
        }

        return !IsWindowCloaked(hWnd);
    } catch (...) {
        return false;
    }
}

bool GetVisibleWindowBoundsForDockReveal(HWND hWnd, RECT* bounds) noexcept {
    if (!bounds) {
        return false;
    }
    return (SUCCEEDED(DwmGetWindowAttribute(
                hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, bounds,
                sizeof(*bounds))) &&
            bounds->left < bounds->right && bounds->top < bounds->bottom) ||
           (GetWindowRect(hWnd, bounds) && bounds->left < bounds->right &&
            bounds->top < bounds->bottom);
}

bool ValidateTaskbarForDockReveal(void*, HWND hWnd, bool secondary) {
    if (!hWnd || !IsWindow(hWnd)) {
        return false;
    }

    DWORD processId = 0;
    if (!GetWindowThreadProcessId(hWnd, &processId) ||
        processId != GetCurrentProcessId()) {
        return false;
    }

    WCHAR className[32]{};
    if (!GetClassNameW(hWnd, className, ARRAYSIZE(className))) {
        return false;
    }
    return _wcsicmp(className, secondary ? L"Shell_SecondaryTrayWnd"
                                         : L"Shell_TrayWnd") == 0;
}

bool GetShownTaskbarRectForDockReveal(void*,
                                      HMONITOR monitor,
                                      RECT* rect) {
    return rect && GetTaskbarRectForMonitor(monitor, rect);
}

BOOL GetWindowRectForDockReveal(void*, HWND hWnd, RECT* rect) {
    return GetWindowRect(hWnd, rect);
}

HMONITOR GetWindowMonitorForDockReveal(void*, HWND hWnd) {
    return MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
}

TaskbarRevealForegroundQueryResult QueryForegroundForDockReveal(
    void*,
    HWND taskbar,
    HMONITOR taskbarMonitor,
    TaskbarRevealForegroundSnapshot* snapshot) {
    if (!snapshot || !taskbarMonitor) {
        return TaskbarRevealForegroundQueryResult::InspectionFailed;
    }
    const HWND foreground = GetForegroundWindow();
    if (!foreground || foreground == taskbar ||
        !IsEligibleApplicationWindowForDockReveal(foreground)) {
        return TaskbarRevealForegroundQueryResult::NoneEligible;
    }

    MONITORINFO monitorInfo{.cbSize = sizeof(MONITORINFO)};
    WINDOWPLACEMENT placement{.length = sizeof(WINDOWPLACEMENT)};
    RECT windowRect{};
    if (!GetMonitorInfoW(taskbarMonitor, &monitorInfo) ||
        !GetVisibleWindowBoundsForDockReveal(foreground, &windowRect) ||
        !GetWindowPlacement(foreground, &placement)) {
        return TaskbarRevealForegroundQueryResult::InspectionFailed;
    }

    *snapshot = {
        .monitor = MonitorFromWindow(foreground, MONITOR_DEFAULTTONEAREST),
        .style = GetWindowLongPtrW(foreground, GWL_STYLE),
        .showCommand = placement.showCmd,
        .windowRect = windowRect,
        .monitorRect = monitorInfo.rcMonitor,
    };
    return TaskbarRevealForegroundQueryResult::SnapshotReady;
}

BOOL SetWindowPosForDockReveal(void*,
                               HWND hWnd,
                               HWND insertAfter,
                               int x,
                               int y,
                               int cx,
                               int cy,
                               UINT flags) {
    return SetWindowPos(hWnd, insertAfter, x, y, cx, cy, flags);
}

BOOL PostMessageForDockReveal(void*,
                              HWND hWnd,
                              UINT message,
                              WPARAM wParam,
                              LPARAM lParam) {
    return PostMessageW(hWnd, message, wParam, lParam);
}

const TaskbarRevealZOrderOperations& GetTaskbarRevealZOrderOperations() {
#ifdef WH_EDITING
    if (g_taskbarRevealZOrderOperationsOverride) {
        return *g_taskbarRevealZOrderOperationsOverride;
    }
#endif
    static const TaskbarRevealZOrderOperations operations{
        .context = nullptr,
        .validateTaskbar = ValidateTaskbarForDockReveal,
        .getShownTaskbarRect = GetShownTaskbarRectForDockReveal,
        .getWindowRect = GetWindowRectForDockReveal,
        .getWindowMonitor = GetWindowMonitorForDockReveal,
        .queryForeground = QueryForegroundForDockReveal,
        .setWindowPos = SetWindowPosForDockReveal,
        .postMessage = PostMessageForDockReveal,
    };
    return operations;
}

bool NativeModernDockRevealRepairActiveLocked(bool secondary) noexcept {
    const auto& settings = g_settingsSnapshot;
    return g_initialized.load(std::memory_order_acquire) &&
           g_acceptDockWork &&
           g_taskbarBackend == TaskbarBackend::NativeModern &&
           g_initialHookQueueState.load(std::memory_order_acquire) ==
               InitialHookQueueState::Completed &&
           settings && settings->mode == Mode::dock &&
           (!secondary || !settings->primaryMonitorOnly);
}

bool DockRevealForegroundAllowsRepair(
    HWND hWnd,
    HMONITOR monitor,
    const TaskbarRevealZOrderOperations& operations) noexcept {
    try {
        if (!hWnd || !monitor || !operations.queryForeground) {
            return false;
        }

        TaskbarRevealForegroundSnapshot foreground;
        const TaskbarRevealForegroundQueryResult queryResult =
            operations.queryForeground(operations.context, hWnd, monitor,
                                       &foreground);
        if (queryResult ==
            TaskbarRevealForegroundQueryResult::InspectionFailed) {
            return false;
        }
        if (queryResult ==
                TaskbarRevealForegroundQueryResult::SnapshotReady &&
            foreground.monitor == monitor &&
            IsFullscreenWindowShapeForDockReveal(
                foreground.style, foreground.showCommand,
                foreground.windowRect, foreground.monitorRect)) {
            return false;
        }
        return true;
    } catch (...) {
        return false;
    }
}

bool ApplyTaskbarRevealZOrder(
    HWND hWnd,
    const TaskbarRevealZOrderOperations& operations) noexcept {
    try {
        if (!operations.setWindowPos) {
            return false;
        }
        constexpr UINT kRevealZOrderFlags =
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER |
            SWP_NOSENDCHANGING;
        if (!operations.setWindowPos(operations.context, hWnd, HWND_TOPMOST,
                                     0, 0, 0, 0, kRevealZOrderFlags)) {
            Wh_Log(L"SetWindowPos(taskbar reveal z-order) failed, error %u",
                   GetLastError());
            return false;
        }
        return true;
    } catch (...) {
        Wh_Log(L"Error: Couldn't restore taskbar reveal z-order");
        return false;
    }
}

bool DockRevealEnrollmentMatchesToken(
    const DockRevealEnrollment& enrollment,
    const DockRevealRepairToken& token) noexcept {
    return enrollment.settingsEpoch == token.settingsEpoch &&
           enrollment.dockEpoch == token.dockEpoch &&
           enrollment.backendEpoch == token.backendEpoch &&
           enrollment.generation == token.generation;
}

bool DockRevealRepairTokenIsCurrentLocked(
    HWND hWnd,
    const DockRevealRepairToken& token) noexcept {
    auto pending = g_pendingDockRevealRepairs.find(hWnd);
    auto enrollment = g_dockRevealEnrollments.find(hWnd);
    auto generation = g_taskbarGenerations.find(hWnd);
    return pending != g_pendingDockRevealRepairs.end() &&
           pending->second.token == token &&
           enrollment != g_dockRevealEnrollments.end() &&
           enrollment->second.scopeApplicable &&
           DockRevealEnrollmentMatchesToken(enrollment->second, token) &&
           generation != g_taskbarGenerations.end() &&
           generation->second == token.generation &&
           token.settingsEpoch == g_settingsEpoch &&
           token.dockEpoch == g_dockEpoch &&
           token.backendEpoch == g_taskbarBackendEpoch &&
           NativeModernDockRevealRepairActiveLocked(
               enrollment->second.secondary) &&
           g_dockRevealRepairRegisteredMsg != 0 &&
           !g_dockRevealRepairsInProgress.contains(hWnd);
}

void RetireExactDockRevealRepairLocked(
    HWND hWnd,
    const DockRevealRepairToken& token,
    bool resetShownObservation) noexcept {
    auto pending = g_pendingDockRevealRepairs.find(hWnd);
    if (pending != g_pendingDockRevealRepairs.end() &&
        pending->second.token == token) {
        g_pendingDockRevealRepairs.erase(pending);
    }
    auto deferred = g_deferredDockRevealRepairs.find(hWnd);
    if (deferred != g_deferredDockRevealRepairs.end() &&
        deferred->second.token == token) {
        g_deferredDockRevealRepairs.erase(deferred);
    }
    if (!resetShownObservation) {
        return;
    }
    auto enrollment = g_dockRevealEnrollments.find(hWnd);
    if (enrollment != g_dockRevealEnrollments.end() &&
        DockRevealEnrollmentMatchesToken(enrollment->second, token) &&
        enrollment->second.physicalState ==
            DockRevealPhysicalState::Shown) {
        enrollment->second.physicalState = DockRevealPhysicalState::Unknown;
    }
}

bool DeferExactDockRevealRepairLocked(
    HWND hWnd,
    const DockRevealRepairToken& token) noexcept {
    try {
        g_deferredDockRevealRepairs[hWnd] = {
            .token = token,
            .threadId = GetCurrentThreadId(),
        };
        g_taskbarMutationDeferralHasDeferredDelivery = true;
        return true;
    } catch (...) {
        return false;
    }
}

void FlushDeferredDockRevealRepairsForCurrentThread() noexcept {
    if (g_taskbarMutationDeferralDepth != 0) {
        return;
    }

    const DWORD threadId = GetCurrentThreadId();
    g_taskbarMutationDeferralHasDeferredDelivery = false;
    for (;;) {
        HWND hWnd = nullptr;
        DockRevealRepairToken token;
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto deferred = std::find_if(
                g_deferredDockRevealRepairs.begin(),
                g_deferredDockRevealRepairs.end(),
                [threadId](const auto& entry) {
                    return entry.second.threadId == threadId;
                });
            if (deferred == g_deferredDockRevealRepairs.end()) {
                return;
            }
            hWnd = deferred->first;
            token = deferred->second.token;
            g_deferredDockRevealRepairs.erase(deferred);
            if (!DockRevealRepairTokenIsCurrentLocked(hWnd, token)) {
                RetireExactDockRevealRepairLocked(hWnd, token, true);
                continue;
            }
        } catch (...) {
            // A later ordinary taskbar frame will retry the bounded scan.
            g_taskbarMutationDeferralHasDeferredDelivery = true;
            return;
        }

        const auto& operations = GetTaskbarRevealZOrderOperations();
        BOOL posted = FALSE;
        try {
            posted = operations.postMessage && operations.postMessage(
                operations.context, hWnd,
                g_dockRevealRepairRegisteredMsg,
                static_cast<WPARAM>(token.request),
                static_cast<LPARAM>(token.generation));
        } catch (...) {
        }
        if (posted) {
            continue;
        }

        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            RetireExactDockRevealRepairLocked(hWnd, token, true);
        } catch (...) {
        }
    }
}

bool DockRevealRectNearlyEqual(const RECT& left,
                               const RECT& right) noexcept {
    // The native taskbar can land one or two physical pixels from its stuck
    // rectangle under fractional DPI scaling. Auto-hide moves it by almost
    // its full thickness, so a two-pixel tolerance cannot classify hidden
    // geometry as shown.
    constexpr LONG kTolerance = 2;
    auto close = [](LONG a, LONG b) noexcept {
        const int64_t delta = static_cast<int64_t>(a) - b;
        return delta >= -kTolerance && delta <= kTolerance;
    };
    return close(left.left, right.left) && close(left.top, right.top) &&
           close(left.right, right.right) &&
           close(left.bottom, right.bottom);
}

void ClearDockTaskbarRevealLifecycleLocked(HWND hWnd) {
    g_dockRevealEnrollments.erase(hWnd);
    g_pendingDockRevealRepairs.erase(hWnd);
    g_dockRevealRepairsInProgress.erase(hWnd);
    g_deferredDockRevealRepairs.erase(hWnd);
}

void ClearAllDockTaskbarRevealStateLocked() {
    g_dockRevealEnrollments.clear();
    g_pendingDockRevealRepairs.clear();
    g_dockRevealRepairsInProgress.clear();
    g_deferredDockRevealRepairs.clear();
}

bool EnrollDockTaskbarReveal(HWND hWnd,
                             HMONITOR monitor,
                             bool secondary) noexcept {
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (!hWnd || !monitor || !g_acceptDockWork ||
            !NativeModernDockRevealRepairActiveLocked(secondary) ||
            g_dockRevealRepairRegisteredMsg == 0) {
            ClearDockTaskbarRevealLifecycleLocked(hWnd);
            return false;
        }
    } catch (...) {
        return false;
    }

    const auto& operations = GetTaskbarRevealZOrderOperations();
    RECT shownRect{};
    try {
        if (!hWnd || !monitor || !operations.validateTaskbar ||
            !operations.getShownTaskbarRect ||
            !operations.validateTaskbar(operations.context, hWnd, secondary) ||
            !operations.getShownTaskbarRect(operations.context, monitor,
                                             &shownRect) ||
            shownRect.left >= shownRect.right ||
            shownRect.top >= shownRect.bottom) {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            ClearDockTaskbarRevealLifecycleLocked(hWnd);
            return false;
        }
    } catch (...) {
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            ClearDockTaskbarRevealLifecycleLocked(hWnd);
        } catch (...) {
        }
        return false;
    }

    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (!g_acceptDockWork ||
            !NativeModernDockRevealRepairActiveLocked(secondary) ||
            g_dockRevealRepairRegisteredMsg == 0) {
            ClearDockTaskbarRevealLifecycleLocked(hWnd);
            return false;
        }

        const uint64_t generation = EnsureTaskbarGenerationLocked(hWnd);
        DockRevealPhysicalState physicalState =
            DockRevealPhysicalState::Unknown;
        if (auto previous = g_dockRevealEnrollments.find(hWnd);
            previous != g_dockRevealEnrollments.end() &&
            previous->second.settingsEpoch == g_settingsEpoch &&
            previous->second.dockEpoch == g_dockEpoch &&
            previous->second.backendEpoch == g_taskbarBackendEpoch &&
            previous->second.generation == generation &&
            previous->second.monitor == monitor &&
            previous->second.secondary == secondary &&
            DockRevealRectNearlyEqual(previous->second.shownRect,
                                      shownRect)) {
            physicalState = previous->second.physicalState;
        }
        g_dockRevealEnrollments[hWnd] = {
            .settingsEpoch = g_settingsEpoch,
            .dockEpoch = g_dockEpoch,
            .backendEpoch = g_taskbarBackendEpoch,
            .generation = generation,
            .monitor = monitor,
            .shownRect = shownRect,
            .secondary = secondary,
            .scopeApplicable = true,
            .physicalState = physicalState,
        };
        return true;
    } catch (...) {
        Wh_Log(L"Error: Couldn't enroll taskbar reveal repair");
        return false;
    }
}

bool QueueDockTaskbarRevealRepair(HWND hWnd) noexcept {
    DockRevealRepairToken token;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto enrollment = g_dockRevealEnrollments.find(hWnd);
        auto generation = g_taskbarGenerations.find(hWnd);
        if (enrollment == g_dockRevealEnrollments.end() ||
            generation == g_taskbarGenerations.end() ||
            !enrollment->second.scopeApplicable ||
            !NativeModernDockRevealRepairActiveLocked(
                enrollment->second.secondary) ||
            enrollment->second.settingsEpoch != g_settingsEpoch ||
            enrollment->second.dockEpoch != g_dockEpoch ||
            enrollment->second.backendEpoch != g_taskbarBackendEpoch ||
            generation->second != enrollment->second.generation ||
            g_dockRevealRepairRegisteredMsg == 0 ||
            g_dockRevealRepairsInProgress.contains(hWnd)) {
            return false;
        }
        if (g_pendingDockRevealRepairs.contains(hWnd)) {
            return false;
        }
        uint64_t request = g_nextDockRevealRepairRequest++;
        if (!request) {
            request = g_nextDockRevealRepairRequest++;
        }
        token = {
            .settingsEpoch = g_settingsEpoch,
            .dockEpoch = g_dockEpoch,
            .backendEpoch = g_taskbarBackendEpoch,
            .generation = enrollment->second.generation,
            .request = request,
        };
        g_pendingDockRevealRepairs[hWnd] = {.token = token};
        if (g_taskbarMutationDeferralDepth != 0) {
            if (DeferExactDockRevealRepairLocked(hWnd, token)) {
                return true;
            }
            RetireExactDockRevealRepairLocked(hWnd, token, true);
            return false;
        }
    } catch (...) {
        if (token.request != 0) {
            try {
                std::lock_guard<std::mutex> lock(g_stateMutex);
                RetireExactDockRevealRepairLocked(hWnd, token, true);
            } catch (...) {
            }
        }
        return false;
    }

    const auto& operations = GetTaskbarRevealZOrderOperations();
    BOOL posted = FALSE;
    try {
        posted = operations.postMessage && operations.postMessage(
            operations.context, hWnd, g_dockRevealRepairRegisteredMsg,
            static_cast<WPARAM>(token.request),
            static_cast<LPARAM>(token.generation));
    } catch (...) {
    }
    if (posted) {
        return true;
    }

    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto pending = g_pendingDockRevealRepairs.find(hWnd);
        if (pending != g_pendingDockRevealRepairs.end() &&
            pending->second.token == token) {
            RetireExactDockRevealRepairLocked(hWnd, token, true);
        }
    } catch (...) {
    }
    return false;
}

DockRevealPhysicalState QueryDockRevealPhysicalState(
    const TaskbarRevealZOrderOperations& operations,
    HWND hWnd,
    HMONITOR monitor,
    const RECT& shownRect) noexcept {
    try {
        RECT currentRect{};
        if (!operations.getWindowRect || !operations.getWindowMonitor ||
            !operations.getWindowRect(operations.context, hWnd,
                                      &currentRect) ||
            operations.getWindowMonitor(operations.context, hWnd) != monitor) {
            return DockRevealPhysicalState::Unknown;
        }
        return DockRevealRectNearlyEqual(currentRect, shownRect)
                   ? DockRevealPhysicalState::Shown
                   : DockRevealPhysicalState::Hidden;
    } catch (...) {
        return DockRevealPhysicalState::Unknown;
    }
}

bool ObserveDockTaskbarWindowPositionAfterOriginal(HWND hWnd,
                                                   bool secondary) noexcept {
    DockRevealEnrollment snapshot;
    bool inRepair = false;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto enrollment = g_dockRevealEnrollments.find(hWnd);
        if (enrollment == g_dockRevealEnrollments.end() ||
            enrollment->second.secondary != secondary) {
            return false;
        }
        snapshot = enrollment->second;
        inRepair = g_dockRevealRepairsInProgress.contains(hWnd);
    } catch (...) {
        return false;
    }

    const auto& operations = GetTaskbarRevealZOrderOperations();
    const DockRevealPhysicalState current = QueryDockRevealPhysicalState(
        operations, hWnd, snapshot.monitor, snapshot.shownRect);
    bool shouldQueue = false;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto enrollment = g_dockRevealEnrollments.find(hWnd);
        if (enrollment == g_dockRevealEnrollments.end() ||
            enrollment->second.settingsEpoch != snapshot.settingsEpoch ||
            enrollment->second.dockEpoch != snapshot.dockEpoch ||
            enrollment->second.backendEpoch != snapshot.backendEpoch ||
            enrollment->second.generation != snapshot.generation) {
            return false;
        }
        const DockRevealPhysicalState previous =
            enrollment->second.physicalState;
        enrollment->second.physicalState = current;
        shouldQueue = !inRepair &&
                      current == DockRevealPhysicalState::Shown &&
                      previous != DockRevealPhysicalState::Shown;
    } catch (...) {
        return false;
    }
    return shouldQueue && QueueDockTaskbarRevealRepair(hWnd);
}

bool ProcessDockTaskbarRevealRepairMessage(HWND hWnd,
                                           WPARAM wParam,
                                           LPARAM lParam) noexcept {
    DockRevealRepairToken token;
    DockRevealEnrollment enrollment;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto pending = g_pendingDockRevealRepairs.find(hWnd);
        if (pending == g_pendingDockRevealRepairs.end() ||
            pending->second.token.request != static_cast<uint64_t>(wParam) ||
            pending->second.token.generation !=
                static_cast<uint64_t>(lParam)) {
            return true;
        }
        token = pending->second.token;
        if (g_taskbarMutationDeferralDepth != 0) {
            // Sent messages can be pumped reentrantly by native Unhide/XAML.
            // Keep the exact token pending; the outermost mutation scope will
            // revalidate and repost it after every native frame has unwound.
            if (!DeferExactDockRevealRepairLocked(hWnd, token)) {
                RetireExactDockRevealRepairLocked(hWnd, token, true);
            }
            return true;
        }
        g_pendingDockRevealRepairs.erase(pending);

        auto enrolled = g_dockRevealEnrollments.find(hWnd);
        auto generation = g_taskbarGenerations.find(hWnd);
        if (enrolled == g_dockRevealEnrollments.end() ||
            generation == g_taskbarGenerations.end() ||
            !DockRevealEnrollmentMatchesToken(enrolled->second, token) ||
            token.settingsEpoch != g_settingsEpoch ||
            token.dockEpoch != g_dockEpoch ||
            token.backendEpoch != g_taskbarBackendEpoch ||
            generation->second != token.generation ||
            !NativeModernDockRevealRepairActiveLocked(
                enrolled->second.secondary) ||
            g_dockRevealRepairsInProgress.contains(hWnd)) {
            return true;
        }
        enrollment = enrolled->second;
        g_dockRevealRepairsInProgress[hWnd] = token;
    } catch (...) {
        return true;
    }

    struct RepairGuard {
        HWND hWnd;
        DockRevealRepairToken token;
        ~RepairGuard() {
            try {
                std::lock_guard<std::mutex> lock(g_stateMutex);
                auto inProgress = g_dockRevealRepairsInProgress.find(hWnd);
                if (inProgress != g_dockRevealRepairsInProgress.end() &&
                    inProgress->second == token) {
                    g_dockRevealRepairsInProgress.erase(inProgress);
                }
            } catch (...) {
            }
        }
    } guard{hWnd, token};

    const auto& operations = GetTaskbarRevealZOrderOperations();
    try {
        const DockRevealPhysicalState physicalState =
            QueryDockRevealPhysicalState(operations, hWnd,
                                         enrollment.monitor,
                                         enrollment.shownRect);
        if (!operations.validateTaskbar ||
            !operations.validateTaskbar(operations.context, hWnd,
                                        enrollment.secondary) ||
            physicalState != DockRevealPhysicalState::Shown) {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto current = g_dockRevealEnrollments.find(hWnd);
            if (current != g_dockRevealEnrollments.end() &&
                DockRevealEnrollmentMatchesToken(current->second, token)) {
                current->second.physicalState = physicalState;
            }
            return true;
        }
        {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto current = g_dockRevealEnrollments.find(hWnd);
            if (current == g_dockRevealEnrollments.end() ||
                !DockRevealEnrollmentMatchesToken(current->second, token) ||
                token.settingsEpoch != g_settingsEpoch ||
                token.dockEpoch != g_dockEpoch ||
                token.backendEpoch != g_taskbarBackendEpoch) {
                return true;
            }
            current->second.physicalState = DockRevealPhysicalState::Shown;
        }
        if (!DockRevealForegroundAllowsRepair(hWnd, enrollment.monitor,
                                              operations)) {
            return true;
        }
        {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto current = g_dockRevealEnrollments.find(hWnd);
            auto generation = g_taskbarGenerations.find(hWnd);
            auto inProgress = g_dockRevealRepairsInProgress.find(hWnd);
            if (current == g_dockRevealEnrollments.end() ||
                generation == g_taskbarGenerations.end() ||
                inProgress == g_dockRevealRepairsInProgress.end() ||
                !DockRevealEnrollmentMatchesToken(current->second, token) ||
                generation->second != token.generation ||
                inProgress->second != token ||
                token.settingsEpoch != g_settingsEpoch ||
                token.dockEpoch != g_dockEpoch ||
                token.backendEpoch != g_taskbarBackendEpoch ||
                !NativeModernDockRevealRepairActiveLocked(
                    current->second.secondary)) {
                return true;
            }
        }
        // The state lock is intentionally released before calling User32. A
        // concurrent transition can still win the tiny unlock-to-call race,
        // but no potentially blocking foreground/DWM work occurs in it.
        if (!ApplyTaskbarRevealZOrder(hWnd, operations)) {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            auto current = g_dockRevealEnrollments.find(hWnd);
            if (current != g_dockRevealEnrollments.end() &&
                DockRevealEnrollmentMatchesToken(current->second, token)) {
                current->second.physicalState =
                    DockRevealPhysicalState::Unknown;
            }
        }
    } catch (...) {
        Wh_Log(L"Error: Deferred taskbar reveal repair failed");
    }
    return true;
}

TaskbarDispatch g_nativeTaskbarDispatch;
TaskbarDispatch g_explorerPatcherTaskbarDispatch{
    .backend = TaskbarBackend::ExplorerPatcherLegacy,
};

bool IsTaskbarDispatchActive(const TaskbarDispatch& dispatch) noexcept {
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        if (g_taskbarBackend != dispatch.backend) {
            return false;
        }
        bool active;
        if (dispatch.backend == TaskbarBackend::ExplorerPatcherLegacy) {
            active = g_explorerPatcherInstallState.load(
                         std::memory_order_acquire) ==
                     ExplorerPatcherInstallState::Active;
        } else {
            active = g_initialHookQueueState.load(
                         std::memory_order_acquire) ==
                     InitialHookQueueState::Completed;
        }
        return active;
    } catch (...) {
        return false;
    }
}

void TrayUI__Hide_HookImpl(const TaskbarDispatch& dispatch,
                           void* pThis) noexcept {
    TrayUI__Hide_t original = dispatch.trayUIHideOriginal;
    if (!original) {
        return;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        original(pThis);
        return;
    }
    bool active = IsTaskbarDispatchActive(dispatch);
    InvokeBackendActiveCheckCompletedTestHook();
    if (!active) {
        original(pThis);
        return;
    }

    Wh_Log(L">");
    auto it = g_taskbarsKeptShown.find(pThis);
    if (it != g_taskbarsKeptShown.end()) {
        KillTimer(it->second, kTrayUITimerHide);
        return;
    }
    original(pThis);
}

void WINAPI Native_TrayUI__Hide_Hook(void* pThis) noexcept {
    TrayUI__Hide_HookImpl(g_nativeTaskbarDispatch, pThis);
}

void WINAPI ExplorerPatcher_TrayUI__Hide_Hook(void* pThis) noexcept {
    TrayUI__Hide_HookImpl(g_explorerPatcherTaskbarDispatch, pThis);
}

void CSecondaryTray__AutoHide_HookImpl(const TaskbarDispatch& dispatch,
                                       void* pThis,
                                       bool param1) noexcept {
    CSecondaryTray__AutoHide_t original =
        dispatch.secondaryTrayAutoHideOriginal;
    if (!original) {
        return;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        original(pThis, param1);
        return;
    }
    bool active = IsTaskbarDispatchActive(dispatch);
    InvokeBackendActiveCheckCompletedTestHook();
    if (!active) {
        original(pThis, param1);
        return;
    }

    Wh_Log(L">");
    auto it = g_taskbarsKeptShown.find(pThis);
    if (it != g_taskbarsKeptShown.end()) {
        KillTimer(it->second, kTrayUITimerHide);
        return;
    }
    original(pThis, param1);
}

void WINAPI Native_CSecondaryTray__AutoHide_Hook(void* pThis,
                                                  bool param1) noexcept {
    CSecondaryTray__AutoHide_HookImpl(g_nativeTaskbarDispatch, pThis, param1);
}

void WINAPI ExplorerPatcher_CSecondaryTray__AutoHide_Hook(
    void* pThis,
    bool param1) noexcept {
    CSecondaryTray__AutoHide_HookImpl(g_explorerPatcherTaskbarDispatch, pThis,
                                      param1);
}

void WINAPI Native_TrayUI_Unhide_Hook(void* pThis,
                                      int flags,
                                      int request) noexcept {
    const TrayUI_Unhide_t original =
        g_nativeTaskbarDispatch.trayUIUnhideOriginal;
    if (!original) {
        return;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        original(pThis, flags, request);
        return;
    }
    TaskbarMutationDeferralScope mutationScope;
    original(pThis, flags, request);
}

void WINAPI Native_CSecondaryTray__Unhide_Hook(void* pThis,
                                                int flags,
                                                int request) noexcept {
    const CSecondaryTray__Unhide_t original =
        g_nativeTaskbarDispatch.secondaryTrayUnhideOriginal;
    if (!original) {
        return;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        original(pThis, flags, request);
        return;
    }
    TaskbarMutationDeferralScope mutationScope;
    original(pThis, flags, request);
}

bool CallDirectTaskbarUnhide(
    const TaskbarDispatch& dispatch,
    void* pThis,
    HWND hWnd,
    bool secondary) {
    TrayUI_Unhide_t primaryOriginal = nullptr;
    CSecondaryTray__Unhide_t secondaryOriginal = nullptr;
    if (secondary) {
        secondaryOriginal = dispatch.secondaryTrayUnhideOriginal;
        if (!secondaryOriginal) {
            return false;
        }
    } else {
        primaryOriginal = dispatch.trayUIUnhideOriginal;
        if (!primaryOriginal) {
            return false;
        }
    }

    bool queueAfterOriginal = false;
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        auto enrollment = g_dockRevealEnrollments.find(hWnd);
        if (enrollment != g_dockRevealEnrollments.end() &&
            enrollment->second.secondary == secondary &&
            NativeModernDockRevealRepairActiveLocked(secondary)) {
            queueAfterOriginal = true;
        }
    } catch (...) {
    }

    TaskbarMutationDeferralScope mutationScope;
    if (secondary) {
        secondaryOriginal(pThis, 0, 0);
    } else {
        primaryOriginal(pThis, 0, 0);
    }
    return queueAfterOriginal;
}

struct TaskbarWndProcPreprocessResult {
    bool suppressOriginal = false;
    LRESULT result = 0;
    bool observeDockRevealAfterOriginal = false;
    bool queueDockRevealAfterOriginal = false;
    HWND dockRevealTaskbar = nullptr;
    bool dockRevealSecondary = false;
};

#ifdef WH_EDITING
using TaskbarWndProcPreprocessTestHook = void (*)();
TaskbarWndProcPreprocessTestHook g_taskbarWndProcPreprocessTestHook;
#endif

void InvokeTaskbarWndProcPreprocessTestHook() {
#ifdef WH_EDITING
    if (g_taskbarWndProcPreprocessTestHook) {
        g_taskbarWndProcPreprocessTestHook();
    }
#endif
}

template <typename Preprocess, typename CallOriginal>
LRESULT RunTaskbarWndProcBoundary(Preprocess&& preprocess,
                                  CallOriginal&& callOriginal) noexcept {
    TaskbarWndProcPreprocessResult preprocessResult;
    try {
        preprocessResult = preprocess();
    } catch (...) {
        Wh_Log(L"Error: Taskbar WndProc preprocessing failed");
    }
    g_hookRescanWorker.RetryPending();
    TrySchedulePendingFullPolicyReevaluation();
    if (preprocessResult.suppressOriginal) {
        return preprocessResult.result;
    }
    const LRESULT result = callOriginal();
    if (preprocessResult.observeDockRevealAfterOriginal) {
        ObserveDockTaskbarWindowPositionAfterOriginal(
            preprocessResult.dockRevealTaskbar,
            preprocessResult.dockRevealSecondary);
    }
    if (preprocessResult.queueDockRevealAfterOriginal) {
        QueueDockTaskbarRevealRepair(preprocessResult.dockRevealTaskbar);
    }
    return result;
}

LRESULT TrayUI_WndProc_HookImpl(const TaskbarDispatch& dispatch,
                                void* pThis,
                                HWND hWnd,
                                UINT Msg,
                                WPARAM wParam,
                                LPARAM lParam,
                                bool* flag) noexcept {
    TrayUI_WndProc_t original = dispatch.trayUIWndProcOriginal;
    if (!original) {
        return 0;
    }
    if (g_dockRevealRepairRegisteredMsg != 0 &&
        Msg == g_dockRevealRepairRegisteredMsg) {
        PolicyCallbackScope policyCallback;
        if (policyCallback) {
            ProcessDockTaskbarRevealRepairMessage(hWnd, wParam, lParam);
        }
        return 0;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        return original(pThis, hWnd, Msg, wParam, lParam, flag);
    }
    TaskbarMutationDeferralScope mutationScope;
    bool active = IsTaskbarDispatchActive(dispatch);
    InvokeBackendActiveCheckCompletedTestHook();
    if (!active) {
        return original(pThis, hWnd, Msg, wParam, lParam, flag);
    }
    return RunTaskbarWndProcBoundary(
        [&]() -> TaskbarWndProcPreprocessResult {
            InvokeTaskbarWndProcPreprocessTestHook();
            if (Msg == WM_NCCREATE) {
                Wh_Log(L"WM_NCCREATE: %08X", (DWORD)(ULONG_PTR)hWnd);
                ResetDockTaskbarLifecycle(hWnd, false);
                QueueDockRegionRefresh(hWnd);
                AdjustTaskbar(hWnd);
            } else if (Msg == WM_NCDESTROY) {
                Wh_Log(L"WM_NCDESTROY: %08X", (DWORD)(ULONG_PTR)hWnd);
                ResetDockTaskbarLifecycle(hWnd, true);
                g_taskbarsKeptShown.erase(QueryViaVtableBackwards(
                    pThis, dispatch.trayUIVtableInspectable));
                g_taskbarToViewCoordinator.erase(hWnd);
            } else if (Msg == WM_WINDOWPOSCHANGED) {
                QueueDockRegionRefresh(hWnd);
                return {.observeDockRevealAfterOriginal = true,
                        .dockRevealTaskbar = hWnd,
                        .dockRevealSecondary = false};
            } else if (Msg == WM_DPICHANGED || Msg == WM_DISPLAYCHANGE) {
                ResetDockTopology(hWnd);
            } else if (Msg == kHandleTrayPrivateSettingMessage) {
                // Prevent auto-hide from being disabled while the mod is
                // loaded.
                if ((DWORD)wParam == 4) {
                    BOOL bSetAutoHideEnabled = (BOOL)lParam;
                    if (!bSetAutoHideEnabled) {
                        return {true, 0};
                    }
                }
            } else if (Msg == g_getTaskbarRectRegisteredMsg) {
                HMONITOR monitor = (HMONITOR)wParam;
                RECT* rect = (RECT*)lParam;
                if (dispatch.trayUIGetStuckRectForMonitorOriginal) {
                    if (!dispatch.trayUIGetStuckRectForMonitorOriginal(
                            pThis, monitor, rect)) {
                        SetRectEmpty(rect);
                    }
                } else if (
                    dispatch.trayUIGetStuckRectForMonitorWin10Original) {
                    dispatch.trayUIGetStuckRectForMonitorWin10Original(
                        pThis, rect, monitor);
                } else {
                    SetRectEmpty(rect);
                }
            } else if (IsDockRefreshMessage(
                           g_dockRegionRefreshRegisteredMsg, Msg)) {
                HMONITOR monitor =
                    dispatch.trayUIGetStuckMonitorOriginal(pThis);
                ProcessDockRegionRefresh(
                    hWnd, monitor,
                    {static_cast<uint64_t>(wParam),
                     static_cast<uint64_t>(lParam)});
            } else if (g_updateTaskbarStateRegisteredMsg != 0 &&
                       Msg == g_updateTaskbarStateRegisteredMsg) {
                if (!g_wasAutoHideProcessed) {
                    g_wasAutoHideProcessed = true;
                    g_wasAutoHideDisabled =
                        !SendMessage(hWnd, kHandleTrayPrivateSettingMessage,
                                     kTrayPrivateSettingAutoHideGet, 0);
                    if (g_wasAutoHideDisabled) {
                        SendMessage(hWnd, kHandleTrayPrivateSettingMessage,
                                    kTrayPrivateSettingAutoHideSet, TRUE);
                    }
                }

                HMONITOR monitor =
                    dispatch.trayUIGetStuckMonitorOriginal(pThis);
                EnrollDockTaskbarReveal(hWnd, monitor, false);
                bool keepShown = ShouldKeepTaskbarShown(hWnd, monitor);
                bool queueRevealAfterOriginal = false;

                void* pTrayUI_IInspectable = QueryViaVtableBackwards(
                    pThis, dispatch.trayUIVtableInspectable);

                bool keptShown =
                    g_taskbarsKeptShown.contains(pTrayUI_IInspectable);

                if (keepShown != keptShown) {
                    Wh_Log(L"> keepShown=%d", keepShown);

                    if (keepShown) {
                        g_taskbarsKeptShown[pTrayUI_IInspectable] = hWnd;

                        void* pTrayUI_ITrayComponentHost = QueryViaVtable(
                            pThis, dispatch.trayUIVtableTrayComponentHost);
                        queueRevealAfterOriginal = CallDirectTaskbarUnhide(
                            dispatch, pTrayUI_ITrayComponentHost, hWnd, false);

                        UpdateViewCoordinatorIsExpanded(hWnd);
                    } else {
                        g_taskbarsKeptShown.erase(pTrayUI_IInspectable);

                        SetTimer(hWnd, kTrayUITimerHide, 0, nullptr);

                        UpdateViewCoordinatorIsExpanded(hWnd);
                    }
                }

                if (wParam) {
                    RemoveProp(hWnd, kUpdateTaskbarStatePendingTickCount);
                }
                return {.queueDockRevealAfterOriginal =
                            queueRevealAfterOriginal,
                        .dockRevealTaskbar = hWnd,
                        .dockRevealSecondary = false};
            }
            return {};
        },
        [&]() -> LRESULT {
            return original(pThis, hWnd, Msg, wParam, lParam, flag);
        });
}

LRESULT WINAPI Native_TrayUI_WndProc_Hook(void* pThis,
                                          HWND hWnd,
                                          UINT Msg,
                                          WPARAM wParam,
                                          LPARAM lParam,
                                          bool* flag) noexcept {
    return TrayUI_WndProc_HookImpl(g_nativeTaskbarDispatch, pThis, hWnd, Msg,
                                   wParam, lParam, flag);
}

LRESULT WINAPI ExplorerPatcher_TrayUI_WndProc_Hook(
    void* pThis,
    HWND hWnd,
    UINT Msg,
    WPARAM wParam,
    LPARAM lParam,
    bool* flag) noexcept {
    return TrayUI_WndProc_HookImpl(g_explorerPatcherTaskbarDispatch, pThis,
                                   hWnd, Msg, wParam, lParam, flag);
}

LRESULT CSecondaryTray_v_WndProc_HookImpl(const TaskbarDispatch& dispatch,
                                          void* pThis,
                                          HWND hWnd,
                                          UINT Msg,
                                          WPARAM wParam,
                                          LPARAM lParam) noexcept {
    CSecondaryTray_v_WndProc_t original =
        dispatch.secondaryTrayWndProcOriginal;
    if (!original) {
        return 0;
    }
    if (g_dockRevealRepairRegisteredMsg != 0 &&
        Msg == g_dockRevealRepairRegisteredMsg) {
        PolicyCallbackScope policyCallback;
        if (policyCallback) {
            ProcessDockTaskbarRevealRepairMessage(hWnd, wParam, lParam);
        }
        return 0;
    }
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        return original(pThis, hWnd, Msg, wParam, lParam);
    }
    TaskbarMutationDeferralScope mutationScope;
    bool active = IsTaskbarDispatchActive(dispatch);
    InvokeBackendActiveCheckCompletedTestHook();
    if (!active) {
        return original(pThis, hWnd, Msg, wParam, lParam);
    }
    return RunTaskbarWndProcBoundary(
        [&]() -> TaskbarWndProcPreprocessResult {
            InvokeTaskbarWndProcPreprocessTestHook();
            if (Msg == WM_NCCREATE) {
                Wh_Log(L"WM_NCCREATE: %08X", (DWORD)(ULONG_PTR)hWnd);
                ResetDockTaskbarLifecycle(hWnd, false);
                QueueDockRegionRefresh(hWnd);
                AdjustTaskbar(hWnd);
            } else if (Msg == WM_NCDESTROY) {
                Wh_Log(L"WM_NCDESTROY: %08X", (DWORD)(ULONG_PTR)hWnd);
                ResetDockTaskbarLifecycle(hWnd, true);
                g_taskbarsKeptShown.erase(pThis);
                g_taskbarToViewCoordinator.erase(hWnd);
            } else if (Msg == WM_WINDOWPOSCHANGED) {
                QueueDockRegionRefresh(hWnd);
                return {.observeDockRevealAfterOriginal = true,
                        .dockRevealTaskbar = hWnd,
                        .dockRevealSecondary = true};
            } else if (Msg == WM_DPICHANGED || Msg == WM_DISPLAYCHANGE) {
                ResetDockTopology(hWnd);
            } else if (IsDockRefreshMessage(
                           g_dockRegionRefreshRegisteredMsg, Msg)) {
                void* pCSecondaryTray_ISecondaryTray = QueryViaVtable(
                    pThis, dispatch.secondaryTrayVtableSecondaryTray);
                HMONITOR monitor = dispatch.secondaryTrayGetMonitorOriginal(
                    pCSecondaryTray_ISecondaryTray);
                ProcessDockRegionRefresh(
                    hWnd, monitor,
                    {static_cast<uint64_t>(wParam),
                     static_cast<uint64_t>(lParam)});
            } else if (g_updateTaskbarStateRegisteredMsg != 0 &&
                       Msg == g_updateTaskbarStateRegisteredMsg) {
                void* pCSecondaryTray_ISecondaryTray = QueryViaVtable(
                    pThis, dispatch.secondaryTrayVtableSecondaryTray);

                HMONITOR monitor = dispatch.secondaryTrayGetMonitorOriginal(
                    pCSecondaryTray_ISecondaryTray);

                EnrollDockTaskbarReveal(hWnd, monitor, true);
                bool keepShown = ShouldKeepTaskbarShown(hWnd, monitor);
                bool queueRevealAfterOriginal = false;

                bool keptShown = g_taskbarsKeptShown.contains(pThis);

                if (keepShown != keptShown) {
                    Wh_Log(L"> keepShown=%d", keepShown);

                    if (keepShown) {
                        g_taskbarsKeptShown[pThis] = hWnd;

                        queueRevealAfterOriginal = CallDirectTaskbarUnhide(
                            dispatch, pThis, hWnd, true);

                        UpdateViewCoordinatorIsExpanded(hWnd);
                    } else {
                        g_taskbarsKeptShown.erase(pThis);

                        SetTimer(hWnd, kTrayUITimerHide, 0, nullptr);

                        UpdateViewCoordinatorIsExpanded(hWnd);
                    }
                }

                if (wParam) {
                    RemoveProp(hWnd, kUpdateTaskbarStatePendingTickCount);
                }
                return {.queueDockRevealAfterOriginal =
                            queueRevealAfterOriginal,
                        .dockRevealTaskbar = hWnd,
                        .dockRevealSecondary = true};
            }
            return {};
        },
        [&]() -> LRESULT {
            return original(pThis, hWnd, Msg, wParam, lParam);
        });
}

LRESULT WINAPI Native_CSecondaryTray_v_WndProc_Hook(
    void* pThis,
    HWND hWnd,
    UINT Msg,
    WPARAM wParam,
    LPARAM lParam) noexcept {
    return CSecondaryTray_v_WndProc_HookImpl(
        g_nativeTaskbarDispatch, pThis, hWnd, Msg, wParam, lParam);
}

LRESULT WINAPI ExplorerPatcher_CSecondaryTray_v_WndProc_Hook(
    void* pThis,
    HWND hWnd,
    UINT Msg,
    WPARAM wParam,
    LPARAM lParam) noexcept {
    return CSecondaryTray_v_WndProc_HookImpl(
        g_explorerPatcherTaskbarDispatch, pThis, hWnd, Msg, wParam, lParam);
}

void CALLBACK WinEventProc(HWINEVENTHOOK hWinEventHook,
                           DWORD event,
                           HWND hWnd,
                           LONG idObject,
                           LONG idChild,
                           DWORD dwEventThread,
                           DWORD dwmsEventTime) {
    PolicyCallbackScope policyCallback;
    if (!policyCallback) {
        return;
    }

    if (idObject != OBJID_WINDOW ||
        (GetWindowLong(hWnd, GWL_STYLE) & WS_CHILD) || IsTaskbarWindow(hWnd)) {
        return;
    }

    HWND hParentWnd = GetAncestor(hWnd, GA_PARENT);
    if (hParentWnd && hParentWnd != GetDesktopWindow()) {
        return;
    }

    // Check for Multitasking View window state changes.
    {
        auto multitaskingViewType = GetMultitaskingViewType(hWnd);
        bool isTrackedWinTab = hWnd == g_multitaskingViewHwnd;
        bool isTrackedAltTab = hWnd == g_altTabViewHwnd;

        if (multitaskingViewType != MultitaskingViewType::None ||
            isTrackedWinTab || isTrackedAltTab) {
            auto& targetHwnd =
                (multitaskingViewType == MultitaskingViewType::AltTab ||
                 isTrackedAltTab)
                    ? g_altTabViewHwnd
                    : g_multitaskingViewHwnd;

            bool entering = event == EVENT_OBJECT_SHOW ||
                            event == EVENT_OBJECT_UNCLOAKED ||
                            event == EVENT_OBJECT_CREATE;
            bool leaving = event == EVENT_OBJECT_HIDE ||
                           event == EVENT_OBJECT_CLOAKED ||
                           event == EVENT_OBJECT_DESTROY;

            if (entering) {
                HWND expected = nullptr;
                if (!targetHwnd.compare_exchange_strong(expected, hWnd)) {
                    return;
                }
                Wh_Log(
                    L"MultitaskingView entering (%s)",
                    &targetHwnd == &g_altTabViewHwnd ? L"Alt+Tab" : L"Win+Tab");
            } else if (leaving) {
                HWND expected = hWnd;
                if (!targetHwnd.compare_exchange_strong(expected, nullptr)) {
                    return;
                }
                Wh_Log(
                    L"MultitaskingView leaving (%s)",
                    &targetHwnd == &g_altTabViewHwnd ? L"Alt+Tab" : L"Win+Tab");
            } else {
                return;
            }

            // Fall through to trigger timer for all taskbars.
        }
    }

    Wh_Log(
        L"Event %s for %s",
        [](DWORD event) -> PCWSTR {
            switch (event) {
                case EVENT_OBJECT_CREATE:
                    return L"OBJECT_CREATE";
                case EVENT_OBJECT_DESTROY:
                    return L"OBJECT_DESTROY";
                case EVENT_OBJECT_SHOW:
                    return L"OBJECT_SHOW";
                case EVENT_OBJECT_HIDE:
                    return L"OBJECT_HIDE";
                case EVENT_OBJECT_LOCATIONCHANGE:
                    return L"OBJECT_LOCATIONCHANGE";
                case EVENT_OBJECT_CLOAKED:
                    return L"OBJECT_CLOAKED";
                case EVENT_OBJECT_UNCLOAKED:
                    return L"OBJECT_UNCLOAKED";
                case EVENT_SYSTEM_FOREGROUND:
                    return L"SYSTEM_FOREGROUND";
                default:
                    return L"UNKNOWN";
            }
        }(event),
        GetWindowLogInfo(hWnd).c_str());

    if (event == EVENT_OBJECT_LOCATIONCHANGE) {
        auto state = GetSettingsStateSnapshot();
        if (GetProp(hWnd, kCanHideTaskbarEligibilityProp) ==
                kCanHideTaskbarNotEligible &&
            EligibilitySettingsEpochMatches(
                reinterpret_cast<uintptr_t>(GetProp(
                    hWnd, kCanHideTaskbarEligibilityEpochProp)),
                state.settingsEpoch) &&
            EligibilityLoadNonceMatches(
                reinterpret_cast<uintptr_t>(GetProp(
                    hWnd, kCanHideTaskbarEligibilityNonceProp)),
                state.eligibilityLoadNonce)) {
            return;  // not eligible, position change irrelevant
        }
    } else {
        RemoveProp(hWnd, kCanHideTaskbarEligibilityProp);
        RemoveProp(hWnd, kCanHideTaskbarEligibilityEpochProp);
        RemoveProp(hWnd, kCanHideTaskbarEligibilityNonceProp);
    }

    if (g_pendingEventsTimer) {
        return;
    }

    g_pendingEventsTimer = SetTimer(
        nullptr, 0, 200,
        [](HWND hwnd,         // handle of window for timer messages
           UINT uMsg,         // WM_TIMER message
           UINT_PTR idEvent,  // timer identifier
           DWORD dwTime       // current system time
        ) {
            Wh_Log(L">");

            if (!AdjustAllTaskbarsIfNotPending()) {
                Wh_Log(L"Adjustment already pending, will retry later...");
                return;
            }

            KillTimer(nullptr, g_pendingEventsTimer);
            g_pendingEventsTimer = 0;
        });
}

DWORD WINAPI WinEventHookWorker(WinEventThreadContext* context) {
    HWINEVENTHOOK winObjectEventHook1 =
        SetWinEventHook(EVENT_OBJECT_CREATE, EVENT_OBJECT_HIDE, nullptr,
                        WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
    if (!winObjectEventHook1) {
        Wh_Log(L"Error: SetWinEventHook");
    }

    HWINEVENTHOOK winObjectEventHook2 = SetWinEventHook(
        EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE, nullptr,
        WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
    if (!winObjectEventHook2) {
        Wh_Log(L"Error: SetWinEventHook");
    }

    HWINEVENTHOOK winObjectEventHook3 =
        SetWinEventHook(EVENT_OBJECT_CLOAKED, EVENT_OBJECT_UNCLOAKED, nullptr,
                        WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
    if (!winObjectEventHook3) {
        Wh_Log(L"Error: SetWinEventHook");
    }

    HWINEVENTHOOK winSystemEventHook1 = nullptr;
    if (context->ForegroundWindowOnly()) {
        winSystemEventHook1 =
            SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
                            nullptr, WinEventProc, 0, 0, WINEVENT_OUTOFCONTEXT);
        if (!winSystemEventHook1) {
            Wh_Log(L"Error: SetWinEventHook");
        }
    }

    MSG msg{};
    PeekMessageW(&msg, nullptr, WM_USER, WM_USER, PM_NOREMOVE);
    context->SignalReady();

    bool quit = false;
    while (!quit) {
        HANDLE stopEvent = context->StopEvent();
        DWORD waitResult = MsgWaitForMultipleObjectsEx(
            1, &stopEvent, INFINITE, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        if (waitResult == WAIT_OBJECT_0) {
            break;
        }
        if (waitResult == WAIT_FAILED) {
            Wh_Log(L"Error: MsgWaitForMultipleObjectsEx");
            break;
        }

        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                quit = true;
                break;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    if (g_pendingEventsTimer) {
        KillTimer(nullptr, g_pendingEventsTimer);
        g_pendingEventsTimer = 0;
    }

    if (winObjectEventHook1) {
        UnhookWinEvent(winObjectEventHook1);
    }

    if (winObjectEventHook2) {
        UnhookWinEvent(winObjectEventHook2);
    }

    if (winObjectEventHook3) {
        UnhookWinEvent(winObjectEventHook3);
    }

    if (winSystemEventHook1) {
        UnhookWinEvent(winSystemEventHook1);
    }

    g_multitaskingViewHwnd = nullptr;
    g_altTabViewHwnd = nullptr;

    return 0;
}

struct CheckedHookRegistrationOperations {
    void* context = nullptr;
    BOOL (*registerHook)(void* context,
                         void* target,
                         void* detour,
                         void** original) = nullptr;
};

BOOL RegisterCheckedHook(void*,
                         void* target,
                         void* detour,
                         void** original) {
    return Wh_SetFunctionHook(target, detour, original);
}

struct TaskbarViewResolvedSymbols {
    ViewCoordinator_ShouldTaskbarBeExpanded_t shouldTaskbarBeExpanded =
        nullptr;
    ViewCoordinator_UpdateIsExpanded_t updateIsExpanded = nullptr;
};

bool RegisterTaskbarViewResolvedHooksWith(
    TaskbarViewResolvedSymbols resolved,
    const CheckedHookRegistrationOperations& operations) noexcept {
    try {
        if (!operations.registerHook) {
            return false;
        }

        // Windhawk retains the original-function output slot until hook
        // operations are applied. Keep both pass-through targets in
        // process-lifetime storage before the first registration so a partial
        // transaction remains callable and the retained slot never dangles.
        ViewCoordinator_ShouldTaskbarBeExpanded_Original =
            resolved.shouldTaskbarBeExpanded;
        ViewCoordinator_UpdateIsExpanded_Original = resolved.updateIsExpanded;
        if (resolved.shouldTaskbarBeExpanded &&
            !operations.registerHook(
                operations.context,
                reinterpret_cast<void*>(resolved.shouldTaskbarBeExpanded),
                reinterpret_cast<void*>(
                    ViewCoordinator_ShouldTaskbarBeExpanded_Hook),
                reinterpret_cast<void**>(
                    &ViewCoordinator_ShouldTaskbarBeExpanded_Original))) {
            return false;
        }
        return true;
    } catch (...) {
        return false;
    }
}

bool ResolveTaskbarViewDllSymbols(HMODULE module,
                                  TaskbarViewResolvedSymbols* resolved) {
    if (!resolved) {
        return false;
    }
    // Taskbar.View.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(public: bool __cdecl winrt::Taskbar::implementation::ViewCoordinator::ShouldTaskbarBeExpanded(unsigned __int64,bool))"},
            &resolved->shouldTaskbarBeExpanded,
            nullptr,
            true,
        },
        {
            {LR"(public: void __cdecl winrt::Taskbar::implementation::ViewCoordinator::UpdateIsExpanded(unsigned __int64,enum TaskbarTipTest::TaskbarExpandCollapseReason))"},
            &resolved->updateIsExpanded,
            nullptr,
            true,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"Taskbar.View symbol resolution failed");
        return false;
    }
    return true;
}

bool HookTaskbarViewDllSymbols(HMODULE module) {
    TaskbarViewResolvedSymbols resolved;
    if (!ResolveTaskbarViewDllSymbols(module, &resolved)) {
        return false;
    }
    CheckedHookRegistrationOperations operations{nullptr, RegisterCheckedHook};
    if (!RegisterTaskbarViewResolvedHooksWith(resolved, operations)) {
        Wh_Log(L"Taskbar.View hook registration failed");
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

enum class TaskbarViewInstallResult {
    Pending,
    Active,
    Failed,
};

TaskbarBackend GetTaskbarBackend() noexcept {
    try {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        return g_taskbarBackend;
    } catch (...) {
        return TaskbarBackend::NativeLegacy;
    }
}

TaskbarViewInstallResult TryInstallTaskbarViewHooksInternal(
    HMODULE module,
    bool preInit,
    bool waitForHookSerializer) noexcept {
    TaskbarViewInstallState state =
        g_taskbarViewInstallState.load(std::memory_order_acquire);
    if (state == TaskbarViewInstallState::Active) {
        return TaskbarViewInstallResult::Active;
    }
    if (state == TaskbarViewInstallState::FailedTerminal) {
        return TaskbarViewInstallResult::Failed;
    }
    if (state == TaskbarViewInstallState::Installing ||
        state == TaskbarViewInstallState::RegisteredPendingInitialApply) {
        return TaskbarViewInstallResult::Pending;
    }

    std::unique_lock<std::mutex> hookLock(g_hookOperationsMutex,
                                          std::defer_lock);
    try {
        if (waitForHookSerializer) {
            hookLock.lock();
        } else {
            hookLock.try_lock();
        }
    } catch (...) {
        return TaskbarViewInstallResult::Failed;
    }
    if (!hookLock.owns_lock()) {
        return TaskbarViewInstallResult::Pending;
    }

    state = g_taskbarViewInstallState.load(std::memory_order_relaxed);
    if (state == TaskbarViewInstallState::Active) {
        return TaskbarViewInstallResult::Active;
    }
    if (state == TaskbarViewInstallState::FailedTerminal) {
        return TaskbarViewInstallResult::Failed;
    }
    if (state != TaskbarViewInstallState::Unseen) {
        return TaskbarViewInstallResult::Pending;
    }

    g_taskbarViewInstallState.store(TaskbarViewInstallState::Installing,
                                    std::memory_order_release);
    try {
        if (!module || !HookTaskbarViewDllSymbols(module)) {
            g_taskbarViewInstallState.store(
                TaskbarViewInstallState::FailedTerminal,
                std::memory_order_release);
            return TaskbarViewInstallResult::Failed;
        }

        if (preInit) {
            g_taskbarViewInstallState.store(
                TaskbarViewInstallState::RegisteredPendingInitialApply,
                std::memory_order_release);
            return TaskbarViewInstallResult::Pending;
        }

        if (!Wh_ApplyHookOperations()) {
            g_taskbarViewInstallState.store(
                TaskbarViewInstallState::FailedTerminal,
                std::memory_order_release);
            return TaskbarViewInstallResult::Failed;
        }

        g_taskbarViewInstallState.store(TaskbarViewInstallState::Active,
                                        std::memory_order_release);
        return TaskbarViewInstallResult::Active;
    } catch (...) {
        g_taskbarViewInstallState.store(TaskbarViewInstallState::FailedTerminal,
                                        std::memory_order_release);
        Wh_Log(L"Error: Taskbar.View hook installation failed");
        return TaskbarViewInstallResult::Failed;
    }
}

TaskbarViewInstallResult TryInstallTaskbarViewHooks(HMODULE module,
                                                     bool preInit) noexcept {
    return TryInstallTaskbarViewHooksInternal(module, preInit, false);
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

constexpr PCSTR kExplorerPatcherSymbols[] = {
    R"(??_7TrayUI@@6BITrayDeskBand@@@)",
    R"(??_7TrayUI@@6BITrayComponentHost@@@)",
    R"(??_7CSecondaryTray@@6BISecondaryTray@@@)",
    R"(?GetStuckMonitor@TrayUI@@UEAAPEAUHMONITOR__@@XZ)",
    R"(?GetMonitor@CSecondaryTray@@UEAAPEAUHMONITOR__@@XZ)",
    R"(?GetStuckRectForMonitor@TrayUI@@UEAA_NPEAUHMONITOR__@@PEAUtagRECT@@@Z)",
    R"(?_Hide@TrayUI@@QEAAXXZ)",
    R"(?_AutoHide@CSecondaryTray@@AEAAX_N@Z)",
    R"(?Unhide@TrayUI@@UEAAXW4TrayUnhideFlags@TrayCommon@@W4UnhideRequest@3@@Z)",
    R"(?_Unhide@CSecondaryTray@@AEAAXW4TrayUnhideFlags@TrayCommon@@W4UnhideRequest@3@@Z)",
    R"(?WndProc@TrayUI@@UEAA_JPEAUHWND__@@I_K_JPEA_N@Z)",
    // Available in versions newer than 67.1.
    R"(?v_WndProc@CSecondaryTray@@EEAA_JPEAUHWND__@@I_K_J@Z)",
};
constexpr size_t kExplorerPatcherRequiredSymbolCount = 11;
constexpr size_t kExplorerPatcherSymbolCount =
    ARRAYSIZE(kExplorerPatcherSymbols);

enum class ExplorerPatcherInstallResult {
    Pending,
    Active,
    Failed,
};

struct ExplorerPatcherInstallOperations {
    void* context;
    void* (*resolve)(void* context, PCSTR symbol);
    BOOL (*registerHook)(void* context,
                         void* target,
                         void* hook,
                         void** original);
    BOOL (*apply)(void* context);
    void (*postCommit)(void* context);
};

ExplorerPatcherInstallResult TryInstallExplorerPatcherWithInternal(
    const ExplorerPatcherInstallOperations& operations,
    bool preInit,
    bool waitForHookSerializer) noexcept {
    ExplorerPatcherInstallState state =
        g_explorerPatcherInstallState.load(std::memory_order_acquire);
    if (state == ExplorerPatcherInstallState::Active) {
        return ExplorerPatcherInstallResult::Active;
    }
    if (state == ExplorerPatcherInstallState::FailedTerminal) {
        return ExplorerPatcherInstallResult::Failed;
    }
    if (state == ExplorerPatcherInstallState::Installing ||
        state ==
            ExplorerPatcherInstallState::RegisteredPendingInitialApply) {
        return ExplorerPatcherInstallResult::Pending;
    }

    std::unique_lock<std::mutex> installLock(g_hookOperationsMutex,
                                             std::defer_lock);
    try {
        if (waitForHookSerializer) {
            installLock.lock();
        } else {
            installLock.try_lock();
        }
    } catch (...) {
        Wh_Log(L"Error: Couldn't acquire the hook-operation serializer");
        return ExplorerPatcherInstallResult::Failed;
    }
    if (!installLock.owns_lock()) {
        return ExplorerPatcherInstallResult::Pending;
    }

    state = g_explorerPatcherInstallState.load(std::memory_order_relaxed);
    if (state == ExplorerPatcherInstallState::Active) {
        return ExplorerPatcherInstallResult::Active;
    }
    if (state == ExplorerPatcherInstallState::FailedTerminal) {
        return ExplorerPatcherInstallResult::Failed;
    }
    if (state != ExplorerPatcherInstallState::Unseen) {
        return ExplorerPatcherInstallResult::Pending;
    }

    g_explorerPatcherInstallState.store(
        ExplorerPatcherInstallState::Installing, std::memory_order_release);
    bool registrationAttempted = false;
    try {
        if (!operations.resolve || !operations.registerHook ||
            !operations.apply) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::Unseen,
                std::memory_order_release);
            return ExplorerPatcherInstallResult::Failed;
        }

        void* resolved[kExplorerPatcherSymbolCount]{};
        for (size_t i = 0; i < kExplorerPatcherSymbolCount; i++) {
            resolved[i] =
                operations.resolve(operations.context,
                                   kExplorerPatcherSymbols[i]);
        }
        for (size_t i = 0; i < kExplorerPatcherRequiredSymbolCount; i++) {
            if (!resolved[i]) {
                Wh_Log(L"ExplorerPatcher required symbol doesn't exist: %S",
                       kExplorerPatcherSymbols[i]);
                g_explorerPatcherInstallState.store(
                    ExplorerPatcherInstallState::Unseen,
                    std::memory_order_release);
                return ExplorerPatcherInstallResult::Failed;
            }
        }

        TaskbarDispatch dispatch{
            .backend = TaskbarBackend::ExplorerPatcherLegacy,
            .trayUIVtableInspectable = resolved[0],
            .trayUIVtableTrayComponentHost = resolved[1],
            .secondaryTrayVtableSecondaryTray = resolved[2],
            .trayUIGetStuckMonitorOriginal =
                reinterpret_cast<TrayUI_GetStuckMonitor_t>(resolved[3]),
            .secondaryTrayGetMonitorOriginal =
                reinterpret_cast<CSecondaryTray_GetMonitor_t>(resolved[4]),
            .trayUIGetStuckRectForMonitorOriginal =
                reinterpret_cast<TrayUI_GetStuckRectForMonitor_t>(resolved[5]),
            .trayUIHideOriginal =
                reinterpret_cast<TrayUI__Hide_t>(resolved[6]),
            .secondaryTrayAutoHideOriginal =
                reinterpret_cast<CSecondaryTray__AutoHide_t>(resolved[7]),
            .trayUIUnhideOriginal =
                reinterpret_cast<TrayUI_Unhide_t>(resolved[8]),
            .secondaryTrayUnhideOriginal =
                reinterpret_cast<CSecondaryTray__Unhide_t>(resolved[9]),
            .trayUIWndProcOriginal =
                reinterpret_cast<TrayUI_WndProc_t>(resolved[10]),
            .secondaryTrayWndProcOriginal =
                reinterpret_cast<CSecondaryTray_v_WndProc_t>(resolved[11]),
        };
        g_explorerPatcherTaskbarDispatch = dispatch;

        struct HookRegistration {
            void* target;
            void* detour;
            void** original;
        } registrations[] = {
            {resolved[6],
             reinterpret_cast<void*>(ExplorerPatcher_TrayUI__Hide_Hook),
             reinterpret_cast<void**>(
                 &g_explorerPatcherTaskbarDispatch.trayUIHideOriginal)},
            {resolved[7],
             reinterpret_cast<void*>(
                 ExplorerPatcher_CSecondaryTray__AutoHide_Hook),
             reinterpret_cast<void**>(
                 &g_explorerPatcherTaskbarDispatch
                      .secondaryTrayAutoHideOriginal)},
            {resolved[10],
             reinterpret_cast<void*>(ExplorerPatcher_TrayUI_WndProc_Hook),
             reinterpret_cast<void**>(
                 &g_explorerPatcherTaskbarDispatch.trayUIWndProcOriginal)},
            {resolved[11],
             reinterpret_cast<void*>(
                 ExplorerPatcher_CSecondaryTray_v_WndProc_Hook),
             reinterpret_cast<void**>(
                 &g_explorerPatcherTaskbarDispatch
                      .secondaryTrayWndProcOriginal)},
        };

        for (const auto& registration : registrations) {
            if (!registration.target) {
                continue;
            }
            registrationAttempted = true;
            if (!operations.registerHook(
                    operations.context, registration.target,
                    registration.detour, registration.original)) {
                g_explorerPatcherInstallState.store(
                    ExplorerPatcherInstallState::FailedTerminal,
                    std::memory_order_release);
                return ExplorerPatcherInstallResult::Failed;
            }
        }

        if (preInit) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::RegisteredPendingInitialApply,
                std::memory_order_release);
            return ExplorerPatcherInstallResult::Pending;
        }

        if (!operations.apply(operations.context)) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::FailedTerminal,
                std::memory_order_release);
            return ExplorerPatcherInstallResult::Failed;
        }
        installLock.unlock();
        if (!CommitExplorerPatcherBackend()) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::FailedTerminal,
                std::memory_order_release);
            return ExplorerPatcherInstallResult::Failed;
        }
        g_explorerPatcherInstallState.store(
            ExplorerPatcherInstallState::Active,
            std::memory_order_release);
        if (operations.postCommit) {
            try {
                operations.postCommit(operations.context);
            } catch (...) {
                Wh_Log(L"Error: ExplorerPatcher post-commit scheduling failed");
            }
        }
        return ExplorerPatcherInstallResult::Active;
    } catch (...) {
        g_explorerPatcherInstallState.store(
            registrationAttempted
                ? ExplorerPatcherInstallState::FailedTerminal
                : ExplorerPatcherInstallState::Unseen,
            std::memory_order_release);
        Wh_Log(L"Error: ExplorerPatcher installation failed");
        return ExplorerPatcherInstallResult::Failed;
    }
}

ExplorerPatcherInstallResult TryInstallExplorerPatcherWith(
    const ExplorerPatcherInstallOperations& operations,
    bool preInit) noexcept {
    return TryInstallExplorerPatcherWithInternal(operations, preInit, false);
}

enum class InitialHookCompletionResult {
    NoPending,
    Completed,
    Failed,
};

struct InitialHookCompletionOperations {
    void* applyContext;
    BOOL (*apply)(void* context);
    void* postCommitContext;
    void (*postExplorerPatcherCommit)(void* context);
};

InitialHookCompletionResult CompletePendingInitialHookOperationsWith(
    const InitialHookCompletionOperations& operations) noexcept {
    std::unique_lock<std::mutex> hookLock(g_hookOperationsMutex,
                                          std::defer_lock);
    try {
        hookLock.lock();
        InitialHookQueueState queueState =
            g_initialHookQueueState.load(std::memory_order_relaxed);
        if (queueState == InitialHookQueueState::Applying ||
            queueState == InitialHookQueueState::Completed) {
            return InitialHookCompletionResult::NoPending;
        }
        if (queueState == InitialHookQueueState::FailedTerminal) {
            return InitialHookCompletionResult::Failed;
        }
        const bool explorerPatcherPending =
            g_explorerPatcherInstallState.load(std::memory_order_relaxed) ==
            ExplorerPatcherInstallState::RegisteredPendingInitialApply;
        const bool taskbarViewPending =
            g_taskbarViewInstallState.load(std::memory_order_relaxed) ==
            TaskbarViewInstallState::RegisteredPendingInitialApply;
        g_initialHookQueueState.store(InitialHookQueueState::Applying,
                                      std::memory_order_release);
        bool applied = false;
        try {
            applied = operations.apply &&
                      operations.apply(operations.applyContext);
        } catch (...) {
            Wh_Log(L"Error: Initial hook apply operation failed");
        }
        if (!applied) {
            g_initialHookQueueState.store(
                InitialHookQueueState::FailedTerminal,
                std::memory_order_release);
            if (explorerPatcherPending) {
                g_explorerPatcherInstallState.store(
                    ExplorerPatcherInstallState::FailedTerminal,
                    std::memory_order_release);
            }
            if (taskbarViewPending) {
                g_taskbarViewInstallState.store(
                    TaskbarViewInstallState::FailedTerminal,
                    std::memory_order_release);
            }
            return InitialHookCompletionResult::Failed;
        }

        hookLock.unlock();
        if (explorerPatcherPending) {
            if (!CommitExplorerPatcherBackend()) {
                g_explorerPatcherInstallState.store(
                    ExplorerPatcherInstallState::FailedTerminal,
                    std::memory_order_release);
                if (taskbarViewPending) {
                    g_taskbarViewInstallState.store(
                        TaskbarViewInstallState::FailedTerminal,
                        std::memory_order_release);
                }
                g_initialHookQueueState.store(
                    InitialHookQueueState::FailedTerminal,
                    std::memory_order_release);
                return InitialHookCompletionResult::Failed;
            }
        }
        if (taskbarViewPending) {
            g_taskbarViewInstallState.store(TaskbarViewInstallState::Active,
                                            std::memory_order_release);
        }
        if (explorerPatcherPending) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::Active,
                std::memory_order_release);
        }
        g_initialHookQueueState.store(InitialHookQueueState::Completed,
                                      std::memory_order_release);

        if (explorerPatcherPending &&
            operations.postExplorerPatcherCommit) {
            try {
                operations.postExplorerPatcherCommit(
                    operations.postCommitContext);
            } catch (...) {
                Wh_Log(
                    L"Error: ExplorerPatcher post-commit scheduling failed");
            }
        }
        return InitialHookCompletionResult::Completed;
    } catch (...) {
        g_initialHookQueueState.store(InitialHookQueueState::FailedTerminal,
                                      std::memory_order_release);
        if (g_explorerPatcherInstallState.load(std::memory_order_relaxed) ==
            ExplorerPatcherInstallState::RegisteredPendingInitialApply) {
            g_explorerPatcherInstallState.store(
                ExplorerPatcherInstallState::FailedTerminal,
                std::memory_order_release);
        }
        if (g_taskbarViewInstallState.load(std::memory_order_relaxed) ==
            TaskbarViewInstallState::RegisteredPendingInitialApply) {
            g_taskbarViewInstallState.store(
                TaskbarViewInstallState::FailedTerminal,
                std::memory_order_release);
        }
        Wh_Log(L"Error: Initial hook queue completion failed");
        return InitialHookCompletionResult::Failed;
    }
}

void* ResolveExplorerPatcherExport(void* context, PCSTR symbol) {
    return reinterpret_cast<void*>(
        GetProcAddress(static_cast<HMODULE>(context), symbol));
}

BOOL RegisterExplorerPatcherHook(void*,
                                 void* target,
                                 void* hook,
                                 void** original) {
    return Wh_SetFunctionHook(target, hook, original);
}

BOOL ApplyExplorerPatcherHooks(void*) {
    return Wh_ApplyHookOperations();
}

void ScheduleExplorerPatcherPolicyReevaluation(void*) {
    TrySchedulePendingFullPolicyReevaluation();
}

ExplorerPatcherInstallResult HookExplorerPatcherSymbols(
    HMODULE explorerPatcherModule,
    bool waitForHookSerializer = false) {
    ExplorerPatcherInstallOperations operations{
        explorerPatcherModule, ResolveExplorerPatcherExport,
        RegisterExplorerPatcherHook, ApplyExplorerPatcherHooks,
        ScheduleExplorerPatcherPolicyReevaluation};
    return TryInstallExplorerPatcherWithInternal(
        operations, !g_initialized.load(std::memory_order_acquire),
        waitForHookSerializer);
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

enum class ExplorerPatcherScanResult {
    NotFound,
    Pending,
    Active,
    Failed,
    ScanFailed,
};

struct ExplorerPatcherScanOperations {
    void* context = nullptr;
    BOOL (*enumerate)(void* context,
                      HMODULE* modules,
                      DWORD bufferBytes,
                      DWORD* neededBytes) = nullptr;
    bool (*isExplorerPatcher)(void* context, HMODULE module) = nullptr;
    ExplorerPatcherInstallResult (*install)(void* context,
                                            HMODULE module,
                                            bool waitForHookSerializer) =
        nullptr;
};

ExplorerPatcherScanResult HandleLoadedExplorerPatcherWith(
    const ExplorerPatcherScanOperations& operations,
    bool waitForHookSerializer) noexcept {
    try {
        if (!operations.enumerate || !operations.isExplorerPatcher ||
            !operations.install) {
            return ExplorerPatcherScanResult::ScanFailed;
        }

        std::vector<HMODULE> modules(64);
        DWORD neededBytes = 0;
        for (;;) {
            DWORD bufferBytes =
                static_cast<DWORD>(modules.size() * sizeof(HMODULE));
            if (!operations.enumerate(operations.context, modules.data(),
                                      bufferBytes, &neededBytes)) {
                return ExplorerPatcherScanResult::ScanFailed;
            }
            if (neededBytes <= bufferBytes) {
                break;
            }
            if (neededBytes % sizeof(HMODULE)) {
                return ExplorerPatcherScanResult::ScanFailed;
            }
            modules.resize(neededBytes / sizeof(HMODULE));
        }
        if (neededBytes % sizeof(HMODULE)) {
            return ExplorerPatcherScanResult::ScanFailed;
        }

        size_t moduleCount = neededBytes / sizeof(HMODULE);
        for (size_t i = 0; i < moduleCount; i++) {
            if (!operations.isExplorerPatcher(operations.context,
                                              modules[i])) {
                continue;
            }
            switch (operations.install(operations.context, modules[i],
                                       waitForHookSerializer)) {
                case ExplorerPatcherInstallResult::Pending:
                    return ExplorerPatcherScanResult::Pending;
                case ExplorerPatcherInstallResult::Active:
                    return ExplorerPatcherScanResult::Active;
                case ExplorerPatcherInstallResult::Failed:
                    return ExplorerPatcherScanResult::Failed;
            }
        }
        return ExplorerPatcherScanResult::NotFound;
    } catch (...) {
        return ExplorerPatcherScanResult::ScanFailed;
    }
}

BOOL EnumerateExplorerPatcherModules(void*,
                                     HMODULE* modules,
                                     DWORD bufferBytes,
                                     DWORD* neededBytes) {
    return EnumProcessModules(GetCurrentProcess(), modules, bufferBytes,
                              neededBytes);
}

bool ClassifyExplorerPatcherModule(void*, HMODULE module) {
    return IsExplorerPatcherModule(module);
}

ExplorerPatcherInstallResult InstallScannedExplorerPatcher(
    void*,
    HMODULE module,
    bool waitForHookSerializer) {
    return HookExplorerPatcherSymbols(module, waitForHookSerializer);
}

ExplorerPatcherScanResult HandleLoadedExplorerPatcher(
    bool waitForHookSerializer = false) noexcept {
    ExplorerPatcherScanOperations operations{
        nullptr, EnumerateExplorerPatcherModules,
        ClassifyExplorerPatcherModule, InstallScannedExplorerPatcher};
    return HandleLoadedExplorerPatcherWith(operations, waitForHookSerializer);
}

void ProcessDeferredHookRescans(void*, unsigned bits) {
    if ((bits & kHookRescanTaskbarView) &&
        g_actualWinVersion.load(std::memory_order_acquire) >=
            WinVersion::Win11 &&
        GetTaskbarBackend() == TaskbarBackend::NativeModern) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            TryInstallTaskbarViewHooksInternal(
                taskbarViewModule,
                !g_initialized.load(std::memory_order_acquire), true);
        }
    }

    if (bits & kHookRescanExplorerPatcher) {
        if (HandleLoadedExplorerPatcher(true) ==
            ExplorerPatcherScanResult::ScanFailed) {
            struct RetryDeferredHookRescan {};
            throw RetryDeferredHookRescan{};
        }
    }
}

BOOL SignalDeferredHookRescan(void*, HANDLE workEvent) {
    return SetEvent(workEvent);
}

struct LoadedModuleHookDeferralOperations {
    void* context = nullptr;
    unsigned (*classify)(void* context,
                         HMODULE module,
                         LPCWSTR fileName) = nullptr;
    bool (*request)(void* context, unsigned bits) = nullptr;
};

unsigned ClassifyLoadedModuleForHookRescan(void*,
                                           HMODULE module,
                                           LPCWSTR fileName) noexcept {
    if (!module || !fileName) {
        return 0;
    }

    const wchar_t* baseName = wcsrchr(fileName, L'\\');
    if (const wchar_t* forwardSlash = wcsrchr(fileName, L'/');
        forwardSlash && (!baseName || forwardSlash > baseName)) {
        baseName = forwardSlash;
    }
    baseName = baseName ? baseName + 1 : fileName;
    if (_wcsnicmp(baseName, L"ep_taskbar.", 11) == 0) {
        return kHookRescanExplorerPatcher;
    }
    if (_wcsicmp(baseName, L"Taskbar.View.dll") == 0 ||
        _wcsicmp(baseName, L"ExplorerExtensions.dll") == 0) {
        return kHookRescanTaskbarView;
    }
    return 0;
}

bool RequestLoadedModuleHookRescan(void*, unsigned bits) noexcept {
    return g_hookRescanWorker.Request(bits);
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;

HMODULE LoadLibraryExW_HookWith(
    LPCWSTR lpLibFileName,
    HANDLE hFile,
    DWORD dwFlags,
    LoadLibraryExW_t original,
    const LoadedModuleHookDeferralOperations& operations) noexcept {
    HMODULE module = original(lpLibFileName, hFile, dwFlags);
    if (!module || !operations.classify || !operations.request) {
        return module;
    }

    try {
        unsigned bits =
            operations.classify(operations.context, module, lpLibFileName);
        if (bits) {
            operations.request(operations.context, bits);
        }
    } catch (...) {
        Wh_Log(L"Error: Couldn't defer loaded-module hook rescan");
    }
    return module;
}

HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    LoadedModuleHookDeferralOperations operations{
        nullptr, ClassifyLoadedModuleForHookRescan,
        RequestLoadedModuleHookRescan};
    return LoadLibraryExW_HookWith(lpLibFileName, hFile, dwFlags,
                                   LoadLibraryExW_Original, operations);
}

bool RegisterNativeTaskbarResolvedHooksWith(
    TaskbarDispatch resolved,
    const CheckedHookRegistrationOperations& operations) noexcept {
    try {
        if (!operations.registerHook) {
            return false;
        }

        // A queued hook can outlive this helper's stack frame even when a
        // later registration fails. Publish raw pass-through targets first,
        // then let Windhawk write trampolines directly into these stable
        // process-lifetime slots. Never overwrite those outputs afterward.
        g_nativeTaskbarDispatch = resolved;
        struct Registration {
            void* target;
            void* detour;
            void** original;
            bool modernOnly = false;
        } registrations[]{
            {reinterpret_cast<void*>(resolved.trayUIHideOriginal),
             reinterpret_cast<void*>(Native_TrayUI__Hide_Hook),
             reinterpret_cast<void**>(
                 &g_nativeTaskbarDispatch.trayUIHideOriginal)},
            {reinterpret_cast<void*>(resolved.secondaryTrayAutoHideOriginal),
             reinterpret_cast<void*>(Native_CSecondaryTray__AutoHide_Hook),
             reinterpret_cast<void**>(
                 &g_nativeTaskbarDispatch.secondaryTrayAutoHideOriginal)},
            {reinterpret_cast<void*>(resolved.trayUIWndProcOriginal),
             reinterpret_cast<void*>(Native_TrayUI_WndProc_Hook),
             reinterpret_cast<void**>(
                 &g_nativeTaskbarDispatch.trayUIWndProcOriginal)},
            {reinterpret_cast<void*>(resolved.secondaryTrayWndProcOriginal),
             reinterpret_cast<void*>(
                  Native_CSecondaryTray_v_WndProc_Hook),
             reinterpret_cast<void**>(
                  &g_nativeTaskbarDispatch.secondaryTrayWndProcOriginal)},
            {reinterpret_cast<void*>(resolved.trayUIUnhideOriginal),
             reinterpret_cast<void*>(Native_TrayUI_Unhide_Hook),
             reinterpret_cast<void**>(
                 &g_nativeTaskbarDispatch.trayUIUnhideOriginal),
             true},
            {reinterpret_cast<void*>(resolved.secondaryTrayUnhideOriginal),
             reinterpret_cast<void*>(
                 Native_CSecondaryTray__Unhide_Hook),
             reinterpret_cast<void**>(
                 &g_nativeTaskbarDispatch.secondaryTrayUnhideOriginal),
             true},
        };
        for (const auto& registration : registrations) {
            if (registration.modernOnly &&
                resolved.backend != TaskbarBackend::NativeModern) {
                continue;
            }
            if (registration.target &&
                !operations.registerHook(
                    operations.context, registration.target,
                    registration.detour, registration.original)) {
                return false;
            }
        }
        return true;
    } catch (...) {
        return false;
    }
}

bool HookTaskbarSymbols() {
    TaskbarBackend backend = GetTaskbarBackend();
    if (backend == TaskbarBackend::ExplorerPatcherLegacy) {
        Wh_Log(L"Refusing to install native hooks for ExplorerPatcher backend");
        return false;
    }

    HMODULE module;
    if (backend == TaskbarBackend::NativeLegacy) {
        module = GetModuleHandle(nullptr);
    } else {
        module = LoadLibraryEx(L"taskbar.dll", nullptr,
                               LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (!module) {
            Wh_Log(L"Couldn't load taskbar.dll");
            return false;
        }
    }

    TaskbarDispatch resolved{.backend = backend};

    // Taskbar.dll, explorer.exe
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {
                // Windows 11.
                LR"(const TrayUI::`vftable'{for `IInspectable'})",

                // Windows 10.
                LR"(const TrayUI::`vftable'{for `ITrayDeskBand'})",
            },
            &resolved.trayUIVtableInspectable,
        },
        {
            {LR"(const TrayUI::`vftable'{for `ITrayComponentHost'})"},
            &resolved.trayUIVtableTrayComponentHost,
        },
        {
            {LR"(const CSecondaryTray::`vftable'{for `ISecondaryTray'})"},
            &resolved.secondaryTrayVtableSecondaryTray,
        },
        {
            {LR"(public: virtual struct HMONITOR__ * __cdecl TrayUI::GetStuckMonitor(void))"},
            &resolved.trayUIGetStuckMonitorOriginal,
        },
        {
            {LR"(public: virtual struct HMONITOR__ * __cdecl CSecondaryTray::GetMonitor(void))"},
            &resolved.secondaryTrayGetMonitorOriginal,
        },
        {
            {LR"(public: virtual bool __cdecl TrayUI::GetStuckRectForMonitor(struct HMONITOR__ *,struct tagRECT *))"},
            &resolved.trayUIGetStuckRectForMonitorOriginal,
            nullptr,
            true,  // Windows 11.
        },
        {
            {LR"(public: virtual struct tagRECT __cdecl TrayUI::GetStuckRectForMonitor(struct HMONITOR__ *))"},
            &resolved.trayUIGetStuckRectForMonitorWin10Original,
            nullptr,
            true,  // Windows 10.
        },
        {
            {LR"(public: void __cdecl TrayUI::_Hide(void))"},
            &resolved.trayUIHideOriginal,
            nullptr,
        },
        {
            {LR"(private: void __cdecl CSecondaryTray::_AutoHide(bool))"},
            &resolved.secondaryTrayAutoHideOriginal,
            nullptr,
        },
        {
            {LR"(public: virtual void __cdecl TrayUI::Unhide(enum TrayCommon::TrayUnhideFlags,enum TrayCommon::UnhideRequest))"},
            &resolved.trayUIUnhideOriginal,
        },
        {
            {LR"(private: void __cdecl CSecondaryTray::_Unhide(enum TrayCommon::TrayUnhideFlags,enum TrayCommon::UnhideRequest))"},
            &resolved.secondaryTrayUnhideOriginal,
        },
        {
            {LR"(public: virtual __int64 __cdecl TrayUI::WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64,bool *))"},
            &resolved.trayUIWndProcOriginal,
            nullptr,
        },
        {
            {LR"(private: virtual __int64 __cdecl CSecondaryTray::v_WndProc(struct HWND__ *,unsigned int,unsigned __int64,__int64))"},
            &resolved.secondaryTrayWndProcOriginal,
            nullptr,
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        return false;
    }
    CheckedHookRegistrationOperations operations{nullptr, RegisterCheckedHook};
    return RegisterNativeTaskbarResolvedHooksWith(resolved, operations);
}

Settings BuildSettings() {
    Settings settings;
    PCWSTR mode = Wh_GetStringSetting(L"mode");
    if (wcscmp(mode, L"maximized") == 0) {
        settings.mode = Mode::maximized;
    } else if (wcscmp(mode, L"dock") == 0) {
        settings.mode = Mode::dock;
    } else if (wcscmp(mode, L"fullscreen") == 0) {
        settings.mode = Mode::fullscreen;
    } else if (wcscmp(mode, L"never") == 0) {
        settings.mode = Mode::never;
    }
    Wh_FreeStringSetting(mode);

    settings.foregroundWindowOnly =
        Wh_GetIntSetting(L"foregroundWindowOnly");

    for (int i = 0;; i++) {
        PCWSTR program = Wh_GetStringSetting(L"excludedPrograms[%d]", i);

        bool hasProgram = *program;
        if (hasProgram) {
            std::wstring programUpper = program;
            LCMapStringEx(
                LOCALE_NAME_USER_DEFAULT, LCMAP_UPPERCASE, &programUpper[0],
                static_cast<int>(programUpper.length()), &programUpper[0],
                static_cast<int>(programUpper.length()), nullptr, nullptr, 0);

            settings.excludedPrograms.insert(std::move(programUpper));
        }

        Wh_FreeStringSetting(program);

        if (!hasProgram) {
            break;
        }
    }

    settings.primaryMonitorOnly = Wh_GetIntSetting(L"primaryMonitorOnly");
    settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");
    return settings;
}

void PublishSettings(Settings settings) {
    auto snapshot =
        std::make_shared<const Settings>(std::move(settings));
    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_settingsSnapshot = std::move(snapshot);
        g_settingsEpoch++;
        g_dockEpoch++;
        g_dockRegionCaches.clear();
        g_pendingDockRefreshes.clear();
        ClearAllDockTaskbarRevealStateLocked();
    }
}

BOOL Wh_ModInit() {
    Wh_Log(L">");

    ResetPolicyCallbackTeardown();
    g_initialHookQueueState.store(InitialHookQueueState::AwaitingApply,
                                  std::memory_order_release);
    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_eligibilityLoadNonce = GenerateEligibilityLoadNonce();
    }
    PublishSettings(BuildSettings());
    auto settings = GetSettingsStateSnapshot().settings;

    HMODULE hUser32Module =
        LoadLibraryEx(L"user32.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (hUser32Module) {
        pIsWindowArranged = (IsWindowArranged_t)GetProcAddress(
            hUser32Module, "IsWindowArranged");
        pGetWindowBand =
            (GetWindowBand_t)GetProcAddress(hUser32Module, "GetWindowBand");
    }

    WinVersion actualWinVersion = GetExplorerVersion();
    g_actualWinVersion.store(actualWinVersion, std::memory_order_release);
    if (actualWinVersion == WinVersion::Unsupported) {
        Wh_Log(L"Unsupported Windows version");
        return FALSE;
    }

    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_taskbarBackend =
            actualWinVersion >= WinVersion::Win11 &&
                    !settings->oldTaskbarOnWin11
                ? TaskbarBackend::NativeModern
                : TaskbarBackend::NativeLegacy;
        g_taskbarBackendEpoch++;
        ClearAllDockTaskbarRevealStateLocked();
    }

    if (settings->oldTaskbarOnWin11) {
        bool hasWin10Taskbar = actualWinVersion < WinVersion::Win11_24H2;

        if (hasWin10Taskbar && !HookTaskbarSymbols()) {
            return FALSE;
        }
    } else if (actualWinVersion >= WinVersion::Win11) {
        if (HMODULE taskbarViewModule = GetTaskbarViewModuleHandle()) {
            if (TryInstallTaskbarViewHooks(taskbarViewModule, true) ==
                TaskbarViewInstallResult::Failed) {
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

    ExplorerPatcherScanResult explorerPatcherScan =
        HandleLoadedExplorerPatcher(true);
    if (explorerPatcherScan == ExplorerPatcherScanResult::Failed ||
        explorerPatcherScan == ExplorerPatcherScanResult::ScanFailed ||
        (explorerPatcherScan == ExplorerPatcherScanResult::Pending &&
         g_explorerPatcherInstallState.load(std::memory_order_acquire) !=
             ExplorerPatcherInstallState::RegisteredPendingInitialApply)) {
        Wh_Log(L"HandleLoadedExplorerPatcher failed");
        return FALSE;
    }

    HookRescanWorkerOperations hookRescanOperations{
        .context = nullptr,
        .process = ProcessDeferredHookRescans,
        .signal = SignalDeferredHookRescan,
    };
    if (!g_hookRescanWorker.Start(hookRescanOperations)) {
        Wh_Log(L"Couldn't start the hook-rescan worker");
        return FALSE;
    }

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    auto pKernelBaseLoadLibraryExW = (decltype(&LoadLibraryExW))GetProcAddress(
        kernelBaseModule, "LoadLibraryExW");
    if (!pKernelBaseLoadLibraryExW ||
        !WindhawkUtils::SetFunctionHook(pKernelBaseLoadLibraryExW,
                                        LoadLibraryExW_Hook,
                                        &LoadLibraryExW_Original)) {
        Wh_Log(L"Couldn't hook LoadLibraryExW");
        StopHookRescanWorkerConfirmed();
        return FALSE;
    }

    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_acceptDockWork = true;
    }
    auto initializedState = GetSettingsStateSnapshot();
    g_winEventThreadLifecycle.SetAdmission(
        initializedState.settings->mode != Mode::never,
        initializedState.settingsEpoch);
    g_initialized.store(true, std::memory_order_release);
    SetPolicyCallbacksAccepted(true);

    return TRUE;
}

void Wh_ModAfterInit() {
    Wh_Log(L">");

    InitialHookCompletionOperations initialOperations{
        nullptr, ApplyExplorerPatcherHooks, nullptr,
        ScheduleExplorerPatcherPolicyReevaluation};
    if (CompletePendingInitialHookOperationsWith(initialOperations) ==
        InitialHookCompletionResult::Failed) {
        Wh_Log(L"Couldn't complete pending initial hooks");
        ClosePolicyCallbacksForTeardown();
        g_hookRescanWorker.CloseAdmission();
        StopHookRescanWorkerConfirmed();
        g_initialized.store(false, std::memory_order_release);
        g_winEventThreadLifecycle.SetAdmission(false, 0);
        g_winEventThreadLifecycle.Stop();
        InvokeInitialHookFailureBeforePolicyDrainTestHook();
        WaitForPolicyCallbacks();
        try {
            std::lock_guard<std::mutex> lock(g_stateMutex);
            g_acceptDockWork = false;
            g_dockEpoch++;
            g_pendingDockRefreshes.clear();
            g_dockRegionCaches.clear();
            ClearAllDockTaskbarRevealStateLocked();
            g_fullPolicyReevaluationPending = false;
        } catch (...) {
            Wh_Log(L"Error: Couldn't finish initial-hook failure cleanup");
        }
        return;
    }

    // Retry retained loader notifications and close the startup race with a
    // fresh worker-side rescan of both independently installable hook sets.
    if (!g_hookRescanWorker.Request(kHookRescanExplorerPatcher |
                                    kHookRescanTaskbarView)) {
        Wh_Log(L"Couldn't signal the post-init hook rescan");
    }

    TrySchedulePendingFullPolicyReevaluation();

    WNDCLASS wndclass;
    if (GetClassInfo(GetModuleHandle(nullptr), L"Shell_TrayWnd", &wndclass)) {
        QueueAllDockRegionRefreshes();
        AdjustAllTaskbars();
    }
}

void ClearCanHideTaskbarForWindowProps() {
    EnumWindows(
        [](HWND hWnd, LPARAM) -> BOOL {
            RemoveProp(hWnd, kCanHideTaskbarEligibilityProp);
            RemoveProp(hWnd, kCanHideTaskbarEligibilityEpochProp);
            RemoveProp(hWnd, kCanHideTaskbarEligibilityNonceProp);
            return TRUE;
        },
        0);
}

void Wh_ModBeforeUninit() {
    Wh_Log(L">");

    ClosePolicyCallbacksForTeardown();
    g_hookRescanWorker.CloseAdmission();
    StopHookRescanWorkerConfirmed();
    g_initialized.store(false, std::memory_order_release);
    WaitForPolicyCallbacks();
}

void Wh_ModUninit() {
    Wh_Log(L">");

    g_winEventThreadLifecycle.SetAdmission(false, 0);
    g_initialized.store(false, std::memory_order_release);
    ClosePolicyCallbacksForTeardown();
    StopHookRescanWorkerConfirmed();
    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_acceptDockWork = false;
        g_dockEpoch++;
        g_pendingDockRefreshes.clear();
        g_dockRegionCaches.clear();
        g_taskbarGenerations.clear();
        ClearAllDockTaskbarRevealStateLocked();
        g_fullPolicyReevaluationPending = false;
    }

    g_winEventThreadLifecycle.Stop();
    WaitForPolicyCallbacks();
    ClearCanHideTaskbarForWindowProps();
    {
        std::lock_guard<std::mutex> lock(g_stateMutex);
        g_eligibilityLoadNonce = 0;
    }

    if (g_wasAutoHideDisabled) {
        HWND hTaskbarWnd = FindCurrentProcessTaskbarWnd();
        if (hTaskbarWnd) {
            SendMessage(hTaskbarWnd, kHandleTrayPrivateSettingMessage,
                        kTrayPrivateSettingAutoHideSet, FALSE);
        }
    }
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    auto previous = GetSettingsStateSnapshot().settings;
    Settings next = BuildSettings();
    bool oldTaskbarOnWin11Changed =
        previous->oldTaskbarOnWin11 != next.oldTaskbarOnWin11;
    g_winEventThreadLifecycle.SetAdmission(false, 0);
    PublishSettings(std::move(next));
    g_winEventThreadLifecycle.Stop();

    if (oldTaskbarOnWin11Changed) {
        *bReload = TRUE;
        return TRUE;
    }

    ClearCanHideTaskbarForWindowProps();

    auto state = GetSettingsStateSnapshot();
    g_winEventThreadLifecycle.SetAdmission(
        state.settings->mode != Mode::never, state.settingsEpoch);

    QueueAllDockRegionRefreshes();
    AdjustAllTaskbars();

    return TRUE;
}
