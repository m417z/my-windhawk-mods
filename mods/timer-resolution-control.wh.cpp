// ==WindhawkMod==
// @id           timer-resolution-control
// @name         Timer Resolution Control
// @description  Prevent programs from changing the Windows timer resolution and increasing power consumption
// @version      1.1
// @author       m417z
// @github       https://github.com/m417z
// @twitter      https://twitter.com/m417z
// @homepage     https://m417z.com/
// @include      *
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
# Timer Resolution Control

The default timer resolution on Windows is 15.6 milliseconds. When programs increase the timer
frequency they increase power consumption and harm battery life. This mod provides configuration
to determine which programs are allowed to change the timer resolution.

More details:
[Windows Timer Resolution: Megawatts Wasted](https://randomascii.wordpress.com/2013/07/08/windows-timer-resolution-megawatts-wasted/)

The changes will apply immediately when the mod loads or when settings are changed.

Authors: m417z, levitation
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- DefaultConfig: allow
  $name: Default configuration
  $description: Configuration for all programs, can be overridden for specific programs below
  $options:
  - allow: Allow changing the timer resolution
  - block: Disallow changing the timer resolution
  - limit: Limit changing the timer resolution
- DefaultLimit: 10
  $name: Default timer resolution limit (for the limit configuration)
  $description: The lowest possible delay between timer events, in milliseconds
- PerProgramConfig:
  - - Name: notepad.exe
      $name: Program name or path
    - Config: allow
      $name: Program configuration
      $options:
      - allow: Allow changing the timer resolution
      - block: Disallow changing the timer resolution
      - limit: Limit changing the timer resolution
    - Limit: 10
      $name: Program timer resolution limit (for the limit configuration)
  $name: Per-program configuration
*/
// ==/WindhawkModSettings==

#include <ntdef.h>
#include <ntstatus.h>

enum class Config {
    allow,
    block,
    limit
};

ULONG g_minimumResolution;
ULONG g_maximumResolution;
ULONG g_limitResolution;
ULONG g_lastDesiredResolution;

Config ConfigFromString(PCWSTR string) {
    if (wcscmp(string, L"block") == 0) {
        return Config::block;
    }

    if (wcscmp(string, L"limit") == 0) {
        return Config::limit;
    }

    return Config::allow;
}

typedef NTSTATUS (WINAPI *NtSetTimerResolution_t)(ULONG, BOOLEAN, PULONG);
NtSetTimerResolution_t pOriginalNtSetTimerResolution;

NTSTATUS WINAPI NtSetTimerResolutionHook(ULONG DesiredResolution, BOOLEAN SetResolution, PULONG CurrentResolution)
{
    if (!SetResolution) {
        Wh_Log(L"< SetResolution is FALSE");
        return pOriginalNtSetTimerResolution(DesiredResolution, SetResolution, CurrentResolution);
    }

    Wh_Log(L"> DesiredResolution: %f milliseconds", (double)DesiredResolution / 10000.0);

    g_lastDesiredResolution = DesiredResolution;

    ULONG limitResolution = g_limitResolution;

    if (limitResolution == ULONG_MAX) {
        Wh_Log(L"* Blocking resolution change");
        return STATUS_SUCCESS;
    }

    if (DesiredResolution < limitResolution) {
        Wh_Log(L"* Overriding resolution: %f milliseconds", (double)limitResolution / 10000.0);
        DesiredResolution = limitResolution;
    }

    return pOriginalNtSetTimerResolution(DesiredResolution, SetResolution, CurrentResolution);
}

void LoadSettings()
{
    WCHAR programPath[1024];
    DWORD dwSize = ARRAYSIZE(programPath);
    if (!QueryFullProcessImageName(GetCurrentProcess(), 0, programPath, &dwSize)) {
        *programPath = L'\0';
    }

    PCWSTR programFileName = wcsrchr(programPath, L'\\');
    if (programFileName) {
        programFileName++;
        if (!*programFileName) {
            programFileName = nullptr;
        }
    }

    bool matched = false;
    Config config = Config::allow;
    int limit = 0;

    for (int i = 0; ; i++) {
        PCWSTR name = Wh_GetStringSetting(L"PerProgramConfig[%d].Name", i);
        bool hasName = *name;
        if (hasName) {
            if (programFileName && wcsicmp(programFileName, name) == 0) {
                matched = true;
            }
            else if (wcsicmp(programPath, name) == 0) {
                matched = true;
            }
        }

        Wh_FreeStringSetting(name);

        if (!hasName) {
            break;
        }

        if (matched) {
            PCWSTR configString = Wh_GetStringSetting(L"PerProgramConfig[%d].Config", i);
            config = ConfigFromString(configString);
            Wh_FreeStringSetting(configString);

            if (config == Config::limit) {
                limit = Wh_GetIntSetting(L"PerProgramConfig[%d].Limit", i);
            }

            break;
        }
    }

    if (!matched) {
        PCWSTR configString = Wh_GetStringSetting(L"DefaultConfig");
        config = ConfigFromString(configString);
        Wh_FreeStringSetting(configString);

        if (config == Config::limit) {
            limit = Wh_GetIntSetting(L"DefaultLimit");
        }
    }

    if (config == Config::block) {
        Wh_Log(L"Config loaded: Disallowing changes");
        g_limitResolution = ULONG_MAX;
    }
    else if (config == Config::limit) {
        ULONG limitResolution = limit * 10000;
        if (limitResolution > g_minimumResolution) {
            limitResolution = g_minimumResolution;
        }
        else if (limitResolution < g_maximumResolution) {
            limitResolution = g_maximumResolution;
        }

        Wh_Log(L"Config loaded: Limiting to %f milliseconds", (double)limitResolution / 10000.0);
        g_limitResolution = limitResolution;
    }
    else {
        Wh_Log(L"Config loaded: Allowing changes");
        g_limitResolution = 0;
    }
}

BOOL Wh_ModInit()
{
    Wh_Log(L"Init");

    HMODULE hNtdll = GetModuleHandle(L"ntdll.dll");
    if (!hNtdll) {
        return FALSE;
    }

    FARPROC pNtQueryTimerResolution = GetProcAddress(hNtdll, "NtQueryTimerResolution");
    if (!pNtQueryTimerResolution) {
        return FALSE;
    }

    FARPROC pNtSetTimerResolution = GetProcAddress(hNtdll, "NtSetTimerResolution");
    if (!pNtSetTimerResolution) {
        return FALSE;
    }

    ULONG MinimumResolution;
    ULONG MaximumResolution;
    ULONG CurrentResolution;
    NTSTATUS status = ((NTSTATUS(WINAPI*)(PULONG, PULONG, PULONG))pNtQueryTimerResolution)(
        &MinimumResolution, &MaximumResolution, &CurrentResolution);
    if (!NT_SUCCESS(status)) {
        Wh_Log(L"NtQueryTimerResolution failed with status: 0x%X", status);
        return FALSE;
    }

    Wh_Log(L"NtQueryTimerResolution: min=%f, max=%f, current=%f",
        (double)MinimumResolution / 10000.0,
        (double)MaximumResolution / 10000.0,
        (double)CurrentResolution / 10000.0);
    g_minimumResolution = MinimumResolution;
    g_maximumResolution = MaximumResolution;
    g_lastDesiredResolution = CurrentResolution;    //TODO: During mod load there is a small race condition where the hook is not yet set, the g_lastDesiredResolution is already set, and then the program tries to set a new desired resolution which will not be saved into g_lastDesiredResolution.

    LoadSettings();

    Wh_SetFunctionHook((void*)pNtSetTimerResolution, (void*)NtSetTimerResolutionHook, (void**)&pOriginalNtSetTimerResolution);

    return TRUE;
}

void EnforceLimits() 
{
    ULONG CurrentResolution;
    NTSTATUS status = pOriginalNtSetTimerResolution(0, FALSE, &CurrentResolution);
    if (status == STATUS_TIMER_RESOLUTION_NOT_SET) {
        Wh_Log(L"NtSetTimerResolution has not been called by the program");
        status = STATUS_SUCCESS;
        CurrentResolution = g_lastDesiredResolution;    //Take the value obtained from NtQueryTimerResolution, which does not fail with STATUS_TIMER_RESOLUTION_NOT_SET. Even if the program has not set the resolution, I have observed that the system itself may have set the resolution on program start to 1ms for some reason.
    }
    if (NT_SUCCESS(status)) {
        Wh_Log(L"NtGetTimerResolution: current=%f", (double)CurrentResolution / 10000.0);

        ULONG limitResolution = g_limitResolution;
        ULONG lastDesiredResolution = g_lastDesiredResolution;
        if (CurrentResolution < limitResolution
            || CurrentResolution < lastDesiredResolution) {

            Wh_Log(L"> CurrentResolution: %f milliseconds", (double)CurrentResolution / 10000.0);

            if (lastDesiredResolution < limitResolution) {
                Wh_Log(L"* Overriding resolution: %f milliseconds", (double)limitResolution / 10000.0);
                pOriginalNtSetTimerResolution(limitResolution, TRUE, &CurrentResolution);
            }
            else {
                Wh_Log(L"* Restoring resolution: %f milliseconds", (double)lastDesiredResolution / 10000.0);
                pOriginalNtSetTimerResolution(lastDesiredResolution, TRUE, &CurrentResolution);
            }
        }
    }
    else {
        Wh_Log(L"NtGetTimerResolution failed with status: 0x%X", status);
    }
}

void Wh_ModAfterInit() 
{  
    //Force mod timer resolution settings if the resolution was already changed by the program. NB! In order to avoid any race conditions, do this only after the hook is activated. Else there might be a situation that the mod overrides the current resolution while hook is not yet set, and then the program changes it again before the hook is finally set.

    EnforceLimits();
}

void Wh_ModSettingsChanged() 
{
    Wh_Log(L"SettingsChanged");

    LoadSettings();

    EnforceLimits();
}

void Wh_ModUninit() 
{
    Wh_Log(L"Uniniting...");

    //Lift all limits and restore original resolution set by the program
    ULONG lastDesiredResolution = g_lastDesiredResolution;
    Wh_Log(L"Restoring desired resolution: %f milliseconds", (double)lastDesiredResolution / 10000.0);
    ULONG CurrentResolution;
    pOriginalNtSetTimerResolution(g_lastDesiredResolution, TRUE, &CurrentResolution);
}

