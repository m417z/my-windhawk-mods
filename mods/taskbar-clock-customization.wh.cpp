// ==WindhawkMod==
// @id              taskbar-clock-customization
// @name            Taskbar Clock Customization
// @description     Customize the taskbar clock: define a custom date/time format, add a news feed, customize fonts and colors, and more
// @version         1.4
// @author          m417z
// @github          https://github.com/m417z
// @twitter         https://twitter.com/m417z
// @homepage        https://m417z.com/
// @include         explorer.exe
// @architecture    x86-64
// @compilerOptions -D__USE_MINGW_ANSI_STDIO=0 -lole32 -loleaut32 -lruntimeobject -lversion -lwininet
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
# Taskbar Clock Customization

Customize the taskbar clock: define a custom date/time format, add a news feed,
customize fonts and colors, and more.

Only Windows 10 64-bit and Windows 11 are supported.

**Note:** To customize the old taskbar on Windows 11 (if using ExplorerPatcher
or a similar tool), enable the relevant option in the mod's settings.

![Screenshot](https://i.imgur.com/gM9kbH5.png)

## Available patterns

Supported fields - top line, bottom line, middle line (Windows 10 only), tooltip
extra line - can be configured with text that contains patterns. The following
patterns can be used:

* `%time%` - the time as configured by the time format in settings.
  * `%time<n>%` - additional time formats which can be specified by separating
    the time format string with `;`. `<n>` is the additional time format number,
    starting with 2.
* `%date%` - the date as configured by the date format in settings.
  * `%date<n>%` - additional date formats which can be specified by separating
    the date format string with `;`. `<n>` is the additional date format number,
    starting with 2.
* `%weekday%` - the week day as configured by the week day format in settings.
  * `%weekday<n>%` - additional week day formats which can be specified by
    separating the week day format string with `;`. `<n>` is the additional week
    day format number, starting with 2.
* `%weekday_num%` - the week day number according to the [first day of
   week](https://superuser.com/q/61002) system configuration. For example, if
   first day of week is Sunday, then the week day number is 1 for Sunday, 2 for
   Monday, ..., 7 for Saturday.
* `%weeknum%` - the week number, calculated as following: The week containing 1
  January is defined as week 1 of the year. Subsequent weeks start on first day
  of week according to the system configuration.
* `%weeknum_iso%` - the [ISO week
  number](https://en.wikipedia.org/wiki/ISO_week_date).
* `%dayofyear%` - the day of year starting from January 1st.
* `%timezone%` - the time zone in ISO 8601 format.
* `%web<n>%` - the web contents as configured in settings, truncated with
  ellipsis, where `<n>` is the web contents number.
* `%web<n>_full%` - the full web contents as configured in settings, where `<n>`
  is the web contents number.
* `%newline%` - a newline.

## Text styles

For Windows 11 version 22H2 and newer, the mod allows to change the clock text
styles, such as the font color and size.

![Screenshot](https://i.imgur.com/3JiXwjT.png)
*/
// ==/WindhawkModReadme==

// ==WindhawkModSettings==
/*
- ShowSeconds: true
  $name: Show seconds
- TimeFormat: >-
    hh':'mm':'ss tt
  $name: Time format
  $description: >-
    Leave empty for the default format, for syntax refer to the following page:

    https://docs.microsoft.com/en-us/windows/win32/api/datetimeapi/nf-datetimeapi-gettimeformatex#remarks
- DateFormat: >-
    ddd',' MMM dd yyyy
  $name: Date format
  $description: >-
    Leave empty for the default format, for syntax refer to the following page:

    https://docs.microsoft.com/en-us/windows/win32/intl/day--month--year--and-era-format-pictures
- WeekdayFormat: dddd
  $name: Week day format
  $description: >-
    Leave empty for the default format, for syntax refer to the following page:

    https://docs.microsoft.com/en-us/windows/win32/intl/day--month--year--and-era-format-pictures
- TopLine: '%date% | %time%'
  $name: Top line
  $description: >-
    Text to be shown on the first line, set to "-" for the default value, refer
    to the mod details for list of patterns that can be used
- BottomLine: '%web1%'
  $name: Bottom line
  $description: >-
    Only shown if the taskbar is large enough, set to "-" for the default value
- MiddleLine: '%weekday%'
  $name: Middle line (Windows 10 only)
  $description: >-
    Only shown if the taskbar is large enough, set to "-" for the default value
- TooltipLine: '%web1_full%'
  $name: Tooltip extra line
- Width: 180
  $name: Clock width (Windows 10 only)
- Height: 60
  $name: Clock height (Windows 10 only)
- MaxWidth: 0
  $name: Clock max width (Windows 11 only)
  $description: Set to zero to have no max width
- TextSpacing: 0
  $name: Line spacing
  $description: >-
    Set 0 for the default system value. A negative value can be used for
    negative spacing.
- WebContentsItems:
  - - Url: https://feeds.bbci.co.uk/news/world/rss.xml
      $name: Web content URL
    - BlockStart: '<item>'
      $name: Web content block start
      $description: The string in the webpage to start from
    - Start: '<title><![CDATA['
      $name: Web content start
      $description: The string just before the content
    - End: ']]></title>'
      $name: Web content end
      $description: The string just after the content
    - MaxLength: 28
      $name: Web content maximum length
      $description: Longer strings will be truncated with ellipsis
  $name: Web content items
  $description: >-
    Will be used to fetch data displayed in place of the %web<n>%,
    %web<n>_full% patterns, where <n> is the web contents number
- WebContentsUpdateInterval: 10
  $name: Web content update interval
  $description: The update interval, in minutes, of the web content items
- TimeStyle:
  - Visible: true
  - TextColor: ""
    $name: Text color
    $description: >-
      Can be a color name (Red, Black, ...) or an RGB/ARGB color code (like
      #00FF00, #CC00FF00, ...)
  - TextAlignment: ""
    $name: Text alignment
    $options:
    - "": Default
    - Right: Right
    - Center: Center
    - Left: Left
  - FontSize: 0
    $name: Font size
    $description: Set to zero for the default size
  - FontFamily: ""
    $name: Font family
    $description: >-
      For a list of fonts that are shipped with Windows 11, refer to the
      following page:

      https://learn.microsoft.com/en-us/typography/fonts/windows_11_font_list
  - FontWeight: ""
    $name: Font weight
    $options:
    - "": Default
    - Thin: Thin
    - ExtraLight: Extra light
    - Light: Light
    - SemiLight: Semi light
    - Normal: Normal
    - Medium: Medium
    - SemiBold: Semi bold
    - Bold: Bold
    - ExtraBold: Extra bold
    - Black: Black
    - ExtraBlack: Extra black
  - FontStyle: ""
    $name: Font style
    $options:
    - "": Default
    - Normal: Normal
    - Oblique: Oblique
    - Italic: Italic
  - FontStretch: ""
    $name: Font stretch
    $description: Only supported for some fonts
    $options:
    - "": Default
    - Undefined: Undefined
    - UltraCondensed: Ultra condensed
    - ExtraCondensed: Extra condensed
    - Condensed: Condensed
    - SemiCondensed: Semi condensed
    - Normal: Normal
    - SemiExpanded: Semi expanded
    - Expanded: Expanded
    - ExtraExpanded: Extra expanded
    - UltraExpanded: Ultra expanded
  - CharacterSpacing: 0
    $name: Character spacing
    $description: Can be a positive or a negative number
  $name: Top line style (Windows 11 version 22H2 and newer)
- DateStyle:
  - Visible: true
  - TextColor: ""
    $name: Text color
    $description: >-
      Can be a color name (Red, Black, ...) or an RGB/ARGB color code (like
      #00FF00, #CC00FF00, ...)
  - TextAlignment: ""
    $name: Text alignment
    $options:
    - "": Default
    - Right: Right
    - Center: Center
    - Left: Left
  - FontSize: 0
    $name: Font size
    $description: Set to zero for the default size
  - FontFamily: ""
    $name: Font family
    $description: >-
      For a list of fonts that are shipped with Windows 11, refer to the
      following page:

      https://learn.microsoft.com/en-us/typography/fonts/windows_11_font_list
  - FontWeight: ""
    $name: Font weight
    $options:
    - "": Default
    - Thin: Thin
    - ExtraLight: Extra light
    - Light: Light
    - SemiLight: Semi light
    - Normal: Normal
    - Medium: Medium
    - SemiBold: Semi bold
    - Bold: Bold
    - ExtraBold: Extra bold
    - Black: Black
    - ExtraBlack: Extra black
  - FontStyle: ""
    $name: Font style
    $options:
    - "": Default
    - Normal: Normal
    - Oblique: Oblique
    - Italic: Italic
  - FontStretch: ""
    $name: Font stretch
    $description: Only supported for some fonts
    $options:
    - "": Default
    - Undefined: Undefined
    - UltraCondensed: Ultra condensed
    - ExtraCondensed: Extra condensed
    - Condensed: Condensed
    - SemiCondensed: Semi condensed
    - Normal: Normal
    - SemiExpanded: Semi expanded
    - Expanded: Expanded
    - ExtraExpanded: Extra expanded
    - UltraExpanded: Ultra expanded
  - CharacterSpacing: 0
    $name: Character spacing
    $description: Can be a positive or a negative number
  $name: Bottom line style (Windows 11 version 22H2 and newer)
- oldTaskbarOnWin11: false
  $name: Customize the old taskbar on Windows 11
  $description: >-
    Enable this option to customize the old taskbar on Windows 11 (if using
    ExplorerPatcher or a similar tool).
*/
// ==/WindhawkModSettings==

#include <windhawk_utils.h>

#include <atomic>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

using namespace std::string_view_literals;

#include <psapi.h>
#include <wininet.h>

#undef GetCurrentTime

#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.UI.Xaml.Controls.h>
#include <winrt/Windows.UI.Xaml.Interop.h>
#include <winrt/Windows.UI.Xaml.Markup.h>
#include <winrt/Windows.UI.Xaml.Media.h>

using namespace winrt::Windows::UI::Xaml;

class StringSetting {
   public:
    StringSetting() = default;
    StringSetting(PCWSTR valueName) : m_stringSetting(valueName) {}
    operator PCWSTR() const { return m_stringSetting.get(); }
    PCWSTR get() const { return m_stringSetting.get(); }

   private:
    // https://stackoverflow.com/a/51274008
    template <auto fn>
    struct deleter_from_fn {
        template <typename T>
        constexpr void operator()(T* arg) const {
            fn(arg);
        }
    };
    using string_setting_unique_ptr =
        std::unique_ptr<const WCHAR[], deleter_from_fn<Wh_FreeStringSetting>>;

    string_setting_unique_ptr m_stringSetting;
};

struct WebContentsSettings {
    StringSetting url;
    StringSetting blockStart;
    StringSetting start;
    StringSetting end;
    int maxLength;
};

struct TextStyleSettings {
    bool visible;
    StringSetting textColor;
    StringSetting textAlignment;
    int fontSize;
    StringSetting fontFamily;
    StringSetting fontWeight;
    StringSetting fontStyle;
    StringSetting fontStretch;
    int characterSpacing;
};

struct {
    bool showSeconds;
    StringSetting timeFormat;
    StringSetting dateFormat;
    StringSetting weekdayFormat;
    StringSetting topLine;
    StringSetting bottomLine;
    StringSetting middleLine;
    StringSetting tooltipLine;
    int width;
    int height;
    int maxWidth;
    int textSpacing;
    std::vector<WebContentsSettings> webContentsItems;
    int webContentsUpdateInterval;
    TextStyleSettings timeStyle;
    TextStyleSettings dateStyle;
    bool oldTaskbarOnWin11;

    // Kept for compatibility with old settings:
    StringSetting webContentsUrl;
    StringSetting webContentsBlockStart;
    StringSetting webContentsStart;
    StringSetting webContentsEnd;
    int webContentsMaxLength;
} g_settings;

#define FORMATTED_BUFFER_SIZE 256
#define INTEGER_BUFFER_SIZE sizeof("-2147483648")

enum class WinVersion {
    Unsupported,
    Win10,
    Win11,
    Win11_22H2,
    Win11_24H2,
};

WinVersion g_winVersion;

std::atomic<bool> g_initialized;
std::atomic<bool> g_explorerPatcherInitialized;

WCHAR g_timeFormatted[FORMATTED_BUFFER_SIZE];
WCHAR g_dateFormatted[FORMATTED_BUFFER_SIZE];
WCHAR g_weekdayFormatted[FORMATTED_BUFFER_SIZE];
WCHAR g_weekdayNumFormatted[INTEGER_BUFFER_SIZE];
WCHAR g_weeknumFormatted[INTEGER_BUFFER_SIZE];
WCHAR g_weeknumIsoFormatted[INTEGER_BUFFER_SIZE];
WCHAR g_dayOfYearFormatted[INTEGER_BUFFER_SIZE];
WCHAR g_timezoneFormatted[FORMATTED_BUFFER_SIZE];

std::vector<std::wstring> g_timeExtraFormatted;
std::vector<std::wstring> g_dateExtraFormatted;
std::vector<std::wstring> g_weekdayExtraFormatted;

HANDLE g_webContentUpdateThread;
HANDLE g_webContentUpdateRefreshEvent;
HANDLE g_webContentUpdateStopEvent;
std::mutex g_webContentMutex;
std::atomic<bool> g_webContentLoaded;

std::vector<std::optional<std::wstring>> g_webContentStrings;
std::vector<std::optional<std::wstring>> g_webContentStringsFull;

// Kept for compatibility with old settings:
WCHAR g_webContent[FORMATTED_BUFFER_SIZE];
WCHAR g_webContentFull[FORMATTED_BUFFER_SIZE];

struct ClockElementStyleData {
    winrt::weak_ref<FrameworkElement> dateTimeIconContentElement;
    DWORD styleIndex;
    std::optional<int64_t> dateVisibilityPropertyChangedToken;
    std::optional<int64_t> timeVisibilityPropertyChangedToken;
};

std::atomic<bool> g_clockElementStyleEnabled;
std::atomic<DWORD> g_clockElementStyleIndex;
std::vector<ClockElementStyleData> g_clockElementStyleData;

using GetDpiForWindow_t = UINT(WINAPI*)(HWND hwnd);
GetDpiForWindow_t pGetDpiForWindow;

using GetLocalTime_t = decltype(&GetLocalTime);
GetLocalTime_t GetLocalTime_Original;

using GetTimeFormatEx_t = decltype(&GetTimeFormatEx);
GetTimeFormatEx_t GetTimeFormatEx_Original;

using GetDateFormatEx_t = decltype(&GetDateFormatEx);
GetDateFormatEx_t GetDateFormatEx_Original;

using GetDateFormatW_t = decltype(&GetDateFormatW);
GetDateFormatW_t GetDateFormatW_Original;

using SendMessageW_t = decltype(&SendMessageW);
SendMessageW_t SendMessageW_Original;

std::optional<std::wstring> GetUrlContent(PCWSTR lpUrl,
                                          bool failIfNot200 = true) {
    HINTERNET hOpenHandle = InternetOpen(
        L"WindhawkMod", INTERNET_OPEN_TYPE_PRECONFIG, nullptr, nullptr, 0);
    if (!hOpenHandle) {
        return std::nullopt;
    }

    HINTERNET hUrlHandle =
        InternetOpenUrl(hOpenHandle, lpUrl, nullptr, 0,
                        INTERNET_FLAG_NO_AUTH | INTERNET_FLAG_NO_CACHE_WRITE |
                            INTERNET_FLAG_NO_COOKIES | INTERNET_FLAG_NO_UI |
                            INTERNET_FLAG_PRAGMA_NOCACHE | INTERNET_FLAG_RELOAD,
                        0);
    if (!hUrlHandle) {
        InternetCloseHandle(hOpenHandle);
        return std::nullopt;
    }

    if (failIfNot200) {
        DWORD dwStatusCode = 0;
        DWORD dwStatusCodeSize = sizeof(dwStatusCode);
        if (!HttpQueryInfo(hUrlHandle,
                           HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER,
                           &dwStatusCode, &dwStatusCodeSize, nullptr) ||
            dwStatusCode != 200) {
            InternetCloseHandle(hUrlHandle);
            InternetCloseHandle(hOpenHandle);
            return std::nullopt;
        }
    }

    LPBYTE pUrlContent = (LPBYTE)HeapAlloc(GetProcessHeap(), 0, 0x400);
    if (!pUrlContent) {
        InternetCloseHandle(hUrlHandle);
        InternetCloseHandle(hOpenHandle);
        return std::nullopt;
    }

    DWORD dwNumberOfBytesRead;
    InternetReadFile(hUrlHandle, pUrlContent, 0x400, &dwNumberOfBytesRead);
    DWORD dwLength = dwNumberOfBytesRead;

    while (dwNumberOfBytesRead) {
        LPBYTE pNewUrlContent = (LPBYTE)HeapReAlloc(
            GetProcessHeap(), 0, pUrlContent, dwLength + 0x400);
        if (!pNewUrlContent) {
            InternetCloseHandle(hUrlHandle);
            InternetCloseHandle(hOpenHandle);
            HeapFree(GetProcessHeap(), 0, pUrlContent);
            return std::nullopt;
        }

        pUrlContent = pNewUrlContent;
        InternetReadFile(hUrlHandle, pUrlContent + dwLength, 0x400,
                         &dwNumberOfBytesRead);
        dwLength += dwNumberOfBytesRead;
    }

    InternetCloseHandle(hUrlHandle);
    InternetCloseHandle(hOpenHandle);

    // Assume UTF-8.
    int charsNeeded = MultiByteToWideChar(CP_UTF8, 0, (PCSTR)pUrlContent,
                                          dwLength, nullptr, 0);
    std::wstring unicodeContent(charsNeeded, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, (PCSTR)pUrlContent, dwLength,
                        unicodeContent.data(), unicodeContent.size());

    HeapFree(GetProcessHeap(), 0, pUrlContent);

    return unicodeContent;
}

int StringCopyTruncated(PWSTR dest,
                        size_t destSize,
                        PCWSTR src,
                        bool* truncated) {
    if (destSize == 0) {
        *truncated = *src;
        return 0;
    }

    size_t i;
    for (i = 0; i < destSize - 1 && *src; i++) {
        *dest++ = *src++;
    }

    *dest = L'\0';
    *truncated = *src;
    return i;
}

std::wstring ExtractWebContent(const std::wstring& webContent,
                               PCWSTR webContentsBlockStart,
                               PCWSTR webContentsStart,
                               PCWSTR webContentsEnd) {
    auto block = webContent.find(webContentsBlockStart);
    if (block == std::wstring::npos) {
        return std::wstring();
    }

    auto start = webContent.find(webContentsStart, block);
    if (start == std::wstring::npos) {
        return std::wstring();
    }

    start += wcslen(webContentsStart);

    auto end = webContent.find(webContentsEnd, start);
    if (end == std::wstring::npos) {
        return std::wstring();
    }

    return webContent.substr(start, end - start);
}

void UpdateWebContent() {
    int failed = 0;

    std::wstring lastUrl;
    std::optional<std::wstring> urlContent;

    // Kept for compatibility with old settings:
    if (g_settings.webContentsUrl && g_settings.webContentsBlockStart &&
        g_settings.webContentsStart && g_settings.webContentsEnd) {
        lastUrl = g_settings.webContentsUrl;
        urlContent =
            GetUrlContent(g_settings.webContentsUrl, /*failIfNot200=*/false);

        std::wstring extracted;
        if (urlContent) {
            extracted = ExtractWebContent(
                *urlContent, g_settings.webContentsBlockStart,
                g_settings.webContentsStart, g_settings.webContentsEnd);

            std::lock_guard<std::mutex> guard(g_webContentMutex);

            int maxLen = ARRAYSIZE(g_webContent) - 1;
            if (g_settings.webContentsMaxLength > 0 &&
                g_settings.webContentsMaxLength < maxLen) {
                maxLen = g_settings.webContentsMaxLength;
            }

            bool truncated;
            StringCopyTruncated(g_webContent, maxLen + 1, extracted.c_str(),
                                &truncated);
            if (truncated && maxLen >= 3) {
                g_webContent[maxLen - 1] = L'.';
                g_webContent[maxLen - 2] = L'.';
                g_webContent[maxLen - 3] = L'.';
            }

            maxLen = ARRAYSIZE(g_webContentFull) - 1;
            StringCopyTruncated(g_webContentFull, maxLen + 1, extracted.c_str(),
                                &truncated);
            if (truncated && maxLen >= 3) {
                g_webContentFull[maxLen - 1] = L'.';
                g_webContentFull[maxLen - 2] = L'.';
                g_webContentFull[maxLen - 3] = L'.';
            }
        } else {
            failed++;
        }
    }

    for (size_t i = 0; i < g_settings.webContentsItems.size(); i++) {
        const auto& item = g_settings.webContentsItems[i];

        if (item.url.get() != lastUrl) {
            lastUrl = item.url;
            urlContent = GetUrlContent(item.url, /*failIfNot200=*/false);
        }

        if (!urlContent) {
            failed++;
            continue;
        }

        std::wstring extracted = ExtractWebContent(*urlContent, item.blockStart,
                                                   item.start, item.end);

        std::lock_guard<std::mutex> guard(g_webContentMutex);

        if (item.maxLength <= 0 ||
            extracted.length() <= (size_t)item.maxLength) {
            g_webContentStrings[i] = extracted;
        } else {
            std::wstring truncated(extracted.begin(),
                                   extracted.begin() + item.maxLength);
            if (truncated.length() >= 3) {
                truncated[truncated.length() - 1] = L'.';
                truncated[truncated.length() - 2] = L'.';
                truncated[truncated.length() - 3] = L'.';
            }

            g_webContentStrings[i] = std::move(truncated);
        }

        g_webContentStringsFull[i] = std::move(extracted);
    }

    if (failed == 0) {
        g_webContentLoaded = true;
    }
}

DWORD WINAPI WebContentUpdateThread(LPVOID lpThreadParameter) {
    constexpr DWORD kSecondsForQuickRetry = 30;

    HANDLE handles[] = {
        g_webContentUpdateStopEvent,
        g_webContentUpdateRefreshEvent,
    };

    while (true) {
        UpdateWebContent();

        DWORD seconds = g_settings.webContentsUpdateInterval >= 1
                            ? g_settings.webContentsUpdateInterval * 60
                            : 1;
        if (!g_webContentLoaded && seconds > kSecondsForQuickRetry) {
            seconds = kSecondsForQuickRetry;
        }

        DWORD dwWaitResult = WaitForMultipleObjects(ARRAYSIZE(handles), handles,
                                                    FALSE, seconds * 1000);

        if (dwWaitResult == WAIT_FAILED) {
            Wh_Log(L"WAIT_FAILED");
            break;
        }

        if (dwWaitResult == WAIT_OBJECT_0) {
            break;
        }
    }

    return 0;
}

void WebContentUpdateThreadInit() {
    // A fuzzy check to see if any of the lines contain the web content pattern.
    // If not, no need to fire up the thread.
    if (!wcsstr(g_settings.topLine, L"%web") &&
        !wcsstr(g_settings.bottomLine, L"%web") &&
        !wcsstr(g_settings.middleLine, L"%web") &&
        !wcsstr(g_settings.tooltipLine, L"%web")) {
        return;
    }

    std::lock_guard<std::mutex> guard(g_webContentMutex);

    g_webContentLoaded = false;

    *g_webContent = L'\0';
    *g_webContentFull = L'\0';

    g_webContentStrings.clear();
    g_webContentStrings.resize(g_settings.webContentsItems.size());
    g_webContentStringsFull.clear();
    g_webContentStringsFull.resize(g_settings.webContentsItems.size());

    if ((g_settings.webContentsUrl && *g_settings.webContentsUrl) ||
        g_settings.webContentsItems.size() > 0) {
        g_webContentUpdateRefreshEvent =
            CreateEvent(nullptr, FALSE, FALSE, nullptr);
        g_webContentUpdateStopEvent =
            CreateEvent(nullptr, TRUE, FALSE, nullptr);
        g_webContentUpdateThread = CreateThread(
            nullptr, 0, WebContentUpdateThread, nullptr, 0, nullptr);
    }
}

void WebContentUpdateThreadUninit() {
    if (g_webContentUpdateThread) {
        SetEvent(g_webContentUpdateStopEvent);
        WaitForSingleObject(g_webContentUpdateThread, INFINITE);
        CloseHandle(g_webContentUpdateThread);
        g_webContentUpdateThread = nullptr;
        CloseHandle(g_webContentUpdateRefreshEvent);
        g_webContentUpdateRefreshEvent = nullptr;
        CloseHandle(g_webContentUpdateStopEvent);
        g_webContentUpdateStopEvent = nullptr;
    }
}

int CalculateWeeknum(const SYSTEMTIME* time, DWORD startDayOfWeek) {
    SYSTEMTIME secondWeek{
        .wYear = time->wYear,
        .wMonth = 1,
        .wDay = 1,
    };

    // Calculate wDayOfWeek.
    FILETIME fileTime;
    SystemTimeToFileTime(&secondWeek, &fileTime);
    FileTimeToSystemTime(&fileTime, &secondWeek);

    do {
        secondWeek.wDay++;
        secondWeek.wDayOfWeek = (secondWeek.wDayOfWeek + 1) % 7;
    } while (secondWeek.wDayOfWeek != startDayOfWeek);

    FILETIME targetFileTime;
    SystemTimeToFileTime(time, &targetFileTime);
    ULARGE_INTEGER targetFileTimeInt{
        .LowPart = targetFileTime.dwLowDateTime,
        .HighPart = targetFileTime.dwHighDateTime,
    };

    FILETIME secondWeekFileTime;
    SystemTimeToFileTime(&secondWeek, &secondWeekFileTime);
    ULARGE_INTEGER secondWeekFileTimeInt{
        .LowPart = secondWeekFileTime.dwLowDateTime,
        .HighPart = secondWeekFileTime.dwHighDateTime,
    };

    int weeknum = 1;
    if (targetFileTimeInt.QuadPart >= secondWeekFileTimeInt.QuadPart) {
        ULONGLONG diff =
            targetFileTimeInt.QuadPart - secondWeekFileTimeInt.QuadPart;
        ULONGLONG weekIn100Ns = 10000000ULL * 60 * 60 * 24 * 7;
        weeknum += 1 + diff / weekIn100Ns;
    }

    return weeknum;
}

// Adopted from VMime:
// https://github.com/kisli/vmime/blob/fc69321d5304c73be685c890f3b30528aadcfeaf/src/vmime/utility/datetimeUtils.cpp#L239
int CalculateWeeknumIso(const SYSTEMTIME* time) {
    const int year = time->wYear;
    const int month = time->wMonth;
    const int day = time->wDay;
    const bool iso = true;

    // Algorithm from http://personal.ecu.edu/mccartyr/ISOwdALG.txt

    const bool leapYear =
        ((year % 4) == 0 && (year % 100) != 0) || (year % 400) == 0;
    const bool leapYear_1 =
        (((year - 1) % 4) == 0 && ((year - 1) % 100) != 0) ||
        ((year - 1) % 400) == 0;

    // 4. Find the DayOfYearNumber for Y M D
    static const int DAY_OF_YEAR_NUMBER_MAP[12] = {
        0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334};

    int DayOfYearNumber = day + DAY_OF_YEAR_NUMBER_MAP[month - 1];

    if (leapYear && month > 2) {
        DayOfYearNumber += 1;
    }

    // 5. Find the Jan1Weekday for Y (Monday=1, Sunday=7)
    const int YY = (year - 1) % 100;
    const int C = (year - 1) - YY;
    const int G = YY + YY / 4;
    const int Jan1Weekday = 1 + (((((C / 100) % 4) * 5) + G) % 7);

    // 6. Find the Weekday for Y M D
    const int H = DayOfYearNumber + (Jan1Weekday - 1);
    const int Weekday = 1 + ((H - 1) % 7);

    // 7. Find if Y M D falls in YearNumber Y-1, WeekNumber 52 or 53
    int YearNumber = 0, WeekNumber = 0;

    if (DayOfYearNumber <= (8 - Jan1Weekday) && Jan1Weekday > 4) {
        YearNumber = year - 1;

        if (Jan1Weekday == 5 || (Jan1Weekday == 6 && leapYear_1)) {
            WeekNumber = 53;
        } else {
            WeekNumber = 52;
        }

    } else {
        YearNumber = year;
    }

    // 8. Find if Y M D falls in YearNumber Y+1, WeekNumber 1
    if (YearNumber == year) {
        const int I = (leapYear ? 366 : 365);

        if ((I - DayOfYearNumber) < (4 - Weekday)) {
            YearNumber = year + 1;
            WeekNumber = 1;
        }
    }

    // 9. Find if Y M D falls in YearNumber Y, WeekNumber 1 through 53
    if (YearNumber == year) {
        const int J = DayOfYearNumber + (7 - Weekday) + (Jan1Weekday - 1);

        WeekNumber = J / 7;

        if (Jan1Weekday > 4) {
            WeekNumber -= 1;
        }
    }

    if (!iso && (WeekNumber == 1 && month == 12)) {
        WeekNumber = 53;
    }

    return WeekNumber;
}

// Adopted from VMime:
// https://github.com/kisli/vmime/blob/fc69321d5304c73be685c890f3b30528aadcfeaf/src/vmime/utility/datetimeUtils.cpp#L239
int CalculateDayOfYearNumber(const SYSTEMTIME* time) {
    const int year = time->wYear;
    const int month = time->wMonth;
    const int day = time->wDay;

    const bool leapYear =
        ((year % 4) == 0 && (year % 100) != 0) || (year % 400) == 0;

    static const int DAY_OF_YEAR_NUMBER_MAP[12] = {
        0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334};

    int DayOfYearNumber = day + DAY_OF_YEAR_NUMBER_MAP[month - 1];

    if (leapYear && month > 2) {
        DayOfYearNumber += 1;
    }

    return DayOfYearNumber;
}

// Adopted from:
// https://github.com/microsoft/cpp_client_telemetry/blob/25bc0806f21ecb2587154494f073bfa581cd5089/lib/pal/desktop/WindowsEnvironmentInfo.hpp#L39
void GetTimeZone(WCHAR* buffer, size_t bufferSize) {
    long bias;

    TIME_ZONE_INFORMATION timeZone = {};
    if (GetTimeZoneInformation(&timeZone) == TIME_ZONE_ID_DAYLIGHT) {
        bias = timeZone.Bias + timeZone.DaylightBias;
    } else {
        // TODO: [MG] - ref.
        // https://docs.microsoft.com/en-us/windows/win32/api/timezoneapi/nf-timezoneapi-gettimezoneinformation
        // Need to handle the case when API return TIME_ZONE_ID_UNKNOWN.
        // Otherwise we may be reporting invalid timeZone.Bias
        bias = timeZone.Bias + timeZone.StandardBias;
    }

    auto hours = (long long)abs(bias) / 60;
    auto minutes = (long long)abs(bias) % 60;

    // UTC = local time + bias; bias sign should be interved.
    _snwprintf_s(buffer, bufferSize, _TRUNCATE, L"%c%02d:%02d",
                 bias <= 0 ? L'+' : L'-', static_cast<int>(hours),
                 static_cast<int>(minutes));
}

std::vector<std::wstring> SplitTimeFormatString(std::wstring_view s) {
    size_t posStart = 0;
    std::wstring token;
    std::vector<std::wstring> result;

    while (true) {
        size_t posEnd = s.size();
        bool inQuote = false;
        for (size_t i = posStart; i < s.size(); i++) {
            if (s[i] == L'\'') {
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && s[i] == L';') {
                posEnd = i;
                break;
            }
        }

        if (posEnd == s.size()) {
            break;
        }

        token = s.substr(posStart, posEnd - posStart);
        posStart = posEnd + 1;
        result.push_back(std::move(token));
    }

    token = s.substr(posStart);
    result.push_back(std::move(token));
    return result;
}

void InitializeFormattedStrings(const SYSTEMTIME* time) {
    auto timeFormatParts = SplitTimeFormatString(g_settings.timeFormat.get());

    GetTimeFormatEx_Original(
        nullptr, g_settings.showSeconds ? 0 : TIME_NOSECONDS, time,
        !timeFormatParts[0].empty() ? timeFormatParts[0].c_str() : nullptr,
        g_timeFormatted, ARRAYSIZE(g_timeFormatted));

    g_timeExtraFormatted.resize(timeFormatParts.size() - 1);
    for (size_t i = 1; i < timeFormatParts.size(); i++) {
        WCHAR formatted[FORMATTED_BUFFER_SIZE];
        GetTimeFormatEx_Original(
            nullptr, g_settings.showSeconds ? 0 : TIME_NOSECONDS, time,
            !timeFormatParts[i].empty() ? timeFormatParts[i].c_str() : nullptr,
            formatted, ARRAYSIZE(formatted));
        g_timeExtraFormatted[i - 1] = formatted;
    }

    auto dateFormatParts = SplitTimeFormatString(g_settings.dateFormat.get());

    GetDateFormatEx_Original(
        nullptr, DATE_AUTOLAYOUT, time,
        !dateFormatParts[0].empty() ? dateFormatParts[0].c_str() : nullptr,
        g_dateFormatted, ARRAYSIZE(g_dateFormatted), nullptr);

    g_dateExtraFormatted.resize(dateFormatParts.size() - 1);
    for (size_t i = 1; i < dateFormatParts.size(); i++) {
        WCHAR formatted[FORMATTED_BUFFER_SIZE];
        GetDateFormatEx_Original(
            nullptr, DATE_AUTOLAYOUT, time,
            !dateFormatParts[i].empty() ? dateFormatParts[i].c_str() : nullptr,
            formatted, ARRAYSIZE(formatted), nullptr);
        g_dateExtraFormatted[i - 1] = formatted;
    }

    auto weekdayFormatParts =
        SplitTimeFormatString(g_settings.weekdayFormat.get());

    GetDateFormatEx_Original(
        nullptr, DATE_AUTOLAYOUT, time,
        !weekdayFormatParts[0].empty() ? weekdayFormatParts[0].c_str()
                                       : nullptr,
        g_weekdayFormatted, ARRAYSIZE(g_weekdayFormatted), nullptr);

    g_weekdayExtraFormatted.resize(weekdayFormatParts.size() - 1);
    for (size_t i = 1; i < weekdayFormatParts.size(); i++) {
        WCHAR formatted[FORMATTED_BUFFER_SIZE];
        GetDateFormatEx_Original(nullptr, DATE_AUTOLAYOUT, time,
                                 !weekdayFormatParts[i].empty()
                                     ? weekdayFormatParts[i].c_str()
                                     : nullptr,
                                 formatted, ARRAYSIZE(formatted), nullptr);
        g_weekdayExtraFormatted[i - 1] = formatted;
    }

    // https://stackoverflow.com/a/39344961
    DWORD startDayOfWeek;
    GetLocaleInfoEx(
        LOCALE_NAME_USER_DEFAULT, LOCALE_IFIRSTDAYOFWEEK | LOCALE_RETURN_NUMBER,
        (PWSTR)&startDayOfWeek, sizeof(startDayOfWeek) / sizeof(WCHAR));

    // Start from Sunday instead of Monday.
    startDayOfWeek = (startDayOfWeek + 1) % 7;

    swprintf_s(g_weekdayNumFormatted, L"%d",
               1 + (7 + time->wDayOfWeek - startDayOfWeek) % 7);

    swprintf_s(g_weeknumFormatted, L"%d",
               CalculateWeeknum(time, startDayOfWeek));

    swprintf_s(g_weeknumIsoFormatted, L"%d", CalculateWeeknumIso(time));

    swprintf_s(g_dayOfYearFormatted, L"%d", CalculateDayOfYearNumber(time));

    GetTimeZone(g_timezoneFormatted, ARRAYSIZE(g_timezoneFormatted));
}

int ResolveFormatTokenWithDigit(std::wstring_view format,
                                std::wstring_view formatTokenPrefix,
                                std::wstring_view formatTokenSuffix) {
    if (format.size() <
        formatTokenPrefix.size() + 1 + formatTokenSuffix.size()) {
        return 0;
    }

    if (!format.starts_with(formatTokenPrefix)) {
        return 0;
    }

    char digitChar = format[formatTokenPrefix.size()];
    if (digitChar < L'1' || digitChar > L'9') {
        return 0;
    }

    if (!format.substr(formatTokenPrefix.size() + 1)
             .starts_with(formatTokenSuffix)) {
        return 0;
    }

    return digitChar - L'0';
}

size_t ResolveFormatToken(std::wstring_view format, PCWSTR* resolved) {
    struct {
        std::wstring_view token;
        PCWSTR value;
    } formatTokens[] = {
        {L"%time%"sv, g_timeFormatted},
        {L"%date%"sv, g_dateFormatted},
        {L"%weekday%"sv, g_weekdayFormatted},
        {L"%weekday_num%"sv, g_weekdayNumFormatted},
        {L"%weeknum%"sv, g_weeknumFormatted},
        {L"%weeknum_iso%"sv, g_weeknumIsoFormatted},
        {L"%dayofyear%"sv, g_dayOfYearFormatted},
        {L"%timezone%"sv, g_timezoneFormatted},
        {L"%newline%"sv, L"\n"},
    };

    for (const auto& formatToken : formatTokens) {
        if (format.starts_with(formatToken.token)) {
            *resolved = formatToken.value;
            return formatToken.token.size();
        }
    }

    if (auto token = L"%web%"sv; format.starts_with(token)) {
        std::lock_guard<std::mutex> guard(g_webContentMutex);
        *resolved = *g_webContent ? g_webContent : L"Loading...";
        return token.size();
    }

    if (auto token = L"%web_full%"sv; format.starts_with(token)) {
        std::lock_guard<std::mutex> guard(g_webContentMutex);
        *resolved = *g_webContentFull ? g_webContentFull : L"Loading...";
        return token.size();
    }

    struct {
        std::wstring_view prefix;
        const std::vector<std::wstring>& valueVector;
    } formatExtraTokens[] = {
        {L"%time"sv, g_timeExtraFormatted},
        {L"%date"sv, g_dateExtraFormatted},
        {L"%weekday"sv, g_weekdayExtraFormatted},
    };

    for (auto formatExtraToken : formatExtraTokens) {
        int digit = ResolveFormatTokenWithDigit(format, formatExtraToken.prefix,
                                                L"%"sv);
        if (!digit) {
            continue;
        }

        const auto& valueVector = formatExtraToken.valueVector;

        if (digit < 2 || static_cast<size_t>(digit - 2) >= valueVector.size()) {
            *resolved = L"-";
        } else {
            *resolved = valueVector[digit - 2].c_str();
        }

        return formatExtraToken.prefix.size() + 2;
    }

    if (int digit = ResolveFormatTokenWithDigit(format, L"%web"sv, L"%"sv)) {
        size_t index = digit - 1;

        std::lock_guard<std::mutex> guard(g_webContentMutex);

        if (index >= g_webContentStrings.size()) {
            *resolved = L"-";
        } else if (!g_webContentStrings[index]) {
            *resolved = L"Loading...";
        } else {
            *resolved = g_webContentStrings[index]->c_str();
        }

        return "%web1%"sv.size();
    }

    if (int digit =
            ResolveFormatTokenWithDigit(format, L"%web"sv, L"_full%"sv)) {
        size_t index = digit - 1;

        std::lock_guard<std::mutex> guard(g_webContentMutex);

        if (index >= g_webContentStringsFull.size()) {
            *resolved = L"-";
        } else if (!g_webContentStringsFull[index]) {
            *resolved = L"Loading...";
        } else {
            *resolved = g_webContentStringsFull[index]->c_str();
        }

        return "%web1_full%"sv.size();
    }

    return 0;
}

int FormatLine(PWSTR buffer, size_t bufferSize, std::wstring_view format) {
    if (bufferSize == 0) {
        return 0;
    }

    std::wstring_view formatSuffix = format;

    PWSTR bufferStart = buffer;
    PWSTR bufferEnd = bufferStart + bufferSize;
    while (!formatSuffix.empty() && bufferEnd - buffer > 1) {
        if (formatSuffix[0] == L'%') {
            PCWSTR srcStr = nullptr;
            size_t formatTokenLen = ResolveFormatToken(formatSuffix, &srcStr);
            if (formatTokenLen > 0) {
                bool truncated;
                buffer += StringCopyTruncated(buffer, bufferEnd - buffer,
                                              srcStr, &truncated);
                if (truncated) {
                    break;
                }

                formatSuffix = formatSuffix.substr(formatTokenLen);
                continue;
            }
        }

        *buffer++ = formatSuffix[0];
        formatSuffix = formatSuffix.substr(1);
    }

    if (!formatSuffix.empty() && bufferSize >= 4) {
        buffer[-1] = L'.';
        buffer[-2] = L'.';
        buffer[-3] = L'.';
    }

    *buffer = L'\0';

    return buffer - bufferStart;
}

#pragma region Win11Hooks

DWORD g_refreshIconThreadId;
bool g_refreshIconNeedToAdjustTimer;
bool g_inGetTimeToolTipString;

using ClockSystemTrayIconDataModel_RefreshIcon_t = void(WINAPI*)(
    LPVOID pThis,
    LPVOID  // SystemTrayTelemetry::ClockUpdate&
);
ClockSystemTrayIconDataModel_RefreshIcon_t
    ClockSystemTrayIconDataModel_RefreshIcon_Original;

using ClockSystemTrayIconDataModel_GetTimeToolTipString_t =
    LPVOID(WINAPI*)(LPVOID pThis, LPVOID, LPVOID, LPVOID, LPVOID);
ClockSystemTrayIconDataModel_GetTimeToolTipString_t
    ClockSystemTrayIconDataModel_GetTimeToolTipString_Original;

using ClockSystemTrayIconDataModel_GetTimeToolTipString2_t =
    LPVOID(WINAPI*)(LPVOID pThis, LPVOID, LPVOID, LPVOID, LPVOID);
ClockSystemTrayIconDataModel_GetTimeToolTipString2_t
    ClockSystemTrayIconDataModel_GetTimeToolTipString2_Original;

using DateTimeIconContent_OnApplyTemplate_t = void(WINAPI*)(LPVOID pThis);
DateTimeIconContent_OnApplyTemplate_t
    DateTimeIconContent_OnApplyTemplate_Original;

using BadgeIconContent_get_ViewModel_t = HRESULT(WINAPI*)(LPVOID pThis,
                                                          LPVOID pArgs);
BadgeIconContent_get_ViewModel_t BadgeIconContent_get_ViewModel_Original;

using ClockSystemTrayIconDataModel_GetTimeToolTipString_2_t =
    LPVOID(WINAPI*)(LPVOID pThis, LPVOID, LPVOID, LPVOID, LPVOID);
ClockSystemTrayIconDataModel_GetTimeToolTipString_2_t
    ClockSystemTrayIconDataModel_GetTimeToolTipString_2_Original;

using ICalendar_Second_t = int(WINAPI*)(LPVOID pThis);
ICalendar_Second_t ICalendar_Second_Original;

using ThreadPoolTimer_CreateTimer_t = LPVOID(WINAPI*)(LPVOID param1,
                                                      LPVOID param2,
                                                      LPVOID param3,
                                                      LPVOID param4);
ThreadPoolTimer_CreateTimer_t ThreadPoolTimer_CreateTimer_Original;

using ThreadPoolTimer_CreateTimer_lambda_t = LPVOID(WINAPI*)(DWORD_PTR** param1,
                                                             LPVOID param2,
                                                             LPVOID param3);
ThreadPoolTimer_CreateTimer_lambda_t
    ThreadPoolTimer_CreateTimer_lambda_Original;

void WINAPI ClockSystemTrayIconDataModel_RefreshIcon_Hook(LPVOID pThis,
                                                          LPVOID param1) {
    Wh_Log(L">");

    g_refreshIconThreadId = GetCurrentThreadId();
    g_refreshIconNeedToAdjustTimer =
        g_settings.showSeconds || !g_webContentLoaded;

    ClockSystemTrayIconDataModel_RefreshIcon_Original(pThis, param1);

    g_refreshIconThreadId = 0;
    g_refreshIconNeedToAdjustTimer = false;
}

void UpdateToolTipString(LPVOID tooltipPtrPtr) {
    auto separator = L"\r\n\r\n"sv;

    WCHAR extraLine[256];
    size_t extraLength = FormatLine(extraLine, ARRAYSIZE(extraLine),
                                    g_settings.tooltipLine.get());
    if (extraLength == 0) {
        return;
    }

    // Reference:
    // https://github.com/microsoft/cppwinrt/blob/0a6cb062e2151cf6c8f357aa8ef735e359f8a98c/strings/base_string.h
    struct hstring_header {
        uint32_t flags;
        uint32_t length;
        uint32_t padding1;
        uint32_t padding2;
        wchar_t const* ptr;
    };

    struct shared_hstring_header : hstring_header {
        int32_t /*atomic_ref_count*/ count;
        wchar_t buffer[1];
    };

    shared_hstring_header* tooltipHeader =
        *(shared_hstring_header**)tooltipPtrPtr;

    uint64_t bytesRequired =
        sizeof(shared_hstring_header) +
        sizeof(wchar_t) *
            (tooltipHeader->length + separator.length() + extraLength);

    shared_hstring_header* tooltipHeaderNew =
        (shared_hstring_header*)HeapReAlloc(GetProcessHeap(), 0, tooltipHeader,
                                            bytesRequired);
    if (!tooltipHeaderNew) {
        return;
    }

    tooltipHeaderNew->ptr = tooltipHeaderNew->buffer;

    memcpy(tooltipHeaderNew->buffer + tooltipHeaderNew->length,
           separator.data(), sizeof(wchar_t) * separator.length());
    tooltipHeaderNew->length += separator.length();

    memcpy(tooltipHeaderNew->buffer + tooltipHeaderNew->length, extraLine,
           sizeof(wchar_t) * extraLength);
    tooltipHeaderNew->length += extraLength;

    tooltipHeaderNew->buffer[tooltipHeaderNew->length] = L'\0';

    *(shared_hstring_header**)tooltipPtrPtr = tooltipHeaderNew;
}

LPVOID WINAPI
ClockSystemTrayIconDataModel_GetTimeToolTipString_Hook(LPVOID pThis,
                                                       LPVOID param1,
                                                       LPVOID param2,
                                                       LPVOID param3,
                                                       LPVOID param4) {
    Wh_Log(L">");

    g_inGetTimeToolTipString = true;

    LPVOID ret = ClockSystemTrayIconDataModel_GetTimeToolTipString_Original(
        pThis, param1, param2, param3, param4);

    UpdateToolTipString(ret);

    g_inGetTimeToolTipString = false;

    return ret;
}

LPVOID WINAPI
ClockSystemTrayIconDataModel_GetTimeToolTipString2_Hook(LPVOID pThis,
                                                        LPVOID param1,
                                                        LPVOID param2,
                                                        LPVOID param3,
                                                        LPVOID param4) {
    Wh_Log(L">");

    g_inGetTimeToolTipString = true;

    LPVOID ret = ClockSystemTrayIconDataModel_GetTimeToolTipString2_Original(
        pThis, param1, param2, param3, param4);

    UpdateToolTipString(ret);

    g_inGetTimeToolTipString = false;

    return ret;
}

FrameworkElement FindChildByName(FrameworkElement element,
                                 winrt::hstring name) {
    int childrenCount = Media::VisualTreeHelper::GetChildrenCount(element);

    for (int i = 0; i < childrenCount; i++) {
        auto child = Media::VisualTreeHelper::GetChild(element, i)
                         .try_as<FrameworkElement>();
        if (!child) {
            Wh_Log(L"Failed to get child %d of %d", i + 1, childrenCount);
            continue;
        }

        if (child.Name() == name) {
            return child;
        }
    }

    return nullptr;
}

void ApplyStackPanelStyles(Controls::StackPanel stackPanel,
                           int maxWidth,
                           int textSpacing) {
    if (maxWidth) {
        stackPanel.MaxWidth(maxWidth);
    } else {
        stackPanel.as<DependencyObject>().ClearValue(
            FrameworkElement::MaxWidthProperty());
    }

    if (textSpacing) {
        stackPanel.Spacing(textSpacing);
    } else {
        stackPanel.as<DependencyObject>().ClearValue(
            Controls::StackPanel::SpacingProperty());
    }
}

void ApplyTextBlockStyles(
    Controls::TextBlock textBlock,
    const TextStyleSettings* textStyleSettings,
    bool noWrap,
    std::optional<int64_t>* visibilityPropertyChangedToken) {
    if (visibilityPropertyChangedToken->has_value()) {
        textBlock.UnregisterPropertyChangedCallback(
            UIElement::VisibilityProperty(),
            visibilityPropertyChangedToken->value());
    }

    if (textStyleSettings && !textStyleSettings->visible) {
        textBlock.Visibility(Visibility::Collapsed);
        *visibilityPropertyChangedToken =
            textBlock.RegisterPropertyChangedCallback(
                UIElement::VisibilityProperty(),
                [](DependencyObject sender, DependencyProperty property) {
                    auto textBlock = sender.try_as<Controls::TextBlock>();
                    if (!textBlock) {
                        return;
                    }

                    textBlock.Visibility(Visibility::Collapsed);
                });
        return;
    }

    visibilityPropertyChangedToken->reset();

    textBlock.Visibility(Visibility::Visible);

    if (noWrap) {
        textBlock.TextWrapping(TextWrapping::NoWrap);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::TextWrappingProperty());
    }

    if (textStyleSettings && *textStyleSettings->textColor) {
        auto textColor =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<winrt::Windows::UI::Color>(),
                winrt::box_value(textStyleSettings->textColor.get()))
                .as<winrt::Windows::UI::Color>();
        textBlock.Foreground(Media::SolidColorBrush{textColor});
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::ForegroundProperty());
    }

    if (textStyleSettings && *textStyleSettings->textAlignment) {
        auto textAlignment =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<TextAlignment>(),
                winrt::box_value(textStyleSettings->textAlignment.get()))
                .as<TextAlignment>();
        textBlock.TextAlignment(textAlignment);
    } else {
        textBlock.TextAlignment(TextAlignment::End);
    }

    if (textStyleSettings && textStyleSettings->fontSize) {
        textBlock.FontSize(textStyleSettings->fontSize);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::FontSizeProperty());
    }

    if (textStyleSettings && *textStyleSettings->fontFamily) {
        auto fontFamily =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<Media::FontFamily>(),
                winrt::box_value(textStyleSettings->fontFamily.get()))
                .as<Media::FontFamily>();
        textBlock.FontFamily(fontFamily);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::FontFamilyProperty());
    }

    if (textStyleSettings && *textStyleSettings->fontWeight) {
        auto fontWeight =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<winrt::Windows::UI::Text::FontWeight>(),
                winrt::box_value(textStyleSettings->fontWeight.get()))
                .as<winrt::Windows::UI::Text::FontWeight>();
        textBlock.FontWeight(fontWeight);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::FontWeightProperty());
    }

    if (textStyleSettings && *textStyleSettings->fontStyle) {
        auto fontStyle =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<winrt::Windows::UI::Text::FontStyle>(),
                winrt::box_value(textStyleSettings->fontStyle.get()))
                .as<winrt::Windows::UI::Text::FontStyle>();
        textBlock.FontStyle(fontStyle);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::FontStyleProperty());
    }

    if (textStyleSettings && *textStyleSettings->fontStretch) {
        auto fontStretch =
            Markup::XamlBindingHelper::ConvertValue(
                winrt::xaml_typename<winrt::Windows::UI::Text::FontStretch>(),
                winrt::box_value(textStyleSettings->fontStretch.get()))
                .as<winrt::Windows::UI::Text::FontStretch>();
        textBlock.FontStretch(fontStretch);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::FontStretchProperty());
    }

    if (textStyleSettings && textStyleSettings->characterSpacing) {
        textBlock.CharacterSpacing(textStyleSettings->characterSpacing);
    } else {
        textBlock.as<DependencyObject>().ClearValue(
            Controls::TextBlock::CharacterSpacingProperty());
    }
}

void ApplyDateTimeIconContentStyles(
    FrameworkElement dateTimeIconContentElement) {
    ClockElementStyleData* clockElementStyleData = nullptr;

    for (auto it = g_clockElementStyleData.begin();
         it != g_clockElementStyleData.end();) {
        auto& data = *it;
        auto element = data.dateTimeIconContentElement.get();
        if (!element) {
            it = g_clockElementStyleData.erase(it);
            continue;
        }

        if (element == dateTimeIconContentElement) {
            clockElementStyleData = &data;
            break;
        }

        ++it;
    }

    bool clockElementStyleEnabled = g_clockElementStyleEnabled;
    DWORD clockElementStyleIndex = g_clockElementStyleIndex;

    if (!clockElementStyleData && !clockElementStyleEnabled) {
        return;
    }

    if (clockElementStyleData &&
        clockElementStyleData->styleIndex == clockElementStyleIndex) {
        return;
    }

    auto containerGridElement =
        FindChildByName(dateTimeIconContentElement, L"ContainerGrid")
            .as<Controls::Grid>();
    if (!containerGridElement) {
        return;
    }

    auto stackPanel =
        containerGridElement.Children().GetAt(0).as<Controls::StackPanel>();

    Controls::TextBlock dateInnerTextBlock = nullptr;
    Controls::TextBlock timeInnerTextBlock = nullptr;

    for (const auto& child : stackPanel.Children()) {
        auto childTextBlock = child.try_as<Controls::TextBlock>();
        if (!childTextBlock) {
            continue;
        }

        if (childTextBlock.Name() == L"DateInnerTextBlock") {
            dateInnerTextBlock = childTextBlock;
            continue;
        }

        if (childTextBlock.Name() == L"TimeInnerTextBlock") {
            timeInnerTextBlock = childTextBlock;
            continue;
        }
    }

    if (!dateInnerTextBlock || !timeInnerTextBlock) {
        return;
    }

    if (!clockElementStyleData) {
        g_clockElementStyleData.push_back(ClockElementStyleData{
            .dateTimeIconContentElement = dateTimeIconContentElement,
        });
        clockElementStyleData = &g_clockElementStyleData.back();
    }

    int maxWidth = clockElementStyleEnabled ? g_settings.maxWidth : 0;
    int textSpacing = clockElementStyleEnabled ? g_settings.textSpacing : 0;
    bool noWrap = maxWidth;

    ApplyStackPanelStyles(stackPanel, maxWidth, textSpacing);
    ApplyTextBlockStyles(
        dateInnerTextBlock,
        clockElementStyleEnabled ? &g_settings.dateStyle : nullptr, noWrap,
        &clockElementStyleData->dateVisibilityPropertyChangedToken);
    ApplyTextBlockStyles(
        timeInnerTextBlock,
        clockElementStyleEnabled ? &g_settings.timeStyle : nullptr, noWrap,
        &clockElementStyleData->timeVisibilityPropertyChangedToken);

    clockElementStyleData->styleIndex = clockElementStyleIndex;
}

void WINAPI DateTimeIconContent_OnApplyTemplate_Hook(LPVOID pThis) {
    Wh_Log(L">");

    DateTimeIconContent_OnApplyTemplate_Original(pThis);

    IUnknown* dateTimeIconContentElementIUnknownPtr = *((IUnknown**)pThis + 1);
    if (!dateTimeIconContentElementIUnknownPtr) {
        return;
    }

    FrameworkElement dateTimeIconContentElement = nullptr;
    dateTimeIconContentElementIUnknownPtr->QueryInterface(
        winrt::guid_of<FrameworkElement>(),
        winrt::put_abi(dateTimeIconContentElement));
    if (!dateTimeIconContentElement) {
        return;
    }

    try {
        ApplyDateTimeIconContentStyles(dateTimeIconContentElement);
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        Wh_Log(L"Error %08X", hr);
    }
}

HRESULT WINAPI BadgeIconContent_get_ViewModel_Hook(LPVOID pThis, LPVOID pArgs) {
    // Wh_Log(L">");

    HRESULT ret = BadgeIconContent_get_ViewModel_Original(pThis, pArgs);

    try {
        winrt::Windows::Foundation::IInspectable obj = nullptr;
        winrt::check_hresult(
            ((IUnknown*)pThis)
                ->QueryInterface(
                    winrt::guid_of<winrt::Windows::Foundation::IInspectable>(),
                    winrt::put_abi(obj)));

        if (winrt::get_class_name(obj) == L"SystemTray.DateTimeIconContent") {
            auto dateTimeIconContentElement = obj.as<FrameworkElement>();
            if (dateTimeIconContentElement.IsLoaded()) {
                ApplyDateTimeIconContentStyles(dateTimeIconContentElement);
            }
        }
    } catch (...) {
        HRESULT hr = winrt::to_hresult();
        Wh_Log(L"Error %08X", hr);
    }

    return ret;
}

LPVOID WINAPI
ClockSystemTrayIconDataModel_GetTimeToolTipString_2_Hook(LPVOID pThis,
                                                         LPVOID param1,
                                                         LPVOID param2,
                                                         LPVOID param3,
                                                         LPVOID param4) {
    Wh_Log(L">");

    g_inGetTimeToolTipString = true;

    LPVOID ret = ClockSystemTrayIconDataModel_GetTimeToolTipString_2_Original(
        pThis, param1, param2, param3, param4);

    UpdateToolTipString(ret);

    g_inGetTimeToolTipString = false;

    return ret;
}

int WINAPI ICalendar_Second_Hook(LPVOID pThis) {
    Wh_Log(L">");

    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString && g_refreshIconNeedToAdjustTimer) {
        g_refreshIconNeedToAdjustTimer = false;

        // Make the next refresh happen in a second.
        return 59;
    }

    int ret = ICalendar_Second_Original(pThis);

    return ret;
}

// Similar to ThreadPoolTimer_CreateTimer_lambda_Hook. Only one of them is
// called in some Windows versions due to function inlining.
LPVOID WINAPI ThreadPoolTimer_CreateTimer_Hook(LPVOID param1,
                                               LPVOID param2,
                                               LPVOID param3,
                                               LPVOID param4) {
    // In newer Windows 11 versions, there are only 3 arguments, but that's OK
    // because argument 4 is just ignored (register d9).
    ULONGLONG** elapse =
        (ULONGLONG**)(g_winVersion >= WinVersion::Win11_22H2 ? &param3
                                                             : &param4);

    Wh_Log(L"> %zu", **elapse);

    ULONGLONG elapseNew;

    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString && **elapse == 10000000) {
        // Make the next refresh happen next second. Without this hook, the
        // timer was always set one second forward, and so the clock was
        // accumulating a delay, finally caused one second to be skipped.
        SYSTEMTIME time;
        if (GetLocalTime_Original) {
            GetLocalTime_Original(&time);
        } else {
            GetLocalTime(&time);
        }
        elapseNew = 10000ULL * (1000 - time.wMilliseconds);
        *elapse = &elapseNew;
    }

    return ThreadPoolTimer_CreateTimer_Original(param1, param2, param3, param4);
}

// Similar to ThreadPoolTimer_CreateTimer_Hook. Only one of them is called in
// some Windows versions due to function inlining.
LPVOID WINAPI ThreadPoolTimer_CreateTimer_lambda_Hook(DWORD_PTR** param1,
                                                      LPVOID param2,
                                                      LPVOID param3) {
    DWORD_PTR* elapse = param1[1];

    Wh_Log(L"> %zu", *elapse);

    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString && *elapse == 10000000) {
        // Make the next refresh happen next second. Without this hook, the
        // timer was always set one second forward, and so the clock was
        // accumulating a delay, finally caused one second to be skipped.
        SYSTEMTIME time;
        if (GetLocalTime_Original) {
            GetLocalTime_Original(&time);
        } else {
            GetLocalTime(&time);
        }
        *elapse = 10000ULL * (1000 - time.wMilliseconds);
    }

    return ThreadPoolTimer_CreateTimer_lambda_Original(param1, param2, param3);
}

VOID WINAPI GetLocalTime_Hook_Win11(LPSYSTEMTIME lpSystemTime) {
    Wh_Log(L">");

    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString && g_refreshIconNeedToAdjustTimer) {
        g_refreshIconNeedToAdjustTimer = false;

        // Make the next refresh happen in a second.
        memset(lpSystemTime, 0, sizeof(*lpSystemTime));
        lpSystemTime->wSecond = 59;
        return;
    }

    GetLocalTime_Original(lpSystemTime);
}

int WINAPI GetTimeFormatEx_Hook_Win11(LPCWSTR lpLocaleName,
                                      DWORD dwFlags,
                                      CONST SYSTEMTIME* lpTime,
                                      LPCWSTR lpFormat,
                                      LPWSTR lpTimeStr,
                                      int cchTime) {
    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString) {
        if (wcscmp(g_settings.topLine, L"-") != 0) {
            if (!cchTime) {
                // Hopefully a large enough buffer size.
                return FORMATTED_BUFFER_SIZE;
            }

            return FormatLine(lpTimeStr, cchTime, g_settings.topLine.get()) + 1;
        }
    }

    int ret = GetTimeFormatEx_Original(lpLocaleName, dwFlags, lpTime, lpFormat,
                                       lpTimeStr, cchTime);

    return ret;
}

int WINAPI GetDateFormatEx_Hook_Win11(LPCWSTR lpLocaleName,
                                      DWORD dwFlags,
                                      CONST SYSTEMTIME* lpDate,
                                      LPCWSTR lpFormat,
                                      LPWSTR lpDateStr,
                                      int cchDate,
                                      LPCWSTR lpCalendar) {
    if (g_refreshIconThreadId == GetCurrentThreadId() &&
        !g_inGetTimeToolTipString) {
        // Below is a fix for the following situation. The code inside
        // winrt::SystemTray::implementation::ClockSystemTrayIconDataModel::RefreshIcon
        // looks similar to the following (pseudo code):
        //
        // ----------------------------------------
        // if (someFlag) {
        //     SYSTEMTIME t1;
        //     GetLocalTime(&t1);
        //     setTimeout(nextUpdate, calcNextUpdate(t1.wSecond));
        // }
        // SYSTEMTIME t2;
        // GetLocalTime(&t2);
        // GetDateFormatEx(..., &t2, ...);
        // ----------------------------------------
        //
        // We hook GetLocalTime and change wSecond to update the clock every
        // second, which works if someFlag is set. But in case someFlag isn't
        // set, we end up changing the result of the second GetLocalTime call.
        // To handle this, we check whether the time we get here is the time we
        // set in the hook, and if so, we call GetLocalTime explicitly to pass
        // the correct date to GetDateFormatEx.
        //
        // I'm not sure what someFlag means, but it becomes false when the
        // monitor turns off.

        SYSTEMTIME sentinelSystemTime;
        memset(&sentinelSystemTime, 0, sizeof(sentinelSystemTime));
        sentinelSystemTime.wSecond = 59;
        if (memcmp(lpDate, &sentinelSystemTime, sizeof(sentinelSystemTime)) ==
            0) {
            if (GetLocalTime_Original) {
                GetLocalTime_Original(const_cast<SYSTEMTIME*>(lpDate));
            } else {
                GetLocalTime(const_cast<SYSTEMTIME*>(lpDate));
            }
        }

        if (!(dwFlags & DATE_LONGDATE)) {
            if (!cchDate || g_winVersion >= WinVersion::Win11_22H2) {
                // First call, initialize strings.
                InitializeFormattedStrings(lpDate);
            }

            if (wcscmp(g_settings.bottomLine, L"-") != 0) {
                if (!cchDate) {
                    // Hopefully a large enough buffer size.
                    return FORMATTED_BUFFER_SIZE;
                }

                return FormatLine(lpDateStr, cchDate,
                                  g_settings.bottomLine.get()) +
                       1;
            }
        }
    }

    int ret = GetDateFormatEx_Original(lpLocaleName, dwFlags, lpDate, lpFormat,
                                       lpDateStr, cchDate, lpCalendar);

    return ret;
}

LRESULT WINAPI SendMessageW_Hook(HWND hWnd,
                                 UINT Msg,
                                 WPARAM wParam,
                                 LPARAM lParam) {
    LRESULT ret = SendMessageW_Original(hWnd, Msg, wParam, lParam);

    if (Msg != WM_POWERBROADCAST || wParam != PBT_APMQUERYSUSPEND) {
        return ret;
    }

    switch (lParam) {
        case PBT_APMRESUMECRITICAL:
        case PBT_APMRESUMESUSPEND:
        case PBT_APMRESUMEAUTOMATIC:
            break;

        default:
            return ret;
    }

    WCHAR szClassName[64];
    if (GetClassName(hWnd, szClassName, ARRAYSIZE(szClassName)) == 0) {
        return ret;
    }

    if (_wcsicmp(szClassName, L"MSTaskSwWClass") != 0) {
        return ret;
    }

    Wh_Log(L"Resumed, refreshing web contents");

    std::lock_guard<std::mutex> guard(g_webContentMutex);

    HANDLE event = g_webContentUpdateRefreshEvent;
    if (event) {
        g_webContentLoaded = false;
        SetEvent(event);
    }

    return ret;
}

#pragma endregion Win11Hooks

#pragma region Win10Hooks

DWORD g_updateTextStringThreadId;
int g_getDateFormatExCounter;
DWORD g_getTooltipTextThreadId;
WCHAR* g_getTooltipTextBuffer;
int g_getTooltipTextBufferSize;

using ClockButton_UpdateTextStringsIfNecessary_t =
    unsigned int(WINAPI*)(LPVOID pThis, bool*);
ClockButton_UpdateTextStringsIfNecessary_t
    ClockButton_UpdateTextStringsIfNecessary_Original;

using ClockButton_CalculateMinimumSize_t = LPSIZE(WINAPI*)(LPVOID pThis,
                                                           LPSIZE,
                                                           SIZE);
ClockButton_CalculateMinimumSize_t ClockButton_CalculateMinimumSize_Original;

using ClockButton_GetTextSpacingForOrientation_t =
    int(WINAPI*)(LPVOID pThis, bool, DWORD, DWORD, DWORD, DWORD);
ClockButton_GetTextSpacingForOrientation_t
    ClockButton_GetTextSpacingForOrientation_Original;

using ClockButton_v_GetTooltipText_t =
    HRESULT(WINAPI*)(LPVOID pThis, LPVOID, LPVOID, LPVOID, LPVOID);
ClockButton_v_GetTooltipText_t ClockButton_v_GetTooltipText_Original;

using ClockButton_v_OnDisplayStateChange_t = void(WINAPI*)(LPVOID pThis, bool);
ClockButton_v_OnDisplayStateChange_t
    ClockButton_v_OnDisplayStateChange_Original;

HRESULT WINAPI ClockButton_v_GetTooltipText_Hook(LPVOID pThis,
                                                 LPVOID param1,
                                                 LPVOID param2,
                                                 LPVOID param3,
                                                 LPVOID param4) {
    Wh_Log(L">");

    g_getTooltipTextThreadId = GetCurrentThreadId();

    HRESULT ret = ClockButton_v_GetTooltipText_Original(pThis, param1, param2,
                                                        param3, param4);

    if (g_getTooltipTextBuffer) {
        size_t stringLen = wcslen(g_getTooltipTextBuffer);
        WCHAR* p = g_getTooltipTextBuffer + stringLen;
        size_t size = g_getTooltipTextBufferSize + stringLen;
        if (size > 4) {
            wcscpy(p, L"\r\n\r\n");
            FormatLine(p + 4, size - 4, g_settings.tooltipLine.get());
        }
    }

    g_getTooltipTextBuffer = nullptr;
    g_getTooltipTextBufferSize = 0;

    g_getTooltipTextThreadId = 0;

    return ret;
}

unsigned int WINAPI
ClockButton_UpdateTextStringsIfNecessary_Hook(LPVOID pThis, bool* param1) {
    Wh_Log(L">");

    g_updateTextStringThreadId = GetCurrentThreadId();
    g_getDateFormatExCounter = 0;

    unsigned int ret =
        ClockButton_UpdateTextStringsIfNecessary_Original(pThis, param1);

    g_updateTextStringThreadId = 0;

    if (g_settings.showSeconds || !g_webContentLoaded) {
        // Return the time-out value for the time of the next update.
        SYSTEMTIME time;
        GetLocalTime(&time);
        return 1000 - time.wMilliseconds;
    }

    return ret;
}

LPSIZE WINAPI ClockButton_CalculateMinimumSize_Hook(LPVOID pThis,
                                                    LPSIZE param1,
                                                    SIZE param2) {
    Wh_Log(L">");

    LPSIZE ret =
        ClockButton_CalculateMinimumSize_Original(pThis, param1, param2);

    HWND hWnd = *((HWND*)pThis + 1);
    UINT windowDpi = pGetDpiForWindow ? pGetDpiForWindow(hWnd) : 0;

    if (g_settings.width > 0) {
        ret->cx = g_settings.width;
        if (windowDpi) {
            ret->cx = MulDiv(ret->cx, windowDpi, 96);
        }
    }

    if (g_settings.height > 0) {
        ret->cy = g_settings.height;
        if (windowDpi) {
            ret->cy = MulDiv(ret->cy, windowDpi, 96);
        }
    }

    return ret;
}

int WINAPI ClockButton_GetTextSpacingForOrientation_Hook(LPVOID pThis,
                                                         bool horizontal,
                                                         DWORD dwSiteHeight,
                                                         DWORD dwLine1Height,
                                                         DWORD dwLine2Height,
                                                         DWORD dwLine3Height) {
    Wh_Log(L">");

    int textSpacing = g_settings.textSpacing;
    if (textSpacing == 0) {
        return ClockButton_GetTextSpacingForOrientation_Original(
            pThis, horizontal, dwSiteHeight, dwLine1Height, dwLine2Height,
            dwLine3Height);
    }

    // 1 line.
    if (dwLine3Height == 0 && dwLine2Height == 0) {
        return 0;
    }

    // Since 0 is reserved, shift negative values so that any spacing value can
    // be used.
    if (textSpacing < 0) {
        textSpacing++;
    }

    HWND hWnd = *((HWND*)pThis + 1);
    UINT windowDpi = pGetDpiForWindow ? pGetDpiForWindow(hWnd) : 0;

    if (windowDpi) {
        return MulDiv(textSpacing, windowDpi, 96);
    }

    return textSpacing;
}

int WINAPI GetTimeFormatEx_Hook_Win10(LPCWSTR lpLocaleName,
                                      DWORD dwFlags,
                                      CONST SYSTEMTIME* lpTime,
                                      LPCWSTR lpFormat,
                                      LPWSTR lpTimeStr,
                                      int cchTime) {
    if (g_updateTextStringThreadId == GetCurrentThreadId()) {
        InitializeFormattedStrings(lpTime);

        if (wcscmp(g_settings.topLine, L"-") != 0) {
            return FormatLine(lpTimeStr, cchTime, g_settings.topLine.get()) + 1;
        }
    }

    int ret = GetTimeFormatEx_Original(lpLocaleName, dwFlags, lpTime, lpFormat,
                                       lpTimeStr, cchTime);

    return ret;
}

int WINAPI GetDateFormatEx_Hook_Win10(LPCWSTR lpLocaleName,
                                      DWORD dwFlags,
                                      CONST SYSTEMTIME* lpDate,
                                      LPCWSTR lpFormat,
                                      LPWSTR lpDateStr,
                                      int cchDate,
                                      LPCWSTR lpCalendar) {
    if (g_updateTextStringThreadId == GetCurrentThreadId()) {
        g_getDateFormatExCounter++;
        PCWSTR format = g_getDateFormatExCounter > 1 ? g_settings.middleLine
                                                     : g_settings.bottomLine;
        if (wcscmp(format, L"-") != 0) {
            return FormatLine(lpDateStr, cchDate, format) + 1;
        }
    }

    int ret = GetDateFormatEx_Original(lpLocaleName, dwFlags, lpDate, lpFormat,
                                       lpDateStr, cchDate, lpCalendar);

    return ret;
}

int WINAPI GetDateFormatW_Hook_Win10(LCID Locale,
                                     DWORD dwFlags,
                                     CONST SYSTEMTIME* lpDate,
                                     LPCWSTR lpFormat,
                                     LPWSTR lpDateStr,
                                     int cchDate) {
    if (g_getTooltipTextThreadId == GetCurrentThreadId() &&
        !g_getTooltipTextBuffer) {
        g_getTooltipTextBuffer = lpDateStr;
        g_getTooltipTextBufferSize = cchDate;
    }

    int ret = GetDateFormatW_Original(Locale, dwFlags, lpDate, lpFormat,
                                      lpDateStr, cchDate);

    return ret;
}

#pragma endregion Win10Hooks

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
            } else if (build <= 22000) {
                return WinVersion::Win11;
            } else if (build < 26100) {
                return WinVersion::Win11_22H2;
            } else {
                return WinVersion::Win11_24H2;
            }
            break;
    }

    return WinVersion::Unsupported;
}

bool HookExplorerPatcherSymbols(HMODULE explorerPatcherModule) {
    if (g_explorerPatcherInitialized.exchange(true)) {
        return true;
    }

    if (g_winVersion >= WinVersion::Win11) {
        g_winVersion = WinVersion::Win10;
    }

    struct EXPLORER_PATCHER_HOOK {
        PCSTR symbol;
        void** pOriginalFunction;
        void* hookFunction = nullptr;
        bool optional = false;
    };

    EXPLORER_PATCHER_HOOK hooks[] = {
        {R"(?UpdateTextStringsIfNecessary@ClockButton@@AEAAIPEA_N@Z)",
         (void**)&ClockButton_UpdateTextStringsIfNecessary_Original,
         (void*)ClockButton_UpdateTextStringsIfNecessary_Hook},
        {R"(?CalculateMinimumSize@ClockButton@@QEAA?AUtagSIZE@@U2@@Z)",
         (void**)&ClockButton_CalculateMinimumSize_Original,
         (void*)ClockButton_CalculateMinimumSize_Hook},
        {R"(?GetTextSpacingForOrientation@ClockButton@@AEAAH_NHHHH@Z)",
         (void**)&ClockButton_GetTextSpacingForOrientation_Original,
         (void*)ClockButton_GetTextSpacingForOrientation_Hook},
        {R"(?v_GetTooltipText@ClockButton@@MEAAJPEAPEAUHINSTANCE__@@PEAPEAGPEAG_K@Z)",
         (void**)&ClockButton_v_GetTooltipText_Original,
         (void*)ClockButton_v_GetTooltipText_Hook, true},
        {R"(?v_OnDisplayStateChange@ClockButton@@MEAAX_N@Z)",
         (void**)&ClockButton_v_OnDisplayStateChange_Original},
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

    if (g_initialized) {
        Wh_ApplyHookOperations();
    }

    return succeeded;
}

bool HandleModuleIfExplorerPatcher(HMODULE module) {
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

    if (_wcsnicmp(L"ep_taskbar.", moduleFileName, sizeof("ep_taskbar.") - 1) !=
        0) {
        return true;
    }

    Wh_Log(L"ExplorerPatcher taskbar loaded: %s", moduleFileName);
    return HookExplorerPatcherSymbols(module);
}

void HandleLoadedExplorerPatcher() {
    HMODULE hMods[1024];
    DWORD cbNeeded;
    if (EnumProcessModules(GetCurrentProcess(), hMods, sizeof(hMods),
                           &cbNeeded)) {
        for (size_t i = 0; i < cbNeeded / sizeof(HMODULE); i++) {
            HandleModuleIfExplorerPatcher(hMods[i]);
        }
    }
}

using LoadLibraryExW_t = decltype(&LoadLibraryExW);
LoadLibraryExW_t LoadLibraryExW_Original;
HMODULE WINAPI LoadLibraryExW_Hook(LPCWSTR lpLibFileName,
                                   HANDLE hFile,
                                   DWORD dwFlags) {
    HMODULE module = LoadLibraryExW_Original(lpLibFileName, hFile, dwFlags);
    if (module && !((ULONG_PTR)module & 3) && !g_explorerPatcherInitialized) {
        HandleModuleIfExplorerPatcher(module);
    }

    return module;
}

bool HookWin10TaskbarSymbols() {
    WindhawkUtils::SYMBOL_HOOK explorerExeHooks[] = {
        {
            {LR"(private: unsigned int __cdecl ClockButton::UpdateTextStringsIfNecessary(bool *))"},
            (void**)&ClockButton_UpdateTextStringsIfNecessary_Original,
            (void*)ClockButton_UpdateTextStringsIfNecessary_Hook,
        },
        {
            {LR"(public: struct tagSIZE __cdecl ClockButton::CalculateMinimumSize(struct tagSIZE))"},
            (void**)&ClockButton_CalculateMinimumSize_Original,
            (void*)ClockButton_CalculateMinimumSize_Hook,
        },
        {
            {LR"(private: int __cdecl ClockButton::GetTextSpacingForOrientation(bool,int,int,int,int))"},
            (void**)&ClockButton_GetTextSpacingForOrientation_Original,
            (void*)ClockButton_GetTextSpacingForOrientation_Hook,
        },
        {
            {LR"(protected: virtual long __cdecl ClockButton::v_GetTooltipText(struct HINSTANCE__ * *,unsigned short * *,unsigned short *,unsigned __int64))"},
            (void**)&ClockButton_v_GetTooltipText_Original,
            (void*)ClockButton_v_GetTooltipText_Hook,
            true,
        },
        {
            {LR"(protected: virtual void __cdecl ClockButton::v_OnDisplayStateChange(bool))"},
            (void**)&ClockButton_v_OnDisplayStateChange_Original,
        },
    };

    if (!HookSymbols(GetModuleHandle(nullptr), explorerExeHooks,
                     ARRAYSIZE(explorerExeHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
}

bool GetTaskbarViewDllPath(WCHAR path[MAX_PATH]) {
    WCHAR szWindowsDirectory[MAX_PATH];
    if (!GetWindowsDirectory(szWindowsDirectory,
                             ARRAYSIZE(szWindowsDirectory))) {
        Wh_Log(L"GetWindowsDirectory failed");
        return false;
    }

    // Windows 11 version 22H2.
    wcscpy_s(path, MAX_PATH, szWindowsDirectory);
    wcscat_s(
        path, MAX_PATH,
        LR"(\SystemApps\MicrosoftWindows.Client.Core_cw5n1h2txyewy\Taskbar.View.dll)");
    if (GetFileAttributes(path) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    // Windows 11 version 21H2.
    wcscpy_s(path, MAX_PATH, szWindowsDirectory);
    wcscat_s(
        path, MAX_PATH,
        LR"(\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\ExplorerExtensions.dll)");
    if (GetFileAttributes(path) != INVALID_FILE_ATTRIBUTES) {
        return true;
    }

    return false;
}

bool HookTaskbarViewDllSymbols() {
    WCHAR dllPath[MAX_PATH];
    if (!GetTaskbarViewDllPath(dllPath)) {
        Wh_Log(L"Taskbar view module not found");
        return false;
    }

    HMODULE module =
        LoadLibraryEx(dllPath, nullptr, LOAD_WITH_ALTERED_SEARCH_PATH);
    if (!module) {
        Wh_Log(L"Taskbar view module couldn't be loaded");
        return false;
    }

    // Taskbar.View.dll, ExplorerExtensions.dll
    WindhawkUtils::SYMBOL_HOOK symbolHooks[] = {
        {
            {LR"(private: void __cdecl winrt::SystemTray::implementation::ClockSystemTrayIconDataModel::RefreshIcon(class SystemTrayTelemetry::ClockUpdate &))"},
            (void**)&ClockSystemTrayIconDataModel_RefreshIcon_Original,
            (void*)ClockSystemTrayIconDataModel_RefreshIcon_Hook,
        },
        {
            {LR"(private: struct winrt::hstring __cdecl winrt::SystemTray::implementation::ClockSystemTrayIconDataModel::GetTimeToolTipString(struct _SYSTEMTIME const &,struct _SYSTEMTIME const &,class SystemTrayTelemetry::ClockUpdate &))"},
            (void**)&ClockSystemTrayIconDataModel_GetTimeToolTipString_Original,
            (void*)ClockSystemTrayIconDataModel_GetTimeToolTipString_Hook,
            true,
        },
        {
            {LR"(private: struct winrt::hstring __cdecl winrt::SystemTray::implementation::ClockSystemTrayIconDataModel::GetTimeToolTipString2(struct _SYSTEMTIME const &,struct _SYSTEMTIME const &,class SystemTrayTelemetry::ClockUpdate &))"},
            (void**)&ClockSystemTrayIconDataModel_GetTimeToolTipString2_Original,
            (void*)ClockSystemTrayIconDataModel_GetTimeToolTipString2_Hook,
            true,
        },
        {
            {LR"(public: void __cdecl winrt::SystemTray::implementation::DateTimeIconContent::OnApplyTemplate(void))"},
            (void**)&DateTimeIconContent_OnApplyTemplate_Original,
            (void*)DateTimeIconContent_OnApplyTemplate_Hook,
            true,
        },
        {
            {LR"(public: virtual int __cdecl winrt::impl::produce<struct winrt::SystemTray::implementation::BadgeIconContent,struct winrt::SystemTray::IBadgeIconContent>::get_ViewModel(void * *))"},
            (void**)&BadgeIconContent_get_ViewModel_Original,
            (void*)BadgeIconContent_get_ViewModel_Hook,
            true,
        },
        {
            {LR"(private: struct winrt::hstring __cdecl winrt::SystemTray::implementation::ClockSystemTrayIconDataModel::GetTimeToolTipString(struct _SYSTEMTIME *,struct _TIME_DYNAMIC_ZONE_INFORMATION *,class SystemTrayTelemetry::ClockUpdate &))"},
            (void**)&ClockSystemTrayIconDataModel_GetTimeToolTipString_2_Original,
            (void*)ClockSystemTrayIconDataModel_GetTimeToolTipString_2_Hook,
            true,  // Until Windows 11 version 21H2.
        },
        {
            {LR"(public: int __cdecl winrt::impl::consume_Windows_Globalization_ICalendar<struct winrt::Windows::Globalization::ICalendar>::Second(void)const )"},
            (void**)&ICalendar_Second_Original,
            (void*)ICalendar_Second_Hook,
            true,  // Until Windows 11 version 21H2.
        },
        {
            {
                LR"(public: static __cdecl winrt::Windows::System::Threading::ThreadPoolTimer::CreateTimer(struct winrt::Windows::System::Threading::TimerElapsedHandler const &,class std::chrono::duration<__int64,struct std::ratio<1,10000000> > const &))",
                // Windows 11 21H2:
                LR"(public: struct winrt::Windows::System::Threading::ThreadPoolTimer __cdecl winrt::impl::consume_Windows_System_Threading_IThreadPoolTimerStatics<struct winrt::Windows::System::Threading::IThreadPoolTimerStatics>::CreateTimer(struct winrt::Windows::System::Threading::TimerElapsedHandler const &,class std::chrono::duration<__int64,struct std::ratio<1,10000000> > const &)const )",
            },
            (void**)&ThreadPoolTimer_CreateTimer_Original,
            (void*)ThreadPoolTimer_CreateTimer_Hook,
            true,  // Only for more precise clock, see comment in the hook.
        },
        {
            {
                LR"(public: __cdecl <lambda_b19cf72fe9674443383aa89d5c22450b>::operator()(struct winrt::Windows::System::Threading::IThreadPoolTimerStatics const &)const )",
                // Windows 11 21H2:
                LR"(public: struct winrt::Windows::System::Threading::ThreadPoolTimer __cdecl <lambda_b19cf72fe9674443383aa89d5c22450b>::operator()(struct winrt::Windows::System::Threading::IThreadPoolTimerStatics const &)const )",
            },
            (void**)&ThreadPoolTimer_CreateTimer_lambda_Original,
            (void*)ThreadPoolTimer_CreateTimer_lambda_Hook,
            true,  // Only for more precise clock, see comment in the hook.
        },
    };

    if (!HookSymbols(module, symbolHooks, ARRAYSIZE(symbolHooks))) {
        Wh_Log(L"HookSymbols failed");
        return false;
    }

    return true;
}

void LoadSettings() {
    g_settings.showSeconds = Wh_GetIntSetting(L"ShowSeconds");
    g_settings.timeFormat = Wh_GetStringSetting(L"TimeFormat");
    g_settings.dateFormat = Wh_GetStringSetting(L"DateFormat");
    g_settings.weekdayFormat = Wh_GetStringSetting(L"WeekdayFormat");
    g_settings.topLine = Wh_GetStringSetting(L"TopLine");
    g_settings.bottomLine = Wh_GetStringSetting(L"BottomLine");
    g_settings.middleLine = Wh_GetStringSetting(L"MiddleLine");
    g_settings.tooltipLine = Wh_GetStringSetting(L"TooltipLine");
    g_settings.width = Wh_GetIntSetting(L"Width");
    g_settings.height = Wh_GetIntSetting(L"Height");
    g_settings.maxWidth = Wh_GetIntSetting(L"MaxWidth");
    g_settings.textSpacing = Wh_GetIntSetting(L"TextSpacing");

    g_settings.webContentsItems.clear();
    for (int i = 0;; i++) {
        WebContentsSettings item;
        item.url = Wh_GetStringSetting(L"WebContentsItems[%d].Url", i);
        if (*item.url == '\0') {
            break;
        }

        item.blockStart =
            Wh_GetStringSetting(L"WebContentsItems[%d].BlockStart", i);
        item.start = Wh_GetStringSetting(L"WebContentsItems[%d].Start", i);
        item.end = Wh_GetStringSetting(L"WebContentsItems[%d].End", i);
        item.maxLength = Wh_GetIntSetting(L"WebContentsItems[%d].MaxLength", i);

        g_settings.webContentsItems.push_back(std::move(item));
    }

    g_settings.webContentsUpdateInterval =
        Wh_GetIntSetting(L"WebContentsUpdateInterval");

    g_settings.timeStyle.visible = Wh_GetIntSetting(L"TimeStyle.Visible");
    g_settings.timeStyle.textColor =
        Wh_GetStringSetting(L"TimeStyle.TextColor");
    g_settings.timeStyle.textAlignment =
        Wh_GetStringSetting(L"TimeStyle.TextAlignment");
    g_settings.timeStyle.fontSize = Wh_GetIntSetting(L"TimeStyle.FontSize");
    g_settings.timeStyle.fontFamily =
        Wh_GetStringSetting(L"TimeStyle.FontFamily");
    g_settings.timeStyle.fontWeight =
        Wh_GetStringSetting(L"TimeStyle.FontWeight");
    g_settings.timeStyle.fontStyle =
        Wh_GetStringSetting(L"TimeStyle.FontStyle");
    g_settings.timeStyle.fontStretch =
        Wh_GetStringSetting(L"TimeStyle.FontStretch");
    g_settings.timeStyle.characterSpacing =
        Wh_GetIntSetting(L"TimeStyle.CharacterSpacing");

    g_settings.dateStyle.visible = Wh_GetIntSetting(L"DateStyle.Visible");
    g_settings.dateStyle.textColor =
        Wh_GetStringSetting(L"DateStyle.TextColor");
    g_settings.dateStyle.textAlignment =
        Wh_GetStringSetting(L"DateStyle.TextAlignment");
    g_settings.dateStyle.fontSize = Wh_GetIntSetting(L"DateStyle.FontSize");
    g_settings.dateStyle.fontFamily =
        Wh_GetStringSetting(L"DateStyle.FontFamily");
    g_settings.dateStyle.fontWeight =
        Wh_GetStringSetting(L"DateStyle.FontWeight");
    g_settings.dateStyle.fontStyle =
        Wh_GetStringSetting(L"DateStyle.FontStyle");
    g_settings.dateStyle.fontStretch =
        Wh_GetStringSetting(L"DateStyle.FontStretch");
    g_settings.dateStyle.characterSpacing =
        Wh_GetIntSetting(L"DateStyle.CharacterSpacing");

    g_clockElementStyleEnabled =
        (g_settings.maxWidth || g_settings.textSpacing ||
         !g_settings.timeStyle.visible || *g_settings.timeStyle.textColor ||
         *g_settings.timeStyle.textAlignment || g_settings.timeStyle.fontSize ||
         *g_settings.timeStyle.fontFamily || *g_settings.timeStyle.fontWeight ||
         *g_settings.timeStyle.fontStyle || *g_settings.timeStyle.fontStretch ||
         g_settings.timeStyle.characterSpacing ||
         !g_settings.dateStyle.visible || *g_settings.dateStyle.textColor ||
         *g_settings.dateStyle.textAlignment || g_settings.dateStyle.fontSize ||
         *g_settings.dateStyle.fontFamily || *g_settings.dateStyle.fontWeight ||
         *g_settings.dateStyle.fontStyle || *g_settings.dateStyle.fontStretch ||
         g_settings.dateStyle.characterSpacing);
    g_clockElementStyleIndex++;

    g_settings.oldTaskbarOnWin11 = Wh_GetIntSetting(L"oldTaskbarOnWin11");

    // Kept for compatibility with old settings:
    if (wcsstr(g_settings.topLine, L"%web%") ||
        wcsstr(g_settings.bottomLine, L"%web%") ||
        wcsstr(g_settings.middleLine, L"%web%") ||
        wcsstr(g_settings.tooltipLine, L"%web%") ||
        wcsstr(g_settings.topLine, L"%web_full%") ||
        wcsstr(g_settings.bottomLine, L"%web_full%") ||
        wcsstr(g_settings.middleLine, L"%web_full%") ||
        wcsstr(g_settings.tooltipLine, L"%web_full%")) {
        g_settings.webContentsUrl = Wh_GetStringSetting(L"WebContentsUrl");
        g_settings.webContentsBlockStart =
            Wh_GetStringSetting(L"WebContentsBlockStart");
        g_settings.webContentsStart = Wh_GetStringSetting(L"WebContentsStart");
        g_settings.webContentsEnd = Wh_GetStringSetting(L"WebContentsEnd");
        g_settings.webContentsMaxLength =
            Wh_GetIntSetting(L"WebContentsMaxLength");
    }
}

void ApplySettingsWin11() {
    DWORD dwProcessId;
    DWORD dwCurrentProcessId = GetCurrentProcessId();

    HWND hTaskbarWnd = FindWindow(L"Shell_TrayWnd", nullptr);
    if (hTaskbarWnd && GetWindowThreadProcessId(hTaskbarWnd, &dwProcessId) &&
        dwProcessId == dwCurrentProcessId) {
        // Touch a registry value to trigger a watcher for a clock update. Do so
        // only if the current explorer.exe instance owns the taskbar.
        constexpr WCHAR kTempValueName[] =
            L"_temp_windhawk_taskbar-taskbar-clock-customization";
        HKEY hSubKey;
        LONG result = RegOpenKeyEx(HKEY_CURRENT_USER,
                                   L"Control Panel\\TimeDate\\AdditionalClocks",
                                   0, KEY_WRITE, &hSubKey);
        if (result == ERROR_SUCCESS) {
            if (RegSetValueEx(hSubKey, kTempValueName, 0, REG_SZ,
                              (const BYTE*)L"",
                              sizeof(WCHAR)) != ERROR_SUCCESS) {
                Wh_Log(L"Failed to create temp value");
            } else if (RegDeleteValue(hSubKey, kTempValueName) !=
                       ERROR_SUCCESS) {
                Wh_Log(L"Failed to remove temp value");
            }

            RegCloseKey(hSubKey);
        } else {
            Wh_Log(L"Failed to open subkey: %d", result);
        }
    }
}

void ApplySettingsWin10() {
    DWORD dwProcessId;
    DWORD dwCurrentProcessId = GetCurrentProcessId();

    HWND hTaskbarWnd = FindWindow(L"Shell_TrayWnd", nullptr);
    if (hTaskbarWnd && GetWindowThreadProcessId(hTaskbarWnd, &dwProcessId) &&
        dwProcessId == dwCurrentProcessId) {
        // Apply size.
        RECT rc;
        if (GetClientRect(hTaskbarWnd, &rc)) {
            SendMessage(hTaskbarWnd, WM_SIZE, SIZE_RESTORED,
                        MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
        }

        // Apply text.
        HWND hTrayNotifyWnd =
            FindWindowEx(hTaskbarWnd, nullptr, L"TrayNotifyWnd", nullptr);
        if (hTrayNotifyWnd) {
            HWND hTrayClockWWnd = FindWindowEx(hTrayNotifyWnd, nullptr,
                                               L"TrayClockWClass", nullptr);
            if (hTrayClockWWnd) {
                LONG_PTR lpTrayClockWClassLongPtr =
                    GetWindowLongPtr(hTrayClockWWnd, 0);
                if (lpTrayClockWClassLongPtr) {
                    ClockButton_v_OnDisplayStateChange_Original(
                        (LPVOID)lpTrayClockWClassLongPtr, true);
                }
            }
        }
    }

    HWND hSecondaryTaskbarWnd = FindWindow(L"Shell_SecondaryTrayWnd", nullptr);
    while (hSecondaryTaskbarWnd &&
           GetWindowThreadProcessId(hSecondaryTaskbarWnd, &dwProcessId) &&
           dwProcessId == dwCurrentProcessId) {
        // Apply size.
        RECT rc;
        if (GetClientRect(hSecondaryTaskbarWnd, &rc)) {
            WINDOWPOS windowpos;
            windowpos.hwnd = hSecondaryTaskbarWnd;
            windowpos.hwndInsertAfter = nullptr;
            windowpos.x = 0;
            windowpos.y = 0;
            windowpos.cx = rc.right - rc.left;
            windowpos.cy = rc.bottom - rc.top;
            windowpos.flags = SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE;

            SendMessage(hSecondaryTaskbarWnd, WM_WINDOWPOSCHANGED, 0,
                        (LPARAM)&windowpos);
        }

        // Apply text.
        HWND hClockButtonWnd = FindWindowEx(hSecondaryTaskbarWnd, nullptr,
                                            L"ClockButton", nullptr);
        if (hClockButtonWnd) {
            LONG_PTR lpClockButtonLongPtr =
                GetWindowLongPtr(hClockButtonWnd, 0);
            if (lpClockButtonLongPtr) {
                ClockButton_v_OnDisplayStateChange_Original(
                    (LPVOID)lpClockButtonLongPtr, true);
            }
        }

        hSecondaryTaskbarWnd = FindWindowEx(nullptr, hSecondaryTaskbarWnd,
                                            L"Shell_SecondaryTrayWnd", nullptr);
    }
}

void ApplySettings() {
    if (g_winVersion >= WinVersion::Win11) {
        ApplySettingsWin11();
    } else {
        ApplySettingsWin10();
    }
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

        if (hasWin10Taskbar && !HookWin10TaskbarSymbols()) {
            return FALSE;
        }
    } else if (g_winVersion >= WinVersion::Win11) {
        if (!HookTaskbarViewDllSymbols()) {
            return FALSE;
        }
    } else {
        if (!HookWin10TaskbarSymbols()) {
            return FALSE;
        }
    }

    HandleLoadedExplorerPatcher();

    HMODULE kernelBaseModule = GetModuleHandle(L"kernelbase.dll");
    FARPROC pKernelBaseLoadLibraryExW =
        GetProcAddress(kernelBaseModule, "LoadLibraryExW");
    Wh_SetFunctionHook((void*)pKernelBaseLoadLibraryExW,
                       (void*)LoadLibraryExW_Hook,
                       (void**)&LoadLibraryExW_Original);

    HMODULE hUser32 = LoadLibrary(L"user32.dll");
    if (hUser32) {
        pGetDpiForWindow =
            (GetDpiForWindow_t)GetProcAddress(hUser32, "GetDpiForWindow");
    }

    // Must use GetProcAddress for the functions below, otherwise the stubs in
    // kernel32.dll are being hooked.
    FARPROC pGetTimeFormatEx =
        GetProcAddress(kernelBaseModule, "GetTimeFormatEx");
    if (!pGetTimeFormatEx) {
        return FALSE;
    }

    FARPROC pGetDateFormatEx =
        GetProcAddress(kernelBaseModule, "GetDateFormatEx");
    if (!pGetDateFormatEx) {
        return FALSE;
    }

    if (g_winVersion <= WinVersion::Win10) {
        Wh_SetFunctionHook((void*)pGetTimeFormatEx,
                           (void*)GetTimeFormatEx_Hook_Win10,
                           (void**)&GetTimeFormatEx_Original);
        Wh_SetFunctionHook((void*)pGetDateFormatEx,
                           (void*)GetDateFormatEx_Hook_Win10,
                           (void**)&GetDateFormatEx_Original);

        FARPROC pGetDateFormatW =
            GetProcAddress(kernelBaseModule, "GetDateFormatW");
        if (pGetDateFormatW) {
            Wh_SetFunctionHook((void*)pGetDateFormatW,
                               (void*)GetDateFormatW_Hook_Win10,
                               (void**)&GetDateFormatW_Original);
        }
    } else {
        if (g_winVersion >= WinVersion::Win11_22H2) {
            FARPROC pGetLocalTime =
                GetProcAddress(kernelBaseModule, "GetLocalTime");
            if (!pGetLocalTime) {
                return FALSE;
            }

            Wh_SetFunctionHook((void*)pGetLocalTime,
                               (void*)GetLocalTime_Hook_Win11,
                               (void**)&GetLocalTime_Original);
        }

        Wh_SetFunctionHook((void*)pGetTimeFormatEx,
                           (void*)GetTimeFormatEx_Hook_Win11,
                           (void**)&GetTimeFormatEx_Original);
        Wh_SetFunctionHook((void*)pGetDateFormatEx,
                           (void*)GetDateFormatEx_Hook_Win11,
                           (void**)&GetDateFormatEx_Original);
        Wh_SetFunctionHook((void*)SendMessageW, (void*)SendMessageW_Hook,
                           (void**)&SendMessageW_Original);
    }

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

    WebContentUpdateThreadInit();

    ApplySettings();
}

void Wh_ModBeforeUninit() {
    if (g_winVersion >= WinVersion::Win11 &&
        g_clockElementStyleEnabled.exchange(false)) {
        DWORD styleIndex = ++g_clockElementStyleIndex;

        ApplySettings();

        // Wait for styles to be restored.
        for (int i = 0; i < 20; i++) {
            bool allRestored = true;
            for (const auto& data : g_clockElementStyleData) {
                if (data.styleIndex < styleIndex) {
                    allRestored = false;
                    break;
                }
            }

            if (allRestored) {
                break;
            }

            Sleep(100);
        }
    }
}

void Wh_ModUninit() {
    Wh_Log(L">");

    WebContentUpdateThreadUninit();

    ApplySettings();
}

BOOL Wh_ModSettingsChanged(BOOL* bReload) {
    Wh_Log(L">");

    WebContentUpdateThreadUninit();

    bool prevOldTaskbarOnWin11 = g_settings.oldTaskbarOnWin11;

    LoadSettings();

    *bReload = g_settings.oldTaskbarOnWin11 != prevOldTaskbarOnWin11;
    if (*bReload) {
        return TRUE;
    }

    WebContentUpdateThreadInit();

    ApplySettings();

    return TRUE;
}
