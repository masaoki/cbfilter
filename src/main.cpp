/**
 * @file main.cpp
 * @brief Main entry point and UI implementation for Windows clipboard filter utility
 * 
 * This application provides a system tray utility that transforms clipboard content
 * using AI models (OpenAI API compatible). Users can define filters that process
 * text or images through various AI models and replace clipboard content with results.
 */

#include "clipboard_processor.h"

#include <cwctype>
#include <cstring>
#include <string>
#include <vector>
#include <cstdio>
#include <regex>
#include <format>
#include "resource.h"
#include <windows.h>
#include <objidl.h>
#include <shellapi.h>
#include <commctrl.h>
#include <windowsx.h>
#include <winhttp.h>
#include <gdiplus.h>
#include <wincrypt.h>
#include <shlobj.h>
#include <winrt/base.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <winrt/Windows.Data.Json.h>

using namespace std;

// Window class names for different dialogs
constexpr wchar_t kClassName[] = L"CbFilterHidden";        // Hidden main window
constexpr wchar_t kSettingsClass[] = L"CbFilterSettings";  // Settings dialog window
constexpr wchar_t kEditClass[] = L"CbFilterEdit";          // Filter edit dialog window
constexpr wchar_t kModelClass[] = L"CbFilterModel";        // Model configuration dialog window
constexpr wchar_t kProgressClass[] = L"CbFilterProgress";  // Progress window
constexpr wchar_t kSetupClass[] = L"CbFilterSetup";        // First-run setup dialog
constexpr wchar_t kHotkeyInputClass[] = L"CbFilterHotkeyInput"; // Hotkey input dialog
constexpr wchar_t kFilterMenuClass[] = L"CbFilterMenu";    // Filter menu window

// Hotkey identifier
constexpr int HOTKEY_ID = 1;

// Global hotkey configuration
UINT g_hotkeyModifiers = MOD_WIN | MOD_ALT;  // Default: Win+Alt
UINT g_hotkeyKey = 'V';                       // Default: V key
wstring g_language = L"ja";              // Default language: Japanese

// Custom window messages
constexpr UINT WM_APP_TRAY = WM_APP + 10;  // System tray notification message
constexpr UINT WM_APP_FILTER_COMPLETE = WM_APP + 11;  // Filter execution complete message
constexpr UINT WM_APP_MENU_CLOSE = WM_APP + 12;  // Filter menu close message (sent to parent)
constexpr UINT WM_APP_MENU_SELECTED = WM_APP + 13;  // Filter menu item selected (sent to menu window itself)

// Timer ID for progress window
constexpr UINT_PTR TIMER_ID_PROGRESS = 1;

// Menu item IDs for system tray menu
constexpr UINT MENU_ID_SETTINGS = 4001;
constexpr UINT MENU_ID_EXIT = 4002;

// Control IDs for settings dialog
constexpr int IDC_LIST = 301;
constexpr int IDC_BTN_ADD = 302;
constexpr int IDC_BTN_EDIT = 303;
constexpr int IDC_BTN_DELETE = 304;
constexpr int IDC_BTN_COPY = 305;
constexpr int IDC_BTN_CLOSE = 306;

// Global application state
HINSTANCE g_hInst = nullptr;              // Application instance handle
HWND g_settingsWnd = nullptr;             // Settings window handle
HWND g_editWnd = nullptr;                 // Filter edit dialog handle
HWND g_modelWnd = nullptr;                // Model configuration dialog handle
HWND g_progressWnd = nullptr;             // Progress window handle
HWND g_filterMenuWnd = nullptr;           // Filter menu window handle
WNDPROC g_promptOldProc = nullptr;        // Original window procedure for prompt edit control
WNDPROC g_listOldProc = nullptr;          // Original window procedure for list view control
ULONG_PTR g_gdiplusToken = 0;             // GDI+ initialization token
const wchar_t kCtrlAProp[] = L"cbfilter_oldproc_ctrlA";
HFONT GetUIFont();
void SetUIFont(HWND hwnd);

/**
 * @enum IOType
 * @brief Input/output type for filter operations
 */
enum class IOType { Text, Image };

/**
 * @struct ModelConfig
 * @brief Configuration for an AI model endpoint
 */
struct ModelConfig {
    wstring name;        // Display name for the model
    wstring serverUrl;   // API server URL (e.g., https://api.openai.com/v1)
    wstring modelName;   // Model identifier (e.g., gpt-4o-mini)
    wstring apiKey;      // API authentication key
    wstring providerId;  // API provider id
};

/**
 * @struct FilterDefinition
 * @brief Definition of a clipboard transformation filter
 */
struct FilterDefinition {
    wstring title;      // Display name for the filter
    IOType input;            // Input type (Text or Image)
    IOType output;           // Output type (Text or Image)
    size_t modelIndex;       // Index into g_models vector
    wstring prompt;     // Prompt text to send to the AI model
};

/**
 * @struct TemplateDefinition
 * @brief API request template definition loaded from apidef/<provider>.json
 */
struct TemplateDefinition {
    wstring id;
    wstring providerId;
    IOType input;
    IOType output;
    wstring endpoint;
    wstring resultPath;
    vector<pair<wstring, wstring>> headers; // key/value with placeholders
    wstring payload;                                        // JSON text with placeholders
};

struct ApiProvider {
    wstring id;
    wstring defaultEndpoint;
    vector<TemplateDefinition> templates;
    wstring modelsEndpoint;
    wstring modelsMethod;
    vector<pair<wstring, wstring>> modelsHeaders;
    wstring modelsPayload;
    wstring modelsResultPath;
};

// Global model configurations (loaded from config.ini on startup)
vector<ModelConfig> g_models {
    {L"Translate", L"https://api.openai.com/v1", L"gpt-5.1", L"You are a translator.", L"OpenAI"},
};

// Global filter definitions (loaded from config.ini on startup)
// Note: Default filter titles will be loaded from language resources after LoadConfig()
vector<FilterDefinition> g_filters;
// API providers (loaded from apidef/*.json)
vector<ApiProvider> g_providers;

/**
 * @brief Convert IOType enum to display string
 * @param t The IOType to convert
 * @return Localized string representation
 */
/**
 * @brief Get the path to the language resource file
 * @return Full path to lang.ini in the application directory
 */
wstring GetLangPath() {
    wchar_t buf[MAX_PATH]; GetModuleFileNameW(nullptr, buf, MAX_PATH);
    wstring path(buf); size_t pos = path.find_last_of(L"\\/");
    if (pos != wstring::npos) path = path.substr(0, pos + 1);
    path += L"lang.ini"; return path;
}

/**
 * @brief Get localized string from language resource file (UTF-8 encoded)
 * @param key String key to look up
 * @return Localized string, or key itself if not found
 */
wstring GetString(const wstring& key) {
    wstring langPath = GetLangPath();
    // Read file as UTF-8
    FILE* fp = nullptr;
    if (_wfopen_s(&fp, langPath.c_str(), L"r, ccs=UTF-8") != 0 || !fp) {
        return key; // Return key if file cannot be opened
    }
    
    wstring section = g_language;
    wstring result = key; // Default to key if not found
    bool inTargetSection = false;
    wchar_t line[2048];
    
    while (fgetws(line, 2048, fp)) {
        // Remove trailing newline
        size_t len = wcslen(line);
        if (len > 0 && line[len - 1] == L'\n') {
            line[len - 1] = L'\0';
            if (len > 1 && line[len - 2] == L'\r') line[len - 2] = L'\0';
        }
        
        // Skip empty lines and comments
        if (line[0] == L'\0' || line[0] == L';') continue;
        
        // Check for section header [section]
        if (line[0] == L'[') {
            size_t end = wcscspn(line + 1, L"]");
            if (end > 0 && line[end + 1] == L']') {
                wchar_t sec[256];
                wcsncpy_s(sec, line + 1, end);
                sec[end] = L'\0';
                inTargetSection = (wcscmp(sec, section.c_str()) == 0);
            }
            continue;
        }
        
        // Parse key=value if in target section
        if (inTargetSection) {
            wchar_t* eq = wcschr(line, L'=');
            if (eq) {
                *eq = L'\0';
                // Trim whitespace from key
                wchar_t* keyStart = line;
                while (*keyStart == L' ' || *keyStart == L'\t') keyStart++;
                wchar_t* keyEnd = keyStart + wcslen(keyStart) - 1;
                while (keyEnd > keyStart && (*keyEnd == L' ' || *keyEnd == L'\t')) keyEnd--;
                *(keyEnd + 1) = L'\0';
                
                if (wcscmp(keyStart, key.c_str()) == 0) {
                    // Found the key, get the value
                    wchar_t* valueStart = eq + 1;
                    while (*valueStart == L' ' || *valueStart == L'\t') valueStart++;
                    result = valueStart;
                    break;
                }
            }
        }
    }
    
    fclose(fp);
    return result;
}

wstring GetDefaultLanguageCode() {
    // default language can be extended later; falls back to current g_language
    return g_language;
}

const vector<pair<wstring, wstring>>& SupportedLanguages() {
    static vector<pair<wstring, wstring>> langs{
        {L"ja", L"日本語"},
        {L"en", L"English"},
        {L"zh", L"中文"},
        {L"ko", L"한국어"},
        {L"vi", L"Tiếng Việt"},
        {L"th", L"ไทย"},
        {L"es", L"Español"},
        {L"de", L"Deutsch"},
        {L"fr", L"Français"},
        {L"it", L"Italiano"},
        {L"nl", L"Nederlands"},
        {L"pt", L"Português"},
        {L"ru", L"Русский"},
    };
    return langs;
}

HFONT GetUIFont() {
    static HFONT font = []() -> HFONT {
        HFONT f = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        if (f) return f;
        HDC hdc = GetDC(nullptr);
        int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
        if (hdc) ReleaseDC(nullptr, hdc);
        int height = -MulDiv(12, dpi, 72);
        return CreateFontW(height, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    }();
    return font;
}

void SetUIFont(HWND hwnd) {
    if (!hwnd) return;
    SendMessageW(hwnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetUIFont()), TRUE);
}

/**
 * @brief Convert IOType enum to display string
 * @param t The IOType to convert
 * @return Localized string representation
 */
wstring IOTypeToString(IOType t) {
    return t == IOType::Text ? GetString(L"text_type") : GetString(L"image_type");
}

LRESULT CALLBACK CtrlAEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_KEYDOWN && wParam == 'A' && (GetKeyState(VK_CONTROL) & 0x8000)) {
        SendMessageW(hwnd, EM_SETSEL, 0, -1);
        return 0;
    }
    WNDPROC oldProc = reinterpret_cast<WNDPROC>(GetPropW(hwnd, kCtrlAProp));
    return CallWindowProcW(oldProc, hwnd, msg, wParam, lParam);
}

/**
 * @brief Enable Ctrl+A shortcut for edit control
 * @param edit Edit control handle
 */
void EnableCtrlA(HWND edit) {
    if (!edit) return;
    WNDPROC oldProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(edit, GWLP_WNDPROC));
    SetPropW(edit, kCtrlAProp, reinterpret_cast<HANDLE>(oldProc));
    SetWindowLongPtrW(edit, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(CtrlAEditProc));
}

/**
 * @brief Convert wide string to UTF-8 string
 */
string ToUtf8(const wstring& w) {
    string s; int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len > 1) { s.resize(len - 1); WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], len - 1, nullptr, nullptr); }
    return s;
}

/**
 * @brief Convert UTF-8 string to wide string
 */
wstring FromUtf8(const string& s) {
    wstring w; int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), static_cast<int>(s.size()), nullptr, 0);
    if (len > 0) { w.resize(len); MultiByteToWideChar(CP_UTF8, 0, s.c_str(), static_cast<int>(s.size()), &w[0], len); }
    return w;
}

/**
 * @brief Check if file exists
 * @param path Path to file
 * @return True if file exists, false otherwise
 */
bool FileExists(const wstring& path) {
    DWORD attr = GetFileAttributesW(path.c_str());
    return attr != INVALID_FILE_ATTRIBUTES && !(attr & FILE_ATTRIBUTE_DIRECTORY);
}

/**
 * @brief Get per-user config directory under %APPDATA%\cbfilter, creating it if necessary
 */
wstring GetConfigDirectory() {
    PWSTR known = nullptr;
    wstring base;
    if (SHGetKnownFolderPath(FOLDERID_RoamingAppData, KF_FLAG_CREATE, nullptr, &known) == S_OK && known) {
        base = known;
        CoTaskMemFree(known);
    }
    if (base.empty()) {
        wchar_t buf[MAX_PATH]; GetModuleFileNameW(nullptr, buf, MAX_PATH);
        wstring path(buf); size_t pos = path.find_last_of(L"\\/");
        if (pos != wstring::npos) path = path.substr(0, pos + 1);
        base = path;
    }
    if (!base.empty() && base.back() != L'\\' && base.back() != L'/') base += L"\\";
    base += L"cbfilter\\";
    CreateDirectoryW(base.c_str(), nullptr);  // ok if already exists
    return base;
}
/**
 * @brief Get the path to the configuration file
 * @return Full path to config.json under %APPDATA%\cbfilter
 */
wstring GetConfigPath() {
    wstring dir = GetConfigDirectory();
    return dir + L"config.json";
}

/**
 * @brief Get the directory that contains API provider definitions
 */
wstring GetApiDefDirectory() {
    wchar_t buf[MAX_PATH]; GetModuleFileNameW(nullptr, buf, MAX_PATH);
    wstring path(buf); size_t pos = path.find_last_of(L"\\/");
    if (pos != wstring::npos) path = path.substr(0, pos + 1);
    if (!path.empty() && path.back() != L'\\' && path.back() != L'/') path += L"\\";
    path += L"apidef\\";
    return path;
}

/**
 * @brief Get the path to the bundled default configuration file
 */
wstring GetDefaultConfigPath() {
    wchar_t buf[MAX_PATH]; GetModuleFileNameW(nullptr, buf, MAX_PATH);
    wstring path(buf); size_t pos = path.find_last_of(L"\\/");
    if (pos != wstring::npos) path = path.substr(0, pos + 1);
    path += L"defconf.json";
    return path;
}

#if DEBUG
/**
 * @brief Get the path to the log file
 * @return Full path to cbfilter.log in the application directory
 */
wstring GetLogPath() {
    wchar_t buf[MAX_PATH]; GetModuleFileNameW(nullptr, buf, MAX_PATH);
    wstring path(buf); size_t pos = path.find_last_of(L"\\/");
    if (pos != wstring::npos) path = path.substr(0, pos + 1);
    path += L"cbfilter.log"; return path;
}

/**
 * @brief Write a log message with timestamp to the log file
 * @param msg Message to log
 */
void LogLine(const wstring& msg) {
    wstring path = GetLogPath();
    FILE* fp = nullptr;
    if (_wfopen_s(&fp, path.c_str(), L"a, ccs=UTF-8") != 0 || !fp) return;
    SYSTEMTIME st; GetLocalTime(&st);
    fwprintf(fp, L"[%04d-%02d-%02d %02d:%02d:%02d] %s\n",
             st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, msg.c_str());
    fclose(fp);
}
#else
#define LogLine(msg) ((void)0)
#endif

/**
 * @brief Read UTF-8 text file into wide string
 */
wstring ReadUtf8File(const wstring& path) {
    FILE* fp = nullptr;
    if (_wfopen_s(&fp, path.c_str(), L"rb") != 0 || !fp) return L"";
    fseek(fp, 0, SEEK_END);
    long size = ftell(fp);
    fseek(fp, 0, SEEK_SET);
    string buf; if (size > 0) { buf.resize(size); fread(&buf[0], 1, size, fp); }
    fclose(fp);
    return FromUtf8(buf);
}

/**
 * @brief Write wide string as UTF-8 file (overwrite)
 */
bool WriteUtf8File(const wstring& path, const wstring& content) {
    FILE* fp = nullptr;
    if (_wfopen_s(&fp, path.c_str(), L"wb") != 0 || !fp) return false;
    string utf8 = ToUtf8(content);
    if (!utf8.empty()) fwrite(utf8.data(), 1, utf8.size(), fp);
    fclose(fp);
    return true;
}

wstring ReplaceAll(wstring s, const wstring& from, const wstring& to) {
    size_t pos = 0;
    while ((pos = s.find(from, pos)) != wstring::npos) {
        s.replace(pos, from.length(), to);
        pos += to.length();
    }
    return s;
}

wstring JsonEscape(const wstring& s) {
    wstring out;
    for (wchar_t c : s) {
        switch (c) {
        case L'\\': out += L"\\\\"; break;
        case L'"': out += L"\\\""; break;
        case L'\n': out += L"\\n"; break;
        case L'\r': out += L"\\r"; break;
        case L'\t': out += L"\\t"; break;
        default: out += c; break;
        }
    }
    return out;
}

wstring ReplacePlaceholders(const wstring& src, const ModelConfig& m, const wstring& systemPrompt, const wstring& prompt, const wstring& imageB64, const wstring& imageDataUrl, bool jsonEsc) {
    auto esc = [&](const wstring& v) { return jsonEsc ? JsonEscape(v) : v; };
    wstring out = src;
    out = ReplaceAll(out, L"<<model>>", esc(m.modelName));
    out = ReplaceAll(out, L"<<system_prompt>>", esc(systemPrompt));
    out = ReplaceAll(out, L"<<prompt>>", esc(prompt));
    out = ReplaceAll(out, L"<<input_text>>", esc(prompt));
    out = ReplaceAll(out, L"<<api_key>>", esc(m.apiKey));
    out = ReplaceAll(out, L"<<image_url>>", esc(imageDataUrl));
    out = ReplaceAll(out, L"<<image>>", esc(imageB64));
    return out;
}

bool ContainsNoCase(const wstring& hay, const wstring& needle) {
    auto toLow = [](wstring s) { for (auto& c : s) c = static_cast<wchar_t>(towlower(c)); return s; };
    return toLow(hay).find(toLow(needle)) != wstring::npos;
}

IOType ParseIOType(const wstring& s) {
    wstring lower = s;
    for (auto& c : lower) c = static_cast<wchar_t>(towlower(c));
    return (lower == L"image") ? IOType::Image : IOType::Text;
}
wstring IOTypeToConfig(IOType t) { return t == IOType::Image ? L"image" : L"text"; }

const TemplateDefinition* FindTemplateById(const wstring& id) {
    for (const auto& p : g_providers) {
        for (const auto& t : p.templates) if (t.id == id) return &t;
    }
    return nullptr;
}

wstring NormalizeProviderId(const wstring& raw) {
    if (raw.empty()) return raw;
    size_t pos = raw.find(L'-');
    if (pos != wstring::npos) return raw.substr(0, pos);
    return raw;
}

const ApiProvider* FindProviderById(const wstring& id) {
    for (const auto& p : g_providers) if (p.id == id) return &p;
    return nullptr;
}

/**
 * @brief Find a template by input and output type
 * @param provider API provider
 * @param input Input type
 * @param output Output type
 * @return Template definition, or nullptr if not found
 */
const TemplateDefinition* FindTemplateByIO(const ApiProvider& provider, IOType input, IOType output) {
    for (const auto& t : provider.templates) if (t.input == input && t.output == output) return &t;
    return nullptr;
}

/**
 * @brief Find a template by input and output type
 * @param input Input type
 * @param output Output type
 * @return Template definition, or nullptr if not found
 */
const TemplateDefinition* FindTemplateAny(IOType input, IOType output) {
    for (const auto& p : g_providers) {
        if (const auto* t = FindTemplateByIO(p, input, output)) return t;
    }
    return nullptr;
}

/**
 * @brief Ensure model providers are set
 */
void EnsureModelProviders() {
    if (g_providers.empty()) return;
    const wstring first = g_providers.front().id;
    for (auto& m : g_models) if (m.providerId.empty()) m.providerId = first;
}

/**
 * @brief Load API definitions from JSON files
 */
void LoadApiDefinitions() {
    using namespace winrt::Windows::Data::Json;
    g_providers.clear();
    vector<ApiProvider> providers;
    wstring dir = GetApiDefDirectory();
    wstring pattern = dir + L"*.json";
    WIN32_FIND_DATAW fd{};
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) {
        LogLine(L"apidef directory missing or empty: " + dir);
        return;
    }
    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        wstring fileName = fd.cFileName;
        wstring fullPath = dir + fileName;
        wstring text = ReadUtf8File(fullPath);
        if (text.empty()) { LogLine(L"apidef file missing or empty: " + fullPath); continue; }
        try {
            JsonObject root = JsonObject::Parse(text);
            ApiProvider provider{};
            size_t dot = fileName.find_last_of(L'.');
            provider.id = (dot == wstring::npos) ? fileName : fileName.substr(0, dot);
            if (root.HasKey(L"default-endpoint")) provider.defaultEndpoint = wstring(root.GetNamedString(L"default-endpoint", L"").c_str());
            for (auto const& kv : root) {
                if (kv.Key() == L"models") {
                    if (kv.Value().ValueType() != JsonValueType::Object) continue;
                    JsonObject mobj = kv.Value().GetObject();
                    provider.modelsEndpoint = wstring(mobj.GetNamedString(L"endpoint", L"").c_str());
                    provider.modelsMethod = wstring(mobj.GetNamedString(L"method", L"GET").c_str());
                    provider.modelsResultPath = wstring(mobj.GetNamedString(L"result", L"data").c_str());
                    if (mobj.HasKey(L"headers")) {
                        JsonObject h = mobj.GetNamedObject(L"headers");
                        for (auto const& hk : h) provider.modelsHeaders.push_back({ wstring(hk.Key().c_str()), wstring(hk.Value().GetString().c_str()) });
                    }
                    if (mobj.HasKey(L"payload")) {
                        JsonValue val = mobj.GetNamedValue(L"payload");
                        provider.modelsPayload = val.Stringify().c_str();
                    }
                    continue;
                }
                if (kv.Value().ValueType() != JsonValueType::Object) continue;
                wstring key = kv.Key().c_str();
                JsonObject obj = kv.Value().GetObject();
                TemplateDefinition t{};
                t.id = key;
                t.providerId = provider.id;
                size_t sep = key.find(L'-');
                wstring in = sep == wstring::npos ? key : key.substr(0, sep);
                wstring out = sep == wstring::npos ? key : key.substr(sep + 1);
                t.input = ParseIOType(in);
                t.output = ParseIOType(out);
                t.endpoint = wstring(obj.GetNamedString(L"endpoint", L"/").c_str());
                t.resultPath = wstring(obj.GetNamedString(L"result", L"").c_str());
                if (obj.HasKey(L"headers")) {
                    JsonObject h = obj.GetNamedObject(L"headers");
                    for (auto const& hk : h) {
                        t.headers.push_back({ wstring(hk.Key().c_str()), wstring(hk.Value().GetString().c_str()) });
                    }
                }
                if (obj.HasKey(L"payload")) {
                    JsonValue val = obj.GetNamedValue(L"payload");
                    t.payload = val.Stringify().c_str();
                }
                if (!t.id.empty()) provider.templates.push_back(move(t));
            }
            if (!provider.id.empty() && !provider.templates.empty()) providers.push_back(move(provider));
        } catch (const winrt::hresult_error& e) {
            wstring msg = L"LoadApiDefinitions parse failed for ";
            msg += fullPath;
            msg += L": ";
            msg += e.message().c_str();
            LogLine(msg);
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
    if (!providers.empty()) g_providers = move(providers);
}

struct ApiCallResult { wstring text; HBITMAP image{}; };

/**
 * @brief Encrypt API key using DPAPI and encode as base64 with prefix
 * @param plain Plaintext API key
 * @return Encoded string starting with "dpapi:" or empty on failure
 */
wstring ProtectApiKey(const wstring& plain) {
    if (plain.empty()) return L"";
    DATA_BLOB in{ static_cast<DWORD>(plain.size() * sizeof(wchar_t)), reinterpret_cast<BYTE*>(const_cast<wchar_t*>(plain.data())) };
    DATA_BLOB out{};
    if (!CryptProtectData(&in, L"cbfilter", nullptr, nullptr, nullptr, CRYPTPROTECT_UI_FORBIDDEN, &out)) {
        LogLine(L"CryptProtectData failed: " + to_wstring(GetLastError()));
        return L"";
    }
    DWORD b64Len = 0;
    if (!CryptBinaryToStringW(out.pbData, out.cbData, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &b64Len)) {
        LocalFree(out.pbData);
        LogLine(L"CryptBinaryToStringW length failed: " + to_wstring(GetLastError()));
        return L"";
    }
    wstring b64(b64Len, L'\0');
    if (!CryptBinaryToStringW(out.pbData, out.cbData, CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, b64.data(), &b64Len)) {
        LocalFree(out.pbData);
        LogLine(L"CryptBinaryToStringW failed: " + to_wstring(GetLastError()));
        return L"";
    }
    LocalFree(out.pbData);
    b64.resize(b64Len);
    return L"dpapi:" + b64;
}

/**
 * @brief Decrypt API key stored with ProtectApiKey
 * @param stored Stored string (possibly plaintext for backward compatibility)
 * @return Decrypted plaintext API key, or empty on failure
 */
wstring UnprotectApiKey(const wstring& stored) {
    if (stored.empty()) return L"";
    constexpr wchar_t kPrefix[] = L"dpapi:";
    const wstring prefix = kPrefix;
    if (stored.rfind(prefix, 0) != 0) return stored;  // Legacy plaintext
    wstring b64 = stored.substr(prefix.size());
    DWORD binSize = 0;
    if (!CryptStringToBinaryW(b64.c_str(), 0, CRYPT_STRING_BASE64, nullptr, &binSize, nullptr, nullptr)) {
        LogLine(L"CryptStringToBinaryW length failed: " + to_wstring(GetLastError()));
        return L"";
    }
    vector<BYTE> buf(binSize);
    if (!CryptStringToBinaryW(b64.c_str(), 0, CRYPT_STRING_BASE64, buf.data(), &binSize, nullptr, nullptr)) {
        LogLine(L"CryptStringToBinaryW failed: " + to_wstring(GetLastError()));
        return L"";
    }
    DATA_BLOB in{ binSize, buf.data() };
    DATA_BLOB out{};
    if (!CryptUnprotectData(&in, nullptr, nullptr, nullptr, nullptr, CRYPTPROTECT_UI_FORBIDDEN, &out)) {
        LogLine(L"CryptUnprotectData failed: " + to_wstring(GetLastError()));
        return L"";
    }
    if (out.cbData % sizeof(wchar_t) != 0) {
        LogLine(L"CryptUnprotectData returned unexpected byte length: " + to_wstring(out.cbData));
        LocalFree(out.pbData);
        return L"";
    }
    wstring plain(reinterpret_cast<wchar_t*>(out.pbData), out.cbData / sizeof(wchar_t));
    LocalFree(out.pbData);
    return plain;
}

/**
 * @brief Create default configuration for first run
 */
 winrt::Windows::Data::Json::JsonObject CreateDefaultConfig() {
    using namespace winrt::Windows::Data::Json;
    wstring defPath = GetDefaultConfigPath();
    wstring text = ReadUtf8File(defPath);
    if (!text.empty()) {
        try {
            return JsonObject::Parse(text);
        } catch (const winrt::hresult_error& e) {
            wstring msg = L"defconf.json parse failed: ";
            msg += e.message().c_str();
            LogLine(msg);
        }
    } else {
        LogLine(L"defconf.json missing or empty: " + defPath);
    }
    JsonObject fallback;
    fallback.SetNamedValue(L"language", JsonValue::CreateStringValue(L"en"));
    JsonObject hotkey;
    hotkey.SetNamedValue(L"modifiers", JsonValue::CreateNumberValue(static_cast<double>(MOD_WIN | MOD_ALT)));
    hotkey.SetNamedValue(L"key", JsonValue::CreateNumberValue(static_cast<double>('V')));
    fallback.SetNamedValue(L"hotkey", hotkey);
    JsonArray models;
    JsonObject model;
    model.SetNamedValue(L"name", JsonValue::CreateStringValue(L"Translate"));
    model.SetNamedValue(L"serverUrl", JsonValue::CreateStringValue(L"https://api.openai.com/v1"));
    model.SetNamedValue(L"modelName", JsonValue::CreateStringValue(L"gpt-5.1"));
    model.SetNamedValue(L"providerId", JsonValue::CreateStringValue(L"OpenAI"));
    models.Append(model);
    fallback.SetNamedValue(L"models", models);
    JsonArray filters;
    JsonObject filter;
    filter.SetNamedValue(L"title", JsonValue::CreateStringValue(L"Translate"));
    filter.SetNamedValue(L"input", JsonValue::CreateStringValue(L"text"));
    filter.SetNamedValue(L"output", JsonValue::CreateStringValue(L"text"));
    filter.SetNamedValue(L"modelIndex", JsonValue::CreateNumberValue(static_cast<double>(0)));
    filter.SetNamedValue(L"prompt", JsonValue::CreateStringValue(L"Translate into English."));
    filters.Append(filter);
    fallback.SetNamedValue(L"filters", filters);
    return fallback;
}

/**
 * @brief Save current model and filter configurations to config.json
 */
void SaveConfig() {
    using namespace winrt::Windows::Data::Json;
    wstring cfg = GetConfigPath();
    EnsureModelProviders();
    JsonObject root;
    root.SetNamedValue(L"language", JsonValue::CreateStringValue(g_language));
    JsonObject hotkey;
    hotkey.SetNamedValue(L"modifiers", JsonValue::CreateNumberValue(static_cast<double>(g_hotkeyModifiers)));
    hotkey.SetNamedValue(L"key", JsonValue::CreateNumberValue(static_cast<double>(g_hotkeyKey)));
    root.SetNamedValue(L"hotkey", hotkey);

    JsonArray models;
    for (const auto& m : g_models) {
        JsonObject obj;
        obj.SetNamedValue(L"name", JsonValue::CreateStringValue(m.name));
        obj.SetNamedValue(L"serverUrl", JsonValue::CreateStringValue(m.serverUrl));
        obj.SetNamedValue(L"modelName", JsonValue::CreateStringValue(m.modelName));
        obj.SetNamedValue(L"providerId", JsonValue::CreateStringValue(m.providerId));
        wstring protectedKey = ProtectApiKey(m.apiKey);
        if (protectedKey.empty() && !m.apiKey.empty()) protectedKey = m.apiKey;  // Fallback to avoid losing key
        obj.SetNamedValue(L"apiKey", JsonValue::CreateStringValue(protectedKey));
        models.Append(obj);
    }
    root.SetNamedValue(L"models", models);

    JsonArray filters;
    for (const auto& f : g_filters) {
        JsonObject obj;
        obj.SetNamedValue(L"title", JsonValue::CreateStringValue(f.title));
        obj.SetNamedValue(L"input", JsonValue::CreateStringValue(f.input == IOType::Text ? L"text" : L"image"));
        obj.SetNamedValue(L"output", JsonValue::CreateStringValue(f.output == IOType::Text ? L"text" : L"image"));
        obj.SetNamedValue(L"modelIndex", JsonValue::CreateNumberValue(static_cast<double>(f.modelIndex)));
        obj.SetNamedValue(L"prompt", JsonValue::CreateStringValue(f.prompt));
        filters.Append(obj);
    }
    root.SetNamedValue(L"filters", filters);

    wstring jsonText = root.Stringify().c_str();
    if (!WriteUtf8File(cfg, jsonText)) {
        LogLine(L"Failed to write config to " + cfg);
    }
}

/**
 * @brief Load filters from JSON object
 * @param root JSON object
 * @param target Vector to store filter definitions
 */
void LoadFiltersFromJson(const winrt::Windows::Data::Json::JsonObject& root, vector<FilterDefinition>& target) {
    using namespace winrt::Windows::Data::Json;
    if (!root.HasKey(L"filters")) return;
    vector<FilterDefinition> v;
    for (auto const& item : root.GetNamedArray(L"filters")) {
        if (item.ValueType() != JsonValueType::Object) continue;
        JsonObject obj = item.GetObject();
        FilterDefinition f{};
        f.title = wstring(obj.GetNamedString(L"title", L"").c_str());
        wstring input = wstring(obj.GetNamedString(L"input", L"text").c_str());
        wstring output = wstring(obj.GetNamedString(L"output", L"text").c_str());
        f.input = ParseIOType(input);
        f.output = ParseIOType(output);
        f.modelIndex = static_cast<size_t>(obj.GetNamedNumber(L"modelIndex", 0));
        f.prompt = wstring(obj.GetNamedString(L"prompt", L"").c_str());
        if (!f.title.empty()) v.push_back(move(f));
    }
    if (!v.empty()) target = move(v);
}

/**
 * @brief Load model and filter configurations from config.json
 */
void LoadConfig() {
    using namespace winrt::Windows::Data::Json;
    wstring cfg = GetConfigPath();
    try {
        JsonObject root = FileExists(cfg) ? JsonObject::Parse(ReadUtf8File(cfg)) : CreateDefaultConfig();
        if (root.HasKey(L"language")) g_language = wstring(root.GetNamedString(L"language").c_str());
        if (root.HasKey(L"hotkey")) {
            JsonObject hotkey = root.GetNamedObject(L"hotkey");
            g_hotkeyModifiers = static_cast<UINT>(hotkey.GetNamedNumber(L"modifiers", g_hotkeyModifiers));
            g_hotkeyKey = static_cast<UINT>(hotkey.GetNamedNumber(L"key", g_hotkeyKey));
        }
        if (root.HasKey(L"models")) {
            vector<ModelConfig> v;
            for (auto const& item : root.GetNamedArray(L"models")) {
                if (item.ValueType() != JsonValueType::Object) continue;
                JsonObject obj = item.GetObject();
                ModelConfig m{};
                m.name = wstring(obj.GetNamedString(L"name", L"").c_str());
                m.serverUrl = wstring(obj.GetNamedString(L"serverUrl", L"").c_str());
                m.modelName = wstring(obj.GetNamedString(L"modelName", L"").c_str());
                wstring provider = wstring(obj.GetNamedString(L"providerId", L"").c_str());
                m.providerId = NormalizeProviderId(provider);
                m.apiKey = UnprotectApiKey(wstring(obj.GetNamedString(L"apiKey", L"").c_str()));
                if (!m.name.empty()) v.push_back(move(m));
            }
            if (!v.empty()) g_models = move(v);
        }
        LoadFiltersFromJson(root, g_filters);
        if (!g_filters.empty()) {
            for (auto& f : g_filters) if (f.modelIndex >= g_models.size()) f.modelIndex = 0;
        }
        EnsureModelProviders();
    } catch (const winrt::hresult_error& e) {
        wstring msg = L"LoadConfig JSON parse failed: ";
        msg += e.message().c_str();
        LogLine(msg);
    }
}

/**
 * @brief Get the CLSID for PNG encoder in GDI+
 * @return Pointer to PNG encoder CLSID, or nullptr if not found
 */
const CLSID* GetPngClsid() {
    static CLSID clsid{};
    static bool init = false;
    if (init) return clsid == CLSID{} ? nullptr : &clsid;
    init = true;
    UINT num = 0, size = 0;
    if (Gdiplus::GetImageEncodersSize(&num, &size) != Gdiplus::Ok || size == 0) return nullptr;
    vector<BYTE> buf(size);
    auto* encoders = reinterpret_cast<Gdiplus::ImageCodecInfo*>(buf.data());
    if (Gdiplus::GetImageEncoders(num, size, encoders) != Gdiplus::Ok) return nullptr;
    for (UINT i = 0; i < num; ++i) {
        if (wcscmp(encoders[i].MimeType, L"image/png") == 0) { clsid = encoders[i].Clsid; return &clsid; }
    }
    return nullptr;
}

/**
 * @brief Convert a bitmap to base64-encoded PNG string
 * @param bmp Bitmap handle to convert
 * @param out Output parameter for base64 string
 * @return true on success, false on failure
 */
bool BitmapToBase64Png(HBITMAP bmp, wstring& out) {
    const CLSID* clsid = GetPngClsid();
    if (!clsid) return false;
    Gdiplus::Bitmap bitmap(bmp, nullptr);
    IStream* stream = nullptr;
    if (CreateStreamOnHGlobal(nullptr, TRUE, &stream) != S_OK) return false;
    if (bitmap.Save(stream, clsid, nullptr) != Gdiplus::Ok) { stream->Release(); return false; }
    HGLOBAL hMem = nullptr;
    if (GetHGlobalFromStream(stream, &hMem) != S_OK) { stream->Release(); return false; }
    SIZE_T size = GlobalSize(hMem);
    BYTE* data = static_cast<BYTE*>(GlobalLock(hMem));
    if (!data || size == 0) { if (data) GlobalUnlock(hMem); stream->Release(); return false; }
    DWORD outLen = 0;
    if (!CryptBinaryToStringW(data, static_cast<DWORD>(size), CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &outLen)) {
        GlobalUnlock(hMem); stream->Release(); return false;
    }
    out.resize(outLen);
    if (!CryptBinaryToStringW(data, static_cast<DWORD>(size), CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, out.data(), &outLen)) {
        GlobalUnlock(hMem); stream->Release(); return false;
    }
    out.resize(outLen);
    GlobalUnlock(hMem);
    stream->Release();
    return true;
}

/**
 * @brief Convert base64-encoded PNG string to bitmap
 * @param b64 Base64-encoded PNG data
 * @return Bitmap handle (caller must DeleteObject), or nullptr on failure
 */
HBITMAP Base64ToBitmap(const wstring& b64) {
    DWORD binSize = 0;
    if (!CryptStringToBinaryW(b64.c_str(), 0, CRYPT_STRING_BASE64, nullptr, &binSize, nullptr, nullptr)) return nullptr;
    vector<BYTE> buf(binSize);
    if (!CryptStringToBinaryW(b64.c_str(), 0, CRYPT_STRING_BASE64, buf.data(), &binSize, nullptr, nullptr)) return nullptr;
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, binSize);
    if (!hMem) return nullptr;
    void* dst = GlobalLock(hMem);
    memcpy(dst, buf.data(), binSize);
    GlobalUnlock(hMem);
    IStream* stream = nullptr;
    if (CreateStreamOnHGlobal(hMem, TRUE, &stream) != S_OK) { GlobalFree(hMem); return nullptr; }
    Gdiplus::Bitmap* bmp = Gdiplus::Bitmap::FromStream(stream);
    stream->Release();
    if (!bmp) { GlobalFree(hMem); return nullptr; }
    HBITMAP hOut = nullptr;
    bmp->GetHBITMAP(static_cast<Gdiplus::ARGB>(0xFFFFFFFF), &hOut);
    delete bmp;
    return hOut;
}

/**
 * @brief Extract content field from OpenAI API JSON response
 * @param json JSON response string
 * @return Extracted content text, or empty string if not found
 */
wstring ExtractContent(const wstring& json) {
    size_t p = json.find(L"\"content\"");
    if (p == wstring::npos) return L"";
    p = json.find(L"\"", p + 9);
    if (p == wstring::npos) return L"";
    ++p;
    wstring out;
    while (p < json.size() && json[p] != L'\"') {
        if (json[p] == L'\\' && p + 1 < json.size()) {
            if (json[p + 1] == L'n') { out += L'\n'; p += 2; continue; }
            if (json[p + 1] == L'\"') { out += L'\"'; p += 2; continue; }
        }
        out += json[p]; ++p;
    }
    return out;
}

/**
 * @brief Extract base64 image data from OpenAI image generation API response
 * @param json JSON response string
 * @return Base64-encoded image data, or empty string if not found
 */
 wstring ExtractB64Image(const wstring& json) {
    size_t p = json.find(L"\"b64_json\"");
    if (p == wstring::npos) return L"";
    p = json.find(L"\"", p + 10);
    if (p == wstring::npos) return L"";
    ++p;
    wstring out;
    while (p < json.size() && json[p] != L'\"') {
        if (json[p] == L'\\' && p + 1 < json.size()) {
            if (json[p + 1] == L'"') { out += L'\"'; p += 2; continue; }
            if (json[p + 1] == L'\\') { out += L'\\'; p += 2; continue; }
            if (json[p + 1] == L'/') { out += L'/'; p += 2; continue; }
        }
        out += json[p]; ++p;
    }
    return out;
}

/**
 * @brief Extract image URL from OpenRouter chat/completions response
 * @param json JSON response string
 * @return Base64-encoded image data (without data:image/png;base64, prefix), or empty string if not found
 */
 wstring ExtractImageFromChatResponse(const wstring& json) {
    // Look for "images" array in the response
    size_t imagesPos = json.find(L"\"images\"");
    if (imagesPos == wstring::npos) {
        // Try alternative: "image_url" directly
        imagesPos = json.find(L"\"image_url\"");
        if (imagesPos == wstring::npos) return L"";
    }
    
    // Find image_url.url field (could be "image_url" or "imageUrl")
    size_t urlPos = json.find(L"\"image_url\"", imagesPos);
    if (urlPos == wstring::npos) {
        urlPos = json.find(L"\"imageUrl\"", imagesPos);
    }
    if (urlPos == wstring::npos) return L"";
    
    // Find url field (could be "url" or nested structure)
    size_t urlFieldPos = json.find(L"\"url\"", urlPos);
    if (urlFieldPos == wstring::npos) return L"";
    
    // Find the value (data:image/png;base64,...)
    // Skip to the colon and opening quote
    size_t colonPos = json.find(L":", urlFieldPos + 5);
    if (colonPos == wstring::npos) return L"";
    
    // Find the opening quote after colon
    size_t valueStart = json.find(L"\"", colonPos);
    if (valueStart == wstring::npos) return L"";
    valueStart++;
    
    // Find the closing quote, but handle escaped quotes
    size_t valueEnd = valueStart;
    while (valueEnd < json.length()) {
        if (json[valueEnd] == L'\"' && (valueEnd == valueStart || json[valueEnd - 1] != L'\\')) {
            break;
        }
        valueEnd++;
    }
    if (valueEnd >= json.length()) return L"";
    
    wstring dataUrl = json.substr(valueStart, valueEnd - valueStart);
    
    // Unescape JSON string
    wstring unescaped;
    for (size_t i = 0; i < dataUrl.length(); ++i) {
        if (dataUrl[i] == L'\\' && i + 1 < dataUrl.length()) {
            if (dataUrl[i + 1] == L'\\') { unescaped += L'\\'; i++; continue; }
            if (dataUrl[i + 1] == L'\"') { unescaped += L'\"'; i++; continue; }
            if (dataUrl[i + 1] == L'n') { unescaped += L'\n'; i++; continue; }
            if (dataUrl[i + 1] == L'r') { unescaped += L'\r'; i++; continue; }
            if (dataUrl[i + 1] == L't') { unescaped += L'\t'; i++; continue; }
        }
        unescaped += dataUrl[i];
    }
    dataUrl = unescaped;
    
    // Extract base64 part (remove "data:image/png;base64," prefix)
    size_t commaPos = dataUrl.find(L",");
    if (commaPos != wstring::npos) {
        return dataUrl.substr(commaPos + 1);
    }
    
    return dataUrl; // Return as-is if no comma found
}

/**
 * @brief Prepare endpoint for HTTP request
 * @param serverUrl Server URL
 * @param tplPath Template path
 * @param host Host name
 * @param path Path
 * @param useHttps Use HTTPS
 * @return true if endpoint is valid, false otherwise
 */
bool PrepareEndpoint(const wstring& serverUrl, const wstring& tplPath, wstring& host, wstring& path, bool& useHttps) {
    host = serverUrl;
    path = tplPath;
    if (path.empty()) path = L"/v1/chat/completions";
    // If template path is absolute URL, override host/path
    if (path.rfind(L"http://", 0) == 0 || path.rfind(L"https://", 0) == 0) {
        host = path;
        path = L"";
    }
    if (host.rfind(L"https://", 0) == 0) { host = host.substr(8); useHttps = true; }
    else if (host.rfind(L"http://", 0) == 0) { host = host.substr(7); useHttps = false; }
    size_t slash = host.find(L'/');
    if (slash != wstring::npos) {
        path = host.substr(slash) + (path.empty() ? L"" : path);
        host = host.substr(0, slash);
    }
    if (!path.empty() && path.front() != L'/') path = L"/" + path;
    return !host.empty();
}

/**
 * @brief Build body from template definition
 * @param tpl Template definition
 * @param m Model configuration
 * @param systemPrompt System prompt
 * @param prompt Prompt text
 * @param imageB64 Base64 encoded image data
 * @param imageDataUrl Image data URL
 * @return Body string
 */
wstring BuildBodyFromTemplate(const TemplateDefinition& tpl, const ModelConfig& m, const wstring& systemPrompt, const wstring& prompt, const wstring& imageB64, const wstring& imageDataUrl) {
    return ReplacePlaceholders(tpl.payload, m, systemPrompt, prompt, imageB64, imageDataUrl, true);
}

/**
 * @brief Build header string from template definition
 * @param tpl Template definition
 * @param m Model configuration
 * @param systemPrompt System prompt
 * @param prompt Prompt text
 * @param imageB64 Base64 encoded image data
 * @param imageDataUrl Image data URL
 * @return Header string
 */
wstring BuildHeaderString(const TemplateDefinition& tpl, const ModelConfig& m, const wstring& systemPrompt, const wstring& prompt, const wstring& imageB64, const wstring& imageDataUrl) {
    wstring header;
    for (const auto& kv : tpl.headers) {
        wstring val = ReplacePlaceholders(kv.second, m, systemPrompt, prompt, imageB64, imageDataUrl, false);
        header += kv.first + L": " + val + L"\r\n";
    }
    return header;
}

/**
 * @brief Build header string from vector of key-value pairs
 * @param headers Vector of key-value pairs
 * @param m Model configuration
 * @return Header string
 */
wstring BuildHeaderString(const vector<pair<wstring, wstring>>& headers, const ModelConfig& m) {
    wstring header;
    for (const auto& kv : headers) {
        wstring val = ReplacePlaceholders(kv.second, m, L"", L"", L"", L"", false);
        header += kv.first + L": " + val + L"\r\n";
    }
    return header;
}

/**
 * @brief Make an HTTP request with headers
 * @param host Host name
 * @param path Path
 * @param useHttps Use HTTPS
 * @param headers Headers
 * @param body Body
 * @param method Method
 * @param err Error message
 * @return Result of the request
 */
wstring HttpRequestWithHeaders(const wstring& host, const wstring& path, bool useHttps, const wstring& headers, const string& body, const wstring& method, wstring* err) {
    wstring result;
    auto setErr = [&](const wstring& m) { if (err) *err = m; };
    HINTERNET hs = WinHttpOpen(L"cbfilter/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hs) { setErr(L"WinHttpOpen failed: " + to_wstring(GetLastError())); return L""; }
    HINTERNET hc = WinHttpConnect(hs, host.c_str(), INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!hc) { setErr(L"WinHttpConnect failed: " + to_wstring(GetLastError())); WinHttpCloseHandle(hs); return L""; }
    DWORD flags = useHttps ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hr = WinHttpOpenRequest(hc, method.c_str(), path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!hr) { setErr(L"WinHttpOpenRequest failed: " + to_wstring(GetLastError())); WinHttpCloseHandle(hc); WinHttpCloseHandle(hs); return L""; }
    const bool hasBody = !body.empty();
    BOOL ok = WinHttpSendRequest(hr, headers.c_str(), (DWORD)headers.size(), hasBody ? (LPVOID)body.data() : nullptr, hasBody ? (DWORD)body.size() : 0, hasBody ? (DWORD)body.size() : 0, 0);
    if (!ok) { setErr(L"WinHttpSendRequest failed: " + to_wstring(GetLastError())); }
    if (ok) ok = WinHttpReceiveResponse(hr, nullptr);
    if (!ok) { setErr(L"WinHttpReceiveResponse failed: " + to_wstring(GetLastError())); }
    if (ok) {
        DWORD status = 0, len = sizeof(status);
        if (WinHttpQueryHeaders(hr, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &len, WINHTTP_NO_HEADER_INDEX)) {
            if (status >= 400) setErr(L"HTTP status " + to_wstring(status));
        }
    }
    if (ok) {
        for (;;) {
            DWORD dwSize = 0;
            if (!WinHttpQueryDataAvailable(hr, &dwSize)) break;
            if (dwSize == 0) break;
            string buf; buf.resize(dwSize);
            DWORD dwDownloaded = 0;
            if (!WinHttpReadData(hr, buf.data(), dwSize, &dwDownloaded) || dwDownloaded == 0) break;
            buf.resize(dwDownloaded);
            wstring wbuf; int wlen = MultiByteToWideChar(CP_UTF8, 0, buf.c_str(), static_cast<int>(buf.size()), nullptr, 0);
            if (wlen > 0) { size_t cur = result.size(); result.resize(cur + wlen); MultiByteToWideChar(CP_UTF8, 0, buf.c_str(), static_cast<int>(buf.size()), result.data() + cur, wlen); }
        }
    }
    WinHttpCloseHandle(hr); WinHttpCloseHandle(hc); WinHttpCloseHandle(hs);
    return result;
}

/**
 * @brief Extract value from JSON by path
 * @param json JSON string
 * @param path Path to extract
 * @return Extracted value, or empty string if not found
 */
wstring ExtractByPath(const wstring& json, const wstring& path) {
    using namespace winrt::Windows::Data::Json;
    try {
        JsonValue val = JsonValue::Parse(json);
        vector<wstring> parts;
        size_t start = 0;
        for (size_t i = 0; i <= path.size(); ++i) {
            if (i == path.size() || path[i] == L'.') { parts.push_back(path.substr(start, i - start)); start = i + 1; }
        }
        IJsonValue cur = val;
        for (const auto& p : parts) {
            wstring key = p; int idx = -1;
            size_t lb = p.find(L'[');
            if (lb != wstring::npos && p.back() == L']') {
                key = p.substr(0, lb);
                idx = _wtoi(p.substr(lb + 1, p.size() - lb - 2).c_str());
            }
            if (cur.ValueType() == JsonValueType::Object) {
                JsonObject obj = cur.GetObject();
                if (!key.empty()) cur = obj.GetNamedValue(key);
                else return L"";
            } else if (cur.ValueType() == JsonValueType::Array) {
                JsonArray arr = cur.GetArray();
                if (idx >= 0 && idx < static_cast<int>(arr.Size())) cur = arr.GetAt(idx);
                else return L"";
                idx = -1;
            }
            if (idx >= 0) {
                JsonArray arr = cur.GetArray();
                if (idx >= 0 && idx < static_cast<int>(arr.Size())) cur = arr.GetAt(idx);
            }
        }
        if (cur.ValueType() == JsonValueType::String) return cur.GetString().c_str();
        return cur.Stringify().c_str();
    } catch (...) {
        return L"";
    }
}

/**
 * @brief Build multipart/form-data body for API request
 * @param boundary Boundary string
 * @param model Model name
 * @param prompt Prompt text
 * @param imageB64 Base64 encoded image data
 * @return Multipart/form-data body string
 */
string BuildMultipartBody(const wstring& boundary, const wstring& model, const wstring& prompt, const wstring& imageB64) {
    vector<BYTE> img;
    if (!imageB64.empty()) {
        DWORD size = 0;
        if (CryptStringToBinaryW(imageB64.c_str(), 0, CRYPT_STRING_BASE64, nullptr, &size, nullptr, nullptr)) {
            img.resize(size);
            if (!CryptStringToBinaryW(imageB64.c_str(), 0, CRYPT_STRING_BASE64, img.data(), &size, nullptr, nullptr)) img.clear();
            else img.resize(size);
        }
    }
    string b = "";
    string bnd = ToUtf8(boundary);
    auto addText = [&](const string& name, const string& value) {
        b += "--" + bnd + "\r\n";
        b += "Content-Disposition: form-data; name=\"" + name + "\"\r\n\r\n";
        b += value + "\r\n";
    };
    addText("model", ToUtf8(model));
    addText("prompt", ToUtf8(prompt));
    if (!img.empty()) {
        b += "--" + bnd + "\r\n";
        b += "Content-Disposition: form-data; name=\"image\"; filename=\"image.png\"\r\n";
        b += "Content-Type: image/png\r\n\r\n";
        b.insert(b.end(), img.begin(), img.end());
        b += "\r\n";
    }
    b += "--" + bnd + "--\r\n";
    return b;
}

/**
 * @brief Call a template API
 * @param tpl Template definition
 * @param m Model configuration
 * @param systemPrompt System prompt
 * @param prompt Prompt text
 * @param imageB64 Base64 encoded image data
 * @param imageDataUrl Image data URL
 * @return API call result
 */
ApiCallResult CallTemplate(const TemplateDefinition& tpl, const ModelConfig& m, const wstring& systemPrompt, const wstring& prompt, const wstring& imageB64, const wstring& imageDataUrl) {
    ApiCallResult result;
    wstring endpoint = ReplacePlaceholders(tpl.endpoint, m, systemPrompt, prompt, imageB64, imageDataUrl, false);
    wstring host, path; bool useHttps = true;
    if (!PrepareEndpoint(m.serverUrl, endpoint, host, path, useHttps)) {
        LogLine(L"PrepareEndpoint failed");
        return result;
    }
    wstring body = BuildBodyFromTemplate(tpl, m, systemPrompt, prompt, imageB64, imageDataUrl);
    wstring headers = BuildHeaderString(tpl, m, systemPrompt, prompt, imageB64, imageDataUrl);
    string utf8;
    wstring adjHeaders = headers;
    if (ContainsNoCase(headers, L"multipart/form-data")) {
        wstring boundary = L"----cbfilterboundary";
        adjHeaders = ReplaceAll(headers, L"multipart/form-data", L"multipart/form-data; boundary=" + boundary);
        utf8 = BuildMultipartBody(boundary, m.modelName, prompt, imageB64);
    } else {
        utf8 = ToUtf8(body);
    }
    LogLine(L"request host: " + host);
    LogLine(L"request path: " + path);
    wstring wbody; int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, nullptr, 0);
    if (wlen > 1) { wbody.resize(wlen - 1); MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, wbody.data(), wlen - 1); }
    LogLine(L"body: " + wbody);
    wstring err;
    wstring resp = HttpRequestWithHeaders(host, path, useHttps, adjHeaders, utf8, L"POST", &err);
    if (!err.empty()) LogLine(L"template request error: " + err);
    if (resp.empty()) return result;
    if (tpl.output == IOType::Text) {
        if (!tpl.resultPath.empty()) result.text = ExtractByPath(resp, tpl.resultPath);
        if (result.text.empty()) result.text = ExtractContent(resp);
        if (result.text.empty()) LogLine(L"template response empty content. resp=" + resp.substr(0, 512));
    } else {
        wstring b64 = tpl.resultPath.empty() ? L"" : ExtractByPath(resp, tpl.resultPath);
        if (b64.find(L"data:image") != wstring::npos) {
            size_t c = b64.find(L",");
            if (c != wstring::npos) b64 = b64.substr(c + 1);
        }
        if (b64.empty()) b64 = ExtractB64Image(resp);
        if (b64.empty()) b64 = ExtractContent(resp);
        if (b64.find(L"data:image") != wstring::npos) {
            size_t c = b64.find(L",");
            if (c != wstring::npos) b64 = b64.substr(c + 1);
        }
        if (!b64.empty()) result.image = Base64ToBitmap(b64);
        if (!result.image) LogLine(L"template response produced no image");
    }
    return result;
}

/**
 * @brief Execute a filter transformation on clipboard content
 * @param f Filter definition to execute
 * @return true on success, false on failure
 * 
 * This function handles four transformation types:
 * - Text -> Text: Text completion/translation
 * - Text -> Image: Image generation from text
 * - Image -> Text: Vision API (image description/analysis)
 * - Image -> Image: Image-to-image transformation
 */
bool RunFilter(const FilterDefinition& f) {
    LogLine(L"RunFilter: " + f.title + L" input=" + IOTypeToString(f.input) + L" output=" + IOTypeToString(f.output));
    auto modelRef = [&]() -> const ModelConfig& { return g_models[f.modelIndex < g_models.size() ? f.modelIndex : 0]; };
    const ModelConfig& m = modelRef();
    const ApiProvider* provider = FindProviderById(m.providerId);
    if (!provider && !g_providers.empty()) provider = &g_providers.front();
    const TemplateDefinition* tpl = provider ? FindTemplateByIO(*provider, f.input, f.output) : nullptr;
    if (!tpl) tpl = FindTemplateAny(f.input, f.output);
    if (!tpl) { LogLine(L"fail: no matching template"); return false; }
    try {
        wstring textInput;
        wstring imageB64;
        if (tpl->input == IOType::Text) {
            textInput = GetClipboardText();
            if (textInput.empty()) { LogLine(L"fail: no text in clipboard"); return false; }
        } else {
            HBITMAP bmp = GetClipboardBitmap();
            if (!bmp) { LogLine(L"fail: no image in clipboard"); return false; }
            bool ok = BitmapToBase64Png(bmp, imageB64); DeleteObject(bmp);
            if (!ok) { LogLine(L"fail: base64 encode image failed"); return false; }
        }
        wstring systemPrompt = [&]() -> auto {
            wstring ithing = f.input == IOType::Text ? L"text" : L"image";
            wstring othing = f.output == IOType::Text ? L"text" : L"image";
            return format(
                L"Follow the instructions strictly and convert the input {0} to the output {1}. "
                L"No additional text or comments are allowed.",
                ithing, othing);
        }();
        wstring promptText = f.prompt + L"\n\n" + textInput;
        wstring imageDataUrl = imageB64.empty() ? L"" : (L"data:image/png;base64," + imageB64);
        ApiCallResult res = CallTemplate(*tpl, m, systemPrompt, promptText, imageB64, imageDataUrl);
        if (tpl->output == IOType::Text) {
            if (res.text.empty()) { LogLine(L"fail: template returned empty text"); return false; }
            SetClipboardText(res.text);
            return true;
        } else {
            if (!res.image) { LogLine(L"fail: template returned no image"); return false; }
            try {
                SetClipboardBitmap(res.image);
            } catch (...) {
                DeleteObject(res.image);
                LogLine(L"fail: SetClipboardBitmap threw"); return false;
            }
            return true;
        }
    } catch (const exception& ex) {
        wstring wmsg;
        int len = MultiByteToWideChar(CP_UTF8, 0, ex.what(), -1, nullptr, 0);
        if (len > 1) {
            wmsg.resize(len - 1); MultiByteToWideChar(CP_UTF8, 0, ex.what(), -1, wmsg.data(), len - 1);
        }
        LogLine(L"exception in RunFilter: " + wmsg);
    } catch (...) {
        LogLine(L"unknown exception in RunFilter");
    }
    return false; // fallback
}

/**
 * @struct ModelDialogState
 * @brief State for model configuration dialog
 */
struct ModelDialogState { ModelConfig* model{}; size_t index{}; int result{}; ModelConfig original{}; HWND hName{}, hServer{}, hModel{}, hKey{}, hProvider{}; };

/**
 * @struct EditDialogState
 * @brief State for filter edit dialog
 */
struct EditDialogState { FilterDefinition* filter{}; FilterDefinition original{}; HWND hName{}, hIn{}, hOut{}, hModel{}, hPrompt{}; };

/**
 * @struct HotkeyInputState
 * @brief State for hotkey input dialog
 */
struct HotkeyInputState {
    int result{};
    UINT vkCode{};
    UINT modifiers{};
    wstring keyName;
    HWND hLabel{};
};

/**
 * @brief Convert virtual key code to display string
 * @param vk Virtual key code
 * @param modifiers Modifier flags (MOD_SHIFT, MOD_CONTROL, etc.)
 * @return String representation of the key combination
 */
wstring VKCodeToString(UINT vk, UINT modifiers) {
    wstring result;
    if (modifiers & MOD_SHIFT) result += L"Shift + ";
    if (modifiers & MOD_CONTROL) result += L"Ctrl + ";
    if (modifiers & MOD_ALT) result += L"Alt + ";
    if (modifiers & MOD_WIN) result += L"Win + ";

    if (vk >= 'A' && vk <= 'Z') {
        result += static_cast<wchar_t>(vk);
    } else if (vk >= '0' && vk <= '9') {
        result += static_cast<wchar_t>(vk);
    } else if (vk >= VK_F1 && vk <= VK_F12) {
        result += L"F" + to_wstring(vk - VK_F1 + 1);
    } else if (vk >= VK_NUMPAD0 && vk <= VK_NUMPAD9) {
        result += L"NumPad" + to_wstring(vk - VK_NUMPAD0);
    } else {
        switch (vk) {
        case VK_RETURN: result += L"Enter"; break;
        case VK_ESCAPE: result += L"Escape"; break;
        case VK_TAB: result += L"Tab"; break;
        case VK_SPACE: result += L"Space"; break;
        case VK_BACK: result += L"Backspace"; break;
        case VK_DELETE: result += L"Delete"; break;
        case VK_INSERT: result += L"Insert"; break;
        case VK_HOME: result += L"Home"; break;
        case VK_END: result += L"End"; break;
        case VK_PRIOR: result += L"PageUp"; break;
        case VK_NEXT: result += L"PageDown"; break;
        case VK_UP: result += L"Up"; break;
        case VK_DOWN: result += L"Down"; break;
        case VK_LEFT: result += L"Left"; break;
        case VK_RIGHT: result += L"Right"; break;
        case VK_SNAPSHOT: result += L"PrintScreen"; break;
        case VK_PAUSE: result += L"Pause"; break;
        case VK_ADD: result += L"NumPad+"; break;
        case VK_SUBTRACT: result += L"NumPad-"; break;
        case VK_MULTIPLY: result += L"NumPad*"; break;
        case VK_DIVIDE: result += L"NumPad/"; break;
        case VK_LBUTTON: result += L"LeftButton"; break;
        case VK_RBUTTON: result += L"RightButton"; break;
        case VK_MBUTTON: result += L"MiddleButton"; break;
        default: result += L"Key" + to_wstring(vk); break;
        }
    }
    return result;
}

/**
 * @brief Check if a virtual key code is valid for hotkey registration
 * @param vk Virtual key code to check
 * @return true if the key is valid for hotkey registration
 */
bool IsValidHotkeyVK(UINT vk) {
    // Mouse buttons cannot be used with RegisterHotKey
    if (vk == VK_LBUTTON || vk == VK_RBUTTON || vk == VK_MBUTTON) return false;

    // Modifier keys alone cannot be used
    if (vk == VK_SHIFT || vk == VK_CONTROL || vk == VK_MENU || vk == VK_LWIN || vk == VK_RWIN) return false;

    // Valid key ranges
    if ((vk >= 'A' && vk <= 'Z') || (vk >= '0' && vk <= '9')) return true;
    if (vk >= VK_F1 && vk <= VK_F24) return true;
    if (vk >= VK_NUMPAD0 && vk <= VK_DIVIDE) return true;

    // Other valid keys
    switch (vk) {
    case VK_RETURN:
    case VK_ESCAPE:
    case VK_TAB:
    case VK_SPACE:
    case VK_BACK:
    case VK_DELETE:
    case VK_INSERT:
    case VK_HOME:
    case VK_END:
    case VK_PRIOR:
    case VK_NEXT:
    case VK_UP:
    case VK_DOWN:
    case VK_LEFT:
    case VK_RIGHT:
    case VK_SNAPSHOT:
    case VK_PAUSE:
        return true;
    }
    return false;
}

/**
 * @struct SetupDialogState
 * @brief State for initial API setup dialog
 */
struct SetupDialogState {
    int result{};
    wstring languageCode;
    bool shift{}, ctrl{}, alt{}, win{};
    UINT vkCode{ 'V' };
    size_t providerIndex{};
    wstring serverUrl;
    wstring apiKey;
    HWND hLang{}, hShift{}, hCtrl{}, hAlt{}, hWin{}, hKeyLabel{}, hKeyButton{}, hProvider{}, hServer{}, hApiKey{};
};

/**
 * @struct SettingsState
 * @brief State for settings window
 */
struct SettingsState {
    HWND hList{};
    HWND hHotkeyLabel{};
    HWND hHotkeyButton{};
};

/**
 * @brief Populate combo box with API templates
 */
void PopulateProviderCombo(HWND combo, const wstring& currentId) {
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    int sel = -1;
    for (size_t i = 0; i < g_providers.size(); ++i) {
        const auto& p = g_providers[i];
        int idx = static_cast<int>(SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(p.id.c_str())));
        SendMessageW(combo, CB_SETITEMDATA, idx, static_cast<LPARAM>(i));
        if (p.id == currentId) sel = idx;
    }
    if (SendMessageW(combo, CB_GETCOUNT, 0, 0) == 0) {
        wstring placeholder = L"(no providers)";
        SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(placeholder.c_str()));
        SendMessageW(combo, CB_SETCURSEL, 0, 0);
    } else {
        if (sel < 0) sel = 0;
        SendMessageW(combo, CB_SETCURSEL, sel, 0);
    }
}

/**
 * @brief Window procedure for model configuration dialog
 */
LRESULT CALLBACK ModelDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<ModelDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = reinterpret_cast<ModelDialogState*>(reinterpret_cast<LPCREATESTRUCT>(lParam)->lpCreateParams);
        st->original = *st->model;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));
        const int m = 10, lw = 110; int y = m;
        wstring strName = GetString(L"name");
        CreateWindowW(L"STATIC", strName.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 20, hwnd, nullptr, nullptr, nullptr);
        st->hName = CreateWindowW(L"EDIT", st->model->name.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                  m + lw + 6, y - 2, 360, 22, hwnd, (HMENU)(INT_PTR)200, nullptr, nullptr);
        EnableCtrlA(st->hName);
        y += 28;
        wstring strServerUrl = GetString(L"server_url");
        CreateWindowW(L"STATIC", strServerUrl.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 20, hwnd, nullptr, nullptr, nullptr);
        st->hServer = CreateWindowW(L"EDIT", st->model->serverUrl.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                    m + lw + 6, y - 2, 360, 22, hwnd, (HMENU)(INT_PTR)201, nullptr, nullptr);
        EnableCtrlA(st->hServer);
        y += 28;
        wstring strModelName = GetString(L"model_name");
        CreateWindowW(L"STATIC", strModelName.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 20, hwnd, nullptr, nullptr, nullptr);
        st->hModel = CreateWindowW(L"EDIT", st->model->modelName.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                   m + lw + 6, y - 2, 360, 22, hwnd, (HMENU)(INT_PTR)202, nullptr, nullptr);
        EnableCtrlA(st->hModel);
        y += 28;
        wstring strProvider = GetString(L"provider");
        CreateWindowW(L"STATIC", strProvider.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 20, hwnd, nullptr, nullptr, nullptr);
        st->hProvider = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                      m + lw + 6, y - 2, 360, 200, hwnd, (HMENU)(INT_PTR)207, nullptr, nullptr);
        PopulateProviderCombo(st->hProvider, st->model->providerId);
        y += 28;
        wstring strApiKey = GetString(L"api_key");
        CreateWindowW(L"STATIC", strApiKey.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 20, hwnd, nullptr, nullptr, nullptr);
        st->hKey = CreateWindowW(L"EDIT", st->model->apiKey.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_PASSWORD | WS_TABSTOP,
                                 m + lw + 6, y - 2, 360, 22, hwnd, (HMENU)(INT_PTR)203, nullptr, nullptr);
        EnableCtrlA(st->hKey);
        y += 36;
        wstring strSave = GetString(L"save");
        wstring strDelete = GetString(L"delete");
        wstring strClose = GetString(L"close");
        CreateWindowW(L"BUTTON", strSave.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, m + 20, y, 90, 26, hwnd,
                      (HMENU)(INT_PTR)204, nullptr, nullptr);
        CreateWindowW(L"BUTTON", strDelete.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, m + 120, y, 90, 26, hwnd,
                      (HMENU)(INT_PTR)205, nullptr, nullptr);
        CreateWindowW(L"BUTTON", strClose.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, m + 220, y, 90, 26, hwnd,
                      (HMENU)(INT_PTR)206, nullptr, nullptr);
        SetFocus(st->hName);
        return 0;
    }
    case WM_COMMAND: {
        WORD id = LOWORD(wParam);
        if (id == 204) {
            wchar_t buf[512];
            GetWindowTextW(st->hName, buf, 512); st->model->name = buf;
            GetWindowTextW(st->hServer, buf, 512); st->model->serverUrl = buf;
            GetWindowTextW(st->hModel, buf, 512); st->model->modelName = buf;
            int tsel = static_cast<int>(SendMessageW(st->hProvider, CB_GETCURSEL, 0, 0));
            if (tsel >= 0) {
                size_t provIdx = static_cast<size_t>(SendMessageW(st->hProvider, CB_GETITEMDATA, tsel, 0));
                if (provIdx < g_providers.size()) st->model->providerId = g_providers[provIdx].id;
            }
            GetWindowTextW(st->hKey, buf, 512); st->model->apiKey = buf;
            st->result = 1; DestroyWindow(hwnd); return 0;
        }
        if (id == 205) { st->result = 2; DestroyWindow(hwnd); return 0; }
        if (id == 206) {
            ModelConfig cur = *st->model; wchar_t buf[512];
            GetWindowTextW(st->hName, buf, 512); cur.name = buf;
            GetWindowTextW(st->hServer, buf, 512); cur.serverUrl = buf;
            GetWindowTextW(st->hModel, buf, 512); cur.modelName = buf;
            int tsel = static_cast<int>(SendMessageW(st->hProvider, CB_GETCURSEL, 0, 0));
            if (tsel >= 0) {
                size_t provIdx = static_cast<size_t>(SendMessageW(st->hProvider, CB_GETITEMDATA, tsel, 0));
                if (provIdx < g_providers.size()) cur.providerId = g_providers[provIdx].id;
            }
            GetWindowTextW(st->hKey, buf, 512); cur.apiKey = buf;
            bool dirty = (cur.name != st->original.name) || (cur.serverUrl != st->original.serverUrl) || (cur.modelName != st->original.modelName) || (cur.apiKey != st->original.apiKey) || (cur.providerId != st->original.providerId);
            if (dirty) {
                wstring strUnsaved = GetString(L"unsaved_changes");
                wstring strConfirm = GetString(L"confirm");
                int r = MessageBoxW(hwnd, strUnsaved.c_str(), strConfirm.c_str(), MB_YESNOCANCEL | MB_ICONQUESTION);
                if (r == IDYES) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(204, BN_CLICKED), 0); return 0; }
                if (r == IDCANCEL) return 0;
            }
            DestroyWindow(hwnd); return 0;
        }
        return 0;
    }
    case WM_KEYDOWN:
        if (wParam == VK_RETURN) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(204, BN_CLICKED), 0); return 0; }
        if (wParam == VK_ESCAPE) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(206, BN_CLICKED), 0); return 0; }
        break;
    case WM_CLOSE:
        PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(206, BN_CLICKED), 0); return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Check if a text matches a regex pattern (case-insensitive)
 * @param text Text to check
 * @param pattern Regex pattern
 * @return true if text matches pattern, false otherwise
 */
bool RegexMatchNoCase(const wstring& text, const wstring& pattern) {
    try {
        std::wregex re(pattern, std::regex_constants::icase);
        return std::regex_search(text, re);
    } catch (...) {
        return false;
    }
}

/**
 * @brief Fetch models from API
 * @param provider API provider
 * @param serverUrl Server URL
 * @param apiKey API key
 * @param models Vector to store model names
 * @param err Error message
 * @return true if models were fetched successfully, false otherwise
 */
bool FetchModels(const ApiProvider& provider, const wstring& serverUrl, const wstring& apiKey, vector<wstring>& models, wstring& err) {
    if (provider.modelsEndpoint.empty()) { err = L"models endpoint not defined"; return false; }
    ModelConfig dummy{ L"", serverUrl, L"", apiKey, provider.id };
    wstring endpoint = ReplacePlaceholders(provider.modelsEndpoint, dummy, L"", L"", L"", L"", false);
    wstring host, path; bool useHttps = true;
    if (!PrepareEndpoint(serverUrl, endpoint, host, path, useHttps)) { err = L"PrepareEndpoint failed"; return false; }
    wstring headers = BuildHeaderString(provider.modelsHeaders, dummy);
    string body = ToUtf8(ReplacePlaceholders(provider.modelsPayload, dummy, L"", L"", L"", L"", false));
    wstring resp;
    if (!provider.modelsMethod.empty() && RegexMatchNoCase(provider.modelsMethod, L"post")) {
        resp = HttpRequestWithHeaders(host, path, useHttps, headers, body, L"POST", &err);
    } else {
        static const string emptyBody;
        resp = HttpRequestWithHeaders(host, path, useHttps, headers, emptyBody, L"GET", &err);
    }
    if (resp.empty()) return false;
    try {
        using namespace winrt::Windows::Data::Json;
        IJsonValue root = JsonValue::Parse(resp);
        auto splitPath = [&](const wstring& s) {
            vector<wstring> parts; size_t start = 0;
            for (size_t i = 0; i <= s.size(); ++i) if (i == s.size() || s[i] == L'.') { parts.push_back(s.substr(start, i - start)); start = i + 1; }
            return parts;
        };
        auto parts = splitPath(provider.modelsResultPath);
        IJsonValue cur = root;
        for (const auto& part : parts) {
            if (part.empty()) continue;
            if (cur.ValueType() == JsonValueType::Object) {
                JsonObject obj = cur.GetObject();
                if (!obj.HasKey(part)) { err = L"models result path missing"; return false; }
                cur = obj.GetNamedValue(part);
            } else {
                err = L"models result path invalid"; return false;
            }
        }
        JsonArray arr;
        if (cur.ValueType() == JsonValueType::Array) arr = cur.GetArray();
        else return false;
        for (auto const& item : arr) {
            if (item.ValueType() == JsonValueType::Object) {
                JsonObject obj = item.GetObject();
                if (obj.HasKey(L"id")) models.push_back(wstring(obj.GetNamedString(L"id").c_str()));
            } else if (item.ValueType() == JsonValueType::String) {
                models.push_back(wstring(item.GetString().c_str()));
            }
        }
    } catch (const winrt::hresult_error& e) {
        err = e.message().c_str();
        return false;
    }
    return !models.empty();
}

/**
 * @brief Pick a model by patterns
 * @param models Vector of model names
 * @param patterns Vector of patterns
 * @return Picked model name, or empty string if no model was found
 */
wstring PickModelByPatterns(const vector<wstring>& models, const vector<wstring>& patterns) {
    for (const auto& pat : patterns) {
        for (const auto& m : models) if (RegexMatchNoCase(m, pat)) return m;
    }
    return models.empty() ? L"" : models.front();
}

/**
 * @brief Perform initial setup
 * @param st Setup dialog state
 * @param err Error message
 * @return true if setup was successful, false otherwise
 */
bool PerformInitialSetup(const SetupDialogState& st, wstring& err) {
    using namespace winrt::Windows::Data::Json;
    if (g_providers.empty()) { err = L"No providers"; return false; }
    if (st.providerIndex >= g_providers.size()) { err = L"Invalid provider selection"; return false; }
    const ApiProvider& provider = g_providers[st.providerIndex];
    vector<wstring> modelList;
    if (!FetchModels(provider, st.serverUrl, st.apiKey, modelList, err)) return false;
    if (modelList.empty()) { err = L"No models"; return false; }
    vector<wstring> patternsLLM{ L"gpt-.*-nano", L"gemini-.*-flash-lite", L"gpt-.*-mini", L"gemini-.*-flash", L"gpt-.*", L"claude-.*-haiku", L"gemini-.*-pro", L"claude-.*-sonnet" };
    vector<wstring> patternsImage{ L"gpt.*image.*mini", L"gemini.*image", L"gpt.*image" };
    wstring tt = PickModelByPatterns(modelList, patternsLLM);
    wstring it = PickModelByPatterns(modelList, patternsLLM);
    wstring ti = PickModelByPatterns(modelList, patternsImage);
    wstring ii = PickModelByPatterns(modelList, patternsImage);
    if (tt.empty()) tt = modelList.front();
    if (it.empty()) it = modelList.front();
    if (ti.empty()) ti = modelList.front();
    if (ii.empty()) ii = modelList.front();

    g_models.clear();
    auto addModel = [&](const wstring& name, const wstring& modelName) {
        g_models.push_back({ name, st.serverUrl, modelName, st.apiKey, provider.id });
    };
    addModel(L"Text/Text", tt);
    addModel(L"Text/Image", ti);
    addModel(L"Image/Text", it);
    addModel(L"Image/Image", ii);

    JsonObject def = CreateDefaultConfig();
    g_language = st.languageCode.empty() ? g_language : st.languageCode;
    if (def.HasKey(L"language")) g_language = wstring(def.GetNamedString(L"language", g_language.c_str()).c_str());
    LoadFiltersFromJson(def, g_filters);
    if (g_filters.empty()) {
        g_filters.push_back({ L"Translate", IOType::Text, IOType::Text, 0, L"Translate into English." });
    }
    auto idxFor = [&](IOType in, IOType out) -> size_t {
        if (in == IOType::Text && out == IOType::Text) return 0;
        if (in == IOType::Text && out == IOType::Image) return 1;
        if (in == IOType::Image && out == IOType::Text) return 2;
        return 3;
    };
    for (auto& f : g_filters) {
        size_t idx = idxFor(f.input, f.output);
        if (idx >= g_models.size()) idx = 0;
        f.modelIndex = idx;
    }
    g_hotkeyModifiers = 0;
    if (st.shift) g_hotkeyModifiers |= MOD_SHIFT;
    if (st.ctrl) g_hotkeyModifiers |= MOD_CONTROL;
    if (st.alt) g_hotkeyModifiers |= MOD_ALT;
    if (st.win) g_hotkeyModifiers |= MOD_WIN;
    g_hotkeyKey = st.vkCode;
    SaveConfig();
    return true;
}

/**
 * @brief Collect setup from UI
 * @param st Setup dialog state
 */
void CollectSetupFromUI(SetupDialogState* st) {
    int selLang = static_cast<int>(SendMessageW(st->hLang, CB_GETCURSEL, 0, 0));
    if (selLang >= 0) {
        size_t idx = static_cast<size_t>(SendMessageW(st->hLang, CB_GETITEMDATA, selLang, 0));
        const auto& langs = SupportedLanguages();
        if (idx < langs.size()) st->languageCode = langs[idx].first;
    }
    st->shift = Button_GetCheck(st->hShift) == BST_CHECKED;
    st->ctrl = Button_GetCheck(st->hCtrl) == BST_CHECKED;
    st->alt = Button_GetCheck(st->hAlt) == BST_CHECKED;
    st->win = Button_GetCheck(st->hWin) == BST_CHECKED;
    // vkCode is now set via the key input dialog, not from a text field
    int psel = static_cast<int>(SendMessageW(st->hProvider, CB_GETCURSEL, 0, 0));
    if (psel >= 0) st->providerIndex = static_cast<size_t>(SendMessageW(st->hProvider, CB_GETITEMDATA, psel, 0));
    wchar_t buf[512];
    GetWindowTextW(st->hServer, buf, 512); st->serverUrl = buf;
    GetWindowTextW(st->hApiKey, buf, 512); st->apiKey = buf;
}

/**
 * @brief Window procedure for hotkey input dialog
 * @param hwnd Window handle
 * @param msg Message
 * @param wParam Message parameter
 * @param lParam Message parameter
 * @return Result of message processing
 */
LRESULT CALLBACK HotkeyInputDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<HotkeyInputState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = reinterpret_cast<HotkeyInputState*>(reinterpret_cast<LPCREATESTRUCT>(lParam)->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));

        int m = 20, y = 20;
        HWND hInfo = CreateWindowW(L"STATIC", L"Press any key combination...\n(Escape to cancel)",
            WS_CHILD | WS_VISIBLE | SS_CENTER, m, y, 360, 40, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hInfo);
        y += 50;

        st->hLabel = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_CENTER,
            m, y, 360, 30, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(st->hLabel);
        HFONT hBigFont = CreateFontW(24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        SendMessageW(st->hLabel, WM_SETFONT, reinterpret_cast<WPARAM>(hBigFont), TRUE);
        return 0;
    }
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN: {
        UINT vk = static_cast<UINT>(wParam);
        if (vk == VK_ESCAPE) {
            st->result = 0;
            DestroyWindow(hwnd);
            return 0;
        }

        if (!IsValidHotkeyVK(vk)) return 0;

        UINT modifiers = 0;
        if (GetKeyState(VK_SHIFT) & 0x8000) modifiers |= MOD_SHIFT;
        if (GetKeyState(VK_CONTROL) & 0x8000) modifiers |= MOD_CONTROL;
        if (GetKeyState(VK_MENU) & 0x8000) modifiers |= MOD_ALT;
        if ((GetKeyState(VK_LWIN) & 0x8000) || (GetKeyState(VK_RWIN) & 0x8000)) modifiers |= MOD_WIN;

        st->vkCode = vk;
        st->modifiers = modifiers;
        st->keyName = VKCodeToString(vk, modifiers);
        SetWindowTextW(st->hLabel, st->keyName.c_str());

        // Auto-close after 500ms delay to show the detected key
        SetTimer(hwnd, 1, 500, nullptr);
        return 0;
    }
    case WM_TIMER:
        if (wParam == 1) {
            KillTimer(hwnd, 1);
            st->result = 1;
            DestroyWindow(hwnd);
        }
        return 0;
    case WM_CLOSE:
        st->result = 0;
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Show hotkey input dialog
 * @param parent Parent window handle
 * @return 1 if key was captured, 0 if cancelled
 */
int ShowHotkeyInputDialog(HWND parent, UINT& vkCode, UINT& modifiers) {
    HotkeyInputState st{};
    st.vkCode = vkCode;
    st.modifiers = modifiers;

    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kHotkeyInputClass,
        L"Press Key", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 420, 180, parent, nullptr, g_hInst, &st);

    ShowWindow(dlg, SW_SHOWNORMAL);
    SetForegroundWindow(dlg);
    SetFocus(dlg);
    EnableWindow(parent, FALSE);

    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        if (msg.message == WM_KEYDOWN && msg.hwnd == dlg) {
            SendMessageW(dlg, WM_KEYDOWN, msg.wParam, msg.lParam);
            continue;
        }
        if (msg.message == WM_SYSKEYDOWN && msg.hwnd == dlg) {
            SendMessageW(dlg, WM_SYSKEYDOWN, msg.wParam, msg.lParam);
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    EnableWindow(parent, TRUE);
    SetForegroundWindow(parent);

    if (st.result == 1) {
        vkCode = st.vkCode;
        modifiers = st.modifiers;
    }
    return st.result;
}

/**
 * @brief Window procedure for setup dialog
 * @param hwnd Window handle
 * @param msg Message
 * @param wParam Message parameter
 * @param lParam Message parameter
 * @return Result of message processing
 */
LRESULT CALLBACK SetupDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<SetupDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = reinterpret_cast<SetupDialogState*>(reinterpret_cast<LPCREATESTRUCT>(lParam)->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));
        int m = 12, y = m, lw = 150, cw = 320;
        wstring strLang = GetString(L"language");
        HWND hLangLbl = CreateWindowW(L"STATIC", strLang.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hLangLbl);
        st->hLang = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, m + lw + 6, y - 2, cw, 200, hwnd, (HMENU)(INT_PTR)300, nullptr, nullptr);
        SetUIFont(st->hLang);
        const auto& langs = SupportedLanguages();
        int langSel = 0;
        for (size_t i = 0; i < langs.size(); ++i) {
            int idx = static_cast<int>(SendMessageW(st->hLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(langs[i].second.c_str())));
            SendMessageW(st->hLang, CB_SETITEMDATA, idx, static_cast<LPARAM>(i));
            if (langs[i].first == st->languageCode) langSel = idx;
        }
        SendMessageW(st->hLang, CB_SETCURSEL, langSel, 0);
        y += 32;

        wstring strHotkey = GetString(L"hotkey");
        HWND hHotkeyLbl = CreateWindowW(L"STATIC", strHotkey.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hHotkeyLbl);
        st->hShift = CreateWindowW(L"BUTTON", GetString(L"hotkey_shift").c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, m + lw + 6, y, 70, 22, hwnd, (HMENU)(INT_PTR)301, nullptr, nullptr);
        st->hCtrl  = CreateWindowW(L"BUTTON", GetString(L"hotkey_ctrl").c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, m + lw + 80, y, 60, 22, hwnd, (HMENU)(INT_PTR)302, nullptr, nullptr);
        st->hAlt   = CreateWindowW(L"BUTTON", GetString(L"hotkey_alt").c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, m + lw + 146, y, 60, 22, hwnd, (HMENU)(INT_PTR)303, nullptr, nullptr);
        st->hWin   = CreateWindowW(L"BUTTON", GetString(L"hotkey_win").c_str(), WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, m + lw + 212, y, 60, 22, hwnd, (HMENU)(INT_PTR)304, nullptr, nullptr);
        SetUIFont(st->hShift); SetUIFont(st->hCtrl); SetUIFont(st->hAlt); SetUIFont(st->hWin);
        CheckDlgButton(hwnd, 301, st->shift ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, 302, st->ctrl ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, 303, st->alt ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, 304, st->win ? BST_CHECKED : BST_UNCHECKED);
        y += 28;

        UINT mods = 0;
        if (st->shift) mods |= MOD_SHIFT;
        if (st->ctrl) mods |= MOD_CONTROL;
        if (st->alt) mods |= MOD_ALT;
        if (st->win) mods |= MOD_WIN;
        wstring currentKey = VKCodeToString(st->vkCode, mods);

        st->hKeyLabel = CreateWindowW(L"STATIC", currentKey.c_str(), WS_CHILD | WS_VISIBLE | SS_CENTER | WS_BORDER,
            m + lw + 6, y, 160, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(st->hKeyLabel);

        st->hKeyButton = CreateWindowW(L"BUTTON", L"Change Key...", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            m + lw + 172, y, 100, 22, hwnd, (HMENU)(INT_PTR)305, nullptr, nullptr);
        SetUIFont(st->hKeyButton);
        y += 32;

        wstring strProvider = GetString(L"provider");
        HWND hProvLbl = CreateWindowW(L"STATIC", strProvider.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hProvLbl);
        st->hProvider = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, m + lw + 6, y - 2, cw, 200, hwnd, (HMENU)(INT_PTR)306, nullptr, nullptr);
        SetUIFont(st->hProvider);
        int psel = 0;
        for (size_t i = 0; i < g_providers.size(); ++i) {
            int idx = static_cast<int>(SendMessageW(st->hProvider, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(g_providers[i].id.c_str())));
            SendMessageW(st->hProvider, CB_SETITEMDATA, idx, static_cast<LPARAM>(i));
            if (i == st->providerIndex) psel = idx;
        }
        SendMessageW(st->hProvider, CB_SETCURSEL, psel, 0);
        if (st->providerIndex < g_providers.size() && !g_providers[st->providerIndex].defaultEndpoint.empty()) {
            st->serverUrl = g_providers[st->providerIndex].defaultEndpoint;
        }
        y += 32;

        wstring strServerUrl = GetString(L"server_url");
        HWND hSrvLbl = CreateWindowW(L"STATIC", strServerUrl.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hSrvLbl);
        st->hServer = CreateWindowW(L"EDIT", st->serverUrl.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP, m + lw + 6, y - 2, cw, 22, hwnd, (HMENU)(INT_PTR)307, nullptr, nullptr);
        SetUIFont(st->hServer);
        EnableCtrlA(st->hServer);
        y += 32;

        wstring strApiKey = GetString(L"api_key");
        HWND hApiLbl = CreateWindowW(L"STATIC", strApiKey.c_str(), WS_CHILD | WS_VISIBLE, m, y, lw, 22, hwnd, nullptr, nullptr, nullptr);
        SetUIFont(hApiLbl);
        st->hApiKey = CreateWindowW(L"EDIT", st->apiKey.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP, m + lw + 6, y - 2, cw, 22, hwnd, (HMENU)(INT_PTR)308, nullptr, nullptr);
        SetUIFont(st->hApiKey);
        EnableCtrlA(st->hApiKey);
        y += 44;

        wstring strCheck = GetString(L"check_connection");
        wstring strExit = GetString(L"exit_app");
        HWND hCheck = CreateWindowW(L"BUTTON", strCheck.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, m, y, 160, 30, hwnd, (HMENU)(INT_PTR)309, nullptr, nullptr);
        HWND hExit = CreateWindowW(L"BUTTON", strExit.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, m + 180, y, 160, 30, hwnd, (HMENU)(INT_PTR)310, nullptr, nullptr);
        SetUIFont(hCheck); SetUIFont(hExit);
        return 0;
    }
    case WM_COMMAND: {
        if (!st) {
            LogLine(L"SetupDlgProc WM_COMMAND: st is null");
            return DefWindowProcW(hwnd, msg, wParam, lParam);
        }
        WORD id = LOWORD(wParam);
        LogLine(L"SetupDlgProc WM_COMMAND: id=" + to_wstring(id) + L" code=" + to_wstring(HIWORD(wParam)));
        if (id == 309) { // Save
            CollectSetupFromUI(st);
            wstring err;
            if (st->providerIndex >= g_providers.size()) err = GetString(L"provider");
            if (st->serverUrl.empty()) err = GetString(L"server_url");
            if (!err.empty()) {
                MessageBoxW(hwnd, err.c_str(), L"cbfilter", MB_OK | MB_ICONWARNING);
                return 0;
            }

            // Test hotkey registration before saving
            UINT testMods = 0;
            if (st->shift) testMods |= MOD_SHIFT;
            if (st->ctrl) testMods |= MOD_CONTROL;
            if (st->alt) testMods |= MOD_ALT;
            if (st->win) testMods |= MOD_WIN;

            if (!RegisterHotKey(hwnd, HOTKEY_ID + 1, testMods | MOD_NOREPEAT, st->vkCode)) {
                wstring keyStr = VKCodeToString(st->vkCode, testMods);
                wstring errMsg = L"Cannot register hotkey: " + keyStr + L"\n"
                               + L"The hotkey may already be in use.\n"
                               + L"Please choose a different key combination.";
                MessageBoxW(hwnd, errMsg.c_str(), L"cbfilter", MB_OK | MB_ICONWARNING);
                return 0;
            }
            UnregisterHotKey(hwnd, HOTKEY_ID + 1);  // Clean up test registration

            if (!PerformInitialSetup(*st, err)) {
                wstring errMsg = GetString(L"connection_failed");
                errMsg += L"\n";
                errMsg += err;
                MessageBoxW(hwnd, errMsg.c_str(), L"cbfilter", MB_OK | MB_ICONERROR);
                return 0;
            }
            wstring okMsg = GetString(L"connection_success");
            MessageBoxW(hwnd, okMsg.c_str(), L"cbfilter", MB_OK | MB_ICONINFORMATION);
            st->result = 1; DestroyWindow(hwnd); return 0;
        }
        if (id == 306 && HIWORD(wParam) == CBN_SELCHANGE) {
            LogLine(L"SetupDlgProc: CBN_SELCHANGE for provider combo (id=306)");
            if (!st->hProvider || !st->hServer) {
                LogLine(L"SetupDlgProc: hProvider or hServer is null");
                return 0;
            }
            int psel = static_cast<int>(SendMessageW(st->hProvider, CB_GETCURSEL, 0, 0));
            LogLine(L"SetupDlgProc: provider selection index=" + to_wstring(psel));
            if (psel >= 0) {
                size_t provIdx = static_cast<size_t>(SendMessageW(st->hProvider, CB_GETITEMDATA, psel, 0));
                LogLine(L"SetupDlgProc: provider index=" + to_wstring(provIdx) + L" total providers=" + to_wstring(g_providers.size()));
                if (provIdx < g_providers.size()) {
                    const auto& prov = g_providers[provIdx];
                    LogLine(L"SetupDlgProc: provider id=" + prov.id + L" defaultEndpoint=" + prov.defaultEndpoint);
                    if (!prov.defaultEndpoint.empty()) {
                        SetWindowTextW(st->hServer, prov.defaultEndpoint.c_str());
                        st->serverUrl = prov.defaultEndpoint;
                        LogLine(L"SetupDlgProc: Updated server URL to " + prov.defaultEndpoint);
                    } else {
                        LogLine(L"SetupDlgProc: provider has empty defaultEndpoint");
                    }
                } else {
                    LogLine(L"SetupDlgProc: provider index out of range");
                }
            } else {
                LogLine(L"SetupDlgProc: invalid selection index");
            }
            return 0;
        }
        if (id == 305) { // Change Key button
            UINT vk = st->vkCode;
            UINT mods = 0;
            if (Button_GetCheck(st->hShift) == BST_CHECKED) mods |= MOD_SHIFT;
            if (Button_GetCheck(st->hCtrl) == BST_CHECKED) mods |= MOD_CONTROL;
            if (Button_GetCheck(st->hAlt) == BST_CHECKED) mods |= MOD_ALT;
            if (Button_GetCheck(st->hWin) == BST_CHECKED) mods |= MOD_WIN;

            if (ShowHotkeyInputDialog(hwnd, vk, mods) == 1) {
                st->vkCode = vk;
                // Update modifier checkboxes
                CheckDlgButton(hwnd, 301, (mods & MOD_SHIFT) ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(hwnd, 302, (mods & MOD_CONTROL) ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(hwnd, 303, (mods & MOD_ALT) ? BST_CHECKED : BST_UNCHECKED);
                CheckDlgButton(hwnd, 304, (mods & MOD_WIN) ? BST_CHECKED : BST_UNCHECKED);
                st->shift = (mods & MOD_SHIFT) != 0;
                st->ctrl = (mods & MOD_CONTROL) != 0;
                st->alt = (mods & MOD_ALT) != 0;
                st->win = (mods & MOD_WIN) != 0;
                // Update key label
                wstring keyStr = VKCodeToString(vk, mods);
                SetWindowTextW(st->hKeyLabel, keyStr.c_str());
            }
            return 0;
        }
        if (id == 310) { st->result = 2; DestroyWindow(hwnd); return 0; }
        return 0;
    }
    case WM_CLOSE:
        st->result = 2; DestroyWindow(hwnd); return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Show modal dialog for editing model configuration
 * @param parent Parent window handle
 * @param model Model configuration to edit (modified in place)
 * @param index Index of model in g_models vector
 * @return 1 if saved, 2 if deleted, 0 if cancelled
 */
int ShowModelDialog(HWND parent, ModelConfig& model, size_t index) {
    ModelDialogState st{ &model, index, 0, model };
    wstring strModelSettings = GetString(L"model_settings");
    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kModelClass, strModelSettings.c_str(),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 520, 260, parent, nullptr, g_hInst, &st);
    g_modelWnd = dlg;
    ShowWindow(dlg, SW_SHOWNORMAL);
    EnableWindow(parent, FALSE);
    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        // Handle Escape key explicitly before IsDialogMessageW
        // Check if the message is for the dialog or any of its child controls
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE && 
            (msg.hwnd == dlg || IsChild(dlg, msg.hwnd))) {
            PostMessageW(dlg, WM_CLOSE, 0, 0);
            continue;
        }
        if (!IsDialogMessageW(dlg, &msg)) { TranslateMessage(&msg); DispatchMessageW(&msg); }
    }
    int res = st.result;
    EnableWindow(parent, TRUE);
    SetForegroundWindow(parent);
    g_modelWnd = nullptr;
    return res;
}

/**
 * @brief Show setup dialog
 * @return Result of setup dialog
 */
int ShowSetupDialog() {
    SetupDialogState st{};
    st.languageCode = g_language;
    st.shift = (g_hotkeyModifiers & MOD_SHIFT) != 0;
    st.ctrl = (g_hotkeyModifiers & MOD_CONTROL) != 0;
    st.alt = (g_hotkeyModifiers & MOD_ALT) != 0;
    st.win = (g_hotkeyModifiers & MOD_WIN) != 0;
    st.vkCode = g_hotkeyKey;
    st.providerIndex = 0;
    if (!g_models.empty()) st.serverUrl = g_models.front().serverUrl;
    st.apiKey = L"";
    wstring title = GetString(L"initial_setup");
    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kSetupClass, title.c_str(),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 520, 330, nullptr, nullptr, g_hInst, &st);
    ShowWindow(dlg, SW_SHOWNORMAL);
    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        // Handle WM_COMMAND messages (like CBN_SELCHANGE) before IsDialogMessageW
        // to ensure they reach the window procedure
        if (msg.message == WM_COMMAND && (msg.hwnd == dlg || IsChild(dlg, msg.hwnd))) {
            LogLine(L"ShowSetupDialog: WM_COMMAND received, id=" + to_wstring(LOWORD(msg.wParam)) + 
                    L" code=" + to_wstring(HIWORD(msg.wParam)) + L" hwnd=" + to_wstring(reinterpret_cast<size_t>(msg.hwnd)));
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
            continue;
        }
        // Handle Escape key explicitly before IsDialogMessageW
        // Check if the message is for the dialog or any of its child controls
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE && 
            (msg.hwnd == dlg || IsChild(dlg, msg.hwnd))) {
            PostMessageW(dlg, WM_CLOSE, 0, 0);
            continue;
        }
        bool handled = IsDialogMessageW(dlg, &msg);
        if (handled && msg.message == WM_COMMAND) {
            LogLine(L"ShowSetupDialog: IsDialogMessageW handled WM_COMMAND, id=" + to_wstring(LOWORD(msg.wParam)) + 
                    L" code=" + to_wstring(HIWORD(msg.wParam)));
        }
        if (!handled) { TranslateMessage(&msg); DispatchMessageW(&msg); }
    }
    return st.result;
}

/**
 * @brief Check if filter definition has been modified
 * @param cur Current filter definition
 * @param orig Original filter definition
 * @return true if any field has changed
 */
bool IsFilterDirty(const FilterDefinition& cur, const FilterDefinition& orig) {
    return cur.title != orig.title || cur.input != orig.input || cur.output != orig.output ||
           cur.modelIndex != orig.modelIndex || cur.prompt != orig.prompt;
}

/**
 * @brief Collect filter definition from dialog UI controls
 * @param st Dialog state containing control handles
 * @param out Output filter definition
 */
void CollectFilterFromUI(EditDialogState* st, FilterDefinition& out) {
    wchar_t buf[1024];
    GetWindowTextW(st->hName, buf, 1024);
    out.title = buf;
    out.input = SendMessageW(st->hIn, CB_GETCURSEL, 0, 0) == 0 ? IOType::Text : IOType::Image;
    out.output = SendMessageW(st->hOut, CB_GETCURSEL, 0, 0) == 0 ? IOType::Text : IOType::Image;
    int msel = static_cast<int>(SendMessageW(st->hModel, CB_GETCURSEL, 0, 0));
    if (msel >= 0 && msel < static_cast<int>(g_models.size())) out.modelIndex = static_cast<size_t>(msel);
    int len = GetWindowTextLengthW(st->hPrompt);
    if (len < 0) len = 0;
    wstring tmp; tmp.resize(static_cast<size_t>(len) + 1);
    GetWindowTextW(st->hPrompt, tmp.data(), len + 1);
    tmp.resize(static_cast<size_t>(len));
    out.prompt = tmp;
}

/**
 * @brief Reassign model indices when a model is deleted
 * @param idx Index of deleted model
 * 
 * Updates all filters that reference the deleted model or models after it
 */
void ReassignModelOnDelete(size_t idx) {
    for (auto& f : g_filters) {
        if (f.modelIndex == idx) f.modelIndex = 0;
        else if (f.modelIndex > idx) --f.modelIndex;
    }
}

/**
 * @brief Populate combo box with model names
 * @param combo Combo box control handle
 */
void PopulateModelCombo(HWND combo) {
    SendMessageW(combo, CB_RESETCONTENT, 0, 0);
    for (const auto& m : g_models) SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(m.name.c_str()));
    wstring strAddModel = GetString(L"add_language_model");
    SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strAddModel.c_str()));
}

/**
 * @brief Subclassed window procedure for prompt edit control
 * Prevents Tab key from moving focus away from multiline edit
 */
LRESULT CALLBACK PromptEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_GETDLGCODE) {
        LRESULT code = CallWindowProcW(g_promptOldProc, hwnd, msg, wParam, lParam);
        return code & ~(DLGC_WANTTAB | DLGC_WANTALLKEYS);
    }
    if (msg == WM_KEYDOWN && wParam == 'A' && (GetKeyState(VK_CONTROL) & 0x8000)) {
        SendMessageW(hwnd, EM_SETSEL, 0, -1);
        return 0;
    }
    return CallWindowProcW(g_promptOldProc, hwnd, msg, wParam, lParam);
}

/**
 * @brief Subclassed window procedure for list view control
 * Handles Enter key to edit selected filter, Escape to close dialog
 */
LRESULT CALLBACK ListViewProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_KEYDOWN) {
        if (wParam == VK_RETURN) {
            PostMessageW(GetParent(hwnd), WM_COMMAND, MAKEWPARAM(IDC_BTN_EDIT, BN_CLICKED), 0);
            return 0;
        }
        if (wParam == VK_ESCAPE) {
            PostMessageW(GetParent(hwnd), WM_COMMAND, MAKEWPARAM(IDC_BTN_CLOSE, BN_CLICKED), 0);
            return 0;
        }
    }
    return CallWindowProcW(g_listOldProc, hwnd, msg, wParam, lParam);
}

/**
 * @brief Window procedure for filter edit dialog
 */
LRESULT CALLBACK EditDlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<EditDialogState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = reinterpret_cast<EditDialogState*>(reinterpret_cast<LPCREATESTRUCT>(lParam)->lpCreateParams);
        st->original = *st->filter;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));
        const int margin = 12, labelW = 160; int y = margin;
        wstring strFilterName = GetString(L"filter_name");
        CreateWindowW(L"STATIC", strFilterName.c_str(), WS_CHILD | WS_VISIBLE, margin, y, labelW, 22, hwnd, nullptr, nullptr, nullptr);
    st->hName = CreateWindowW(L"EDIT", st->filter->title.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
        margin + labelW + 6, y - 2, 400, 24, hwnd, (HMENU)(INT_PTR)100, nullptr, nullptr);
    EnableCtrlA(st->hName);
        y += 32;
        wstring strInputType = GetString(L"input_type");
        CreateWindowW(L"STATIC", strInputType.c_str(), WS_CHILD | WS_VISIBLE, margin, y, labelW, 22, hwnd, nullptr, nullptr, nullptr);
        st->hIn = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
            margin + labelW + 6, y - 2, 180, 300, hwnd, (HMENU)(INT_PTR)101, nullptr, nullptr);
        wstring strText = GetString(L"text_type");
        wstring strImage = GetString(L"image_type");
        SendMessageW(st->hIn, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strText.c_str()));
        SendMessageW(st->hIn, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strImage.c_str()));
        SendMessageW(st->hIn, CB_SETCURSEL, st->filter->input == IOType::Text ? 0 : 1, 0);
        wstring strOutputType = GetString(L"output_type");
        CreateWindowW(L"STATIC", strOutputType.c_str(), WS_CHILD | WS_VISIBLE, margin, y + 32, labelW, 22, hwnd, nullptr, nullptr, nullptr);
        st->hOut = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
            margin + labelW + 6, y + 30, 180, 300, hwnd, (HMENU)(INT_PTR)102, nullptr, nullptr);
        SendMessageW(st->hOut, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strText.c_str()));
        SendMessageW(st->hOut, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strImage.c_str()));
        SendMessageW(st->hOut, CB_SETCURSEL, st->filter->output == IOType::Text ? 0 : 1, 0);
        wstring strLanguageModel = GetString(L"language_model");
        CreateWindowW(L"STATIC", strLanguageModel.c_str(), WS_CHILD | WS_VISIBLE, margin, y + 64, labelW, 22, hwnd, nullptr, nullptr, nullptr);
        st->hModel = CreateWindowW(L"COMBOBOX", nullptr, WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
            margin + labelW + 6, y + 62, 300, 300, hwnd, (HMENU)(INT_PTR)103, nullptr, nullptr);
        for (const auto& m : g_models) SendMessageW(st->hModel, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(m.name.c_str()));
        wstring strAddModel = GetString(L"add_language_model");
        SendMessageW(st->hModel, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(strAddModel.c_str()));
        SendMessageW(st->hModel, CB_SETCURSEL, st->filter->modelIndex, 0);
        wstring strModelSettings = GetString(L"model_settings");
        CreateWindowW(L"BUTTON", strModelSettings.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, margin + labelW + 6 + 310, y + 61, 110, 26, hwnd,
            (HMENU)(INT_PTR)104, nullptr, nullptr);
        wstring strPrompt = GetString(L"prompt");
        CreateWindowW(L"STATIC", strPrompt.c_str(), WS_CHILD | WS_VISIBLE, margin, y + 96, labelW, 22, hwnd, nullptr, nullptr, nullptr);
    st->hPrompt = CreateWindowW(L"EDIT", st->filter->prompt.c_str(), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_LEFT | ES_MULTILINE | ES_AUTOVSCROLL | WS_TABSTOP,
        margin, y + 120, 600, 180, hwnd, (HMENU)(INT_PTR)105, nullptr, nullptr);
        g_promptOldProc = reinterpret_cast<WNDPROC>(SetWindowLongPtr(st->hPrompt, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(PromptEditProc)));
        wstring strSave = GetString(L"save");
        wstring strClose = GetString(L"close");
        CreateWindowW(L"BUTTON", strSave.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, margin + 360, y + 310, 100, 30, hwnd,
            (HMENU)(INT_PTR)106, nullptr, nullptr);
        CreateWindowW(L"BUTTON", strClose.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, margin + 470, y + 310, 100, 30, hwnd,
            (HMENU)(INT_PTR)107, nullptr, nullptr);
        SetFocus(st->hName);
        return 0;
    }
    case WM_COMMAND: {
        WORD id = LOWORD(wParam), code = HIWORD(wParam);
        if ((code == BN_CLICKED || code == 0) && id == 107) {
            FilterDefinition cur = *st->filter;
            CollectFilterFromUI(st, cur);
            if (IsFilterDirty(cur, st->original)) {
                wstring strUnsaved = GetString(L"unsaved_changes");
                wstring strConfirm = GetString(L"confirm");
                int r = MessageBoxW(hwnd, strUnsaved.c_str(), strConfirm.c_str(), MB_YESNOCANCEL | MB_ICONQUESTION);
                if (r == IDYES) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(106, BN_CLICKED), 0); return 0; }
                if (r == IDCANCEL) return 0;
            }
            DestroyWindow(hwnd); return 0;
        }
        if ((id == 101 || id == 102) && code == CBN_SELCHANGE) {
            return 0;
        }
        if (id == 103 && code == CBN_SELCHANGE) {
            int sel = static_cast<int>(SendMessageW(st->hModel, CB_GETCURSEL, 0, 0));
            if (sel == static_cast<int>(g_models.size())) {
                wstring strNewModel = GetString(L"new_model");
                ModelConfig m{ strNewModel, L"", L"", L"", L"" };
                g_models.push_back(m);
                size_t idx = g_models.size() - 1;
                int res = ShowModelDialog(hwnd, g_models[idx], idx);
                if (res == 2) g_models.pop_back();
                PopulateModelCombo(st->hModel);
                size_t newSel = (res == 1) ? idx : st->filter->modelIndex;
                if (newSel >= g_models.size()) newSel = 0;
                SendMessageW(st->hModel, CB_SETCURSEL, static_cast<WPARAM>(newSel), 0);
                SaveConfig();
            }
            return 0;
        }
        switch (id) {
        case 104: {
            int sel = static_cast<int>(SendMessageW(st->hModel, CB_GETCURSEL, 0, 0));
            if (sel >= 0 && sel < static_cast<int>(g_models.size())) {
                int res = ShowModelDialog(hwnd, g_models[sel], sel);
                if (res == 2 && g_models.size() > 1) {
                    g_models.erase(g_models.begin() + sel);
                    ReassignModelOnDelete(sel);
                    PopulateModelCombo(st->hModel);
                    int newSel = sel; if (newSel >= static_cast<int>(g_models.size())) newSel = 0;
                    SendMessageW(st->hModel, CB_SETCURSEL, newSel, 0);
                } else {
                    PopulateModelCombo(st->hModel);
                    int newSel = (sel < static_cast<int>(g_models.size())) ? sel : 0;
                    SendMessageW(st->hModel, CB_SETCURSEL, newSel, 0);
                }
                SaveConfig();
            }
            return 0;
        }
        case 106: {
            CollectFilterFromUI(st, *st->filter);
            SaveConfig(); DestroyWindow(hwnd); return 0;
        }
        }
        return 0;
    }
    case WM_KEYDOWN:
        if (wParam == VK_RETURN) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(106, BN_CLICKED), 0); return 0; }
        if (wParam == VK_ESCAPE) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(107, BN_CLICKED), 0); return 0; }
        break;
    case WM_CLOSE:
        PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(107, BN_CLICKED), 0); return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Show modal dialog for editing filter definition
 * @param parent Parent window handle
 * @param filter Filter definition to edit (modified in place)
 */
void ShowEditDialog(HWND parent, FilterDefinition& filter) {
    EditDialogState st{ &filter, filter };
    wstring strFilterEdit = GetString(L"filter_edit");
    HWND dlg = CreateWindowExW(WS_EX_DLGMODALFRAME | WS_EX_CONTROLPARENT, kEditClass, strFilterEdit.c_str(),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, CW_USEDEFAULT, CW_USEDEFAULT, 650, 430, parent, nullptr, g_hInst, &st);
    g_editWnd = dlg;
    ShowWindow(dlg, SW_SHOWNORMAL);
    EnableWindow(parent, FALSE);
    MSG msg;
    while (IsWindow(dlg) && GetMessageW(&msg, nullptr, 0, 0)) {
        // Handle Escape key explicitly before IsDialogMessageW
        // Check if the message is for the dialog or any of its child controls
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE && 
            (msg.hwnd == dlg || IsChild(dlg, msg.hwnd))) {
            PostMessageW(dlg, WM_CLOSE, 0, 0);
            continue;
        }
        if (!IsDialogMessageW(dlg, &msg)) { TranslateMessage(&msg); DispatchMessageW(&msg); }
    }
    EnableWindow(parent, TRUE); SetForegroundWindow(parent); g_editWnd = nullptr;
}
/**
 * @brief Update list view with current filter definitions
 * @param list List view control handle
 */
void UpdateListView(HWND list) {
    ListView_DeleteAllItems(list);
    int idx = 0; for (const auto& f : g_filters) {
        LVITEMW item{}; item.mask = LVIF_TEXT; item.iItem = idx; item.pszText = const_cast<LPWSTR>(f.title.c_str()); ListView_InsertItem(list, &item);
        ListView_SetItemText(list, idx, 1, const_cast<LPWSTR>(IOTypeToString(f.input).c_str()));
        ListView_SetItemText(list, idx, 2, const_cast<LPWSTR>(IOTypeToString(f.output).c_str()));
        ListView_SetItemText(list, idx, 3, const_cast<LPWSTR>(g_models[f.modelIndex].name.c_str()));
        ++idx;
    }
}

/**
 * @brief Window procedure for settings window
 */
LRESULT CALLBACK SettingsWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<SettingsState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = new SettingsState(); SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));
        wstring strFilterList = GetString(L"filter_list");
        CreateWindowW(L"STATIC", strFilterList.c_str(), WS_CHILD | WS_VISIBLE, 16, 10, 200, 20, hwnd, nullptr, nullptr, nullptr);
        st->hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, nullptr,
                                   WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS | WS_TABSTOP,
                                   16, 32, 560, 230, hwnd, (HMENU)(INT_PTR)IDC_LIST, g_hInst, nullptr);
        g_listOldProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(st->hList, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(ListViewProc)));
        ListView_SetExtendedListViewStyle(st->hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
        LVCOLUMNW col{}; col.mask = LVCF_TEXT | LVCF_WIDTH;
        wstring strFilter = GetString(L"filter");
        wstring strInput = GetString(L"input");
        wstring strOutput = GetString(L"output");
        wstring strModel = GetString(L"model");
        col.pszText = const_cast<LPWSTR>(strFilter.c_str()); col.cx = 200; ListView_InsertColumn(st->hList, 0, &col);
        col.pszText = const_cast<LPWSTR>(strInput.c_str()); col.cx = 80; ListView_InsertColumn(st->hList, 1, &col);
        col.pszText = const_cast<LPWSTR>(strOutput.c_str()); col.cx = 80; ListView_InsertColumn(st->hList, 2, &col);
        col.pszText = const_cast<LPWSTR>(strModel.c_str()); col.cx = 180; ListView_InsertColumn(st->hList, 3, &col);
        wstring strCopy = GetString(L"copy");
        wstring strAdd = GetString(L"add");
        wstring strEdit = GetString(L"edit");
        wstring strDelete = GetString(L"delete");
        wstring strClose = GetString(L"close");
        CreateWindowW(L"BUTTON", strCopy.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 16, 270, 80, 26, hwnd, (HMENU)(INT_PTR)IDC_BTN_COPY, g_hInst, nullptr);
        CreateWindowW(L"BUTTON", strAdd.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 104, 270, 80, 26, hwnd, (HMENU)(INT_PTR)IDC_BTN_ADD, g_hInst, nullptr);
        CreateWindowW(L"BUTTON", strEdit.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 192, 270, 80, 26, hwnd, (HMENU)(INT_PTR)IDC_BTN_EDIT, g_hInst, nullptr);
        CreateWindowW(L"BUTTON", strDelete.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 280, 270, 80, 26, hwnd, (HMENU)(INT_PTR)IDC_BTN_DELETE, g_hInst, nullptr);
        CreateWindowW(L"BUTTON", strClose.c_str(), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 496, 270, 80, 26, hwnd, (HMENU)(INT_PTR)IDC_BTN_CLOSE, g_hInst, nullptr);

        // Add hotkey display and change button
        wstring strCurrentHotkey = L"Current Hotkey:";
        CreateWindowW(L"STATIC", strCurrentHotkey.c_str(), WS_CHILD | WS_VISIBLE, 16, 310, 120, 20, hwnd, nullptr, nullptr, nullptr);
        wstring hotkeyStr = VKCodeToString(g_hotkeyKey, g_hotkeyModifiers);
        st->hHotkeyLabel = CreateWindowW(L"STATIC", hotkeyStr.c_str(), WS_CHILD | WS_VISIBLE | SS_CENTER | WS_BORDER,
            140, 308, 200, 22, hwnd, nullptr, nullptr, nullptr);
        st->hHotkeyButton = CreateWindowW(L"BUTTON", L"Change Hotkey...", WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            350, 308, 120, 22, hwnd, (HMENU)(INT_PTR)307, g_hInst, nullptr);

        UpdateListView(st->hList);
        return 0;
    }
    case WM_NOTIFY: {
        LPNMHDR hdr = reinterpret_cast<LPNMHDR>(lParam);
        if (hdr->idFrom == IDC_LIST && hdr->code == NM_DBLCLK) {
            wstring strUseEdit = GetString(L"use_edit_button");
            wstring strHint = GetString(L"hint");
            MessageBoxW(hwnd, strUseEdit.c_str(), strHint.c_str(), MB_OK | MB_ICONINFORMATION); return TRUE;
        }
        return 0;
    }
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDC_BTN_CLOSE: DestroyWindow(hwnd); return 0;
        case IDC_BTN_COPY: {
            int sel = ListView_GetNextItem(st->hList, -1, LVNI_SELECTED);
            if (sel >= 0 && sel < static_cast<int>(g_filters.size())) {
                FilterDefinition copy = g_filters[sel]; copy.title += GetString(L"copy_suffix"); g_filters.push_back(copy);
                UpdateListView(st->hList); int idx = static_cast<int>(g_filters.size() - 1); ListView_SetItemState(st->hList, idx, LVIS_SELECTED, LVIS_SELECTED); SaveConfig();
            }
            return 0;
        }
        case IDC_BTN_ADD: {
            wstring strNewFilter = GetString(L"new_filter");
            FilterDefinition def{ strNewFilter, IOType::Text, IOType::Text, 0, L"" };
            g_filters.push_back(def); UpdateListView(st->hList);
            int idx = static_cast<int>(g_filters.size() - 1); ListView_SetItemState(st->hList, idx, LVIS_SELECTED, LVIS_SELECTED);
            ShowEditDialog(hwnd, g_filters.back()); UpdateListView(st->hList);
            int reselection = idx; if (reselection >= static_cast<int>(g_filters.size())) reselection = static_cast<int>(g_filters.size()) - 1;
            if (reselection >= 0) ListView_SetItemState(st->hList, reselection, LVIS_SELECTED, LVIS_SELECTED);
            SaveConfig(); return 0;
        }
        case IDC_BTN_EDIT: {
            int sel = ListView_GetNextItem(st->hList, -1, LVNI_SELECTED);
            if (sel >= 0 && sel < static_cast<int>(g_filters.size())) {
                ShowEditDialog(hwnd, g_filters[sel]); UpdateListView(st->hList);
                int reselection = sel; if (reselection >= static_cast<int>(g_filters.size())) reselection = static_cast<int>(g_filters.size()) - 1;
                if (reselection >= 0) ListView_SetItemState(st->hList, reselection, LVIS_SELECTED, LVIS_SELECTED); SaveConfig();
            }
            return 0;
        }
        case IDC_BTN_DELETE: {
            int sel = ListView_GetNextItem(st->hList, -1, LVNI_SELECTED);
            if (sel >= 0 && sel < static_cast<int>(g_filters.size())) { g_filters.erase(g_filters.begin() + sel); UpdateListView(st->hList); SaveConfig(); }
            return 0;
        }
        case 307: { // Change Hotkey button
            UINT vk = g_hotkeyKey;
            UINT mods = g_hotkeyModifiers;
            if (ShowHotkeyInputDialog(hwnd, vk, mods) == 1) {
                // Test if the new hotkey can be registered
                if (!RegisterHotKey(hwnd, HOTKEY_ID + 2, mods | MOD_NOREPEAT, vk)) {
                    wstring keyStr = VKCodeToString(vk, mods);
                    wstring errMsg = L"Cannot register hotkey: " + keyStr + L"\n"
                                   + L"The hotkey may already be in use.";
                    MessageBoxW(hwnd, errMsg.c_str(), L"cbfilter", MB_OK | MB_ICONWARNING);
                    return 0;
                }
                UnregisterHotKey(hwnd, HOTKEY_ID + 2);

                // Update global hotkey settings
                g_hotkeyKey = vk;
                g_hotkeyModifiers = mods;

                // Update display
                wstring keyStr = VKCodeToString(vk, mods);
                SetWindowTextW(st->hHotkeyLabel, keyStr.c_str());

                // Re-register hotkey on main window
                HWND mainWnd = FindWindowW(kClassName, L"cbfilter");
                if (mainWnd) {
                    UnregisterHotKey(mainWnd, HOTKEY_ID);
                    if (!RegisterHotKey(mainWnd, HOTKEY_ID, mods | MOD_NOREPEAT, vk)) {
                        MessageBoxW(hwnd, L"Failed to re-register hotkey on main window.", L"cbfilter", MB_OK | MB_ICONERROR);
                    }
                }

                SaveConfig();
            }
            return 0;
        }
        default: return 0;
        }
    }
    case WM_KEYDOWN:
        if (wParam == VK_RETURN) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_EDIT, BN_CLICKED), 0); return 0; }
        if (wParam == VK_ESCAPE) { DestroyWindow(hwnd); return 0; }
        break;
    case WM_CLOSE: DestroyWindow(hwnd); return 0;
    case WM_NCDESTROY: delete st; SetWindowLongPtr(hwnd, GWLP_USERDATA, 0); if (g_settingsWnd == hwnd) g_settingsWnd = nullptr; return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}
/**
 * @brief Show or activate settings window
 * @param hInst Application instance handle
 */
void ShowSettingsWindow(HINSTANCE hInst) {
    if (g_settingsWnd && IsWindow(g_settingsWnd)) { ShowWindow(g_settingsWnd, SW_SHOWNORMAL); SetForegroundWindow(g_settingsWnd); return; }
    wstring strSettings = GetString(L"settings");
    g_settingsWnd = CreateWindowExW(WS_EX_CONTROLPARENT, kSettingsClass, strSettings.c_str(), WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 620, 400, nullptr, nullptr, hInst, nullptr);
    ShowWindow(g_settingsWnd, SW_SHOW);
}

/**
 * @brief Add application icon to system tray
 * @param hwnd Window handle to receive tray notifications
 */
void AddTrayIcon(HWND hwnd) {
    NOTIFYICONDATA nid{};
    nid.cbSize = sizeof(NOTIFYICONDATA);
    nid.hWnd = hwnd;
    nid.uID = 1;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_APP_TRAY;
    nid.hIcon = static_cast<HICON>(LoadImageW(g_hInst, MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR));
    lstrcpyW(nid.szTip, L"cbfilter");
    Shell_NotifyIconW(NIM_ADD, &nid);
}

/**
 * @brief Remove application icon from system tray
 * @param hwnd Window handle
 */
void RemoveTrayIcon(HWND hwnd) { NOTIFYICONDATA nid{}; nid.cbSize = sizeof(NOTIFYICONDATA); nid.hWnd = hwnd; nid.uID = 1; Shell_NotifyIconW(NIM_DELETE, &nid); }

/**
 * @brief Show context menu for system tray icon
 * @param hwnd Window handle
 */
void ShowTrayMenu(HWND hwnd) {
    POINT pt; GetCursorPos(&pt); HMENU tray = CreatePopupMenu();
    wstring strSettings = GetString(L"settings");
    wstring strExit = GetString(L"exit");
    InsertMenuW(tray, 0, MF_BYPOSITION | MF_STRING, MENU_ID_SETTINGS, strSettings.c_str());
    InsertMenuW(tray, 1, MF_BYPOSITION | MF_STRING, MENU_ID_EXIT, strExit.c_str());
    SetForegroundWindow(hwnd);
    UINT cmd = TrackPopupMenu(tray, TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, hwnd, nullptr); DestroyMenu(tray);
    if (cmd == MENU_ID_SETTINGS) ShowSettingsWindow(g_hInst); else if (cmd == MENU_ID_EXIT) PostMessageW(hwnd, WM_CLOSE, 0, 0);
}

// Structure to hold progress window state
struct ProgressWindowState {
    FilterDefinition* filter;
    HWND hwndPreviousActive;
    DWORD startTime;
    bool result;
    HWND hwndProgress;
    HANDLE hThread;  // Thread handle for filter execution
};

// Thread function to run filter asynchronously
DWORD WINAPI RunFilterThread(LPVOID lpParam) {
    ProgressWindowState* state = static_cast<ProgressWindowState*>(lpParam);
    state->result = RunFilter(*state->filter);
    PostMessageW(state->hwndProgress, WM_APP_FILTER_COMPLETE, 0, reinterpret_cast<LPARAM>(state));
    return 0;
}

/**
 * @brief Window procedure for progress window
 * Displays filter execution progress with elapsed time
 */
LRESULT CALLBACK ProgressWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_CREATE) {
        CREATESTRUCT* cs = reinterpret_cast<CREATESTRUCT*>(lParam);
        ProgressWindowState* state = static_cast<ProgressWindowState*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
        
        // Create static text for message
        wstring strExecuting = GetString(L"executing_filter");
        CreateWindowW(L"STATIC", strExecuting.c_str(), WS_VISIBLE | WS_CHILD | SS_LEFT,
            20, 20, 300, 20, hwnd, nullptr, g_hInst, nullptr);
        
        // Create static text for elapsed time
        CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
            20, 50, 300, 20, hwnd, reinterpret_cast<HMENU>(1001), g_hInst, nullptr);
        
        // Start timer to update elapsed time (every 100ms)
        SetTimer(hwnd, TIMER_ID_PROGRESS, 100, nullptr);
        
        // Start filter execution in background thread
        state->startTime = GetTickCount();
        state->hwndProgress = hwnd;
        state->hThread = CreateThread(nullptr, 0, RunFilterThread, state, 0, nullptr);
        if (!state->hThread) {
            // Failed to create thread, cleanup and close window
            delete state;
            DestroyWindow(hwnd);
            return -1;
        }
        
        return 0;
    }
    
    ProgressWindowState* state = reinterpret_cast<ProgressWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (!state) return DefWindowProcW(hwnd, msg, wParam, lParam);
    
    if (msg == WM_TIMER && wParam == TIMER_ID_PROGRESS) {
        // Update elapsed time
        DWORD elapsed = (GetTickCount() - state->startTime) / 1000; // Convert to seconds
        wstring strElapsed = GetString(L"elapsed_time");
        // Replace {0} with elapsed time
        size_t pos = strElapsed.find(L"{0}");
        if (pos != wstring::npos) {
            strElapsed.replace(pos, 3, to_wstring(elapsed));
        } else {
            // Fallback if {0} not found
            strElapsed += L" " + to_wstring(elapsed);
            if (g_language == L"ja") strElapsed += L"秒";
            else strElapsed += L" seconds";
        }
        SetWindowTextW(GetDlgItem(hwnd, 1001), strElapsed.c_str());
        return 0;
    }
    
    if (msg == WM_APP_FILTER_COMPLETE) {
        // Filter execution completed
        KillTimer(hwnd, TIMER_ID_PROGRESS);
        // Close thread handle (thread has already completed)
        if (state->hThread) {
            CloseHandle(state->hThread);
            state->hThread = nullptr;
        }
        DestroyWindow(hwnd);
        return 0;
    }
    
    if (msg == WM_DESTROY) {
        // Restore focus and send Ctrl+V if successful
        if (state->result) {
            if (state->hwndPreviousActive && IsWindow(state->hwndPreviousActive)) {
                SetForegroundWindow(state->hwndPreviousActive);
                SetFocus(state->hwndPreviousActive);
                Sleep(80);
            }
            SendCtrlV();
        } else {
            wstring strFilterFailed = GetString(L"filter_execution_failed");
            MessageBoxW(hwnd, strFilterFailed.c_str(), L"cbfilter", MB_OK | MB_ICONERROR);
        }
        // Ensure thread handle is closed (safety check in case WM_APP_FILTER_COMPLETE was missed)
        if (state->hThread) {
            CloseHandle(state->hThread);
            state->hThread = nullptr;
        }
        delete state;
        g_progressWnd = nullptr;
        return 0;
    }
    
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

/**
 * @brief Show progress window and execute filter asynchronously
 * @param hwnd Window handle
 * @param filter Filter to execute
 * @param hwndPreviousActive Previous active window to restore focus to
 */
void ShowProgressAndRunFilter(HWND hwnd, const FilterDefinition& filter, HWND hwndPreviousActive) {
    ProgressWindowState* state = new ProgressWindowState();
    state->filter = const_cast<FilterDefinition*>(&filter);
    state->hwndPreviousActive = hwndPreviousActive;
    state->result = false;
    state->hThread = nullptr;
    
    // Get screen center position for progress window
    int screenWidth = GetSystemMetrics(SM_CXSCREEN);
    int screenHeight = GetSystemMetrics(SM_CYSCREEN);
    int windowWidth = 350;
    int windowHeight = 120;
    int x = (screenWidth - windowWidth) / 2;
    int y = (screenHeight - windowHeight) / 2;
    
    // Create progress window (non-modal, handled by main message loop)
    HWND progressWnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_DLGMODALFRAME,
        kProgressClass,
        L"cbfilter",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_POPUP,
        x, y,
        windowWidth, windowHeight,
        hwnd,
        nullptr,
        g_hInst,
        state
    );
    
    if (progressWnd) {
        g_progressWnd = progressWnd;
        ShowWindow(progressWnd, SW_SHOWNORMAL);
        UpdateWindow(progressWnd);
        // Window will be handled by main message loop, no need for separate loop here
    } else {
        delete state;
    }
}

/**
 * @struct FilterMenuState
 * @brief State for filter selection menu window
 */
struct FilterMenuState {
    vector<int> filterIndices;  // Indices into g_filters
    int selectedIndex{};        // Currently selected item (0-based)
    int result{-1};             // Selected filter index or -1 if cancelled
    HWND hwndPreviousActive{};  // Window to restore focus to
    HWND hwndParent{};          // Parent window (main window) to send close message to
};

/**
 * @brief Window procedure for filter selection menu
 */
LRESULT CALLBACK FilterMenuWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* st = reinterpret_cast<FilterMenuState*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        st = reinterpret_cast<FilterMenuState*>(reinterpret_cast<LPCREATESTRUCT>(lParam)->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(st));
        return 0;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        // Draw background
        HBRUSH hBrush = CreateSolidBrush(RGB(255, 255, 255));
        FillRect(hdc, &rc, hBrush);
        DeleteObject(hBrush);

        // Draw menu items
        int itemHeight = 30;
        for (size_t i = 0; i < st->filterIndices.size(); ++i) {
            RECT itemRect = { 4, static_cast<LONG>(i * itemHeight + 4), rc.right - 4, static_cast<LONG>((i + 1) * itemHeight) };

            // Highlight selected item
            if (static_cast<int>(i) == st->selectedIndex) {
                HBRUSH hSelBrush = CreateSolidBrush(RGB(0, 120, 215));
                FillRect(hdc, &itemRect, hSelBrush);
                DeleteObject(hSelBrush);
                SetTextColor(hdc, RGB(255, 255, 255));
            } else {
                SetTextColor(hdc, RGB(0, 0, 0));
            }

            SetBkMode(hdc, TRANSPARENT);
            HFONT hFont = GetUIFont();
            SelectObject(hdc, hFont);

            // Draw number (1-9, 0 for 10th)
            wstring numStr;
            if (i < 9) numStr = to_wstring(i + 1) + L". ";
            else if (i == 9) numStr = L"0. ";
            else numStr = L"   ";

            const auto& filter = g_filters[st->filterIndices[i]];
            wstring text = numStr + filter.title;
            DrawTextW(hdc, text.c_str(), -1, &itemRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        }

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_KEYDOWN: {
        if (wParam == VK_ESCAPE) {
            st->result = -1;
            DestroyWindow(hwnd);
            return 0;
        }
        if (wParam == VK_RETURN) {
            if (st->selectedIndex >= 0 && st->selectedIndex < static_cast<int>(st->filterIndices.size())) {
                st->result = st->filterIndices[st->selectedIndex];
                DestroyWindow(hwnd);
            }
            return 0;
        }
        if (wParam == VK_UP) {
            if (st->selectedIndex > 0) {
                st->selectedIndex--;
                InvalidateRect(hwnd, nullptr, FALSE);
            } else {
                st->selectedIndex = static_cast<int>(st->filterIndices.size()) - 1;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        }
        if (wParam == VK_DOWN) {
            if (st->selectedIndex < static_cast<int>(st->filterIndices.size()) - 1) {
                st->selectedIndex++;
                InvalidateRect(hwnd, nullptr, FALSE);
            } else {
                st->selectedIndex = 0;
                InvalidateRect(hwnd, nullptr, FALSE);
            }
            return 0;
        }
        // Handle number keys 1-9, 0
        if (wParam >= '1' && wParam <= '9') {
            int idx = static_cast<int>(wParam - '1');
            if (idx < static_cast<int>(st->filterIndices.size())) {
                st->result = st->filterIndices[idx];
                DestroyWindow(hwnd);
            }
            return 0;
        }
        if (wParam == '0') {
            if (st->filterIndices.size() >= 10) {
                st->result = st->filterIndices[9];
                DestroyWindow(hwnd);
            }
            return 0;
        }
        break;
    }
    case WM_LBUTTONDOWN: {
        int y = GET_Y_LPARAM(lParam);
        int itemHeight = 30;
        int idx = (y - 4) / itemHeight;
        if (idx >= 0 && idx < static_cast<int>(st->filterIndices.size())) {
            st->result = st->filterIndices[idx];
            LogLine(L"FilterMenuWndProc: WM_LBUTTONDOWN selected index=" + to_wstring(st->result));
            // Post message to self to exit message loop before destroying window
            PostMessageW(hwnd, WM_APP_MENU_SELECTED, 0, 0);
        }
        return 0;
    }
    case WM_APP_MENU_SELECTED:
        // Menu item was selected, close the window
        LogLine(L"FilterMenuWndProc: WM_APP_MENU_SELECTED received, destroying window");
        DestroyWindow(hwnd);
        return 0;
    case WM_ACTIVATE:
        // Close the menu when it loses focus (but only if no item was selected)
        if (LOWORD(wParam) == WA_INACTIVE && st->result < 0) {
            st->result = -1;
            if (st->hwndParent) {
                PostMessageW(st->hwndParent, WM_APP_MENU_CLOSE, 0, 0);  // Notify menu close
            }
            DestroyWindow(hwnd);
        }
        return 0;
    case WM_CLOSE:
        // Only set result to -1 if no item was selected
        if (st->result < 0) {
            st->result = -1;
        }
        if (st && st->hwndParent) {
            PostMessageW(st->hwndParent, WM_APP_MENU_CLOSE, 0, 0);  // Notify menu close
        }
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        LogLine(L"FilterMenuWndProc: WM_DESTROY");
        if (g_filterMenuWnd == hwnd) g_filterMenuWnd = nullptr;
        return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Show filter selection menu and execute selected filter
 * @param hwnd Window handle
 * @param hwndPreviousActive Previous active window to restore focus to (optional)
 * @return true if filter was executed successfully
 *
 * Shows a custom menu window with filters compatible with current clipboard content,
 * then executes the selected filter and pastes the result.
 */
bool ShowFilterMenuAndRun(HWND hwnd, HWND hwndPreviousActive = nullptr) {
    // Check if a filter is already running
    if (g_progressWnd && IsWindow(g_progressWnd)) {
        SetForegroundWindow(g_progressWnd);
        MessageBeep(MB_ICONWARNING);
        return false;
    }

    // Close existing filter menu if one is already open
    if (g_filterMenuWnd && IsWindow(g_filterMenuWnd)) {
        DestroyWindow(g_filterMenuWnd);
        g_filterMenuWnd = nullptr;
    }

    ClipboardType ct = DetectClipboard();
    FilterMenuState st;
    st.hwndPreviousActive = hwndPreviousActive;
    st.hwndParent = hwnd;

    // Build list of compatible filters
    for (size_t i = 0; i < g_filters.size(); ++i) {
        const auto& f = g_filters[i];
        if ((ct == ClipboardType::Text && f.input != IOType::Text) || (ct == ClipboardType::Bitmap && f.input != IOType::Image)) continue;
        st.filterIndices.push_back(static_cast<int>(i));
    }

    if (st.filterIndices.empty()) {
        wstring strNoFilters = GetString(L"no_compatible_filters");
        MessageBoxW(hwnd, strNoFilters.c_str(), L"cbfilter", MB_OK | MB_ICONINFORMATION);
        return false;
    }

    // Get cursor position and calculate window size
    POINT pt;
    GetCursorPos(&pt);
    int itemHeight = 30;
    int windowWidth = 300;
    int windowHeight = static_cast<int>(st.filterIndices.size()) * itemHeight + 8;

    // Create custom filter menu window
    HWND menuWnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        kFilterMenuClass,
        L"",
        WS_POPUP | WS_BORDER,
        pt.x, pt.y,
        windowWidth, windowHeight,
        hwnd,
        nullptr,
        g_hInst,
        &st
    );

    if (!menuWnd) return false;

    g_filterMenuWnd = menuWnd;  // Track the menu window globally
    ShowWindow(menuWnd, SW_SHOWNORMAL);
    SetForegroundWindow(menuWnd);
    SetFocus(menuWnd);

    // Modal message loop
    MSG msg;
    LogLine(L"ShowFilterMenuAndRun: Starting message loop");
    bool menuClosed = false;
    while (IsWindow(menuWnd) && !menuClosed) {
        BOOL bRet = GetMessageW(&msg, nullptr, 0, 0);
        if (bRet == 0) {
            LogLine(L"ShowFilterMenuAndRun: GetMessage returned 0 (WM_QUIT)");
            break;  // WM_QUIT
        }
        if (bRet == -1) {
            LogLine(L"ShowFilterMenuAndRun: GetMessage returned -1 (error)");
            break;  // Error
        }
        LogLine(L"ShowFilterMenuAndRun: Received message=" + to_wstring(msg.message));
        if (msg.message == WM_APP_MENU_CLOSE || msg.message == WM_APP_MENU_SELECTED) {
            LogLine(L"ShowFilterMenuAndRun: Received menu close message, breaking loop");
            menuClosed = true;
            // Still dispatch the message so WM_APP_MENU_SELECTED can be handled
            if (msg.message == WM_APP_MENU_SELECTED) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            break;  // Menu closed
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    LogLine(L"ShowFilterMenuAndRun: Message loop ended, st.result=" + to_wstring(st.result));

    // Clear the global menu window reference
    if (g_filterMenuWnd == menuWnd) {
        g_filterMenuWnd = nullptr;
    }

    // Check if a filter was selected
    if (st.result >= 0 && st.result < static_cast<int>(g_filters.size())) {
        LogLine(L"ShowFilterMenuAndRun: Executing filter index=" + to_wstring(st.result));
        ShowProgressAndRunFilter(hwnd, g_filters[st.result], hwndPreviousActive);
        return true;
    }

    LogLine(L"ShowFilterMenuAndRun: No filter selected, returning false");
    return false;
}

/**
 * @brief Window procedure for hidden main window
 * Handles hotkey notifications and system tray messages
 */
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        // Register hotkey with configured modifiers and key
        if (!RegisterHotKey(hwnd, HOTKEY_ID, g_hotkeyModifiers | MOD_NOREPEAT, g_hotkeyKey)) {
            wstring keyStr = VKCodeToString(g_hotkeyKey, g_hotkeyModifiers);
            wstring errMsg = L"Failed to register hotkey: " + keyStr + L"\n"
                           + L"The hotkey may already be in use by another application.\n"
                           + L"Please change the hotkey in Settings.";
            MessageBoxW(hwnd, errMsg.c_str(), L"cbfilter - Hotkey Registration Failed", MB_OK | MB_ICONWARNING);
        }
        return 0;
    }
    case WM_APP_TRAY:
        if (lParam == WM_RBUTTONUP || lParam == WM_LBUTTONUP || lParam == WM_CONTEXTMENU) { ShowTrayMenu(hwnd); return 0; }
        else if (lParam == WM_LBUTTONDBLCLK) { ShowSettingsWindow(g_hInst); return 0; }
        break;
    case WM_HOTKEY: {
        // Check if a filter is already running
        if (g_progressWnd && IsWindow(g_progressWnd)) {
            // Bring the progress window to front instead of starting a new filter
            SetForegroundWindow(g_progressWnd);
            MessageBeep(MB_ICONWARNING);
            return 0;
        }
        // Save the currently active window before showing menu
        HWND hwndActive = GetForegroundWindow();
        ShowFilterMenuAndRun(hwnd, hwndActive);
        return 0;
    }
    case WM_DESTROY: UnregisterHotKey(hwnd, HOTKEY_ID); RemoveTrayIcon(hwnd); PostQuitMessage(0); return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/**
 * @brief Register a window class
 * @param hInst Application instance handle
 * @param kClass Window class name
 * @param proc Window procedure
 */
void RegWindowClass(HINSTANCE hInst, LPCWSTR kClass, WNDPROC proc) {
    WNDCLASSEXW wc{ sizeof(WNDCLASSEXW) };
    wc.lpfnWndProc = proc;
    wc.hInstance = hInst;
    wc.lpszClassName = kClass;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    RegisterClassExW(&wc);
}

/**
 * @brief Application entry point
 * Initializes GDI+, registers window classes, creates hidden window,
 * and runs message loop
 */
int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int) {
    winrt::init_apartment();
    g_hInst = hInst;
    using namespace winrt::Windows::Data::Json;
    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    if (Gdiplus::GdiplusStartup(&g_gdiplusToken, &gdiplusStartupInput, nullptr) != Gdiplus::Ok) return 1;
    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_WIN95_CLASSES }; InitCommonControlsEx(&icc);
    INITCOMMONCONTROLSEX icc2{ sizeof(icc2), ICC_LISTVIEW_CLASSES }; InitCommonControlsEx(&icc2);
    LoadApiDefinitions();
    JsonObject defCfg = CreateDefaultConfig();
    if (defCfg.HasKey(L"language")) g_language = wstring(defCfg.GetNamedString(L"language", g_language.c_str()).c_str());
    if (defCfg.HasKey(L"hotkey")) {
        JsonObject hk = defCfg.GetNamedObject(L"hotkey");
        g_hotkeyModifiers = static_cast<UINT>(hk.GetNamedNumber(L"modifiers", g_hotkeyModifiers));
        g_hotkeyKey = static_cast<UINT>(hk.GetNamedNumber(L"key", g_hotkeyKey));
    }
    // Register setup dialog class early because it may be shown before config is created
    RegWindowClass(hInst, kSetupClass, SetupDlgProc);
    RegWindowClass(hInst, kHotkeyInputClass, HotkeyInputDlgProc);

    if (!FileExists(GetConfigPath())) {
        int setupRes = ShowSetupDialog();
        if (setupRes != 1) return 0;
    }
    LoadConfig();
    EnsureModelProviders();
    // Initialize default filters if none loaded
    if (g_filters.empty()) {
        wstring strTranslate = GetString(L"translate_to_english");
        wstring strSummarize = GetString(L"summarize");
        g_filters.push_back({ strTranslate, IOType::Text, IOType::Text, 0, L"Translate into English." });
        g_filters.push_back({ strSummarize, IOType::Text, IOType::Text, 0, L"Summarize the following text." });
        SaveConfig();
    }
    RegWindowClass(hInst, kSettingsClass, SettingsWndProc);
    RegWindowClass(hInst, kEditClass, EditDlgProc);
    RegWindowClass(hInst, kModelClass, ModelDlgProc);
    RegWindowClass(hInst, kProgressClass, ProgressWndProc);
    RegWindowClass(hInst, kFilterMenuClass, FilterMenuWndProc);
    RegWindowClass(hInst, kClassName, WndProc);
    HWND hwnd = CreateWindowExW(0, kClassName, L"cbfilter", WS_OVERLAPPED, 0, 0, 0, 0, nullptr, nullptr, hInst, nullptr);
    if (!hwnd) return 1;
    ShowWindow(hwnd, SW_HIDE);
    AddTrayIcon(hwnd);
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        // Handle Escape key explicitly before IsDialogMessageW for all dialogs
        // Check if the message is for a dialog or any of its child controls
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE) {
            HWND targetWnd = nullptr;
            if (g_settingsWnd && (msg.hwnd == g_settingsWnd || IsChild(g_settingsWnd, msg.hwnd))) {
                targetWnd = g_settingsWnd;
            } else if (g_editWnd && (msg.hwnd == g_editWnd || IsChild(g_editWnd, msg.hwnd))) {
                targetWnd = g_editWnd;
            } else if (g_modelWnd && (msg.hwnd == g_modelWnd || IsChild(g_modelWnd, msg.hwnd))) {
                targetWnd = g_modelWnd;
            } else if (g_progressWnd && (msg.hwnd == g_progressWnd || IsChild(g_progressWnd, msg.hwnd))) {
                targetWnd = g_progressWnd;
            }
            if (targetWnd) {
                PostMessageW(targetWnd, WM_CLOSE, 0, 0);
                continue;
            }
        }
        if (g_settingsWnd && IsDialogMessageW(g_settingsWnd, &msg)) continue;
        if (g_editWnd && IsDialogMessageW(g_editWnd, &msg)) continue;
        if (g_modelWnd && IsDialogMessageW(g_modelWnd, &msg)) continue;
        if (g_progressWnd && IsDialogMessageW(g_progressWnd, &msg)) continue;
        TranslateMessage(&msg); DispatchMessageW(&msg);
    }
    if (g_gdiplusToken) Gdiplus::GdiplusShutdown(g_gdiplusToken);
    return 0;
}
