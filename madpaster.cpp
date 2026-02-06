/*
 * MadPaster - Windows Keyboard Paste Utility
 * Author: Dave Cox
 * Version: 3.0.1
 * Date: January 31, 2026
 */

#define UNICODE
#define _UNICODE

#include <windows.h>
#include <commctrl.h>   // For up-down (spin) control
#include <commdlg.h>    // For GetOpenFileName file dialog
#include <shellapi.h>   // For Shell_NotifyIcon (system tray)
#include <gdiplus.h>    // For PNG image loading
#include <mmsystem.h>   // For timeBeginPeriod/timeEndPeriod
#include <string>
#include <vector>

using namespace Gdiplus;

// Link common controls
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "winmm.lib")

// ============================================================================
// Constants and Control IDs
// ============================================================================

const int maxchar = 45000;

// Injection modes for different target types
enum class InjectionMode {
    Unicode,    // KEYEVENTF_UNICODE - works for local apps
    VKScancode, // VK codes with scancodes - better for remote clients
    Hybrid,     // Try VK first, fall back to Unicode
    Auto        // Detect target type and choose mode
};

// Pacing strategies for input injection
enum class PacingStrategy {
    Burst,        // Send chunk, pause after - for local targets
    PerCharacter, // Pause after each complete character - for remote
    PerEvent      // Pause between every INPUT event - most conservative
};

// Input injection constants (legacy burst mode)
const int CHUNK_SIZE = 2;             // Characters per SendInput batch (conservative)
const int INTER_CHUNK_PAUSE_MS = 25;  // Base pause between chunks
const int NEWLINE_PAUSE_MS = 100;     // Pause before/after newlines
const int MAX_RETRY_COUNT = 3;        // Retries on partial SendInput
const int IDLE_WAIT_MS = 50;          // Max wait for WaitForInputIdle

// Per-event pacing constants (new mode)
const int PER_EVENT_DELAY_MS = 2;     // Delay between each INPUT event
const int PER_CHAR_DELAY_MS = 5;      // Delay after each complete character
const int LINE_START_GUARD_CHARS = 3; // Extra delay for first N chars after newline
const int LINE_START_GUARD_MS = 10;   // Extra delay per guard char

// Remote client window classes (null-terminated array)
const wchar_t* REMOTE_WINDOW_CLASSES[] = {
    L"TscShellContainerClass",  // mstsc.exe (RDP)
    L"ICAClientClass",          // Citrix Receiver
    L"RAIL_WINDOW",             // Citrix seamless apps
    L"Transparent Windows Client", // Azure Virtual Desktop
    L"vncviewer",               // VNC clients
    L"TightVNC",
    L"RealVNC",
    L"MozillaWindowClass",      // Firefox (noVNC)
    L"Chrome_WidgetWin_1",      // Chrome/Edge (noVNC, Azure Bastion)
    nullptr
};

// Window dimensions
const int WINDOW_WIDTH = 400;
const int WINDOW_HEIGHT = 439;

// Control IDs
#define IDC_RADIO_CLIPBOARD     101
#define IDC_RADIO_FILE          102
#define IDC_EDIT_DELAY          103
#define IDC_SPIN_DELAY          104
#define IDC_BUTTON_ARM          105
#define IDC_BUTTON_BROWSE       106
#define IDC_STATIC_FILEPATH     107
#define IDC_STATIC_STATUS       108
#define IDC_EDIT_KEYSTROKE      109
#define IDC_SPIN_KEYSTROKE      110
#define IDC_COMBO_MODE          111
#define IDC_CHECK_DIAG          112
#define IDC_PROGRESS            113
#define IDC_CHECK_SILENT        114

// Timer IDs
#define IDT_COUNTDOWN           201

// Icons
#define IDI_APPICON             100  // Embedded resource icon

// Tray icon
#define IDI_TRAY                301
#define WM_TRAYICON            (WM_USER + 1)

// Tray menu items
#define IDM_TRAY_ARM            401
#define IDM_TRAY_SHOW           402
#define IDM_TRAY_EXIT           403

// Hotkey IDs
#define IDH_PASTE_HOTKEY        501

// Floating progress window
#define FLOATING_PROGRESS_CLASS L"MadPasterFloatingProgress"
#define FLOATING_PROGRESS_WIDTH 300
#define FLOATING_PROGRESS_HEIGHT 70

// ============================================================================
// Global Application State
// ============================================================================

struct AppState {
    HINSTANCE hInstance;
    HWND hwndMain;
    HWND hwndRadioClipboard;
    HWND hwndRadioFile;
    HWND hwndEditDelay;
    HWND hwndSpinDelay;
    HWND hwndEditKeystroke;
    HWND hwndSpinKeystroke;
    HWND hwndComboMode;
    HWND hwndCheckDiag;
    HWND hwndCheckSilent;
    HWND hwndButtonArm;
    HWND hwndButtonBrowse;
    HWND hwndStaticFilePath;
    HWND hwndStaticStatus;
    HWND hwndProgress;
    HWND hwndLogo;

    // Floating progress window (visible when minimized to tray)
    HWND hwndFloatingProgress;
    HWND hwndFloatingProgressBar;
    HWND hwndFloatingLabel;

    NOTIFYICONDATA nid;
    bool minimizedToTray;

    // Custom fonts
    HFONT hFontUI;
    HFONT hFontMono;
    HFONT hFontButton;

    // Custom icon
    HICON hAppIcon;

    // Logo image
    Gdiplus::Image* pLogoImage;
    ULONG_PTR gdiplusToken;

    // Settings
    bool useClipboard;
    int delaySeconds;
    int keystrokeDelayMs;
    std::wstring selectedFilePath;

    // Countdown state
    bool isArmed;
    int countdownRemaining;

    // Injection settings
    InjectionMode injectionMode;
    bool diagnosticMode;
    bool silentMode;
};

static AppState g_app = {};

// ============================================================================
// File Encoding Support
// ============================================================================

enum class FileEncoding {
    UTF8_BOM,
    UTF16_LE_BOM,
    UTF16_BE_BOM,
    ANSI_OR_UTF8
};

// Detect file encoding from BOM
FileEncoding detectEncoding(const std::vector<unsigned char>& buffer) {
    if (buffer.size() >= 3 &&
        buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF) {
        return FileEncoding::UTF8_BOM;
    }
    if (buffer.size() >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE) {
        return FileEncoding::UTF16_LE_BOM;
    }
    if (buffer.size() >= 2 && buffer[0] == 0xFE && buffer[1] == 0xFF) {
        return FileEncoding::UTF16_BE_BOM;
    }
    return FileEncoding::ANSI_OR_UTF8;
}

// Convert UTF-8 to wide string
std::wstring utf8ToWide(const char* utf8Str, size_t len) {
    if (len == 0) return L"";

    int wideLen = MultiByteToWideChar(CP_UTF8, 0, utf8Str,
                                       static_cast<int>(len), nullptr, 0);
    if (wideLen == 0) return L"";

    std::wstring wideStr(wideLen, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8Str, static_cast<int>(len),
                        &wideStr[0], wideLen);
    return wideStr;
}

// Convert ANSI to wide string
std::wstring ansiToWide(const char* ansiStr, size_t len) {
    if (len == 0) return L"";

    int wideLen = MultiByteToWideChar(CP_ACP, 0, ansiStr,
                                       static_cast<int>(len), nullptr, 0);
    if (wideLen == 0) return L"";

    std::wstring wideStr(wideLen, 0);
    MultiByteToWideChar(CP_ACP, 0, ansiStr, static_cast<int>(len),
                        &wideStr[0], wideLen);
    return wideStr;
}

// ============================================================================
// Clipboard Functions
// ============================================================================

bool openClipboard() {
    if (OpenClipboard(nullptr)) {
        return true;
    } else {
        MessageBox(
            nullptr,
            L"Failed to OpenClipboard.",
            L"MadPaster - Error",
            MB_OK | MB_ICONERROR | MB_TOPMOST
        );
        return false;
    }
}

void closeClipboard() {
    CloseClipboard();
}

std::wstring getClipboardText() {
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) {
        MessageBox(
            nullptr,
            L"Clipboard does not contain text.",
            L"MadPaster - Error",
            MB_OK | MB_ICONERROR | MB_TOPMOST
        );
        return L"";
    }

    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    if (hData == nullptr) {
        MessageBox(
            nullptr,
            L"Failed to get clipboard data.",
            L"MadPaster - Error",
            MB_OK | MB_ICONERROR | MB_TOPMOST
        );
        return L"";
    }

    wchar_t* pszText = static_cast<wchar_t*>(GlobalLock(hData));
    if (pszText == nullptr) {
        MessageBox(
            nullptr,
            L"Failed to lock clipboard data.",
            L"MadPaster - Error",
            MB_OK | MB_ICONERROR | MB_TOPMOST
        );
        return L"";
    }

    std::wstring text(pszText);
    GlobalUnlock(hData);
    return text;
}

// ============================================================================
// File Reading Functions
// ============================================================================

std::wstring readFileContents(const std::wstring& filePath, bool& success) {
    success = false;

    HANDLE hFile = CreateFileW(
        filePath.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );

    if (hFile == INVALID_HANDLE_VALUE) {
        DWORD error = GetLastError();
        std::wstring errorMsg = L"Failed to open file.\nError code: " +
                                 std::to_wstring(error) + L"\n\nFile: " + filePath;
        MessageBox(nullptr, errorMsg.c_str(), L"MadPaster - File Error",
                   MB_OK | MB_ICONERROR | MB_TOPMOST);
        return L"";
    }

    LARGE_INTEGER fileSize;
    if (!GetFileSizeEx(hFile, &fileSize)) {
        CloseHandle(hFile);
        MessageBox(nullptr, L"Failed to get file size.", L"MadPaster - File Error",
                   MB_OK | MB_ICONERROR | MB_TOPMOST);
        return L"";
    }

    const LONGLONG maxFileSize = 500 * 1024; // 500 KB
    if (fileSize.QuadPart > maxFileSize) {
        CloseHandle(hFile);
        MessageBox(nullptr, L"File too large.\nMaximum file size: 500KB",
                   L"MadPaster - File Error", MB_OK | MB_ICONERROR | MB_TOPMOST);
        return L"";
    }

    if (fileSize.QuadPart == 0) {
        CloseHandle(hFile);
        MessageBox(nullptr, L"File is empty.",
                   L"MadPaster - File Error", MB_OK | MB_ICONERROR | MB_TOPMOST);
        return L"";
    }

    std::vector<unsigned char> buffer(static_cast<size_t>(fileSize.QuadPart));
    DWORD bytesRead;
    if (!ReadFile(hFile, buffer.data(), static_cast<DWORD>(fileSize.QuadPart),
                  &bytesRead, nullptr)) {
        CloseHandle(hFile);
        MessageBox(nullptr, L"Failed to read file.", L"MadPaster - File Error",
                   MB_OK | MB_ICONERROR | MB_TOPMOST);
        return L"";
    }
    CloseHandle(hFile);

    FileEncoding encoding = detectEncoding(buffer);
    std::wstring result;

    switch (encoding) {
        case FileEncoding::UTF8_BOM:
            result = utf8ToWide(reinterpret_cast<char*>(buffer.data() + 3),
                               bytesRead - 3);
            break;

        case FileEncoding::UTF16_LE_BOM:
            result = std::wstring(
                reinterpret_cast<wchar_t*>(buffer.data() + 2),
                (bytesRead - 2) / sizeof(wchar_t)
            );
            break;

        case FileEncoding::UTF16_BE_BOM:
            {
                size_t charCount = (bytesRead - 2) / sizeof(wchar_t);
                result.resize(charCount);
                for (size_t i = 0; i < charCount; ++i) {
                    unsigned char hi = buffer[2 + i * 2];
                    unsigned char lo = buffer[2 + i * 2 + 1];
                    result[i] = static_cast<wchar_t>((hi << 8) | lo);
                }
            }
            break;

        case FileEncoding::ANSI_OR_UTF8:
        default:
            result = utf8ToWide(reinterpret_cast<char*>(buffer.data()),
                               bytesRead);
            if (result.empty() && bytesRead > 0) {
                result = ansiToWide(reinterpret_cast<char*>(buffer.data()),
                                   bytesRead);
            }
            break;
    }

    success = true;
    return result;
}

// Show file open dialog and return selected path
std::wstring showFileOpenDialog(HWND hwndOwner) {
    wchar_t filePath[MAX_PATH] = {0};

    OPENFILENAMEW ofn = {0};
    ofn.lStructSize = sizeof(OPENFILENAMEW);
    ofn.hwndOwner = hwndOwner;
    ofn.lpstrFilter = L"All Supported Files\0*.txt;*.bat;*.ps1;*.sh;*.json;*.xml;*.yaml;*.yml;*.ini;*.cfg;*.conf;*.log;*.md;*.py;*.js;*.ts;*.cpp;*.c;*.h;*.cs;*.java\0"
                      L"Text Files (*.txt)\0*.txt\0"
                      L"Script Files (*.bat;*.ps1;*.sh)\0*.bat;*.ps1;*.sh\0"
                      L"Config Files (*.json;*.xml;*.yaml;*.yml;*.ini;*.cfg;*.conf)\0*.json;*.xml;*.yaml;*.yml;*.ini;*.cfg;*.conf\0"
                      L"Code Files (*.py;*.js;*.ts;*.cpp;*.c;*.h;*.cs;*.java)\0*.py;*.js;*.ts;*.cpp;*.c;*.h;*.cs;*.java\0"
                      L"All Files (*.*)\0*.*\0";
    ofn.lpstrFile = filePath;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrTitle = L"Select file to send via MadPaster";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetOpenFileNameW(&ofn)) {
        return std::wstring(filePath);
    }

    return L"";
}

// ============================================================================
// Input Injection Subsystem
// ============================================================================

namespace inject {

// RAII guard for high-resolution timer (1ms instead of ~15.6ms default)
struct TimerResolutionGuard {
    TimerResolutionGuard() { timeBeginPeriod(1); }
    ~TimerResolutionGuard() { timeEndPeriod(1); }
};

// Information about detected remote client
struct RemoteClientInfo {
    bool isRemote;
    wchar_t className[256];
    HWND hwnd;
    DWORD threadId;
    DWORD processId;
    HKL keyboardLayout;
};

// Check if window class is a known remote client
bool IsKnownRemoteClass(const wchar_t* className) {
    if (!className || !className[0]) return false;

    for (int i = 0; REMOTE_WINDOW_CLASSES[i] != nullptr; i++) {
        if (_wcsicmp(className, REMOTE_WINDOW_CLASSES[i]) == 0) {
            return true;
        }
    }
    return false;
}

// Detect if foreground window is a remote client
RemoteClientInfo DetectRemoteClient() {
    RemoteClientInfo info = {};
    info.hwnd = GetForegroundWindow();

    if (!info.hwnd) {
        return info;
    }

    // Get window class name
    GetClassNameW(info.hwnd, info.className, 256);

    // Get thread/process info
    info.threadId = GetWindowThreadProcessId(info.hwnd, &info.processId);

    // Get keyboard layout for VK mapping
    info.keyboardLayout = GetKeyboardLayout(info.threadId);

    // Check if this is a known remote class
    info.isRemote = IsKnownRemoteClass(info.className);

    return info;
}

// Send modifier reset fence - releases all modifier keys
void ResetModifiers() {
    INPUT inputs[6] = {};

    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_LSHIFT;
    inputs[0].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = VK_RSHIFT;
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = VK_LCONTROL;
    inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = VK_RCONTROL;
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[4].type = INPUT_KEYBOARD;
    inputs[4].ki.wVk = VK_LMENU;
    inputs[4].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[5].type = INPUT_KEYBOARD;
    inputs[5].ki.wVk = VK_RMENU;
    inputs[5].ki.dwFlags = KEYEVENTF_KEYUP;

    SendInput(6, inputs, sizeof(INPUT));
}

// Forward declaration
void AppendCharacterInputs(std::vector<INPUT>& buffer, wchar_t c);

// VK/Scancode mapping result
struct VKMapping {
    bool success;
    BYTE vk;
    WORD scancode;
    bool needsShift;
};

// Map a character to VK code using VkKeyScanExW
// Only accepts "safe" mappings that require no modifiers or just Shift
// Rejects mappings that need Ctrl/Alt (would trigger shortcuts)
VKMapping MapCharacterToVK(wchar_t ch, HKL layout) {
    VKMapping result = {};

    SHORT vkResult = VkKeyScanExW(ch, layout);
    if (vkResult == -1) {
        // Character cannot be mapped to a VK code
        return result;
    }

    BYTE vk = LOBYTE(vkResult);
    BYTE modifiers = HIBYTE(vkResult);

    // Only accept no modifiers (0) or Shift only (1)
    // Reject Ctrl (2), Alt (4), or combinations
    if (modifiers > 1) {
        return result;
    }

    result.success = true;
    result.vk = vk;
    result.needsShift = (modifiers == 1);

    // Get hardware scancode for VK
    result.scancode = static_cast<WORD>(MapVirtualKeyW(vk, MAPVK_VK_TO_VSC));

    return result;
}

// Append character using VK code with scancode
// Returns number of INPUT events added (2 for simple char, 4 with Shift)
int AppendVKCharacterInputs(std::vector<INPUT>& buffer, wchar_t ch, HKL layout) {
    VKMapping mapping = MapCharacterToVK(ch, layout);
    if (!mapping.success) {
        return 0;  // Caller should fall back to Unicode
    }

    int eventsAdded = 0;

    // Press Shift if needed
    if (mapping.needsShift) {
        INPUT shiftDown = {};
        shiftDown.type = INPUT_KEYBOARD;
        shiftDown.ki.wVk = VK_SHIFT;
        shiftDown.ki.wScan = static_cast<WORD>(MapVirtualKeyW(VK_SHIFT, MAPVK_VK_TO_VSC));
        shiftDown.ki.dwFlags = KEYEVENTF_SCANCODE;
        buffer.push_back(shiftDown);
        eventsAdded++;
    }

    // Key down
    INPUT down = {};
    down.type = INPUT_KEYBOARD;
    down.ki.wVk = mapping.vk;
    down.ki.wScan = mapping.scancode;
    down.ki.dwFlags = KEYEVENTF_SCANCODE;
    buffer.push_back(down);
    eventsAdded++;

    // Key up
    INPUT up = {};
    up.type = INPUT_KEYBOARD;
    up.ki.wVk = mapping.vk;
    up.ki.wScan = mapping.scancode;
    up.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
    buffer.push_back(up);
    eventsAdded++;

    // Release Shift if pressed
    if (mapping.needsShift) {
        INPUT shiftUp = {};
        shiftUp.type = INPUT_KEYBOARD;
        shiftUp.ki.wVk = VK_SHIFT;
        shiftUp.ki.wScan = static_cast<WORD>(MapVirtualKeyW(VK_SHIFT, MAPVK_VK_TO_VSC));
        shiftUp.ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;
        buffer.push_back(shiftUp);
        eventsAdded++;
    }

    return eventsAdded;
}

// Append character using appropriate mode
// Returns true if character was added, false if skipped (should not happen)
bool AppendCharacterWithMode(std::vector<INPUT>& buffer, wchar_t ch,
                             InjectionMode mode, HKL layout) {
    switch (mode) {
        case InjectionMode::Unicode:
            AppendCharacterInputs(buffer, ch);
            return true;

        case InjectionMode::VKScancode: {
            int added = AppendVKCharacterInputs(buffer, ch, layout);
            if (added == 0) {
                // VK mapping failed - fall back to Unicode as last resort
                AppendCharacterInputs(buffer, ch);
            }
            return true;
        }

        case InjectionMode::Hybrid: {
            // Try VK first, fall back to Unicode
            int added = AppendVKCharacterInputs(buffer, ch, layout);
            if (added == 0) {
                AppendCharacterInputs(buffer, ch);
            }
            return true;
        }

        case InjectionMode::Auto:
        default:
            // Auto mode should be resolved before calling this
            // Default to Unicode
            AppendCharacterInputs(buffer, ch);
            return true;
    }
}

// Flush accumulated INPUT events - loops until ALL events are sent
// Returns true if all events were sent, false on unrecoverable failure
// Optional eventsSent pointer to track total events successfully sent
bool FlushInputs(std::vector<INPUT>& buffer, size_t* eventsSent = nullptr) {
    if (buffer.empty()) return true;

    UINT total = static_cast<UINT>(buffer.size());
    UINT offset = 0;
    int consecutiveFailures = 0;

    while (offset < total) {
        UINT remaining = total - offset;
        UINT sent = SendInput(remaining, buffer.data() + offset, sizeof(INPUT));

        if (sent > 0) {
            offset += sent;
            if (eventsSent) *eventsSent += sent;
            consecutiveFailures = 0;
        } else {
            // Complete failure - yield and retry
            consecutiveFailures++;
            if (consecutiveFailures >= MAX_RETRY_COUNT) {
                buffer.clear();
                return false;
            }
            Sleep(1);  // Real yield - allows target to drain input queue
        }
    }

    buffer.clear();
    return true;
}

// Forward declarations for pacing
struct DiagnosticState;

// Pacing configuration for injection
struct PacingConfig {
    PacingStrategy strategy;
    int perEventDelayMs;
    int perCharDelayMs;
    int lineStartGuardChars;
    int lineStartGuardMs;
    int baseKeystrokeDelayMs;  // From UI setting
};

// Get default pacing config based on target type
PacingConfig GetDefaultPacingConfig(bool isRemote) {
    PacingConfig config = {};
    config.perEventDelayMs = PER_EVENT_DELAY_MS;
    config.perCharDelayMs = PER_CHAR_DELAY_MS;
    config.lineStartGuardChars = LINE_START_GUARD_CHARS;
    config.lineStartGuardMs = LINE_START_GUARD_MS;
    config.baseKeystrokeDelayMs = g_app.keystrokeDelayMs;

    if (isRemote) {
        config.strategy = PacingStrategy::PerCharacter;
    } else {
        config.strategy = PacingStrategy::Burst;
    }

    return config;
}

// Flush with per-event pacing - sends events one at a time with delays
// Returns number of events successfully sent
size_t FlushInputsWithPacing(std::vector<INPUT>& buffer, const PacingConfig& config,
                             DiagnosticState* diag) {
    if (buffer.empty()) return 0;

    size_t sent = 0;
    int consecutiveFailures = 0;

    for (size_t i = 0; i < buffer.size(); i++) {
        UINT result = SendInput(1, &buffer[i], sizeof(INPUT));

        if (result > 0) {
            sent++;
            consecutiveFailures = 0;

            // Per-event delay
            if (config.strategy == PacingStrategy::PerEvent && config.perEventDelayMs > 0) {
                Sleep(config.perEventDelayMs);
            }
        } else {
            consecutiveFailures++;
            if (consecutiveFailures >= MAX_RETRY_COUNT) {
                break;  // Abort on repeated failures
            }
            Sleep(1);
            i--;  // Retry this event
        }
    }

    buffer.clear();
    return sent;
}

// Append character using KEYEVENTF_UNICODE (no modifiers involved)
void AppendCharacterInputs(std::vector<INPUT>& buffer, wchar_t c) {
    INPUT down = {};
    down.type = INPUT_KEYBOARD;
    down.ki.wScan = c;
    down.ki.dwFlags = KEYEVENTF_UNICODE;
    buffer.push_back(down);

    INPUT up = {};
    up.type = INPUT_KEYBOARD;
    up.ki.wScan = c;
    up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
    buffer.push_back(up);
}

// Send Enter key using hardware scancode for maximum compatibility
// Unicode CR/LF doesn't create line breaks in Scintilla-based editors
// Using KEYEVENTF_SCANCODE forces hardware-level input that Scintilla handles correctly
void SendEnterKey() {
    INPUT inputs[2] = {};

    // Key down - use scancode mode for hardware-level simulation
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = 0;       // Must be 0 when using KEYEVENTF_SCANCODE
    inputs[0].ki.wScan = 0x1C;  // Hardware scan code for Enter key
    inputs[0].ki.dwFlags = KEYEVENTF_SCANCODE;

    // Key up
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 0;
    inputs[1].ki.wScan = 0x1C;
    inputs[1].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;

    // Send both events atomically
    SendInput(2, inputs, sizeof(INPUT));
}

// Drain the input queue by yielding CPU time repeatedly
// This ensures the target app has time to process pending input before we continue
void DrainInputQueue() {
    // Multiple yields with longer sleeps to let the target process its message queue
    // SwitchToThread yields to any ready thread, Sleep(1) allows scheduler to run others
    for (int i = 0; i < 5; i++) {
        SwitchToThread();
        Sleep(2);
    }
}

// Get process ID of foreground window
DWORD GetForegroundProcessId() {
    HWND fg = GetForegroundWindow();
    if (!fg) return 0;
    DWORD pid = 0;
    GetWindowThreadProcessId(fg, &pid);
    return pid;
}

// Wait for target process to become idle (finished processing input)
// Returns true if idle or on error, false on timeout
bool WaitForTargetIdle(DWORD pid, DWORD maxWaitMs) {
    if (pid == 0) return true;

    HANDLE hProcess = OpenProcess(SYNCHRONIZE, FALSE, pid);
    if (!hProcess) return true;  // Can't open = assume ready

    DWORD result = WaitForInputIdle(hProcess, maxWaitMs);
    CloseHandle(hProcess);

    // 0 = success (idle), WAIT_TIMEOUT = timeout, WAIT_FAILED = error
    return (result != WAIT_TIMEOUT);
}

// Diagnostic state for injection debugging
struct DiagnosticState {
    size_t totalEventsAttempted;
    size_t totalEventsSent;
    size_t totalEventsFailed;
    size_t totalCharsSent;
    size_t totalCharsRequested;
    std::vector<std::pair<DWORD, std::wstring>> foregroundChanges;
    std::vector<std::wstring> errors;
    DWORD startTime;
    DWORD endTime;

    // Context info
    std::wstring injectionModeName;
    std::wstring targetClassName;
    bool targetIsRemote;

    DiagnosticState() : totalEventsAttempted(0), totalEventsSent(0),
                        totalEventsFailed(0), totalCharsSent(0),
                        totalCharsRequested(0), startTime(0), endTime(0),
                        targetIsRemote(false) {}

    void RecordForegroundChange(HWND hwnd) {
        wchar_t className[256] = {};
        if (hwnd) GetClassNameW(hwnd, className, 256);
        foregroundChanges.push_back({GetTickCount(), className});
    }

    void RecordError(const std::wstring& error) {
        errors.push_back(error);
    }

    std::wstring GetSummary(bool forMessageBox = false) {
        std::wstring summary;
        std::wstring nl = forMessageBox ? L"\n" : L"\r\n";

        if (!forMessageBox) {
            // Add timestamp for log file
            SYSTEMTIME st;
            GetLocalTime(&st);
            wchar_t timestamp[64];
            swprintf_s(timestamp, L"[%04d-%02d-%02d %02d:%02d:%02d]",
                st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
            summary += timestamp;
            summary += nl;
        }

        summary += L"MadPaster Injection Report" + nl;
        summary += L"─────────────────────────────" + nl;

        // Target info
        summary += L"Target: " + targetClassName;
        if (targetIsRemote) summary += L" (Remote)";
        summary += nl;

        summary += L"Mode: " + injectionModeName + nl;
        summary += nl;

        // Results
        summary += L"Characters: " + std::to_wstring(totalCharsSent) + L" / " +
                   std::to_wstring(totalCharsRequested);
        if (totalCharsSent == totalCharsRequested) {
            summary += L" ✓";
        } else {
            summary += L" (incomplete)";
        }
        summary += nl;

        summary += L"Events: " + std::to_wstring(totalEventsSent) + L" / " +
                   std::to_wstring(totalEventsAttempted) + L" sent" + nl;

        DWORD duration = endTime - startTime;
        summary += L"Duration: " + std::to_wstring(duration) + L" ms";
        if (duration > 0 && totalCharsSent > 0) {
            double cps = (double)totalCharsSent * 1000.0 / (double)duration;
            wchar_t cpsStr[32];
            swprintf_s(cpsStr, L" (%.1f chars/sec)", cps);
            summary += cpsStr;
        }
        summary += nl;

        // Issues
        if (!foregroundChanges.empty() || !errors.empty()) {
            summary += nl + L"Issues:" + nl;
            if (!foregroundChanges.empty()) {
                summary += L"  • Focus changed " + std::to_wstring(foregroundChanges.size()) +
                           L" time(s) during injection" + nl;
            }
            for (const auto& err : errors) {
                summary += L"  • " + err + nl;
            }
        }

        return summary;
    }
};

// Optional keyboard hook for diagnostic verification
// Counts how many injected events actually reach the system
static HHOOK g_diagHook = nullptr;
static volatile LONG g_hookEventCount = 0;

LRESULT CALLBACK DiagnosticKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0) {
        KBDLLHOOKSTRUCT* pKbd = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
        // Count injected events (LLKHF_INJECTED flag)
        if (pKbd->flags & LLKHF_INJECTED) {
            InterlockedIncrement(&g_hookEventCount);
        }
    }
    return CallNextHookEx(g_diagHook, nCode, wParam, lParam);
}

bool InstallDiagnosticHook() {
    if (g_diagHook) return true;  // Already installed

    g_hookEventCount = 0;
    g_diagHook = SetWindowsHookExW(WH_KEYBOARD_LL, DiagnosticKeyboardProc,
                                    GetModuleHandleW(nullptr), 0);
    return (g_diagHook != nullptr);
}

void RemoveDiagnosticHook() {
    if (g_diagHook) {
        UnhookWindowsHookEx(g_diagHook);
        g_diagHook = nullptr;
    }
}

size_t GetHookEventCount() {
    return static_cast<size_t>(InterlockedExchangeAdd(&g_hookEventCount, 0));
}

void ResetHookEventCount() {
    InterlockedExchange(&g_hookEventCount, 0);
}

// Low-level keyboard hook for abort detection
// Intercepts ESC at system level, works even when Citrix/RDP has focus
static HHOOK g_abortHook = nullptr;
static volatile LONG g_abortRequested = 0;

LRESULT CALLBACK AbortKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)) {
        KBDLLHOOKSTRUCT* pKbd = reinterpret_cast<KBDLLHOOKSTRUCT*>(lParam);
        // Check for ESC key (not injected by us)
        if (pKbd->vkCode == VK_ESCAPE && !(pKbd->flags & LLKHF_INJECTED)) {
            InterlockedExchange(&g_abortRequested, 1);
        }
    }
    return CallNextHookEx(g_abortHook, nCode, wParam, lParam);
}

bool InstallAbortHook() {
    if (g_abortHook) return true;  // Already installed

    g_abortRequested = 0;
    g_abortHook = SetWindowsHookExW(WH_KEYBOARD_LL, AbortKeyboardProc,
                                     GetModuleHandleW(nullptr), 0);
    return (g_abortHook != nullptr);
}

void RemoveAbortHook() {
    if (g_abortHook) {
        UnhookWindowsHookEx(g_abortHook);
        g_abortHook = nullptr;
    }
    g_abortRequested = 0;
}

bool IsAbortRequested() {
    return (InterlockedExchangeAdd(&g_abortRequested, 0) != 0);
}

void ResetAbortFlag() {
    InterlockedExchange(&g_abortRequested, 0);
}

} // namespace inject

// ============================================================================
// Keyboard Simulation
// ============================================================================

// Progress callback type for injection progress reporting
typedef void (*ProgressCallback)(size_t current, size_t total);

// Extended injection function with mode and pacing configuration
size_t sendTextToWindowEx(const std::wstring& text, InjectionMode mode,
                          const inject::PacingConfig& config,
                          inject::DiagnosticState* diag,
                          ProgressCallback progressCallback = nullptr) {
    // Enable high-resolution timer for precise Sleep() calls
    inject::TimerResolutionGuard timerGuard;

    // Install low-level keyboard hook for ESC detection (works even in Citrix/RDP)
    inject::InstallAbortHook();

    if (diag) {
        diag->startTime = GetTickCount();
    }

    // Detect remote client for keyboard layout
    inject::RemoteClientInfo clientInfo = inject::DetectRemoteClient();
    HKL layout = clientInfo.keyboardLayout;

    // Resolve Auto mode - default to Hybrid for best compatibility with remote sessions
    InjectionMode resolvedMode = mode;
    if (mode == InjectionMode::Auto) {
        resolvedMode = InjectionMode::Hybrid;
    }

    // Reset modifiers at start (clean slate)
    inject::ResetModifiers();

    std::vector<INPUT> buffer;
    buffer.reserve(16);  // Larger for VK mode with shift events

    size_t charsSent = 0;
    size_t charsInBuffer = 0;
    size_t charsSinceNewline = 0;  // For line-start guard

    for (size_t i = 0; i < text.size(); ++i) {
        // Check for ESC at chunk boundaries (using low-level hook for Citrix/RDP compatibility)
        if (buffer.empty() && inject::IsAbortRequested()) {
            inject::ResetModifiers();
            inject::RemoveAbortHook();
            if (diag) {
                diag->endTime = GetTickCount();
                diag->totalCharsSent = charsSent;
                diag->RecordError(L"User cancelled with ESC");
            }
            return charsSent;
        }

        wchar_t c = text[i];

        // Skip '\r' in CRLF sequences
        if (c == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n') {
            continue;
        }

        // Handle newlines
        if (c == L'\n' || c == L'\r') {
            // Flush any pending characters
            if (!buffer.empty()) {
                if (diag) diag->totalEventsAttempted += buffer.size();

                if (config.strategy == PacingStrategy::Burst) {
                    if (!inject::FlushInputs(buffer, diag ? &diag->totalEventsSent : nullptr)) {
                        inject::ResetModifiers();
                        inject::RemoveAbortHook();
                        if (diag) {
                            diag->endTime = GetTickCount();
                            diag->totalCharsSent = charsSent;
                            diag->RecordError(L"FlushInputs failed before newline");
                        }
                        return charsSent;
                    }
                } else {
                    size_t sent = inject::FlushInputsWithPacing(buffer, config, diag);
                    if (diag) diag->totalEventsSent += sent;
                }
                charsSent += charsInBuffer;
                charsInBuffer = 0;
                if (progressCallback) progressCallback(charsSent, text.size());
            }

            // Normalized newline handling - single unified pause
            inject::DrainInputQueue();
            Sleep(config.baseKeystrokeDelayMs + NEWLINE_PAUSE_MS / 2);

            // Send Enter key
            inject::SendEnterKey();
            charsSent++;
            charsSinceNewline = 0;  // Reset line-start counter
            if (progressCallback) progressCallback(charsSent, text.size());

            // Brief pause after enter
            Sleep(config.baseKeystrokeDelayMs + NEWLINE_PAUSE_MS / 2);
            inject::DrainInputQueue();
            continue;
        }

        // Accumulate character using appropriate mode
        size_t bufferSizeBefore = buffer.size();
        inject::AppendCharacterWithMode(buffer, c, resolvedMode, layout);
        charsInBuffer++;
        charsSinceNewline++;

        // Determine chunk size based on pacing strategy
        int effectiveChunkSize = (config.strategy == PacingStrategy::Burst) ? CHUNK_SIZE : 1;

        // Flush at chunk boundary
        if (charsInBuffer >= static_cast<size_t>(effectiveChunkSize)) {
            if (diag) diag->totalEventsAttempted += buffer.size();

            if (config.strategy == PacingStrategy::Burst) {
                if (!inject::FlushInputs(buffer, diag ? &diag->totalEventsSent : nullptr)) {
                    inject::ResetModifiers();
                    inject::RemoveAbortHook();
                    if (diag) {
                        diag->endTime = GetTickCount();
                        diag->totalCharsSent = charsSent;
                        diag->RecordError(L"FlushInputs failed");
                    }
                    return charsSent;
                }
            } else {
                size_t sent = inject::FlushInputsWithPacing(buffer, config, diag);
                if (diag) diag->totalEventsSent += sent;
            }

            charsSent += charsInBuffer;
            charsInBuffer = 0;
            if (progressCallback) progressCallback(charsSent, text.size());

            // Calculate pause
            int pauseMs = config.baseKeystrokeDelayMs;

            if (config.strategy == PacingStrategy::Burst) {
                pauseMs += INTER_CHUNK_PAUSE_MS;
            } else if (config.strategy == PacingStrategy::PerCharacter) {
                pauseMs += config.perCharDelayMs;
            }

            // Line-start guard: extra delay for first few chars after newline
            if (charsSinceNewline <= static_cast<size_t>(config.lineStartGuardChars)) {
                pauseMs += config.lineStartGuardMs;
            }

            if (pauseMs > 0) {
                Sleep(pauseMs);
            }

            if (config.strategy == PacingStrategy::Burst) {
                inject::DrainInputQueue();
            }
        }
    }

    // Flush remaining
    if (!buffer.empty()) {
        if (diag) diag->totalEventsAttempted += buffer.size();

        if (config.strategy == PacingStrategy::Burst) {
            if (!inject::FlushInputs(buffer, diag ? &diag->totalEventsSent : nullptr)) {
                inject::ResetModifiers();
                inject::RemoveAbortHook();
                if (diag) {
                    diag->endTime = GetTickCount();
                    diag->totalCharsSent = charsSent;
                    diag->RecordError(L"FlushInputs failed at end");
                }
                return charsSent;
            }
        } else {
            size_t sent = inject::FlushInputsWithPacing(buffer, config, diag);
            if (diag) diag->totalEventsSent += sent;
        }
        charsSent += charsInBuffer;
        if (progressCallback) progressCallback(charsSent, text.size());
    }

    // Reset modifiers at end
    inject::ResetModifiers();
    inject::RemoveAbortHook();

    if (diag) {
        diag->endTime = GetTickCount();
        diag->totalCharsSent = charsSent;
    }

    return charsSent;
}

// Forward declarations for diagnostic logging
std::wstring GetLogPath();
void WriteDiagnosticLog(const std::wstring& content);

// Forward declaration for progress callback
void UpdateProgress(size_t current, size_t total);

// Progress callback wrapper for UpdateProgress
void ProgressCallbackWrapper(size_t current, size_t total) {
    UpdateProgress(current, total);
}

// Original function - wrapper that auto-detects target and selects mode
// Normalize typographic Unicode characters to ASCII equivalents
// Prevents garbled output in remote desktop sessions (Citrix, RDP, VNC)
std::wstring NormalizeSmartCharacters(const std::wstring& input) {
    std::wstring result;
    result.reserve(input.size());
    for (wchar_t c : input) {
        switch (c) {
            case L'\u2018': // LEFT SINGLE QUOTATION MARK
            case L'\u2019': // RIGHT SINGLE QUOTATION MARK
                result += L'\'';
                break;
            case L'\u201C': // LEFT DOUBLE QUOTATION MARK
            case L'\u201D': // RIGHT DOUBLE QUOTATION MARK
                result += L'"';
                break;
            case L'\u2013': // EN DASH
            case L'\u2014': // EM DASH
                result += L'-';
                break;
            case L'\u2026': // HORIZONTAL ELLIPSIS
                result += L"...";
                break;
            default:
                result += c;
                break;
        }
    }
    return result;
}

size_t sendTextToWindow(const std::wstring& text, bool showProgress = false) {
    // Normalize smart quotes/dashes to ASCII for remote desktop compatibility
    std::wstring normalizedText = NormalizeSmartCharacters(text);

    // Detect remote client
    inject::RemoteClientInfo clientInfo = inject::DetectRemoteClient();

    // Get appropriate pacing config
    inject::PacingConfig config = inject::GetDefaultPacingConfig(clientInfo.isRemote);

    // Optional diagnostic state when enabled
    inject::DiagnosticState* diag = nullptr;
    inject::DiagnosticState diagState;
    if (g_app.diagnosticMode) {
        diag = &diagState;

        // Populate context info
        diag->totalCharsRequested = normalizedText.length();
        diag->targetClassName = clientInfo.className;
        diag->targetIsRemote = clientInfo.isRemote;

        // Set injection mode name
        InjectionMode effectiveMode = g_app.injectionMode;
        if (effectiveMode == InjectionMode::Auto) {
            effectiveMode = clientInfo.isRemote ? InjectionMode::Hybrid : InjectionMode::Unicode;
            diag->injectionModeName = L"Auto → ";
        }
        switch (effectiveMode) {
            case InjectionMode::Unicode:
                diag->injectionModeName += L"Unicode";
                break;
            case InjectionMode::VKScancode:
                diag->injectionModeName += L"VK Scancode";
                break;
            case InjectionMode::Hybrid:
                diag->injectionModeName += L"Hybrid";
                break;
            default:
                diag->injectionModeName += L"Auto";
                break;
        }
    }

    // Use configured injection mode (default: Auto)
    InjectionMode mode = g_app.injectionMode;

    // Set up progress callback if requested
    ProgressCallback progressCb = showProgress ? ProgressCallbackWrapper : nullptr;

    size_t result = sendTextToWindowEx(normalizedText, mode, config, diag, progressCb);

    // Log and display diagnostics if enabled
    if (diag) {
        // Write to debug output (for DebugView)
        OutputDebugStringW(diag->GetSummary().c_str());

        // Write to log file
        WriteDiagnosticLog(diag->GetSummary(false));

        // Show message box (use topmost so it appears over other windows)
        MessageBoxW(nullptr, diag->GetSummary(true).c_str(),
                    L"MadPaster - Injection Diagnostics", MB_OK | MB_ICONINFORMATION | MB_TOPMOST);
    }

    return result;
}

// ============================================================================
// Settings Persistence (INI File)
// ============================================================================

std::wstring GetIniPath() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);

    std::wstring iniPath(exePath);
    size_t pos = iniPath.rfind(L".exe");
    if (pos != std::wstring::npos) {
        iniPath.replace(pos, 4, L".ini");
    } else {
        iniPath += L".ini";
    }
    return iniPath;
}

std::wstring GetLogPath() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileNameW(NULL, exePath, MAX_PATH);

    std::wstring logPath(exePath);
    size_t pos = logPath.rfind(L".exe");
    if (pos != std::wstring::npos) {
        logPath.replace(pos, 4, L"-diag.log");
    } else {
        logPath += L"-diag.log";
    }
    return logPath;
}

void WriteDiagnosticLog(const std::wstring& content) {
    std::wstring logPath = GetLogPath();

    // Open file for append (create if doesn't exist)
    HANDLE hFile = CreateFileW(
        logPath.c_str(),
        FILE_APPEND_DATA,
        FILE_SHARE_READ,
        nullptr,
        OPEN_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );

    if (hFile == INVALID_HANDLE_VALUE) {
        return;  // Silently fail if can't write log
    }

    // Convert to UTF-8 for file
    int utf8Len = WideCharToMultiByte(CP_UTF8, 0, content.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (utf8Len > 0) {
        std::vector<char> utf8Buffer(utf8Len);
        WideCharToMultiByte(CP_UTF8, 0, content.c_str(), -1, utf8Buffer.data(), utf8Len, nullptr, nullptr);

        // Write without null terminator, add separator
        DWORD bytesWritten;
        WriteFile(hFile, utf8Buffer.data(), utf8Len - 1, &bytesWritten, nullptr);

        // Use ASCII separator to avoid UTF-8 encoding issues
        const char* separator = "\r\n========================================\r\n\r\n";
        WriteFile(hFile, separator, (DWORD)strlen(separator), &bytesWritten, nullptr);
    }

    CloseHandle(hFile);
}

// Convert string to InjectionMode enum
InjectionMode ParseInjectionMode(const wchar_t* str) {
    if (_wcsicmp(str, L"unicode") == 0) return InjectionMode::Unicode;
    if (_wcsicmp(str, L"vk") == 0) return InjectionMode::VKScancode;
    if (_wcsicmp(str, L"hybrid") == 0) return InjectionMode::Hybrid;
    return InjectionMode::Auto;
}

// Convert InjectionMode to string
const wchar_t* InjectionModeToString(InjectionMode mode) {
    switch (mode) {
        case InjectionMode::Unicode: return L"unicode";
        case InjectionMode::VKScancode: return L"vk";
        case InjectionMode::Hybrid: return L"hybrid";
        case InjectionMode::Auto:
        default: return L"auto";
    }
}

void LoadSettings() {
    std::wstring iniPath = GetIniPath();

    g_app.delaySeconds = GetPrivateProfileIntW(L"Settings", L"Delay", 5, iniPath.c_str());
    if (g_app.delaySeconds < 0) g_app.delaySeconds = 0;
    if (g_app.delaySeconds > 60) g_app.delaySeconds = 60;

    g_app.keystrokeDelayMs = GetPrivateProfileIntW(L"Settings", L"KeystrokeDelay", 3, iniPath.c_str());
    if (g_app.keystrokeDelayMs < 0) g_app.keystrokeDelayMs = 0;
    if (g_app.keystrokeDelayMs > 100) g_app.keystrokeDelayMs = 100;

    wchar_t mode[32];
    GetPrivateProfileStringW(L"Settings", L"Mode", L"clipboard", mode, 32, iniPath.c_str());
    g_app.useClipboard = (wcscmp(mode, L"file") != 0);

    wchar_t filePath[MAX_PATH];
    GetPrivateProfileStringW(L"Settings", L"LastFilePath", L"", filePath, MAX_PATH, iniPath.c_str());
    g_app.selectedFilePath = filePath;

    // Load injection mode
    wchar_t injMode[32];
    GetPrivateProfileStringW(L"Settings", L"InjectionMode", L"auto", injMode, 32, iniPath.c_str());
    g_app.injectionMode = ParseInjectionMode(injMode);

    // Diagnostic mode (default off, usually set via CLI)
    g_app.diagnosticMode = (GetPrivateProfileIntW(L"Settings", L"DiagnosticMode", 0, iniPath.c_str()) != 0);

    // Silent mode (stay in tray after hotkey paste)
    g_app.silentMode = (GetPrivateProfileIntW(L"Settings", L"SilentMode", 0, iniPath.c_str()) != 0);
}

void SaveSettings() {
    std::wstring iniPath = GetIniPath();

    WritePrivateProfileStringW(L"Settings", L"Delay",
        std::to_wstring(g_app.delaySeconds).c_str(), iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"KeystrokeDelay",
        std::to_wstring(g_app.keystrokeDelayMs).c_str(), iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"Mode",
        g_app.useClipboard ? L"clipboard" : L"file", iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"LastFilePath",
        g_app.selectedFilePath.c_str(), iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"InjectionMode",
        InjectionModeToString(g_app.injectionMode), iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"DiagnosticMode",
        g_app.diagnosticMode ? L"1" : L"0", iniPath.c_str());
    WritePrivateProfileStringW(L"Settings", L"SilentMode",
        g_app.silentMode ? L"1" : L"0", iniPath.c_str());
}

// ============================================================================
// System Tray Functions
// ============================================================================

void CreateTrayIcon(HWND hwnd) {
    g_app.nid.cbSize = sizeof(NOTIFYICONDATA);
    g_app.nid.hWnd = hwnd;
    g_app.nid.uID = IDI_TRAY;
    g_app.nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_app.nid.uCallbackMessage = WM_TRAYICON;
    g_app.nid.hIcon = g_app.hAppIcon ? g_app.hAppIcon : LoadIcon(NULL, IDI_APPLICATION);
    wcscpy_s(g_app.nid.szTip, L"MadPaster");
    Shell_NotifyIcon(NIM_ADD, &g_app.nid);
}

void RemoveTrayIcon() {
    Shell_NotifyIcon(NIM_DELETE, &g_app.nid);
}

void ShowTrayMenu(HWND hwnd) {
    POINT pt;
    GetCursorPos(&pt);

    HMENU hMenu = CreatePopupMenu();

    // Build ARM menu item text with current settings info
    std::wstring armText = L"ARM Now";
    if (g_app.useClipboard) {
        armText += L" (Clipboard";
    } else {
        armText += L" (File";
    }
    if (g_app.delaySeconds > 0) {
        armText += L", " + std::to_wstring(g_app.delaySeconds) + L"s";
    }
    armText += L")";

    AppendMenuW(hMenu, MF_STRING, IDM_TRAY_ARM, armText.c_str());
    AppendMenuW(hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(hMenu, MF_STRING, IDM_TRAY_SHOW, L"Show Window");
    AppendMenuW(hMenu, MF_STRING, IDM_TRAY_EXIT, L"Exit");

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

void MinimizeToTray() {
    ShowWindow(g_app.hwndMain, SW_HIDE);
    g_app.minimizedToTray = true;
}

void RestoreFromTray() {
    HWND hwnd = g_app.hwndMain;

    // Attach to foreground thread's input queue for reliable focus
    HWND hwndFg = GetForegroundWindow();
    DWORD fgThread = hwndFg ? GetWindowThreadProcessId(hwndFg, NULL) : 0;
    DWORD myThread = GetCurrentThreadId();
    bool attached = (fgThread && fgThread != myThread)
                    ? AttachThreadInput(myThread, fgThread, TRUE) : false;

    // Restore and bring to front
    ShowWindow(hwnd, SW_RESTORE);
    BringWindowToTop(hwnd);
    SetForegroundWindow(hwnd);

    // HWND_TOPMOST trick to force foreground
    SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

    // Detach and set focus
    if (attached) AttachThreadInput(myThread, fgThread, FALSE);
    SetFocus(hwnd);

    g_app.minimizedToTray = false;
}

// ============================================================================
// ARM Functions
// ============================================================================

void UpdateArmButtonText() {
    wchar_t text[64];
    if (g_app.isArmed && g_app.countdownRemaining > 0) {
        swprintf_s(text, 64, L"ARMED (%d) - Click to Cancel", g_app.countdownRemaining);
    } else if (g_app.isArmed) {
        wcscpy_s(text, L"Executing...");
    } else {
        wcscpy_s(text, L"ARM");
    }
    SetWindowTextW(g_app.hwndButtonArm, text);
    InvalidateRect(g_app.hwndButtonArm, NULL, TRUE);  // Force redraw for owner-drawn button
}

void UpdateStatus(const wchar_t* status) {
    std::wstring fullStatus = L"Status: ";
    fullStatus += status;
    SetWindowTextW(g_app.hwndStaticStatus, fullStatus.c_str());
}

// Floating progress window procedure
LRESULT CALLBACK FloatingProgressProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            RECT rect;
            GetClientRect(hwnd, &rect);
            // Fill with dark gray background
            HBRUSH hBrush = CreateSolidBrush(RGB(45, 45, 48));
            FillRect(hdc, &rect, hBrush);
            DeleteObject(hBrush);
            // Draw border
            HPEN hPen = CreatePen(PS_SOLID, 1, RGB(80, 80, 80));
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
            Rectangle(hdc, rect.left, rect.top, rect.right, rect.bottom);
            SelectObject(hdc, hOldPen);
            SelectObject(hdc, hOldBrush);
            DeleteObject(hPen);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_CTLCOLORSTATIC: {
            HDC hdcStatic = (HDC)wParam;
            SetTextColor(hdcStatic, RGB(255, 255, 255));
            SetBkColor(hdcStatic, RGB(45, 45, 48));
            static HBRUSH hBrushStatic = CreateSolidBrush(RGB(45, 45, 48));
            return (LRESULT)hBrushStatic;
        }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void CreateFloatingProgressWindow() {
    if (g_app.hwndFloatingProgress) return;  // Already created

    // Register the floating progress window class
    WNDCLASSEXW wcFloat = {};
    wcFloat.cbSize = sizeof(WNDCLASSEXW);
    wcFloat.lpfnWndProc = FloatingProgressProc;
    wcFloat.hInstance = g_app.hInstance;
    wcFloat.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcFloat.lpszClassName = FLOATING_PROGRESS_CLASS;
    RegisterClassExW(&wcFloat);

    // Get screen dimensions for centering
    int screenWidth = GetSystemMetrics(SM_CXSCREEN);
    int screenHeight = GetSystemMetrics(SM_CYSCREEN);
    int x = (screenWidth - FLOATING_PROGRESS_WIDTH) / 2;
    int y = (screenHeight - FLOATING_PROGRESS_HEIGHT) / 2;

    // Create the floating window (popup, always on top, no taskbar entry)
    g_app.hwndFloatingProgress = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        FLOATING_PROGRESS_CLASS,
        L"MadPaster",
        WS_POPUP,
        x, y, FLOATING_PROGRESS_WIDTH, FLOATING_PROGRESS_HEIGHT,
        NULL, NULL, g_app.hInstance, NULL);

    // Create progress bar inside the floating window
    g_app.hwndFloatingProgressBar = CreateWindowW(PROGRESS_CLASSW, NULL,
        WS_CHILD | WS_VISIBLE | PBS_SMOOTH,
        10, 10, FLOATING_PROGRESS_WIDTH - 20, 20,
        g_app.hwndFloatingProgress, NULL, g_app.hInstance, NULL);
    SendMessageW(g_app.hwndFloatingProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

    // Create "Press ESC to cancel" label
    g_app.hwndFloatingLabel = CreateWindowW(L"STATIC", L"Press ESC to cancel",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        10, 40, FLOATING_PROGRESS_WIDTH - 20, 20,
        g_app.hwndFloatingProgress, NULL, g_app.hInstance, NULL);
    if (g_app.hFontUI) {
        SendMessageW(g_app.hwndFloatingLabel, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);
    }
}

// Progress bar helper functions (now uses floating window)
void ShowProgress() {
    CreateFloatingProgressWindow();
    if (g_app.hwndFloatingProgress) {
        SendMessageW(g_app.hwndFloatingProgressBar, PBM_SETPOS, 0, 0);
        ShowWindow(g_app.hwndFloatingProgress, SW_SHOWNOACTIVATE);
    }
    // Also update embedded progress bar
    if (g_app.hwndProgress) {
        SendMessageW(g_app.hwndProgress, PBM_SETPOS, 0, 0);
        ShowWindow(g_app.hwndProgress, SW_SHOW);
    }
}

void HideProgress() {
    if (g_app.hwndFloatingProgress) {
        ShowWindow(g_app.hwndFloatingProgress, SW_HIDE);
        SendMessageW(g_app.hwndFloatingProgressBar, PBM_SETPOS, 0, 0);
    }
    if (g_app.hwndProgress) {
        ShowWindow(g_app.hwndProgress, SW_HIDE);
        SendMessageW(g_app.hwndProgress, PBM_SETPOS, 0, 0);
    }
}

void UpdateProgress(size_t current, size_t total) {
    if (total > 0) {
        int percent = static_cast<int>((current * 100) / total);
        // Update floating progress bar
        if (g_app.hwndFloatingProgressBar) {
            SendMessageW(g_app.hwndFloatingProgressBar, PBM_SETPOS, percent, 0);
        }
        // Update embedded progress bar
        if (g_app.hwndProgress) {
            SendMessageW(g_app.hwndProgress, PBM_SETPOS, percent, 0);
        }
    }
    // Pump messages to keep UI responsive
    MSG msg;
    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
}

void ResetArmState() {
    g_app.isArmed = false;
    g_app.countdownRemaining = 0;

    KillTimer(g_app.hwndMain, IDT_COUNTDOWN);

    // Re-enable controls
    EnableWindow(g_app.hwndRadioClipboard, TRUE);
    EnableWindow(g_app.hwndRadioFile, TRUE);
    EnableWindow(g_app.hwndButtonBrowse, !g_app.useClipboard);
    EnableWindow(g_app.hwndEditDelay, TRUE);
    EnableWindow(g_app.hwndSpinDelay, TRUE);
    EnableWindow(g_app.hwndEditKeystroke, TRUE);
    EnableWindow(g_app.hwndSpinKeystroke, TRUE);
    EnableWindow(g_app.hwndComboMode, TRUE);
    EnableWindow(g_app.hwndCheckDiag, TRUE);
    EnableWindow(g_app.hwndCheckSilent, TRUE);

    UpdateArmButtonText();
    UpdateStatus(L"Ready - ARM Starts MadPaster  ESC Interrupts MadPaster");
}

void ExecutePaste() {
    UpdateStatus(L"Executing...");
    UpdateArmButtonText();

    // Minimize to tray before pasting
    MinimizeToTray();

    // Wait for focus to stabilize on target window
    // Requires 3 consecutive checks where a different window has focus
    HWND hwndSelf = g_app.hwndMain;
    int stableCount = 0;
    DWORD startTime = GetTickCount();
    const DWORD FOCUS_TIMEOUT_MS = 1000;  // Max wait time

    while (stableCount < 3) {
        Sleep(50);
        HWND hwndFg = GetForegroundWindow();
        if (hwndFg != hwndSelf && hwndFg != NULL) {
            stableCount++;
        } else {
            stableCount = 0;
        }
        // Timeout check
        if (GetTickCount() - startTime > FOCUS_TIMEOUT_MS) {
            break;
        }
    }

    // Get text content
    std::wstring textContent;
    bool success = false;

    if (g_app.useClipboard) {
        if (openClipboard()) {
            textContent = getClipboardText();
            closeClipboard();
            success = !textContent.empty();
        }
    } else {
        // Verify file still exists
        DWORD attrs = GetFileAttributesW(g_app.selectedFilePath.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES) {
            MessageBox(NULL, L"Selected file no longer exists.", L"MadPaster - Error",
                       MB_OK | MB_ICONERROR | MB_TOPMOST);
            ResetArmState();
            return;
        }
        textContent = readFileContents(g_app.selectedFilePath, success);
    }

    if (success && !textContent.empty()) {
        if (textContent.length() < static_cast<size_t>(maxchar)) {
            ShowProgress();
            size_t charsSent = sendTextToWindow(textContent, true);
            HideProgress();
            if (charsSent < textContent.length()) {
                // User pressed ESC - restore window and show progress
                RestoreFromTray();
                ResetArmState();
                std::wstring msg = L"Interrupted at " + std::to_wstring(charsSent) +
                                   L" / " + std::to_wstring(textContent.length()) + L" characters";
                UpdateStatus(msg.c_str());
                return;
            }
        } else {
            std::wstring message = L"Text exceeds maximum length (" +
                std::to_wstring(maxchar) + L" characters).\n\nCurrent length: " +
                std::to_wstring(textContent.length()) + L" characters.";
            MessageBox(NULL, message.c_str(), L"MadPaster - Error",
                MB_OK | MB_ICONWARNING | MB_TOPMOST);
        }
    }

    HideProgress();
    ResetArmState();
}

// Execute immediate paste from global hotkey (CTRL+ALT+V)
void ExecuteImmediatePaste() {
    // Don't interrupt if already armed
    if (g_app.isArmed) return;

    // Get text content based on mode
    std::wstring text;

    if (g_app.useClipboard) {
        if (!openClipboard()) return;
        text = getClipboardText();
        closeClipboard();
    } else {
        if (g_app.selectedFilePath.empty()) return;
        DWORD attrs = GetFileAttributesW(g_app.selectedFilePath.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) return;
        bool success = false;
        text = readFileContents(g_app.selectedFilePath, success);
        if (!success) return;
    }

    if (text.empty()) return;
    if (text.length() >= static_cast<size_t>(maxchar)) {
        MessageBox(NULL, L"Text exceeds maximum length.",
                   L"MadPaster - Error", MB_OK | MB_ICONWARNING | MB_TOPMOST);
        return;
    }

    // Minimize to tray before pasting
    MinimizeToTray();

    // Wait for focus to stabilize on target window
    HWND hwndSelf = g_app.hwndMain;
    int stableCount = 0;
    DWORD startTime = GetTickCount();
    const DWORD FOCUS_TIMEOUT_MS = 1000;

    while (stableCount < 3) {
        Sleep(50);
        HWND hwndFg = GetForegroundWindow();
        if (hwndFg != hwndSelf && hwndFg != NULL) {
            stableCount++;
        } else {
            stableCount = 0;
        }
        if (GetTickCount() - startTime > FOCUS_TIMEOUT_MS) {
            break;
        }
    }

    // Show progress bar and inject with ESC handling enabled
    ShowProgress();
    size_t charsSent = sendTextToWindow(text, true);
    HideProgress();

    // Handle interruption
    if (charsSent < text.length()) {
        RestoreFromTray();
        std::wstring msg = L"Interrupted at " + std::to_wstring(charsSent) +
                           L" / " + std::to_wstring(text.length()) + L" characters";
        UpdateStatus(msg.c_str());
        return;
    }

    // Restore from tray after successful paste (unless silent mode)
    if (!g_app.silentMode) {
        RestoreFromTray();
    }
}

void StartArmCountdown() {
    // Get current delay value from control
    wchar_t delayStr[16];
    GetWindowTextW(g_app.hwndEditDelay, delayStr, 16);
    g_app.delaySeconds = _wtoi(delayStr);
    if (g_app.delaySeconds < 0) g_app.delaySeconds = 0;
    if (g_app.delaySeconds > 60) g_app.delaySeconds = 60;

    // Get current keystroke delay value from control
    wchar_t keystrokeStr[16];
    GetWindowTextW(g_app.hwndEditKeystroke, keystrokeStr, 16);
    g_app.keystrokeDelayMs = _wtoi(keystrokeStr);
    if (g_app.keystrokeDelayMs < 0) g_app.keystrokeDelayMs = 0;
    if (g_app.keystrokeDelayMs > 100) g_app.keystrokeDelayMs = 100;

    // Validate: if file mode, check that a file is selected
    if (!g_app.useClipboard) {
        if (g_app.selectedFilePath.empty()) {
            MessageBox(g_app.hwndMain, L"Please select a file first.",
                       L"MadPaster", MB_OK | MB_ICONWARNING);
            return;
        }
        // Check file exists
        DWORD attrs = GetFileAttributesW(g_app.selectedFilePath.c_str());
        if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) {
            MessageBox(g_app.hwndMain, L"Selected file does not exist.",
                       L"MadPaster", MB_OK | MB_ICONWARNING);
            return;
        }
    }

    // Save settings when ARMing
    SaveSettings();

    g_app.isArmed = true;
    g_app.countdownRemaining = g_app.delaySeconds;

    // Disable controls
    EnableWindow(g_app.hwndRadioClipboard, FALSE);
    EnableWindow(g_app.hwndRadioFile, FALSE);
    EnableWindow(g_app.hwndButtonBrowse, FALSE);
    EnableWindow(g_app.hwndEditDelay, FALSE);
    EnableWindow(g_app.hwndSpinDelay, FALSE);
    EnableWindow(g_app.hwndEditKeystroke, FALSE);
    EnableWindow(g_app.hwndSpinKeystroke, FALSE);
    EnableWindow(g_app.hwndComboMode, FALSE);
    EnableWindow(g_app.hwndCheckDiag, FALSE);
    EnableWindow(g_app.hwndCheckSilent, FALSE);

    if (g_app.delaySeconds > 0) {
        UpdateArmButtonText();
        UpdateStatus(L"Armed - switch to target window!");
        SetTimer(g_app.hwndMain, IDT_COUNTDOWN, 1000, NULL);
    } else {
        ExecutePaste();
    }
}

void CancelArm() {
    ResetArmState();
    UpdateStatus(L"Cancelled");
}

// ============================================================================
// Logo Painting
// ============================================================================

void PaintLogo(HWND hwnd) {
    if (!g_app.pLogoImage || g_app.pLogoImage->GetLastStatus() != Gdiplus::Ok) {
        return;
    }

    RECT rect;
    GetClientRect(hwnd, &rect);

    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);

    Gdiplus::Graphics graphics(hdc);
    graphics.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
    graphics.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);

    // Draw image scaled to fit the control
    int width = rect.right - rect.left;
    int height = rect.bottom - rect.top;
    graphics.DrawImage(g_app.pLogoImage, 0, 0, width, height);

    EndPaint(hwnd, &ps);
}

// Subclass procedure for logo static control
LRESULT CALLBACK LogoProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                          UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    if (msg == WM_PAINT) {
        PaintLogo(hwnd);
        return 0;
    } else if (msg == WM_NCDESTROY) {
        RemoveWindowSubclass(hwnd, LogoProc, uIdSubclass);
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

// ============================================================================
// Window Procedure
// ============================================================================

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE: {
            HINSTANCE hInst = g_app.hInstance;

            // Load logo image
            wchar_t logoPath[MAX_PATH];
            GetModuleFileNameW(NULL, logoPath, MAX_PATH);
            std::wstring logoPathStr(logoPath);
            size_t pos = logoPathStr.rfind(L"\\");
            if (pos != std::wstring::npos) {
                logoPathStr = logoPathStr.substr(0, pos + 1) + L"MadPaster.png";
            }
            g_app.pLogoImage = Gdiplus::Image::FromFile(logoPathStr.c_str());

            // Create custom fonts
            g_app.hFontUI = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");

            g_app.hFontMono = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");

            g_app.hFontButton = CreateFontW(-16, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");

            // Logo display area (top right)
            g_app.hwndLogo = CreateWindowW(L"STATIC", L"",
                WS_CHILD | WS_VISIBLE | SS_NOTIFY,
                240, 10, 136, 136, hwnd, NULL, hInst, NULL);
            SetWindowSubclass(g_app.hwndLogo, LogoProc, 0, 0);

            // Source group label
            HWND hwndSourceLabel = CreateWindowW(L"STATIC", L"Source:",
                WS_CHILD | WS_VISIBLE,
                24, 20, 60, 20, hwnd, NULL, hInst, NULL);
            SendMessageW(hwndSourceLabel, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Clipboard radio button
            g_app.hwndRadioClipboard = CreateWindowW(L"BUTTON", L"Clipboard",
                WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP,
                40, 44, 100, 20, hwnd, (HMENU)IDC_RADIO_CLIPBOARD, hInst, NULL);
            SendMessageW(g_app.hwndRadioClipboard, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // File radio button
            g_app.hwndRadioFile = CreateWindowW(L"BUTTON", L"File:",
                WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                40, 70, 60, 20, hwnd, (HMENU)IDC_RADIO_FILE, hInst, NULL);
            SendMessageW(g_app.hwndRadioFile, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Browse button
            g_app.hwndButtonBrowse = CreateWindowW(L"BUTTON", L"Browse...",
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                105, 67, 90, 26, hwnd, (HMENU)IDC_BUTTON_BROWSE, hInst, NULL);
            SendMessageW(g_app.hwndButtonBrowse, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // File path display (shortened to not overlap logo)
            g_app.hwndStaticFilePath = CreateWindowW(L"STATIC", L"(no file selected)",
                WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP | SS_PATHELLIPSIS,
                40, 100, 190, 16, hwnd, (HMENU)IDC_STATIC_FILEPATH, hInst, NULL);
            SendMessageW(g_app.hwndStaticFilePath, WM_SETFONT, (WPARAM)g_app.hFontMono, TRUE);

            // Delay label
            HWND hwndDelayLabel = CreateWindowW(L"STATIC", L"Delay (seconds):",
                WS_CHILD | WS_VISIBLE,
                24, 140, 120, 20, hwnd, NULL, hInst, NULL);
            SendMessageW(hwndDelayLabel, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Delay edit box
            g_app.hwndEditDelay = CreateWindowW(L"EDIT", L"5",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_RIGHT,
                175, 137, 60, 26, hwnd, (HMENU)IDC_EDIT_DELAY, hInst, NULL);
            SendMessageW(g_app.hwndEditDelay, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Spin control for delay
            g_app.hwndSpinDelay = CreateWindowW(UPDOWN_CLASSW, NULL,
                WS_CHILD | WS_VISIBLE | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS,
                0, 0, 0, 0, hwnd, (HMENU)IDC_SPIN_DELAY, hInst, NULL);
            SendMessageW(g_app.hwndSpinDelay, UDM_SETBUDDY, (WPARAM)g_app.hwndEditDelay, 0);
            SendMessageW(g_app.hwndSpinDelay, UDM_SETRANGE32, 0, 60);

            // Keystroke delay label
            HWND hwndKeystrokeLabel = CreateWindowW(L"STATIC", L"Keystroke Delay (ms):",
                WS_CHILD | WS_VISIBLE,
                24, 167, 165, 20, hwnd, NULL, hInst, NULL);
            SendMessageW(hwndKeystrokeLabel, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Keystroke delay edit box
            g_app.hwndEditKeystroke = CreateWindowW(L"EDIT", L"3",
                WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER | ES_RIGHT,
                175, 164, 60, 26, hwnd, (HMENU)IDC_EDIT_KEYSTROKE, hInst, NULL);
            SendMessageW(g_app.hwndEditKeystroke, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Spin control for keystroke delay
            g_app.hwndSpinKeystroke = CreateWindowW(UPDOWN_CLASSW, NULL,
                WS_CHILD | WS_VISIBLE | UDS_SETBUDDYINT | UDS_ALIGNRIGHT | UDS_ARROWKEYS,
                0, 0, 0, 0, hwnd, (HMENU)IDC_SPIN_KEYSTROKE, hInst, NULL);
            SendMessageW(g_app.hwndSpinKeystroke, UDM_SETBUDDY, (WPARAM)g_app.hwndEditKeystroke, 0);
            SendMessageW(g_app.hwndSpinKeystroke, UDM_SETRANGE32, 0, 100);

            // Injection mode label
            HWND hwndModeLabel = CreateWindowW(L"STATIC", L"Injection Mode:",
                WS_CHILD | WS_VISIBLE,
                24, 197, 110, 20, hwnd, NULL, hInst, NULL);
            SendMessageW(hwndModeLabel, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Injection mode combo box
            g_app.hwndComboMode = CreateWindowW(L"COMBOBOX", NULL,
                WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                140, 194, 95, 120, hwnd, (HMENU)IDC_COMBO_MODE, hInst, NULL);
            SendMessageW(g_app.hwndComboMode, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);
            SendMessageW(g_app.hwndComboMode, CB_ADDSTRING, 0, (LPARAM)L"Auto");
            SendMessageW(g_app.hwndComboMode, CB_ADDSTRING, 0, (LPARAM)L"Unicode");
            SendMessageW(g_app.hwndComboMode, CB_ADDSTRING, 0, (LPARAM)L"VK Scancode");
            SendMessageW(g_app.hwndComboMode, CB_ADDSTRING, 0, (LPARAM)L"Hybrid");

            // Diagnostic mode checkbox
            g_app.hwndCheckDiag = CreateWindowW(L"BUTTON", L"Diagnostics",
                WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                250, 196, 110, 20, hwnd, (HMENU)IDC_CHECK_DIAG, hInst, NULL);
            SendMessageW(g_app.hwndCheckDiag, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Silent mode checkbox (stay in tray after hotkey paste)
            g_app.hwndCheckSilent = CreateWindowW(L"BUTTON", L"Silent (stay in tray after paste)",
                WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                24, 218, 340, 20, hwnd, (HMENU)IDC_CHECK_SILENT, hInst, NULL);
            SendMessageW(g_app.hwndCheckSilent, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // ARM button (large, prominent, owner-drawn)
            g_app.hwndButtonArm = CreateWindowW(L"BUTTON", L"ARM",
                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                24, 254, 352, 50, hwnd, (HMENU)IDC_BUTTON_ARM, hInst, NULL);

            // Progress bar (hidden by default, shown during paste)
            g_app.hwndProgress = CreateWindowW(PROGRESS_CLASSW, NULL,
                WS_CHILD | PBS_SMOOTH,
                24, 309, 352, 20, hwnd, (HMENU)IDC_PROGRESS, hInst, NULL);
            SendMessageW(g_app.hwndProgress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

            // Status label
            g_app.hwndStaticStatus = CreateWindowW(L"STATIC", L"Status: Ready - ARM Starts MadPaster  ESC Interrupts MadPaster",
                WS_CHILD | WS_VISIBLE,
                24, 334, 352, 20, hwnd, (HMENU)IDC_STATIC_STATUS, hInst, NULL);
            SendMessageW(g_app.hwndStaticStatus, WM_SETFONT, (WPARAM)g_app.hFontUI, TRUE);

            // Apply saved settings to controls
            SendMessageW(g_app.useClipboard ? g_app.hwndRadioClipboard : g_app.hwndRadioFile,
                BM_SETCHECK, BST_CHECKED, 0);
            SetWindowTextW(g_app.hwndEditDelay, std::to_wstring(g_app.delaySeconds).c_str());
            SetWindowTextW(g_app.hwndEditKeystroke, std::to_wstring(g_app.keystrokeDelayMs).c_str());
            EnableWindow(g_app.hwndButtonBrowse, !g_app.useClipboard);

            if (!g_app.selectedFilePath.empty()) {
                SetWindowTextW(g_app.hwndStaticFilePath, g_app.selectedFilePath.c_str());
            }

            // Apply injection mode setting to combo box
            int modeIndex = 0;  // Auto
            switch (g_app.injectionMode) {
                case InjectionMode::Auto: modeIndex = 0; break;
                case InjectionMode::Unicode: modeIndex = 1; break;
                case InjectionMode::VKScancode: modeIndex = 2; break;
                case InjectionMode::Hybrid: modeIndex = 3; break;
            }
            SendMessageW(g_app.hwndComboMode, CB_SETCURSEL, modeIndex, 0);

            // Apply diagnostic mode setting
            SendMessageW(g_app.hwndCheckDiag, BM_SETCHECK,
                g_app.diagnosticMode ? BST_CHECKED : BST_UNCHECKED, 0);

            // Apply silent mode setting
            SendMessageW(g_app.hwndCheckSilent, BM_SETCHECK,
                g_app.silentMode ? BST_CHECKED : BST_UNCHECKED, 0);

            // Create tray icon
            CreateTrayIcon(hwnd);

            // Register global hotkey (CTRL+ALT+V)
            if (!RegisterHotKey(hwnd, IDH_PASTE_HOTKEY, MOD_CONTROL | MOD_ALT | MOD_NOREPEAT, 'V')) {
                DWORD err = GetLastError();
                wchar_t msg[128];
                swprintf_s(msg, L"Failed to register CTRL+ALT+V hotkey (error %lu). Another app may have it.", err);
                MessageBoxW(hwnd, msg, L"MadPaster - Warning", MB_OK | MB_ICONWARNING);
            }

            break;
        }

        case WM_HOTKEY:
            if (wParam == IDH_PASTE_HOTKEY) {
                ExecuteImmediatePaste();
            }
            break;

        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case IDC_RADIO_CLIPBOARD:
                    g_app.useClipboard = true;
                    EnableWindow(g_app.hwndButtonBrowse, FALSE);
                    break;

                case IDC_RADIO_FILE:
                    g_app.useClipboard = false;
                    EnableWindow(g_app.hwndButtonBrowse, TRUE);
                    break;

                case IDC_BUTTON_BROWSE: {
                    std::wstring path = showFileOpenDialog(hwnd);
                    if (!path.empty()) {
                        g_app.selectedFilePath = path;
                        SetWindowTextW(g_app.hwndStaticFilePath, path.c_str());
                    }
                    break;
                }

                case IDC_COMBO_MODE:
                    if (HIWORD(wParam) == CBN_SELCHANGE) {
                        int sel = (int)SendMessageW(g_app.hwndComboMode, CB_GETCURSEL, 0, 0);
                        switch (sel) {
                            case 0: g_app.injectionMode = InjectionMode::Auto; break;
                            case 1: g_app.injectionMode = InjectionMode::Unicode; break;
                            case 2: g_app.injectionMode = InjectionMode::VKScancode; break;
                            case 3: g_app.injectionMode = InjectionMode::Hybrid; break;
                        }
                    }
                    break;

                case IDC_CHECK_DIAG:
                    g_app.diagnosticMode = (SendMessageW(g_app.hwndCheckDiag, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    break;

                case IDC_CHECK_SILENT:
                    g_app.silentMode = (SendMessageW(g_app.hwndCheckSilent, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    break;

                case IDC_BUTTON_ARM:
                    if (g_app.isArmed) {
                        CancelArm();
                    } else {
                        StartArmCountdown();
                    }
                    break;

                case IDM_TRAY_ARM:
                    // ARM from tray using last settings
                    if (!g_app.isArmed) {
                        StartArmCountdown();
                    }
                    break;

                case IDM_TRAY_SHOW:
                    RestoreFromTray();
                    break;

                case IDM_TRAY_EXIT:
                    SaveSettings();
                    RemoveTrayIcon();
                    DestroyWindow(hwnd);
                    break;
            }
            break;

        case WM_DRAWITEM: {
            LPDRAWITEMSTRUCT dis = (LPDRAWITEMSTRUCT)lParam;
            if (dis->CtlID == IDC_BUTTON_ARM) {
                // Determine button state
                bool isPressed = (dis->itemState & ODS_SELECTED);
                bool isDisabled = (dis->itemState & ODS_DISABLED);

                // Set colors based on state
                COLORREF bgColor, textColor;
                if (isDisabled) {
                    bgColor = RGB(200, 200, 200);
                    textColor = RGB(128, 128, 128);
                } else if (g_app.isArmed) {
                    // Armed state - orange/amber
                    bgColor = RGB(221, 107, 32);  // #DD6B20
                    textColor = RGB(255, 255, 255);
                } else {
                    // Normal state - dark blue-gray
                    bgColor = RGB(45, 55, 72);    // #2D3748
                    textColor = RGB(255, 255, 255);
                }

                // Draw rounded rectangle background
                HBRUSH hBrush = CreateSolidBrush(bgColor);
                HPEN hPen = CreatePen(PS_SOLID, 1, bgColor);
                HBRUSH hOldBrush = (HBRUSH)SelectObject(dis->hDC, hBrush);
                HPEN hOldPen = (HPEN)SelectObject(dis->hDC, hPen);

                RoundRect(dis->hDC, dis->rcItem.left, dis->rcItem.top,
                          dis->rcItem.right, dis->rcItem.bottom, 8, 8);

                SelectObject(dis->hDC, hOldBrush);
                SelectObject(dis->hDC, hOldPen);
                DeleteObject(hBrush);
                DeleteObject(hPen);

                // Draw text
                wchar_t text[64];
                GetWindowTextW(dis->hwndItem, text, 64);

                SetBkMode(dis->hDC, TRANSPARENT);
                SetTextColor(dis->hDC, textColor);
                HFONT hOldFont = (HFONT)SelectObject(dis->hDC, g_app.hFontButton);

                DrawTextW(dis->hDC, text, -1, &dis->rcItem,
                          DT_CENTER | DT_VCENTER | DT_SINGLELINE);

                SelectObject(dis->hDC, hOldFont);

                return TRUE;
            }
            break;
        }

        case WM_TIMER:
            if (wParam == IDT_COUNTDOWN) {
                g_app.countdownRemaining--;
                if (g_app.countdownRemaining <= 0) {
                    KillTimer(hwnd, IDT_COUNTDOWN);
                    ExecutePaste();
                } else {
                    UpdateArmButtonText();
                }
            }
            break;

        case WM_TRAYICON:
            switch (lParam) {
                case WM_LBUTTONUP:
                case WM_LBUTTONDBLCLK:
                    RestoreFromTray();
                    break;
                case WM_RBUTTONUP:
                    ShowTrayMenu(hwnd);
                    break;
            }
            break;

        case WM_SIZE:
            if (wParam == SIZE_MINIMIZED) {
                MinimizeToTray();
            }
            break;

        case WM_CLOSE:
            UnregisterHotKey(hwnd, IDH_PASTE_HOTKEY);
            SaveSettings();
            RemoveTrayIcon();
            DestroyWindow(hwnd);
            break;

        case WM_DESTROY:
            // Clean up floating progress window
            if (g_app.hwndFloatingProgress) {
                DestroyWindow(g_app.hwndFloatingProgress);
                g_app.hwndFloatingProgress = NULL;
            }

            // Clean up custom fonts
            if (g_app.hFontUI) DeleteObject(g_app.hFontUI);
            if (g_app.hFontMono) DeleteObject(g_app.hFontMono);
            if (g_app.hFontButton) DeleteObject(g_app.hFontButton);

            // Clean up custom icon
            if (g_app.hAppIcon) DestroyIcon(g_app.hAppIcon);

            // Clean up logo image and GDI+
            if (g_app.pLogoImage) delete g_app.pLogoImage;
            GdiplusShutdown(g_app.gdiplusToken);

            PostQuitMessage(0);
            break;

        default:
            return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    return 0;
}

// ============================================================================
// Command Line Parsing
// ============================================================================

// Parse command line arguments
// Supports: --diag, --mode=vk|hybrid|unicode|auto
void ParseCommandLine() {
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    if (!argv) return;

    for (int i = 1; i < argc; i++) {
        // --diag flag
        if (_wcsicmp(argv[i], L"--diag") == 0) {
            g_app.diagnosticMode = true;
            continue;
        }

        // --mode=value
        if (_wcsnicmp(argv[i], L"--mode=", 7) == 0) {
            const wchar_t* modeStr = argv[i] + 7;
            g_app.injectionMode = ParseInjectionMode(modeStr);
            continue;
        }
    }

    LocalFree(argv);
}

// ============================================================================
// Entry Point
// ============================================================================

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance;
    (void)lpCmdLine;

    g_app.hInstance = hInstance;

    // Initialize common controls (for spin control)
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_UPDOWN_CLASS;
    InitCommonControlsEx(&icex);

    // Load settings before creating window
    LoadSettings();

    // Parse command line (overrides INI settings)
    ParseCommandLine();

    // Initialize GDI+
    GdiplusStartupInput gdiplusStartupInput;
    GdiplusStartup(&g_app.gdiplusToken, &gdiplusStartupInput, NULL);

    // Load custom icon - try embedded resource first, then file
    g_app.hAppIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    if (!g_app.hAppIcon) {
        // Try loading from file MadPaster.ico
        wchar_t iconPath[MAX_PATH];
        GetModuleFileNameW(NULL, iconPath, MAX_PATH);
        std::wstring path(iconPath);
        size_t pos = path.rfind(L"\\");
        if (pos != std::wstring::npos) {
            path = path.substr(0, pos + 1) + L"MadPaster.ico";
        }

        g_app.hAppIcon = (HICON)LoadImageW(NULL, path.c_str(), IMAGE_ICON, 0, 0,
                                            LR_LOADFROMFILE | LR_DEFAULTSIZE);
        if (!g_app.hAppIcon) {
            g_app.hAppIcon = LoadIcon(NULL, IDI_APPLICATION); // fallback
        }
    }

    // Register window class
    WNDCLASSEXW wc = {};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"MadPasterWindowClass";
    wc.hIcon = g_app.hAppIcon;
    wc.hIconSm = g_app.hAppIcon;

    if (!RegisterClassExW(&wc)) {
        MessageBox(NULL, L"Failed to register window class.", L"MadPaster - Error",
                   MB_OK | MB_ICONERROR);
        return 1;
    }

    // Create main window (centered on screen)
    int screenWidth = GetSystemMetrics(SM_CXSCREEN);
    int screenHeight = GetSystemMetrics(SM_CYSCREEN);
    int x = (screenWidth - WINDOW_WIDTH) / 2;
    int y = (screenHeight - WINDOW_HEIGHT) / 2;

    g_app.hwndMain = CreateWindowExW(
        0,
        L"MadPasterWindowClass",
        L"MadPaster",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        x, y, WINDOW_WIDTH, WINDOW_HEIGHT,
        NULL, NULL, hInstance, NULL
    );

    if (!g_app.hwndMain) {
        MessageBox(NULL, L"Failed to create window.", L"MadPaster - Error",
                   MB_OK | MB_ICONERROR);
        return 1;
    }

    ShowWindow(g_app.hwndMain, nCmdShow);
    UpdateWindow(g_app.hwndMain);

    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}
