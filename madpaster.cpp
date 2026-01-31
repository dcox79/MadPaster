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
#include <string>
#include <vector>

using namespace Gdiplus;

// Link common controls
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdiplus.lib")

// ============================================================================
// Constants and Control IDs
// ============================================================================

const int maxchar = 45000;

// Window dimensions
const int WINDOW_WIDTH = 400;
const int WINDOW_HEIGHT = 320;

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
    HWND hwndButtonArm;
    HWND hwndButtonBrowse;
    HWND hwndStaticFilePath;
    HWND hwndStaticStatus;
    HWND hwndLogo;

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
// Keyboard Simulation
// ============================================================================

size_t sendTextToWindow(const std::wstring& text) {
    for (size_t i = 0; i < text.size(); ++i) {
        // Check for ESC key at the start of each iteration
        if (GetAsyncKeyState(VK_ESCAPE) & 0x8000) {
            return i;  // Return count sent before interrupt
        }
        wchar_t c = text[i];
        // Handle Windows line endings: skip '\r' if followed by '\n'
        if (c == L'\r' && i + 1 < text.size() && text[i + 1] == L'\n') {
            continue;
        }
        // If newline, send a real Enter key event
        if (c == L'\n' || c == L'\r') {
            INPUT input[2] = {};
            input[0].type = INPUT_KEYBOARD;
            input[0].ki.wVk = VK_RETURN;
            input[0].ki.dwFlags = 0;

            input[1].type = INPUT_KEYBOARD;
            input[1].ki.wVk = VK_RETURN;
            input[1].ki.dwFlags = KEYEVENTF_KEYUP;

            SendInput(2, input, sizeof(INPUT));
            Sleep(g_app.keystrokeDelayMs);
        } else {
            INPUT input[2] = {};
            input[0].type = INPUT_KEYBOARD;
            input[0].ki.wVk = 0;
            input[0].ki.wScan = c;
            input[0].ki.dwFlags = KEYEVENTF_UNICODE;

            input[1].type = INPUT_KEYBOARD;
            input[1].ki.wVk = 0;
            input[1].ki.wScan = c;
            input[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;

            SendInput(2, input, sizeof(INPUT));
            Sleep(g_app.keystrokeDelayMs);
        }
    }
    return text.size();  // All characters sent
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
    ShowWindow(g_app.hwndMain, SW_SHOW);
    SetForegroundWindow(g_app.hwndMain);
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

    UpdateArmButtonText();
    UpdateStatus(L"Ready - ARM Starts MadPaster  ESC Interrupts MadPaster");
}

void ExecutePaste() {
    UpdateStatus(L"Executing...");
    UpdateArmButtonText();

    // Minimize to tray before pasting
    MinimizeToTray();

    // Brief delay to ensure window is hidden and focus transfers
    Sleep(150);

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
            size_t charsSent = sendTextToWindow(textContent);
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

    ResetArmState();
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

            // ARM button (large, prominent, owner-drawn)
            g_app.hwndButtonArm = CreateWindowW(L"BUTTON", L"ARM",
                WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                24, 190, 352, 50, hwnd, (HMENU)IDC_BUTTON_ARM, hInst, NULL);

            // Status label
            g_app.hwndStaticStatus = CreateWindowW(L"STATIC", L"Status: Ready - ARM Starts MadPaster  ESC Interrupts MadPaster",
                WS_CHILD | WS_VISIBLE,
                24, 255, 352, 20, hwnd, (HMENU)IDC_STATIC_STATUS, hInst, NULL);
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

            // Create tray icon
            CreateTrayIcon(hwnd);

            break;
        }

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
            SaveSettings();
            RemoveTrayIcon();
            DestroyWindow(hwnd);
            break;

        case WM_DESTROY:
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
