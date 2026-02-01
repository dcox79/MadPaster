# MadPaster

A Windows utility that simulates keyboard strokes to paste clipboard or file content. Built primarily for enterprise remote desktop environments where traditional paste (Ctrl+V) doesn't work, such as Azure Virtual Desktop (AVD), Citrix, and Azure Bastion. Also works with noVNC consoles, Proxmox, VMware, and other virtualization platforms.

![MadPaster Logo](MadPaster.png)

## Features

- **Dual Input Modes**: Paste from clipboard or file
- **Countdown Timer**: Configurable 0-60 second delay before pasting
- **Adjustable Speed**: Keystroke delay from 0-100ms for compatibility
- **System Tray Support**: Minimize to tray and quick ARM from tray menu
- **Interrupt Support**: Press ESC to stop pasting mid-operation
- **Wide Encoding Support**: UTF-8, UTF-16 LE/BE (with BOM), and ANSI
- **Settings Persistence**: Automatically saves preferences to INI file
- **Large Content Support**: Handles up to 45,000 characters or 500KB files

## Use Cases

- **Azure Virtual Desktop (AVD)** - When clipboard redirection is disabled or restricted
- **Citrix Virtual Apps/Desktops** - Bypassing clipboard policies in locked-down environments
- **Azure Bastion** - Pasting into browser-based VM consoles
- noVNC/VNC console windows
- Virtual machine consoles (Proxmox, VMware, VirtualBox)
- Serial terminal connections
- Any application where standard paste doesn't work

## Build

### Prerequisites

- MinGW-w64 with g++ (recommended) or MSVC
- Windows SDK

### PATH Setup (MinGW)

Add the MinGW bin directory to your system PATH:
```
C:\msys64\ucrt64\bin
```

### Compile

**Standard build:**
```bash
windres madpaster.rc -o madpaster.res -O coff
g++ -o madpaster.exe madpaster.cpp madpaster.res -mwindows -lcomdlg32 -lcomctl32 -lgdiplus
```

**Standalone build (no DLL dependencies):**
```bash
windres madpaster.rc -o madpaster.res -O coff
g++ -o madpaster.exe madpaster.cpp madpaster.res -mwindows -lcomdlg32 -lcomctl32 -lgdiplus -static
```

## Usage

1. **Select Source**: Choose "Clipboard" or "File" as your paste source
2. **Set Delay**: Configure countdown timer (0-60 seconds)
3. **Adjust Speed**: Set keystroke delay in milliseconds (default: 3ms)
4. **ARM**: Click ARM button to start countdown
5. **Switch Windows**: During countdown, switch to your target window
6. **Auto-Paste**: Application minimizes to tray and begins pasting
7. **Interrupt**: Press ESC at any time to stop pasting

### System Tray

- **Left-click**: Restore window
- **Right-click**: Context menu with ARM and Exit options
- Minimizing the window sends it to the system tray

## Technical Details

### Architecture

- Single-file C++ application using Win32 API
- Global `AppState` struct managing all application state
- Owner-drawn ARM button with visual state indicators
- GDI+ for PNG logo rendering

### Keyboard Simulation

Uses `SendInput` with `KEYEVENTF_UNICODE` for character-by-character simulation. Line breaks are sent as `VK_RETURN` key events. Default 3ms delay per keystroke ensures reliability across different applications.

### File Encoding

Automatic detection and conversion:
- UTF-8 with BOM
- UTF-16 LE with BOM
- UTF-16 BE with BOM
- ANSI fallback

### Settings

Stored in `madpaster.ini` (created automatically):
- Last selected mode (clipboard/file)
- Countdown delay
- Keystroke delay
- Last file path

## Limitations

- Maximum content length: 45,000 characters
- Maximum file size: 500 KB
- Windows-only (uses Win32 API)

## Version History

- **v3.0** - Renamed to MadPaster, updated branding
- Previous versions maintained as MadPaste

## License

This project is provided as-is for personal and educational use.

## Author

Dave Cox

---

**Note**: MadPaster simulates keyboard input at a low level. Some applications with strict security policies may block or detect this behavior.
