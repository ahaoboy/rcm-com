# RCM-COM — Windows Shell Extension Context Menu Logger

A Rust-based COM Shell Extension DLL that captures right-click context menu information (cursor position, directory path, selected files) and writes it to a `log.txt` file next to the DLL.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│  Windows Explorer (explorer.exe)                                │
│                                                                 │
│  User right-clicks a file/folder/background                     │
│       │                                                         │
│       ▼                                                         │
│  Registry lookup: ContextMenuHandlers\RcmContextMenu            │
│       │                                                         │
│       ▼                                                         │
│  CoCreateInstance(CLSID_RCM) ──► DllGetClassObject()            │
│       │                            │                            │
│       │                            ▼                            │
│       │                       ClassFactory::CreateInstance()     │
│       │                            │                            │
│       │                            ▼                            │
│       │                       ContextMenuHandler (new)          │
│       │                                                         │
│       ▼                                                         │
│  QueryInterface(IID_IShellExtInit)                              │
│       │                                                         │
│       ▼                                                         │
│  IShellExtInit::Initialize(pidlFolder, pDataObj, hKeyProgID)    │
│       │                                                         │
│       ├── GetCursorPos()         → cursor (x, y)                │
│       ├── SHGetPathFromIDListW() → folder path                  │
│       ├── IDataObject::GetData() → CF_HDROP                     │
│       │      └── DragQueryFileW() → selected file paths         │
│       ├── GetForegroundWindow()  → window handle                │
│       └── GetClassNameW()        → window class name            │
│                                                                 │
│       ▼                                                         │
│  QueryInterface(IID_IContextMenu)                               │
│       │                                                         │
│       ▼                                                         │
│  IContextMenu::QueryContextMenu()                               │
│       │                                                         │
│       └── Write all captured info to log.txt                    │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

## Project Structure

```
rcm-com/
├── Cargo.toml          # Dependencies (windows crate 0.61)
├── README.md           # This file
└── src/
    ├── lib.rs          # COM DLL: manual COM vtable implementation
    │   ├── ContextMenuInfo   — struct holding all captured data
    │   ├── ClassFactory      — IClassFactory COM object
    │   ├── ContextMenuHandler — IShellExtInit + IContextMenu
    │   ├── DllMain           — DLL entry point
    │   ├── DllGetClassObject — COM class object factory
    │   └── DllCanUnloadNow   — COM lifecycle management
    │
    └── main.rs         # CLI tool for registry registration
        ├── register    — write CLSID + handler keys to registry
        └── unregister  — remove all registry entries
```

## COM Interfaces Implemented

| Interface        | Purpose                                    |
| ---------------- | ------------------------------------------ |
| `IClassFactory`  | Creates `ContextMenuHandler` instances     |
| `IShellExtInit`  | Receives folder PIDL + IDataObject on init |
| `IContextMenu`   | Triggered when context menu is shown       |

## Data Captured (ContextMenuInfo)

| Field            | Source                         |
| ---------------- | ------------------------------ |
| `timestamp`      | `std::time::SystemTime`        |
| `cursor_x/y`     | `GetCursorPos()`               |
| `folder_path`    | `SHGetPathFromIDListW(pidl)`   |
| `selected_files` | `IDataObject` → `DragQueryFileW` |
| `file_count`     | `DragQueryFileW(0xFFFFFFFF)`   |
| `is_background`  | No files selected + folder set |
| `window_handle`  | `GetForegroundWindow()`        |
| `window_class`   | `GetClassNameW()`              |
| `process_id`     | `std::process::id()`           |

## Build & Usage Steps

### Step 1: Build

```bash
cargo build --release
```

Output:
- `target/release/rcm_com.dll` — the COM shell extension DLL
- `target/release/rcm.exe` — the registration CLI tool

### Step 2: Deploy

Copy both files to a **permanent** location (the DLL must stay in place while registered):

```bash
mkdir C:\rcm
copy target\release\rcm_com.dll C:\rcm\
copy target\release\rcm.exe C:\rcm\
```

### Step 3: Register (requires Administrator)

Open an **elevated** command prompt (Run as Administrator):

```bash
C:\rcm\rcm.exe register
```

This writes the following registry keys:
- `HKCR\CLSID\{B8A0E19C-4C6D-4A82-9F3B-6E8E7D1F2A5C}` — COM class registration
- `HKCR\CLSID\{...}\InProcServer32` — DLL path + `ThreadingModel=Apartment`
- `HKCR\*\shellex\ContextMenuHandlers\RcmContextMenu` — file handler
- `HKCR\Directory\shellex\ContextMenuHandlers\RcmContextMenu` — directory handler
- `HKCR\Directory\Background\shellex\ContextMenuHandlers\RcmContextMenu` — background handler
- `HKLM\SOFTWARE\Microsoft\..\Shell Extensions\Approved` — approved extension

### Step 4: Restart Explorer

```bash
taskkill /f /im explorer.exe
start explorer.exe
```

Or log out and log back in.

### Step 5: Test

1. Right-click any file, folder, or empty space in Explorer
2. Check `C:\rcm\log.txt` (same directory as the DLL)

Example log output:
```
[2026-04-04 02:30:15 UTC]
Position: (1024, 768)
Directory: C:\Users\Admin\Desktop
Background: false
File Count: 2
Window: 0x1A2B3C
Window Class: CabinetWClass
Process ID: 12345
Selected Files:
  - C:\Users\Admin\Desktop\readme.txt
  - C:\Users\Admin\Desktop\photo.jpg
---
```

### Step 6: Unregister

```bash
C:\rcm\rcm.exe unregister
taskkill /f /im explorer.exe
start explorer.exe
```

## CLSID

```
{B8A0E19C-4C6D-4A82-9F3B-6E8E7D1F2A5C}
```

## Troubleshooting

- **log.txt not created**: Ensure the DLL directory is writable. Shell extensions run inside `explorer.exe` which may have restricted write access to certain directories.
- **No effect after registration**: Restart Explorer (`taskkill /f /im explorer.exe && start explorer.exe`).
- **Registration fails**: Must run `rcm.exe register` from an elevated (Administrator) command prompt.
- **DLL not loading**: Verify the DLL path in registry matches the actual file location. Check with `reg query "HKCR\CLSID\{B8A0E19C-4C6D-4A82-9F3B-6E8E7D1F2A5C}\InProcServer32"`.
