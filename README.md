# RCM-COM

> ⚠️ **WARNING — Early Development**: This project is in active development.
> APIs, commands, and features are subject to breaking changes at any time.
> **Do not use in production environments.**

A Rust-based Windows Shell Extension that captures right-click context menu
information and sends it to a listening process via a named pipe.

Shortcut (`.lnk`) files are captured as the shortcut file path itself instead
of the linked target path.

---

## Install

Build, then run as **Administrator**:

```bash
rcm install
rcm restart-explorer
```

On Windows 11, switch to the classic context menu first so the extension is
triggered directly:

```bash
rcm menu win10
rcm restart-explorer
```

## Uninstall

```bash
rcm uninstall
rcm restart-explorer
```

## CLI Commands

| Command | Description |
|---|---|
| `rcm install` | Install and register the shell extension (requires admin) |
| `rcm uninstall` | Uninstall and clean up registry entries (requires admin) |
| `rcm start` | Start listening for context menu events via named pipe |
| `rcm status` | Show current registration status and configuration |
| `rcm menu win10` | Switch to Windows 10 classic context menu |
| `rcm menu win11` | Switch back to Windows 11 default context menu |
| `rcm menu default` | Set classic menu as default (`-c false` to disable) |
| `rcm restart-explorer` | Restart Explorer (stop → wait 5s → start) |

## Listening

```bash
rcm start
```

Right-click any file, folder, or empty space to see real-time output:

```
[2026-05-26 10:30:15 UTC]
Position: (1024, 768)
Directory: C:\Users\Admin\Desktop
Background: false
File Count: 2
Window: 0x1A2B3C
Window Class: CabinetWClass
Process ID: 12345
Event: Menu (0 - CMF_NORMAL)
Selected Files:
  - C:\Users\Admin\Desktop\readme.txt
  - C:\Users\Admin\Desktop\photo.jpg
---
```
