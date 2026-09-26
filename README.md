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
| `rcm enable` | Block (hide) the native context menu — the default |
| `rcm disable` | Stop blocking — let the native context menu appear |
| `rcm query` | Print whether menu blocking is currently enabled |
| `rcm log get` | Query the log level of the running shell extension |
| `rcm log set <level>` | Set the log level (see [Logging](#logging)) |
| `rcm client get` | Show which program is registered as using the pipe |
| `rcm client set [path]` | Register a program path (defaults to this executable) |

Every command exits `0` on success and `1` on failure, so scripts and CI can
detect errors.

## Pipe client registration

The extension records which program is using the pipe:

```bash
rcm client get                 # print the registered absolute path
rcm client set                 # register this executable
rcm client set C:\tools\app.exe  # register a specific program
```

`rcm start` registers its own executable automatically when it subscribes, so
`rcm client get` normally reports the running listener.

## Logging

All output — command results, diagnostics, and DLL logs — is emitted through
the [`log`](https://docs.rs/log) facade and controlled by a single level:

| Level | Shows |
|---|---|
| `off` | Nothing |
| `error` | Failures only |
| `warn` | Failures and warnings |
| `info` | Normal command output (default) |
| `debug` | Adds verbose detail (e.g. the full context-menu struct) |
| `trace` | Everything |

```bash
rcm log get        # ask the running extension for its live level
rcm log set debug  # set it (persisted, and pushed to the extension)
```

The level is stored under `HKCU\Software\RcmCom\LogLevel` and applied to newly
started processes. When the shell extension is already loaded, `rcm log set`
also pushes the change to it over the pipe; `rcm log get` reports the local
setting if the extension is not running.

* The `rcm` CLI writes `info`/`debug`/`trace` to **stdout** and
  `warn`/`error` to **stderr**.
* The DLL running inside Explorer has no console, so it appends to
  `rcm.log` next to `rcm_com.dll` (size-capped and with duplicate messages
  suppressed).

For a one-off run you can override the stored level with the `RCM_LOG`
environment variable:

```bash
RCM_LOG=debug rcm status
```

## Listening

```bash
rcm start
```

All CLI ↔ extension traffic uses a **single duplex named pipe**
(`\\.\pipe\rcm_com`), with the shell extension as the server and the CLI as the
client. Because the extension hosts the pipe, `rcm enable` / `rcm disable` /
`rcm query` / `rcm log` / `rcm client` all work on their own — `rcm start` does
not need to be running, the extension just has to be loaded (right-click once).

`rcm start` waits for the extension and then prints each event as it happens.
Right-click any file, folder, or empty space to see real-time output:

```
INFO  [2026-05-26 10:30:15 UTC]
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

Run at `debug` level (`rcm log set debug`) to also receive the full struct for
each event.
