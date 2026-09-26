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
| `rcm shift get` | Show the Shift+right-click behaviour of the running extension |
| `rcm shift set <on\|off>` | Change it for the current Explorer session |

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

## Implementing your own listener

The shell extension hosts the pipe, so your program only needs to be a
**client** — no daemon has to be running first. In Rust:

```rust
use rcm_com::server::{listen, listen_with, ListenOptions};

// Reports every right-click; reconnects automatically if Explorer restarts.
listen(|info| {
    println!("{} -> {:?}", info.ts, info.files);
})
.await?;

// Or pass initialisation parameters when subscribing:
listen_with(
    |info| println!("{info}"),
    ListenOptions {
        // Intercept Shift+right-click too (see below).
        shift_bypass: Some(false),
    },
)
.await?;
```

`listen` never returns; spawn it as a background task if your app has other
work to do. It does not install a `log` backend — call
`rcm_com::logging::init_console()` or your own logger if you want the
library's messages.

### Wire protocol

Any language can implement the listener against `\\.\pipe\rcm_com`. Messages are
**newline-delimited JSON**, one message per line. Send one request to subscribe:

```json
{"cmd":"subscribe","path":"C:\\tools\\my-listener.exe","options":{"shift_bypass":false}}
```

`path` is optional; when given it is recorded as the pipe's registered user
(see `rcm client get`). `options` is optional too — see
[Shift + right-click](#shift--right-click). After that the server pushes one
JSON object per captured right-click:

```json
{"type":"event","event":{"cid":"","ts":"2026-05-26 10:30:15 UTC","x":1024,"y":768,"dir":"C:\\Users\\Admin\\Desktop","files":["C:\\Users\\Admin\\Desktop\\readme.txt"],"bg":false,"hwnd":1715004,"class":"CabinetWClass","pid":12345,"event":{"type":"Menu","flags":0}}}
```

The connection stays open; multiple clients can subscribe at once and each
receives every event. Events captured while nobody is connected are buffered
(the most recent few) and replayed to the next subscriber, so the right-click
that loaded the extension is not lost.

The other request types (`enable`, `disable`, `query`, `get_log`, `set_log`,
`get_client`, `set_client`, `get_shift_bypass`, `set_shift_bypass`) each answer
with a single `{"type":...}` line; see `src/pipe.rs` for the full list.

### Shift + right-click

On Windows 11 the classic (expanded) context menu is reached with
**Shift+right-click**. Because that is an explicit request for the real menu,
by default the extension reports the event but **does not intercept** the
native menu — it opens as Windows would normally show it. A plain right-click
keeps being intercepted while blocking is enabled.

```bash
rcm shift get        # current behaviour (default: on)
rcm shift set off    # intercept Shift+right-click as well
rcm shift set on     # restore the default escape hatch
```

This setting is **not persisted** — it lives in the loaded shell extension
only, so an Explorer restart (or reloading the DLL) returns to the default
`on`. That also means `rcm shift get` / `set` require the extension to be
loaded (right-click once); they fail with an error otherwise.

A **subscriber can set it for its session** via the subscribe `options`, which
is handy when a listener needs the native menu for its own workflow:

```rust
listen_with(cb, ListenOptions { shift_bypass: Some(false) }).await?;
```

Menu blocking is global, so subscription options apply to the running extension
for the session only (not persisted), and the last subscriber to pass a value
wins.

To turn interception off entirely, use `rcm disable`.


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
