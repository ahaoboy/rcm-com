//! Unified logging built on the [`log`] facade.
//!
//! This module is the **single** place that writes output — no other module
//! touches stdout, stderr, or the log file. Three channels cover every case:
//!
//! | Channel | Purpose | Level-controlled? |
//! |---|---|---|
//! | [`output`] | Command results: status reports, queried values, captured events | **No** — always printed |
//! | `log::{info,warn,error,debug,…}` | Diagnostics: progress notes, warnings, failures | Yes |
//! | [`log_load`] | DLL load-stage diagnostics, before the backend exists | **No** — always recorded |
//!
//! Results and diagnostics are deliberately separate: `rcm log set off` must not
//! make `rcm status` print nothing and still exit `0`. Use `log::*` for anything
//! a user could reasonably want silenced, and [`output`] for the payload they
//! asked for.
//!
//! # Targets
//!
//! * **Console** (`rcm` CLI): diagnostics go to stdout for `info` and below, and
//!   to stderr for `warn`/`error` — the conventional split that keeps data on
//!   stdout and diagnostics on stderr.
//! * **File** (`rcm_com.dll` inside Explorer, where no console exists): records
//!   are appended to `rcm.log` next to the DLL, with duplicate suppression and a
//!   **hard 1 MiB cap** (see [`LOG_MAX_BYTES`]).
//!
//! The active level is persisted under `HKCU\Software\RcmCom\LogLevel` so that
//! `rcm log <level>` survives across processes. A running DLL additionally
//! receives the new level over the pipe (see [`crate::control`]).
//!
//! # Load-stage diagnostics
//!
//! Failures while the shell loads the DLL happen *before* the logger exists
//! (or because it could not be set up), and they are what you most need
//! evidence for. [`log_load`] therefore writes straight to the log file and
//! ignores the configured level. It must never be called from `DllMain`.

use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::str::FromStr;
use std::sync::atomic::{AtomicU8, Ordering};
use std::sync::{Mutex, OnceLock};
use std::time::{Duration, Instant};

use log::{Level, LevelFilter, Log, Metadata, Record};
use serde::{Deserialize, Serialize};

use crate::cmd;
use crate::consts::CONFIG_LOG_LEVEL;
use crate::error::Result;

/// Registry value name for the persisted log level.
const LOG_LEVEL_VALUE: &str = CONFIG_LOG_LEVEL;

/// Level used when nothing has been configured.
const DEFAULT_LEVEL: LogLevel = LogLevel::Info;

/// Hard upper bound for the log file. It is never exceeded: room is reserved
/// before writing, so a record cannot push the file past this size.
const LOG_MAX_BYTES: u64 = 1024 * 1024;
/// Headroom kept free for the notice written when the log resets. Subtracting
/// it makes the bound provable: after a reset the file is at most
/// `notice + record` = `LOG_RESERVE + (LOG_MAX_BYTES - LOG_RESERVE)`.
const LOG_RESET_RESERVE: u64 = 256;
/// Largest record accepted, leaving room for a reset notice in the same file.
const LOG_MAX_RECORD_BYTES: u64 = LOG_MAX_BYTES - LOG_RESET_RESERVE;
/// Window during which an identical message is suppressed in the log file.
const LOG_DEDUP_WINDOW: Duration = Duration::from_secs(5);

/// File name of the DLL log.
const LOG_FILE_NAME: &str = "rcm.log";

// =============================================================================
// Log file location
// =============================================================================

/// Resolved log file path, computed once per process.
static LOG_FILE: OnceLock<Option<PathBuf>> = OnceLock::new();

/// Resolve (once) the log file path: `rcm.log` next to the DLL.
///
/// The file is always kept with the DLL, so a single known location holds all
/// diagnostics. Returns `None` when the DLL directory itself cannot be
/// resolved, in which case file logging is unavailable.
fn log_file_path() -> Option<&'static PathBuf> {
    LOG_FILE
        .get_or_init(|| crate::helpers::dll_dir().map(|dir| dir.join(LOG_FILE_NAME)))
        .as_ref()
}

/// Append one already-formatted record, keeping the file within
/// [`LOG_MAX_BYTES`].
///
/// Size is reserved *before* writing. The previous implementation only checked
/// the size already on disk, so a record could push the file past the cap —
/// and a large one (a debug dump naming many files) could overshoot it by a lot.
///
/// When the cap is reached the file is restarted with a notice explaining that
/// older entries were dropped, so a reader is never misled by a silent gap.
fn append_line(path: &Path, line: &str) {
    // A record that could never fit even in an empty file is replaced by a short
    // notice, so pathological input cannot defeat the cap.
    let record = if line.len() as u64 + 1 > LOG_MAX_RECORD_BYTES {
        format!(
            "[{}] WARN  oversized log record discarded ({} bytes)",
            crate::helpers::timestamp(),
            line.len()
        )
    } else {
        line.to_owned()
    };
    let needed = record.len() as u64 + 1;

    // Reserve space up front so the write below cannot exceed the cap.
    if std::fs::metadata(path).is_ok_and(|meta| meta.len() + needed > LOG_MAX_RECORD_BYTES)
        && let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .write(true)
            .truncate(true)
            .open(path)
    {
        let _ = writeln!(
            file,
            "[{}] WARN  log reached its {LOG_MAX_BYTES}-byte cap; older entries were dropped",
            crate::helpers::timestamp()
        );
    }

    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)
    {
        let _ = writeln!(file, "{record}");
    }
}

/// Record a DLL load-stage message directly to the log file.
///
/// Deliberately bypasses the [`log`] facade: load failures happen before the
/// logger is installed — or because it could not be installed — and they are
/// recorded even when the level is `off`, since a DLL that fails to load is
/// exactly when the evidence matters most.
///
/// Must **not** be called from `DllMain`: file I/O while the loader lock is
/// held can deadlock the process. Call it from COM entry points instead.
pub fn log_load(message: impl std::fmt::Display) {
    let Some(path) = log_file_path() else {
        return;
    };
    append_line(
        path,
        &format!("[{}] LOAD  {message}", crate::helpers::timestamp()),
    );
}

// =============================================================================
// LogLevel
// =============================================================================

/// A log level, shared by the CLI, the persisted setting, and the pipe
/// protocol.
///
/// Using an enum instead of a string means an invalid level cannot be
/// constructed: the compiler and the deserialiser both reject it, and the CLI
/// can offer the exact set of values as completions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LogLevel {
    /// Disable all logging.
    Off,
    /// Errors only.
    Error,
    /// Warnings and errors.
    Warn,
    /// Normal output (the default).
    #[default]
    Info,
    /// Verbose diagnostics.
    Debug,
    /// Everything, including traces.
    Trace,
}

impl LogLevel {
    /// Every level, from quietest to most verbose.
    pub const ALL: [LogLevel; 6] = [
        LogLevel::Off,
        LogLevel::Error,
        LogLevel::Warn,
        LogLevel::Info,
        LogLevel::Debug,
        LogLevel::Trace,
    ];

    /// Convert to the [`log`] crate's filter.
    pub const fn to_filter(self) -> LevelFilter {
        match self {
            LogLevel::Off => LevelFilter::Off,
            LogLevel::Error => LevelFilter::Error,
            LogLevel::Warn => LevelFilter::Warn,
            LogLevel::Info => LevelFilter::Info,
            LogLevel::Debug => LevelFilter::Debug,
            LogLevel::Trace => LevelFilter::Trace,
        }
    }

    /// Convert from the [`log`] crate's filter.
    pub const fn from_filter(filter: LevelFilter) -> Self {
        match filter {
            LevelFilter::Off => LogLevel::Off,
            LevelFilter::Error => LogLevel::Error,
            LevelFilter::Warn => LogLevel::Warn,
            LevelFilter::Info => LogLevel::Info,
            LevelFilter::Debug => LogLevel::Debug,
            LevelFilter::Trace => LogLevel::Trace,
        }
    }

    /// Canonical lowercase name.
    pub const fn as_str(self) -> &'static str {
        match self {
            LogLevel::Off => "off",
            LogLevel::Error => "error",
            LogLevel::Warn => "warn",
            LogLevel::Info => "info",
            LogLevel::Debug => "debug",
            LogLevel::Trace => "trace",
        }
    }

    /// Serialise to the compact integer stored atomically.
    const fn to_u8(self) -> u8 {
        self as u8
    }

    /// Inverse of [`LogLevel::to_u8`]; unknown values fall back to the default.
    const fn from_u8(value: u8) -> Self {
        match value {
            0 => LogLevel::Off,
            1 => LogLevel::Error,
            2 => LogLevel::Warn,
            3 => LogLevel::Info,
            4 => LogLevel::Debug,
            5 => LogLevel::Trace,
            _ => DEFAULT_LEVEL,
        }
    }
}

impl std::fmt::Display for LogLevel {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(self.as_str())
    }
}

impl FromStr for LogLevel {
    type Err = String;

    /// Parse a level name, case-insensitively. `warning` and `none` are
    /// accepted as aliases for `warn` and `off`.
    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        match s.trim().to_ascii_lowercase().as_str() {
            "off" | "none" => Ok(LogLevel::Off),
            "error" => Ok(LogLevel::Error),
            "warn" | "warning" => Ok(LogLevel::Warn),
            "info" => Ok(LogLevel::Info),
            "debug" => Ok(LogLevel::Debug),
            "trace" => Ok(LogLevel::Trace),
            other => Err(format!(
                "invalid log level '{other}' (expected one of: off, error, warn, info, debug, trace)"
            )),
        }
    }
}

// =============================================================================
// Persistence (HKCU\Software\RcmCom\LogLevel)
// =============================================================================

/// Read the persisted log level, if any.
fn load_persisted_level() -> Option<LogLevel> {
    cmd::read_config(LOG_LEVEL_VALUE)?.parse().ok()
}

/// Persist a log level for future processes under `HKCU`.
pub fn persist_level(level: LogLevel) -> Result<()> {
    cmd::write_config(LOG_LEVEL_VALUE, level.as_str())
}

// =============================================================================
// Logger backend
// =============================================================================

/// Where log records are written.
enum Target {
    /// CLI: stdout for info and below, stderr for warnings and errors.
    Console,
    /// DLL: append to a file (no console is attached inside Explorer).
    File(PathBuf),
}

/// Last message written to the file target, used to suppress duplicates.
struct Dedupe {
    hash: u64,
    at: Option<Instant>,
}

struct RcmLogger {
    target: Target,
    dedupe: Mutex<Dedupe>,
}

impl RcmLogger {
    fn write_console(&self, record: &Record) {
        let line = format!("{:<5} {}", record.level(), record.args());
        if record.level() <= Level::Warn {
            let _ = writeln!(std::io::stderr(), "{line}");
        } else {
            let _ = writeln!(std::io::stdout(), "{line}");
        }
    }

    fn write_file(&self, path: &Path, record: &Record) {
        let message = record.args().to_string();

        // Suppress an identical message repeated within the window — a broken
        // listener used to make every right-click append another error line.
        let hash = {
            let mut hasher = DefaultHasher::new();
            message.hash(&mut hasher);
            hasher.finish()
        };
        {
            let mut dedupe = self.dedupe.lock().unwrap_or_else(|e| e.into_inner());
            if dedupe.hash == hash
                && dedupe
                    .at
                    .is_some_and(|t| t.elapsed() < LOG_DEDUP_WINDOW)
            {
                return;
            }
            dedupe.hash = hash;
            dedupe.at = Some(Instant::now());
        }

        // Truncate once the file grows past the cap so disk usage is bounded.
        append_line(
            path,
            &format!(
                "[{}] {:<5} {}",
                crate::helpers::timestamp(),
                record.level(),
                message
            ),
        );
    }
}

impl Log for RcmLogger {
    fn enabled(&self, metadata: &Metadata) -> bool {
        match active_level().to_filter().to_level() {
            Some(max) => metadata.level() <= max,
            None => false,
        }
    }

    fn log(&self, record: &Record) {
        if !self.enabled(record.metadata()) {
            return;
        }
        match &self.target {
            Target::Console => self.write_console(record),
            Target::File(path) => self.write_file(path, record),
        }
    }

    fn flush(&self) {}
}

// =============================================================================
// Initialisation
// =============================================================================

static INIT: OnceLock<()> = OnceLock::new();
static LOGGER: OnceLock<&'static RcmLogger> = OnceLock::new();
/// Guards the one-time "DLL loaded" record (it must not repeat per right-click).
static LOAD_LOGGED: OnceLock<()> = OnceLock::new();

/// The level currently in effect, independent of whether a logger has been
/// installed yet.
///
/// Keeping this global (rather than inside [`RcmLogger`]) means
/// [`apply_level`] and [`current_level`] stay correct even before
/// initialisation — e.g. if a lib user calls `start()` without `init_dll()`.
static ACTIVE_LEVEL: AtomicU8 = AtomicU8::new(LogLevel::Info as u8);

fn active_level() -> LogLevel {
    LogLevel::from_u8(ACTIVE_LEVEL.load(Ordering::Relaxed))
}

/// Effective level for a freshly initialised process.
///
/// `RCM_LOG` (if set and valid) overrides the persisted setting, which in turn
/// overrides [`DEFAULT_LEVEL`].
fn initial_level() -> LogLevel {
    if let Ok(raw) = std::env::var("RCM_LOG")
        && let Ok(level) = raw.parse::<LogLevel>()
    {
        return level;
    }
    load_persisted_level().unwrap_or(DEFAULT_LEVEL)
}

fn init(target: Target) {
    INIT.get_or_init(|| {
        let level = initial_level();
        ACTIVE_LEVEL.store(level.to_u8(), Ordering::Relaxed);
        let logger: &'static RcmLogger = Box::leak(Box::new(RcmLogger {
            target,
            dedupe: Mutex::new(Dedupe {
                hash: 0,
                at: None,
            }),
        }));
        let _ = LOGGER.set(logger);
        if log::set_logger(logger).is_ok() {
            // `set_logger` does not touch the global max level (which defaults
            // to `Off`), so it must be set explicitly or nothing is emitted.
            log::set_max_level(level.to_filter());
        } else {
            log_load("failed to install the log backend (a logger was already set)");
        }
    });
}

/// Initialise logging for the `rcm` CLI (writes to stdout/stderr).
///
/// Idempotent: only the first call takes effect.
pub fn init_console() {
    init(Target::Console);
}

/// Initialise logging for the DLL running inside Explorer (writes to a file).
///
/// Uses the same path as [`log_load`] (`rcm.log` next to the DLL) so load
/// diagnostics and runtime logs always land in one file. Falls back to the
/// console target only when the DLL directory cannot be resolved.
///
/// Must be called *after* the loader lock is released (i.e. from COM
/// activation, never from `DllMain`).
pub fn init_dll() {
    match log_file_path() {
        Some(path) => init(Target::File(path.clone())),
        None => init(Target::Console),
    }
    // Record the successful load once per process — `init_dll` runs on every
    // COM activation, which happens once per right-click.
    LOAD_LOGGED.get_or_init(|| {
        log_load(format_args!(
            "rcm_com.dll loaded into process {}",
            std::process::id()
        ));
    });
}

// =============================================================================
// Runtime level control
// =============================================================================

/// Currently active level of this process.
pub fn current_level() -> LogLevel {
    active_level()
}

/// Change the level of this process only (no persistence).
///
/// Used to apply a level pushed over the pipe into a running DLL. Works even
/// before [`init_console`] / [`init_dll`] have run.
pub(crate) fn apply_level(level: LogLevel) {
    ACTIVE_LEVEL.store(level.to_u8(), Ordering::Relaxed);
    log::set_max_level(level.to_filter());
}

/// Write a command **result** to stdout, unconditionally.
///
/// Results — a status report, a queried value, a captured event — are the
/// program's primary output, not diagnostics, so they deliberately ignore the
/// log level: `rcm log set off` must not make `rcm status` silently print
/// nothing and still exit `0`. Use `log::*` for diagnostics instead.
pub fn output(message: impl std::fmt::Display) {
    let _ = writeln!(std::io::stdout(), "{message}");
}

/// Set the level for this process and persist it for future processes.
pub fn set_level(level: LogLevel) -> Result<()> {
    apply_level(level);
    persist_level(level)
}
