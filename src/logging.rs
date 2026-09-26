//! Unified logging built on the [`log`] facade.
//!
//! Every message the program produces — CLI results, diagnostics, and DLL
//! logs — goes through `log`, so a single level controls all output.
//!
//! The backend is [`RcmLogger`], which writes to one of two targets:
//!
//! * **Console** (`rcm` CLI): `info`/`debug`/`trace` go to stdout, while
//!   `warn`/`error` go to stderr — the conventional split that keeps data on
//!   stdout and diagnostics on stderr.
//! * **File** (`rcm_com.dll` inside Explorer, where no console exists):
//!   appended to `rcm.log` next to the DLL, with a size cap and duplicate
//!   suppression so a misbehaving listener cannot grow it without bound.
//!
//! The active level is persisted under `HKCU\Software\RcmCom\LogLevel` so that
//! `rcm log <level>` survives across processes. A running DLL additionally
//! receives the new level over the control pipe (see [`crate::control`]).

use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU8, Ordering};
use std::sync::{Mutex, OnceLock};
use std::time::{Duration, Instant};

use log::{Level, LevelFilter, Log, Metadata, Record};
use windows::Win32::System::Registry::HKEY_CURRENT_USER;

use crate::cmd::{RegKeyGuard, create_key, get_reg_value, open_key, set_reg_value};
use crate::error::Result;

/// Registry key (under `HKCU`) holding user-scoped settings.
const CONFIG_KEY: &str = r"Software\RcmCom";
/// Registry value name for the persisted log level.
const LOG_LEVEL_VALUE: &str = "LogLevel";

/// Level used when nothing has been configured.
const DEFAULT_LEVEL: LevelFilter = LevelFilter::Info;

/// Maximum size of the DLL log file before it is truncated.
const LOG_MAX_BYTES: u64 = 1024 * 1024;
/// Window during which an identical message is suppressed in the log file.
const LOG_DEDUP_WINDOW: Duration = Duration::from_secs(5);

// =============================================================================
// Level <-> string / integer conversion
// =============================================================================

/// Serialise a [`LevelFilter`] to the compact integer stored atomically.
const fn filter_to_u8(filter: LevelFilter) -> u8 {
    match filter {
        LevelFilter::Off => 0,
        LevelFilter::Error => 1,
        LevelFilter::Warn => 2,
        LevelFilter::Info => 3,
        LevelFilter::Debug => 4,
        LevelFilter::Trace => 5,
    }
}

/// Inverse of [`filter_to_u8`].
const fn u8_to_filter(value: u8) -> LevelFilter {
    match value {
        0 => LevelFilter::Off,
        1 => LevelFilter::Error,
        2 => LevelFilter::Warn,
        3 => LevelFilter::Info,
        4 => LevelFilter::Debug,
        _ => LevelFilter::Trace,
    }
}

/// Parse a level name (`off`, `error`, `warn`, `info`, `debug`, `trace`).
///
/// Case-insensitive; `warning` is accepted as an alias for `warn`.
pub fn parse_level(s: &str) -> Option<LevelFilter> {
    Some(match s.trim().to_ascii_lowercase().as_str() {
        "off" | "none" => LevelFilter::Off,
        "error" => LevelFilter::Error,
        "warn" | "warning" => LevelFilter::Warn,
        "info" => LevelFilter::Info,
        "debug" => LevelFilter::Debug,
        "trace" => LevelFilter::Trace,
        _ => return None,
    })
}

/// Canonical name of a [`LevelFilter`] (as accepted by [`parse_level`]).
pub fn level_name(filter: LevelFilter) -> &'static str {
    match filter {
        LevelFilter::Off => "off",
        LevelFilter::Error => "error",
        LevelFilter::Warn => "warn",
        LevelFilter::Info => "info",
        LevelFilter::Debug => "debug",
        LevelFilter::Trace => "trace",
    }
}

// =============================================================================
// Persistence (HKCU\Software\RcmCom\LogLevel)
// =============================================================================

/// Read the persisted log level, if any.
fn load_persisted_level() -> Option<LevelFilter> {
    let key = open_key(HKEY_CURRENT_USER, CONFIG_KEY).ok()?;
    let _guard = RegKeyGuard::new(key);
    let raw = get_reg_value(key, Some(LOG_LEVEL_VALUE)).ok()?;
    parse_level(&raw)
}

/// Persist a log level for future processes under `HKCU`.
pub fn persist_level(filter: LevelFilter) -> Result<()> {
    let key = create_key(HKEY_CURRENT_USER, CONFIG_KEY)?;
    let _guard = RegKeyGuard::new(key);
    set_reg_value(key, Some(LOG_LEVEL_VALUE), level_name(filter))
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
    filter: AtomicU8,
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
        if let Ok(meta) = std::fs::metadata(path)
            && meta.len() > LOG_MAX_BYTES
        {
            let _ = std::fs::OpenOptions::new()
                .write(true)
                .truncate(true)
                .open(path);
        }

        if let Ok(mut file) = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(path)
        {
            let _ = writeln!(
                file,
                "[{}] {:<5} {}",
                crate::helpers::timestamp(),
                record.level(),
                message
            );
        }
    }
}

impl Log for RcmLogger {
    fn enabled(&self, metadata: &Metadata) -> bool {
        match u8_to_filter(self.filter.load(Ordering::Relaxed)).to_level() {
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

/// Effective level for a freshly initialised process.
///
/// `RCM_LOG` (if set and valid) overrides the persisted setting, which in turn
/// overrides [`DEFAULT_LEVEL`].
fn initial_level() -> LevelFilter {
    if let Ok(raw) = std::env::var("RCM_LOG")
        && let Some(filter) = parse_level(&raw)
    {
        return filter;
    }
    load_persisted_level().unwrap_or(DEFAULT_LEVEL)
}

fn init(target: Target) {
    INIT.get_or_init(|| {
        let level = initial_level();
        let logger: &'static RcmLogger = Box::leak(Box::new(RcmLogger {
            filter: AtomicU8::new(filter_to_u8(level)),
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
            log::set_max_level(level);
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
/// Falls back to the console target if the DLL directory cannot be resolved.
/// Must be called *after* the loader lock is released (i.e. from COM
/// activation, never from `DllMain`).
pub fn init_dll() {
    match crate::helpers::dll_dir() {
        Some(dir) => init(Target::File(dir.join("rcm.log"))),
        None => init(Target::Console),
    }
}

// =============================================================================
// Runtime level control
// =============================================================================

/// Currently active level of this process.
pub fn current_level() -> LevelFilter {
    LOGGER
        .get()
        .map(|logger| u8_to_filter(logger.filter.load(Ordering::Relaxed)))
        .unwrap_or(DEFAULT_LEVEL)
}

/// Change the level of this process only (no persistence).
///
/// Used to apply a level pushed over the control pipe into a running DLL.
pub(crate) fn apply_level(filter: LevelFilter) {
    if let Some(logger) = LOGGER.get() {
        logger
            .filter
            .store(filter_to_u8(filter), Ordering::Relaxed);
    }
    log::set_max_level(filter);
}

/// Set the level for this process and persist it for future processes.
pub fn set_level(filter: LevelFilter) -> Result<()> {
    apply_level(filter);
    persist_level(filter)
}
