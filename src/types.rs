//! Data types for right-click context menu information captured by the shell extension.

use serde::{Deserialize, Serialize};
use windows::Win32::UI::Shell::{
    CMF_ASYNCVERBSTATE, CMF_CANRENAME, CMF_DEFAULTONLY, CMF_DISABLEDVERBS, CMF_DONOTPICKDEFAULT,
    CMF_EXPLORE, CMF_EXTENDEDVERBS, CMF_INCLUDESTATIC, CMF_ITEMMENU, CMF_NODEFAULT, CMF_NORMAL,
    CMF_NOVERBS, CMF_OPTIMIZEFORINVOKE, CMF_SYNCCASCADEMENU, CMF_VERBSONLY,
};

/// Human-readable names for every `CMF_*` bit understood by Explorer.
///
/// Single source of truth for flag decoding, kept in sync with the
/// `windows`-crate constants rather than re-declaring magic numbers.
const CMF_FLAGS: [(u32, &str); 14] = [
    (CMF_DEFAULTONLY, "CMF_DEFAULTONLY"),
    (CMF_VERBSONLY, "CMF_VERBSONLY"),
    (CMF_EXPLORE, "CMF_EXPLORE"),
    (CMF_NOVERBS, "CMF_NOVERBS"),
    (CMF_CANRENAME, "CMF_CANRENAME"),
    (CMF_NODEFAULT, "CMF_NODEFAULT"),
    (CMF_INCLUDESTATIC, "CMF_INCLUDESTATIC"),
    (CMF_ITEMMENU, "CMF_ITEMMENU"),
    (CMF_EXTENDEDVERBS, "CMF_EXTENDEDVERBS"),
    (CMF_DISABLEDVERBS, "CMF_DISABLEDVERBS"),
    (CMF_ASYNCVERBSTATE, "CMF_ASYNCVERBSTATE"),
    (CMF_OPTIMIZEFORINVOKE, "CMF_OPTIMIZEFORINVOKE"),
    (CMF_SYNCCASCADEMENU, "CMF_SYNCCASCADEMENU"),
    (CMF_DONOTPICKDEFAULT, "CMF_DONOTPICKDEFAULT"),
];

/// The type of event that triggered the context menu.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum Event {
    Click { flags: u32 },
    Menu { flags: u32 },
    Shift { flags: u32 },
}

impl Default for Event {
    fn default() -> Self {
        Event::Menu { flags: 0 }
    }
}

impl Event {
    /// Return the raw flags bitmask.
    pub fn flags(&self) -> u32 {
        match self {
            Event::Click { flags } => *flags,
            Event::Menu { flags } => *flags,
            Event::Shift { flags } => *flags,
        }
    }

    /// Return a human-readable representation of the flags bitmask.
    pub fn flags_str(&self) -> String {
        let uflags = self.flags();
        if uflags == CMF_NORMAL {
            return "CMF_NORMAL".to_string();
        }
        let names: Vec<&str> = CMF_FLAGS
            .iter()
            .filter(|(bit, _)| uflags & bit != 0)
            .map(|(_, name)| *name)
            .collect();
        names.join(" | ")
    }
}

impl std::fmt::Display for Event {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let name = match self {
            Event::Click { .. } => "Click",
            Event::Menu { .. } => "Menu",
            Event::Shift { .. } => "Shift",
        };
        write!(f, "{} ({} - {})", name, self.flags(), self.flags_str())
    }
}

/// All captured right-click context data sent from the shell extension to the
/// listening process via the named pipe.
///
/// Fields are grouped by meaning (identity/timing, location, owning window,
/// trigger) rather than by size: Rust's default `repr` already reorders fields
/// for the smallest layout, so this order costs nothing and reads better.
///
/// Every field carries `#[serde(default)]` so the wire format stays additive:
/// a payload produced by a different version still deserialises instead of
/// failing on a missing key. Always add new fields the same way.
///
/// **Units:** the timing fields ([`Self::captured`], [`Self::elapsed`]) are
/// integers in **microseconds**. The unit is stated here once rather than
/// repeated in each field name.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct ContextMenuInfo {
    // ── identity & timing ────────────────────────────────────────────────
    /// Process-unique id of this event (monotonic, hex).
    ///
    /// Lets a listener correlate events with each other or with the
    /// extension's log lines.
    #[serde(default)]
    pub cid: String,
    /// Wall-clock instant at which the capture started.
    ///
    /// **Microseconds since the Unix epoch (UTC).** This is an absolute
    /// instant, not a duration: adding [`Self::elapsed`] yields the moment the
    /// record was handed to the pipe.
    ///
    /// The value stays below 2^53 (~year 2255), so it survives a round trip
    /// through JSON consumers that use IEEE-754 doubles.
    #[serde(default)]
    pub captured: u64,
    /// Time spent inside the shell extension producing this record.
    ///
    /// **Microseconds**, measured from `IShellExtInit::Initialize` (entry)
    /// until `IContextMenu::QueryContextMenu` hands the record to the pipe, so
    /// it covers everything the extension costs the shell for one right-click.
    #[serde(default)]
    pub elapsed: u64,

    // ── where the click happened ─────────────────────────────────────────
    /// Cursor position in screen coordinates.
    #[serde(default)]
    pub x: i32,
    /// Cursor position in screen coordinates.
    #[serde(default)]
    pub y: i32,
    /// Folder the menu was invoked on; may be empty for some shell paths.
    #[serde(default)]
    pub dir: String,
    /// Selected items. Empty means the click was on the background.
    #[serde(default)]
    pub files: Vec<String>,
    /// `true` when the click was on empty space rather than on selected items.
    #[serde(default)]
    pub bg: bool,

    // ── owning window ────────────────────────────────────────────────────
    /// Foreground window handle, as an integer.
    #[serde(default)]
    pub hwnd: usize,
    /// Window class of the foreground window, e.g. `CabinetWClass`.
    #[serde(default)]
    pub class: String,
    /// Process id of the process the extension is running in (Explorer).
    #[serde(default)]
    pub pid: u32,

    // ── trigger ──────────────────────────────────────────────────────────
    /// What kind of right-click produced this event, with the raw flags.
    #[serde(default)]
    pub event: Event,
}

impl std::fmt::Display for ContextMenuInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "Event:  {}", self.event)?;
        writeln!(f, "Id:     {}", self.cid)?;
        writeln!(f, "Captured: {} us since Unix epoch", self.captured)?;
        writeln!(f, "Elapsed: {:.3} ms", self.elapsed as f64 / 1000.0)?;
        writeln!(f, "Position: ({}, {})", self.x, self.y)?;
        writeln!(f, "Directory: {}", self.dir)?;
        writeln!(f, "Background: {}", self.bg)?;
        writeln!(f, "File Count: {}", self.files.len())?;
        writeln!(f, "Window: 0x{:X}", self.hwnd)?;
        writeln!(f, "Window Class: {}", self.class)?;
        writeln!(f, "Process ID: {}", self.pid)?;
        if !self.files.is_empty() {
            writeln!(f, "Selected Files:")?;
            for file in &self.files {
                writeln!(f, "  - {file}")?;
            }
        }
        writeln!(f, "---")?;
        Ok(())
    }
}
