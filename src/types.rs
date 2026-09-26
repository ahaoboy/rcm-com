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
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct ContextMenuInfo {
    pub cid: String,
    pub ts: String,
    pub x: i32,
    pub y: i32,
    pub dir: String,
    pub files: Vec<String>,
    pub bg: bool,
    pub hwnd: usize,
    pub class: String,
    pub pid: u32,
    pub event: Event,
}

impl std::fmt::Display for ContextMenuInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "[{}]", self.ts)?;
        writeln!(f, "Position: ({}, {})", self.x, self.y)?;
        writeln!(f, "Directory: {}", self.dir)?;
        writeln!(f, "Background: {}", self.bg)?;
        writeln!(f, "File Count: {}", self.files.len())?;
        writeln!(f, "Window: 0x{:X}", self.hwnd)?;
        writeln!(f, "Window Class: {}", self.class)?;
        writeln!(f, "Process ID: {}", self.pid)?;
        writeln!(f, "Event: {}", self.event)?;
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
