//! Data types for right-click context menu information captured by the shell extension.

use serde::{Deserialize, Serialize};

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
        let mut flags_str = Vec::new();
        if uflags == 0 {
            flags_str.push("CMF_NORMAL");
        }
        if uflags & 0x00000001 != 0 {
            flags_str.push("CMF_DEFAULTONLY");
        }
        if uflags & 0x00000002 != 0 {
            flags_str.push("CMF_VERBSONLY");
        }
        if uflags & 0x00000004 != 0 {
            flags_str.push("CMF_EXPLORE");
        }
        if uflags & 0x00000008 != 0 {
            flags_str.push("CMF_NOVERBS");
        }
        if uflags & 0x00000010 != 0 {
            flags_str.push("CMF_CANRENAME");
        }
        if uflags & 0x00000020 != 0 {
            flags_str.push("CMF_NODEFAULT");
        }
        if uflags & 0x00000040 != 0 {
            flags_str.push("CMF_INCLUDESTATIC");
        }
        if uflags & 0x00000080 != 0 {
            flags_str.push("CMF_ITEMMENU");
        }
        if uflags & 0x00000100 != 0 {
            flags_str.push("CMF_EXTENDEDVERBS");
        }
        if uflags & 0x00000200 != 0 {
            flags_str.push("CMF_DISABLEDVERBS");
        }
        if uflags & 0x00000400 != 0 {
            flags_str.push("CMF_ASYNCVERBSTATE");
        }
        if uflags & 0x00000800 != 0 {
            flags_str.push("CMF_OPTIMIZEFORINVOKE");
        }
        if uflags & 0x00001000 != 0 {
            flags_str.push("CMF_SYNCCASCADEMENU");
        }
        if uflags & 0x00002000 != 0 {
            flags_str.push("CMF_DONOTPICKDEFAULT");
        }
        if uflags & 0x00010000 != 0 {
            flags_str.push("CMF_DVFILE");
        }
        flags_str.join(" | ")
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
