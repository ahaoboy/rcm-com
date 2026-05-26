//! Data types for right-click context menu information captured by the shell extension.

use serde::{Deserialize, Serialize};

/// The type of event that triggered the context menu.
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type")]
pub enum Event {
    LeftClickSelect { flags: u32 },
    RightClickMenu { flags: u32 },
    ShiftSelect { flags: u32 },
}

impl Default for Event {
    fn default() -> Self {
        Event::RightClickMenu { flags: 0 }
    }
}

impl Event {
    /// Return the raw flags bitmask.
    pub fn flags(&self) -> u32 {
        match self {
            Event::LeftClickSelect { flags } => *flags,
            Event::RightClickMenu { flags } => *flags,
            Event::ShiftSelect { flags } => *flags,
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
            Event::LeftClickSelect { .. } => "LeftClickSelect",
            Event::RightClickMenu { .. } => "RightClickMenu",
            Event::ShiftSelect { .. } => "ShiftSelect",
        };
        write!(f, "{} ({} - {})", name, self.flags(), self.flags_str())
    }
}

/// All captured right-click context data sent from the shell extension to the
/// listening process via the named pipe.
#[derive(Debug, Default, Clone, Serialize, Deserialize)]
pub struct ContextMenuInfo {
    pub cid: String,
    pub timestamp: String,
    pub cursor_x: i32,
    pub cursor_y: i32,
    pub folder_path: String,
    pub selected_files: Vec<String>,
    pub file_count: u32,
    pub is_background: bool,
    pub window_handle: usize,
    pub window_class: String,
    pub process_id: u32,
    pub event: Event,
}

impl std::fmt::Display for ContextMenuInfo {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        writeln!(f, "[{}]", self.timestamp)?;
        writeln!(f, "Position: ({}, {})", self.cursor_x, self.cursor_y)?;
        writeln!(f, "Directory: {}", self.folder_path)?;
        writeln!(f, "Background: {}", self.is_background)?;
        writeln!(f, "File Count: {}", self.file_count)?;
        writeln!(f, "Window: 0x{:X}", self.window_handle)?;
        writeln!(f, "Window Class: {}", self.window_class)?;
        writeln!(f, "Process ID: {}", self.process_id)?;
        writeln!(f, "Event: {}", self.event)?;
        if !self.selected_files.is_empty() {
            writeln!(f, "Selected Files:")?;
            for file in &self.selected_files {
                writeln!(f, "  - {file}")?;
            }
        }
        writeln!(f, "---")?;
        Ok(())
    }
}
