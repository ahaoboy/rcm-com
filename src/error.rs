use thiserror::Error;

/// Errors surfaced by the crate.
///
/// Only variants that are actually constructed live here; an unused one is dead
/// weight in the public API.
#[derive(Error, Debug)]
pub enum RcmError {
    #[error("I/O Error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Serialization Error: {0}")]
    Serialization(#[from] serde_json::Error),

    #[error("Registry Error: {0}")]
    Registry(String),

    #[error("Registry Key Not Found: {0}")]
    RegistryKeyNotFound(String),

    #[error("Environment Error: {0}")]
    Environment(String),

    #[cfg(feature = "cli")]
    #[error("rcm-reg Error: {0}")]
    RcmReg(#[from] rcm_reg::RcmError),
}

pub type Result<T> = std::result::Result<T, RcmError>;
