use thiserror::Error;

#[derive(Error, Debug)]
pub enum RcmError {
    #[error("I/O Error: {0}")]
    Io(#[from] std::io::Error),

    #[error("Serialization Error: {0}")]
    Serialization(#[from] serde_json::Error),

    #[error("Registry Error: {0}")]
    Registry(String),

    #[error("Environment Error: {0}")]
    Environment(String),

    #[error("UTF-8 Decoding Error: {0}")]
    Utf8(#[from] std::string::FromUtf8Error),
}

pub type Result<T> = std::result::Result<T, RcmError>;
