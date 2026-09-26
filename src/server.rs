//! Context-menu event listener for the `rcm start` command.
//!
//! The transport (single duplex pipe, framing, reconnection) lives in
//! [`crate::pipe`]; this module is the small public entry point the CLI uses.

use crate::error::Result;
use crate::types::ContextMenuInfo;

/// Optional parameters a listener can send when it subscribes.
///
/// Every field is optional: leave it as `None` to keep the current setting.
#[derive(Debug, Default, Clone)]
pub struct ListenOptions {
    /// Whether Shift+right-click should show the native menu.
    ///
    /// `Some(true)` (the default policy) keeps the Windows 11 escape hatch
    /// working; `Some(false)` intercepts Shift as well. Applied to the running
    /// extension for this session only — menu blocking is global, so the last
    /// subscriber to pass a value wins.
    pub shift_bypass: Option<bool>,
}

/// Stream context-menu events until the process exits.
///
/// Connects to the shell extension and calls `on_message` for every captured
/// event, reconnecting automatically if Explorer (and therefore the pipe
/// server) restarts. While the shell extension has not been loaded yet, this
/// waits and retries.
///
/// This is the entry point for embedding a listener in your own program:
///
/// ```no_run
/// # async fn run() -> Result<(), rcm_com::error::RcmError> {
/// rcm_com::server::listen(|info| {
///     println!("{info}");
/// })
/// .await?;
/// # Ok(())
/// # }
/// ```
///
/// The function never returns, so spawn it as a background task when your
/// program has other work to do. Use [`listen_with`] to pass subscription
/// options. For non-Rust implementations the wire protocol is documented in
/// the README.
pub async fn listen<F>(on_message: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    listen_with(on_message, ListenOptions::default()).await
}

/// Like [`listen`], but sends initialisation parameters when subscribing.
///
/// ```no_run
/// # async fn run() -> Result<(), rcm_com::error::RcmError> {
/// use rcm_com::server::{listen_with, ListenOptions};
///
/// listen_with(
///     |info| println!("{info}"),
///     ListenOptions {
///         // Intercept Shift+right-click too.
///         shift_bypass: Some(false),
///     },
/// )
/// .await?;
/// # Ok(())
/// # }
/// ```
pub async fn listen_with<F>(on_message: F, options: ListenOptions) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    crate::pipe::subscribe(on_message, options.shift_bypass).await
}
