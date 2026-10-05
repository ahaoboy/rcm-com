//! Public entry point for listening to context-menu events.
//!
//! The transport lives in [`crate::events`]; this module is the stable API an
//! embedding program uses.

use crate::error::Result;
use crate::types::ContextMenuInfo;

/// Optional parameters sent to the shell extension when a listener starts.
#[derive(Debug, Default, Clone)]
pub struct ListenOptions {
    /// Whether Shift+right-click should show the native menu.
    ///
    /// `Some(true)` (the extension's default) keeps the Windows 11 classic-menu
    /// escape hatch working; `Some(false)` intercepts Shift as well. Applied to
    /// the running extension for this session only — menu blocking is global, so
    /// the last listener to set it wins.
    pub shift_bypass: Option<bool>,
}

/// Stream context-menu events until the process exits.
///
/// Hosts the event pipe, so the listener owns it: restarting Explorer does not
/// break the channel, and every Explorer process that loads the extension
/// connects independently, so events from all of them arrive here.
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
/// Never returns; spawn it as a background task if your program has other work
/// to do. Use [`listen_with`] to pass [`ListenOptions`].
pub async fn listen<F>(on_message: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    listen_with(on_message, ListenOptions::default()).await
}

/// Like [`listen`], but applies [`ListenOptions`] to the extension first.
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
    // Best effort: the extension may not be loaded yet. Applied once here rather
    // than per event, because menu policy is global state, not something the
    // event stream carries.
    //
    // `try_set_shift_bypass` is blocking pipe I/O, so it is pushed onto a
    // blocking thread — calling it directly would stall an executor worker for
    // up to its timeout.
    if let Some(enabled) = options.shift_bypass {
        let pushed =
            tokio::task::spawn_blocking(move || crate::control::try_set_shift_bypass(enabled))
                .await;
        match pushed {
            Ok(Err(err)) => log::debug!("could not apply shift_bypass yet: {err}"),
            Err(err) => log::debug!("shift_bypass task failed: {err}"),
            Ok(Ok(())) => {}
        }
    }
    crate::events::serve(on_message).await
}
