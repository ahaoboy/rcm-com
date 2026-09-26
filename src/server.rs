//! Context-menu event listener for the `rcm start` command.
//!
//! The transport (single duplex pipe, framing, reconnection) lives in
//! [`crate::pipe`]; this module is the small public entry point the CLI uses.

use crate::error::Result;
use crate::types::ContextMenuInfo;

/// Stream context-menu events until the process exits.
///
/// Connects to the shell extension and calls `on_message` for every captured
/// event, reconnecting automatically if Explorer (and therefore the pipe
/// server) restarts. While the shell extension has not been loaded yet, this
/// waits and retries.
pub async fn listen<F>(on_message: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    crate::pipe::subscribe(on_message).await
}
