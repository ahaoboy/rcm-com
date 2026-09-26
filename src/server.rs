use tokio::io::AsyncReadExt;
use tokio::net::windows::named_pipe::{NamedPipeServer, ServerOptions};

use crate::error::Result;
use crate::{ContextMenuInfo, PIPE_NAME};

/// Maximum accepted size of a single JSON frame (1 MiB). Bounds server memory
/// so a malicious or buggy client cannot stream indefinitely into the server.
const MAX_FRAME_BYTES: u64 = 1024 * 1024;

/// Create the listener pipe with a DACL restricted to the current user and
/// `Local System`.
///
/// `first` guards against pipe-name squatting: it must be `true` only for the
/// very first instance, since the option makes creation fail when an instance
/// already exists. Subsequent iterations reference the freed name again.
fn create_server(first: bool) -> Result<NamedPipeServer> {
    let mut security = crate::helpers::PipeSecurity::new();
    let mut options = ServerOptions::new();
    options.first_pipe_instance(first);
    // Safety: `security` owns a valid SECURITY_ATTRIBUTES (or a null descriptor
    // if the DACL could not be built) that outlives the call.
    let server = unsafe {
        options.create_with_security_attributes_raw(PIPE_NAME, security.as_ptr())
    }?;
    Ok(server)
}

pub async fn listen<F>(mut on_message: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    let mut first = true;
    loop {
        let mut server = create_server(first)?;
        first = false;
        server.connect().await?;

        // Read at most MAX_FRAME_BYTES + 1 so an oversized frame is detected
        // without buffering the whole (otherwise unbounded) stream.
        let mut buf = Vec::new();
        (&mut server)
            .take(MAX_FRAME_BYTES + 1)
            .read_to_end(&mut buf)
            .await?;

        if buf.len() as u64 > MAX_FRAME_BYTES {
            log::warn!("dropped context-menu frame larger than {MAX_FRAME_BYTES} bytes");
            continue;
        }

        match serde_json::from_slice::<ContextMenuInfo>(&buf) {
            Ok(info) => on_message(info),
            Err(e) => log::warn!("ignored malformed context-menu frame: {e}"),
        }
    }
}
