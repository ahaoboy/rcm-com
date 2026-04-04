use tokio::io::AsyncReadExt;
use tokio::net::windows::named_pipe::ServerOptions;

use crate::{ContextMenuInfo, PIPE_NAME};
use crate::error::Result;

pub async fn listen<F>(mut on_message: F) -> Result<()>
where
    F: FnMut(ContextMenuInfo),
{
    let mut server = ServerOptions::new()
        .first_pipe_instance(true)
        .create(PIPE_NAME)?;

    loop {
        server.connect().await?;

        let mut buf = vec![];
        server.read_to_end(&mut buf).await?;

        let json_str = String::from_utf8(buf)?;
        let info = serde_json::from_str::<ContextMenuInfo>(&json_str)?;
        
        on_message(info);

        server = ServerOptions::new().create(PIPE_NAME)?;
    }
}
