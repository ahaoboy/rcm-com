use clap::{Parser, Subcommand};
use tokio::io::AsyncReadExt;
use tokio::net::windows::named_pipe::ServerOptions;

mod cmd;

#[derive(Parser)]
#[command(name = "rcm")]
#[command(about = "RCM Context Menu - Shell Extension Registration Tool", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Install and register the shell extension (requires admin)
    Install,
    /// Uninstall and unregister the shell extension (requires admin)
    Uninstall,
    /// Start listening for context menu events via named pipe
    Start,
}

const PIPE_NAME: &str = r"\\.\pipe\rcm_com_pipe";

#[tokio::main]
async fn main() {
    let cli = Cli::parse();

    let result: Result<(), String> = match cli.command {
        Commands::Install => cmd::register(),
        Commands::Uninstall => cmd::unregister(),
        Commands::Start => {
            if let Err(e) = run_server().await {
                Err(format!("Server error: {e}"))
            } else {
                Ok(())
            }
        }
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}

async fn run_server() -> std::io::Result<()> {
    println!("Listening for Explorer context menu events on pipe: {}", PIPE_NAME);
    
    // Create the first instance of the server pipe
    let mut server = ServerOptions::new()
        .first_pipe_instance(true)
        .create(PIPE_NAME)?;

    loop {
        // Wait for a client to connect
        server.connect().await?;

        // Read all data from the connected client
        let mut buf = vec![];
        server.read_to_end(&mut buf).await?;

        if let Ok(json_str) = String::from_utf8(buf) {
            println!("{}", json_str);
        } else {
            eprintln!("Received invalid UTF-8 payload");
        }

        // Drop the old server instance
        // Create a new pipe server instance for the next client
        server = ServerOptions::new().create(PIPE_NAME)?;
    }
}
