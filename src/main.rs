use clap::{Parser, Subcommand};
use rcm_com::{PIPE_NAME, cmd, server::listen};

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
    /// Show current registration status and configuration
    Status,
}

#[tokio::main]
async fn main() {
    let cli = Cli::parse();

    let result = match cli.command {
        Commands::Install => cmd::register(),
        Commands::Uninstall => cmd::unregister(),
        Commands::Start => {
            println!(
                "Listening for Explorer context menu events on pipe: {}",
                PIPE_NAME
            );
            listen(|info| {
                println!("{:#?}", info);
            })
            .await
        }
        Commands::Status => cmd::status().map(|s| {
            println!("{s}");
        }),
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
    }
}
