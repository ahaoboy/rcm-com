use clap::{Parser, Subcommand};
use rcm_com::{cmd, error::Result, server::listen, PIPE_NAME};

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

#[tokio::main]
async fn main() {
    let cli = Cli::parse();

    let result: Result<()> = match cli.command {
        Commands::Install => cmd::register(),
        Commands::Uninstall => cmd::unregister(),
        Commands::Start => {
            println!("Listening for Explorer context menu events on pipe: {}", PIPE_NAME);
            listen(|info| {
                println!("{:#?}", info);
            })
            .await
        }
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}
