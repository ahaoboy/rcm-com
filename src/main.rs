use clap::{Parser, Subcommand};
use rcm_com::{PIPE_NAME, cmd, get_menu_style, is_default_classic, restart_explorer, server::listen, set_default_classic_menu, set_win11_menu_style};

#[derive(Parser)]
#[command(name = "rcm")]
#[command(about = "RCM Context Menu - Shell Extension Registration Tool", long_about = None)]
#[command(version)]
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
    /// Switch right-click menu or show current style
    Menu {
        #[command(subcommand)]
        action: Option<MenuAction>,
    },
    /// Restart Windows Explorer (stop, wait 5s, start)
    RestartExplorer,
}

#[derive(Subcommand)]
enum MenuAction {
    /// Use Windows 10 classic expanded context menu
    Win10,
    /// Use Windows 11 default compact context menu
    Win11,
    /// Set whether the classic menu is the default
    Default {
        /// Whether to use classic (Win10) style by default
        #[arg(short, long, default_value = "true")]
        classic: bool,
    },
}

#[tokio::main]
async fn main() {
    let cli = Cli::parse();

    // install / uninstall require elevation
    if matches!(cli.command, Commands::Install | Commands::Uninstall) && !is_admin::is_admin() {
        eprintln!("Error: install and uninstall require Administrator privileges.");
        eprintln!("Please run this command from an elevated terminal.");
        std::process::exit(1);
    }

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
        Commands::Menu { action } => match action {
            Some(MenuAction::Win10) => set_win11_menu_style(true),
            Some(MenuAction::Win11) => set_win11_menu_style(false),
            Some(MenuAction::Default { classic }) => set_default_classic_menu(classic),
            None => {
                println!("Menu style:     {}", get_menu_style());
                println!("Default classic: {}", is_default_classic());
                Ok(())
            }
        },
        Commands::RestartExplorer => restart_explorer(),
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
    }
}
