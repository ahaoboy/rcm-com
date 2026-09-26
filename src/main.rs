use clap::{Parser, Subcommand};
use rcm_com::logging::{self, LogLevel};
use rcm_com::{cmd, error::RcmError, server::listen};
use rcm_reg::{MenuStyle, restart_explorer};
use std::path::PathBuf;

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
    /// Stop blocking — let the native context menu appear
    Disable,
    /// Block the native context menu (default behaviour)
    Enable,
    /// Query whether menu blocking is currently enabled
    Query,
    /// Query or change the log level
    Log {
        #[command(subcommand)]
        action: Option<LogAction>,
    },
    /// Register or query the program that is using the pipe
    Client {
        #[command(subcommand)]
        action: Option<ClientAction>,
    },
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

#[derive(Subcommand)]
enum LogAction {
    /// Show the log level currently used by the shell extension
    Get,
    /// Set the log level (persisted, and pushed to the shell extension)
    Set {
        /// The new log level
        level: LogLevel,
    },
}

#[derive(Subcommand)]
enum ClientAction {
    /// Show the absolute path of the program registered as using the pipe
    Get,
    /// Register a program path as the user of the pipe
    Set {
        /// Path to register; defaults to this executable
        path: Option<PathBuf>,
    },
}

/// Show or change the log level.
///
/// `Get` asks the running shell extension for its *live* level and falls back
/// to this process's level (persisted setting) when it is not loaded. `Set`
/// persists the level and pushes it to the extension if it is running.
async fn handle_log(action: Option<LogAction>) -> Result<(), RcmError> {
    match action.unwrap_or(LogAction::Get) {
        LogAction::Get => match rcm_com::get_log_level().await {
            Ok(level) => {
                log::info!("log level: {level} (shell extension)");
                Ok(())
            }
            Err(_) => {
                log::info!(
                    "log level: {} (local; shell extension not running)",
                    logging::current_level()
                );
                Ok(())
            }
        },
        LogAction::Set { level } => {
            logging::set_level(level)?;
            log::info!("log level set to '{level}'");
            if !rcm_com::try_set_remote_log_level(level).await {
                log::warn!("shell extension not updated — the new level applies after it reloads");
            }
            Ok(())
        }
    }
}

/// Show or register the program that is using the pipe.
async fn handle_client(action: Option<ClientAction>) -> Result<(), RcmError> {
    match action.unwrap_or(ClientAction::Get) {
        ClientAction::Get => match rcm_com::get_client().await {
            Ok(Some(path)) => {
                log::info!("{path}");
                Ok(())
            }
            Ok(None) => {
                log::info!("no program registered");
                Ok(())
            }
            Err(e) => Err(e),
        },
        ClientAction::Set { path } => {
            let path = resolve_path(path)?;
            rcm_com::set_client(path.clone()).await?;
            log::info!("registered pipe user: {path}");
            Ok(())
        }
    }
}

/// Resolve a path argument to an absolute path, defaulting to this executable.
fn resolve_path(path: Option<PathBuf>) -> Result<String, RcmError> {
    let path = match path {
        Some(path) if path.is_absolute() => path,
        Some(path) => std::env::current_dir()
            .map_err(|e| RcmError::Environment(format!("cannot read current directory: {e}")))?
            .join(path),
        None => std::env::current_exe()
            .map_err(|e| RcmError::Environment(format!("cannot resolve executable path: {e}")))?,
    };
    Ok(path.to_string_lossy().into_owned())
}

#[tokio::main]
async fn main() {
    // Install the `log` backend first so that every subsequent message — from
    // this crate and from the library — is handled uniformly.
    logging::init_console();

    let cli = Cli::parse();

    // install / uninstall require elevation
    if matches!(cli.command, Commands::Install | Commands::Uninstall) && !is_admin::is_admin() {
        log::error!("install and uninstall require Administrator privileges");
        log::error!("please run this command from an elevated terminal");
        std::process::exit(1);
    }

    let result = match cli.command {
        Commands::Install => cmd::register(),
        Commands::Uninstall => cmd::unregister(),
        Commands::Start => listen(|info| {
            // Human-readable summary on `info`, full struct on `debug`.
            log::info!("{info}");
            log::debug!("{info:#?}");
        })
        .await,
        Commands::Status => cmd::status().map(|s| {
            log::info!("{s}");
        }),
        Commands::Menu { action } => match action {
            Some(MenuAction::Win10) => MenuStyle::Classic.set().map_err(RcmError::from),
            Some(MenuAction::Win11) => MenuStyle::Windows11.set().map_err(RcmError::from),
            Some(MenuAction::Default { classic }) => if classic {
                MenuStyle::Classic.set()
            } else {
                MenuStyle::Windows11.set()
            }
            .map_err(RcmError::from),
            None => {
                let style = MenuStyle::current();
                log::info!("Menu style:      {style}");
                log::info!("Default classic: {}", matches!(style, MenuStyle::Classic));
                Ok(())
            }
        },
        Commands::RestartExplorer => {
            restart_explorer(std::time::Duration::from_secs(5)).map_err(RcmError::from)
        }
        Commands::Enable => rcm_com::enable().await.map(|_| {
            log::info!("Menu blocking ENABLED — native context menu will be hidden.");
        }),
        Commands::Disable => rcm_com::disable().await.map(|_| {
            log::info!("Menu blocking DISABLED — native context menu will be shown.");
        }),
        Commands::Query => match rcm_com::query().await {
            Ok(enabled) => {
                log::info!(
                    "Menu blocking: {}",
                    if enabled { "ENABLED" } else { "DISABLED" }
                );
                Ok(())
            }
            Err(e) => Err(e),
        },
        Commands::Log { action } => handle_log(action).await,
        Commands::Client { action } => handle_client(action).await,
    };

    // A non-zero exit code lets scripts and CI detect failure.
    if let Err(e) = result {
        log::error!("{e}");
        std::process::exit(1);
    }
}
