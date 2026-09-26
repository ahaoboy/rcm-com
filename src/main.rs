use clap::{Parser, Subcommand};
use log::LevelFilter;
use rcm_com::logging;
use rcm_com::{PIPE_NAME, cmd, error::RcmError, server::listen};
use rcm_reg::{MenuStyle, restart_explorer};

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
    /// Show or change the log level (`off`|`error`|`warn`|`info`|`debug`|`trace`)
    Log {
        #[command(subcommand)]
        action: Option<LogAction>,
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
    /// Disable all logging
    Off,
    /// Log errors only
    Error,
    /// Log warnings and errors
    Warn,
    /// Log informational messages and above (default)
    Info,
    /// Log debug messages and above
    Debug,
    /// Log everything, including traces
    Trace,
}

impl LogAction {
    fn filter(&self) -> LevelFilter {
        match self {
            LogAction::Off => LevelFilter::Off,
            LogAction::Error => LevelFilter::Error,
            LogAction::Warn => LevelFilter::Warn,
            LogAction::Info => LevelFilter::Info,
            LogAction::Debug => LevelFilter::Debug,
            LogAction::Trace => LevelFilter::Trace,
        }
    }
}

/// Persist and apply a new log level, then ask the loaded shell extension to
/// adopt it. Failing to reach the extension is not fatal — the level is already
/// stored for the next time it loads.
fn set_log_level(action: LogAction) -> Result<(), RcmError> {
    let filter = action.filter();
    logging::set_level(filter)?;
    let live = rcm_com::try_set_remote_log_level(logging::level_name(filter));
    log::info!("log level set to '{}'", logging::level_name(filter));
    if !live {
        log::warn!("shell extension not updated — the new level applies after it reloads");
    }
    Ok(())
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
        Commands::Start => {
            log::info!("Listening for Explorer context menu events on pipe: {PIPE_NAME}");
            listen(|info| {
                // Human-readable summary on `info`, full struct on `debug`.
                log::info!("{info}");
                log::debug!("{info:#?}");
            })
            .await
        }
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
        Commands::Enable => rcm_com::enable().map(|_| {
            log::info!("Menu blocking ENABLED — native context menu will be hidden.");
        }),
        Commands::Disable => rcm_com::disable().map(|_| {
            log::info!("Menu blocking DISABLED — native context menu will be shown.");
        }),
        Commands::Query => match rcm_com::query() {
            Ok(enabled) => {
                log::info!(
                    "Menu blocking: {}",
                    if enabled { "ENABLED" } else { "DISABLED" }
                );
                Ok(())
            }
            Err(e) => Err(e),
        },
        Commands::Log { action } => match action {
            Some(action) => set_log_level(action),
            None => {
                log::info!("log level: {}", logging::level_name(logging::current_level()));
                Ok(())
            }
        },
    };

    // A non-zero exit code lets scripts and CI detect failure.
    if let Err(e) = result {
        log::error!("{e}");
        std::process::exit(1);
    }
}
