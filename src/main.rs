use clap::{Parser, Subcommand};

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
    /// Register the shell extension (requires admin)
    #[command(alias = "reg")]
    Register {
        /// The program to execute when invoked
        program: String,

        /// Arguments for the program
        #[arg(short, long)]
        args: Option<String>,

        /// CID
        #[arg(short, long)]
        cid: String,
    },
    /// Unregister the shell extension (requires admin)
    #[command(alias = "unreg")]
    Unregister,
}

fn main() {
    let cli = Cli::parse();

    let result = match cli.command {
        Commands::Register { program, args, cid } => cmd::register(program, args, cid),
        Commands::Unregister => cmd::unregister(),
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}
