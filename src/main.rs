use std::process::ExitCode;

use clap::Parser;

use docxai::cli::Cli;

fn main() -> ExitCode {
    let cli = Cli::parse();
    match docxai::run(cli) {
        Ok(()) => ExitCode::SUCCESS,
        Err(err) => {
            eprintln!("{err}");
            ExitCode::from(err.exit_code() as u8)
        }
    }
}
