use anyhow::Result;
use clap::Args;

#[derive(Debug, Args)]
pub struct DoctorArgs {
    /// Emit JSON. The default output is already JSON when WinCOM is checked.
    #[arg(long)]
    pub json: bool,
}

impl DoctorArgs {
    pub fn run(self) -> Result<()> {
        let output = crate::wincom::office_doctor()?;
        println!("{output}");
        Ok(())
    }
}
