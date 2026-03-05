pub mod app;

use anyhow::Result;
use app::App;
use crossterm::event::{EnableMouseCapture, DisableMouseCapture};
use crossterm::execute;

pub async fn run() -> Result<()> {
    let mut terminal = ratatui::init();
    execute!(std::io::stdout(), EnableMouseCapture)?;
    let result = App::new().await?.run(&mut terminal).await;
    execute!(std::io::stdout(), DisableMouseCapture)?;
    ratatui::restore();
    result
}
