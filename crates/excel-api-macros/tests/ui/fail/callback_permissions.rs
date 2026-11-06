use excel_api::{thunk, MacroContext};
use excel_api_macros::excel_command;

#[excel_command(name = "COMMAND", thunk = "command")]
fn command(_context: &MacroContext<'_>) {}

fn main() {
    thunk::with_callback(|scope| {
        scope.with_worksheet_context(|_| ());
        scope.with_thread_safe_context(|_| ());
        scope.with_macro_context(|_| ());
        scope.with_command_context(|_| ());
    });
    __excel_command_thunk_command();
}
