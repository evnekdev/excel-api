use excel_com::AttachedApplication;

fn attached_session_cannot_quit(session: AttachedApplication<'_>) {
    session.quit();
}

fn main() {}
