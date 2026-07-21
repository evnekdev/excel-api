use std::env;

fn main() {
    if let Err(error) = excel_com_pre_add_delta::run(env::args().skip(1).collect()) {
        eprintln!("excel-com-pre-add-delta: {error}");
        std::process::exit(1);
    }
}
