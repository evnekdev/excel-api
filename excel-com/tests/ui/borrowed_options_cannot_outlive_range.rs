use excel_com::{
    Range,
    tables::{ListObjectAddOptions, TableHeaderMode},
};

fn invalid_options<'a>(source: &'a Range) -> ListObjectAddOptions<'static> {
    ListObjectAddOptions {
        source: &source,
        has_headers: TableHeaderMode::YES,
        destination: None,
        table_style_name: None,
    }
}

fn main() {}
