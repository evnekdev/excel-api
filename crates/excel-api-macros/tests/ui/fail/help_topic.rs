use excel_api_macros::excel_function;

#[excel_function(name = "BAD.HELP", thunk = "bad_help", help_topic = "guide.chm")]
fn no_context() -> f64 { 0.0 }

#[excel_function(name = "BAD.EMPTY", thunk = "bad_empty", help_topic = "!0")]
fn empty_resource() -> f64 { 0.0 }

#[excel_function(name = "BAD.ID", thunk = "bad_id", help_topic = "guide.chm!4294967296")]
fn overflow() -> f64 { 0.0 }

#[excel_function(name = "BAD.SIGN", thunk = "bad_sign", help_topic = "guide.chm!-1")]
fn signed_id() -> f64 { 0.0 }

#[excel_function(name = "BAD.DUP", thunk = "bad_dup", help_topic = "guide.chm!1", help_topic = "guide.chm!2")]
fn duplicate() -> f64 { 0.0 }

#[excel_function(name = "BAD.HOST", thunk = "bad_host", help_topic = "https:///help.chm!0")]
fn empty_host() -> f64 { 0.0 }

#[excel_function(name = "BAD.SPACE", thunk = "bad_space", help_topic = "https://example.test/bad path!0")]
fn url_space() -> f64 { 0.0 }

#[excel_function(name = "BAD.CONTROL", thunk = "bad_control", help_topic = "bad\nfile.chm!0")]
fn file_control() -> f64 { 0.0 }

fn main() {}
