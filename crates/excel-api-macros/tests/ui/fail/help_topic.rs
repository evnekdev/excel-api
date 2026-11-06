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

fn main() {}
