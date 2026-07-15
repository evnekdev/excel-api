# Registration Architecture

## Descriptor

```rust
FunctionRegistration {
    rust_symbol,
    excel_name,
    argument_names,
    category,
    descriptions,
    help_topic,
    signature,
    flags,
}
```

## Signature model

The descriptor does not store an arbitrary type-text string as the source of
truth. It stores typed metadata from which verified type text is generated.

## General argument families

The modern C API distinguishes reference-preserving and value-only general
arguments.

The registration layer must preserve this choice explicitly because it changes
what Excel passes to the function.

## Flags

- volatile;
- thread-safe;
- macro-sheet equivalent;
- cluster-safe where supported;
- modify-in-place only in a future expert API.

Reject incompatible combinations, especially thread-safe plus macro-sheet
permissions.

## Registration call

Build arguments to `xlfRegister` in a stable owned vector and call through
`Excel12v`.

Store registration IDs for later unregistration.

## Commands

Commands use separate descriptor/registration semantics and are not worksheet
functions with a flag.

M12 models `CommandRegistration` separately from `FunctionRegistration`.
`#[excel_command]` accepts only a `&MacroContext`, produces a no-argument
`short WINAPI` export, and records the verified `I` return ABI with
`pxMacroType = 2`. A command returns 1 on success and 0 for an ordinary error
or panic. Function and command registration IDs remain separate and rollback
in reverse registration order.

## Function Wizard

Support:

- category;
- function description;
- argument names;
- argument help;
- help topic.

## Explicit registry first

Use an explicit add-in descriptor/list before introducing linker-section or
inventory-based distributed registration.

## M9A metadata generation

`#[excel_function]` now generates one deterministic hidden typed descriptor
while preserving the annotated Rust function. Attribute text supplies Excel
name, category, description, future thunk symbol, flags, and complete argument
help. The Rust signature supplies the closed argument/result families; an
explicit return registration override is available for the frozen M8 `Q`
return oracle. No thunk, export, callback code, or unsafe code is generated.

## M9B generated thunks

The macro's closed argument/result kinds now drive both the registration
families and raw ABI signature. `A` uses a 16-bit signed Boolean, `B` uses
`f64`, `J` uses `i32`, `C%`/`D%` use UTF-16 pointers, and `Q`/`U` use
`XLOPER12*`, matching Microsoft's xlfRegister type table. The `thunk` literal
must be a non-empty ASCII x64 export identifier and is emitted exactly with
`export_name`; a duplicate export fails at link time.

The minimal XLL's production descriptor list now consists of generated
metadata. Test-only handwritten descriptors remain the frozen registration
oracle, and PE inspection proves the final DLL contains exactly the same 12
named exports as M8.

## M8 implementation

`FunctionSignature` is the source of truth. It distinguishes scalar `B/A/J`,
value-only `Q`, reference-preserving `U`, counted `D%`, and NUL-terminated
`C%`. Return type, arguments, then canonical `!`, `#`, `$`, `&` modifiers form
the type text; macro-sheet plus thread-safe is rejected.

Registration owns every counted UTF-16 argument buffer and raw root until the
synchronous call returns. The fields are module, procedure, type text, Excel
name, comma-separated argument names, macro type, category, shortcut, help
topic, function description, and argument descriptions. Successful IDs are
stored in runtime state; failure rolls them back in reverse order. Close first
deletes each `pxFunctionText` hidden name through one-argument `xlfSetName`, as
required by Microsoft, then unregisters its stored ID.

## M16 asynchronous registration

`ExcelReturnType::AsyncVoid` emits the documented leading `>` and
`ExcelArgumentType::AsyncHandle` emits `X`. Validation requires the pair with
exactly one handle, excludes the hidden handle from Function Wizard argument
metadata, and rejects cluster-safe async registration. The sample's
`RUST.ASYNC.DOUBLE` is exactly `>BX$`.
