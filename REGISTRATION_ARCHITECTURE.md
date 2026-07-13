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
