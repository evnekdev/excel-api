# Packaging Architecture

Status: planned.

Responsibilities:

- build `cdylib`;
- produce `.xll`;
- verify exports;
- generate/link `.def` when required;
- include version resources;
- package x64 artifacts;
- optional code signing;
- emit diagnostics and reproducible metadata.
