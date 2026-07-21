# Typelib validation status

The ordinary `excel-com-typelib-audit check` deliberately uses `not-recorded` environment inputs and therefore reports the historical `typelib/SOURCE_MANIFEST.toml` stale. The companion `check-historical` command reads only the committed historical environment labels, then re-inspects the current registered typelib without writing any evidence. This preserves the 05B–05E artifact while validating its deterministic type-library content.
