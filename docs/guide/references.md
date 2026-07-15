# References

```rust,no_run
use excel_api::ExcelReferenceArg;

fn keep_reference<'call>(input: ExcelReferenceArg<'call>) {
    let _value = input.into_inner(); // valid only for this Excel callback
}
```

References are not values. A borrowed `ExcelReference<'call>` represents either
a single area or a multi-area reference, and remains valid only during its
callback. `ExcelReferenceArg<'call>` is the macro-level `U` argument mapping;
it preserves reference semantics instead of forcing a value conversion.

Do not send a reference to a worker, retain its raw pointer, or infer that a
reference is safe in every callback context. If a later operation needs data,
obtain a legal value under the correct typed Excel context and copy it into an
owned value.
