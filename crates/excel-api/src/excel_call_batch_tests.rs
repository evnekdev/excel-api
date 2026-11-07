use super::*;
use crate::{ConversionLimits, ExcelValue, OwnedExcelReference, ReferenceArea};
use std::sync::{Mutex, atomic::AtomicBool};

#[derive(Debug, PartialEq)]
enum Argument {
    Reference(OwnedExcelReference),
    Value(ExcelValue),
}
#[derive(Clone, Copy)]
enum Reply {
    Boolean(bool),
    Number,
    Text,
    Error,
    Empty,
    Array,
    InvalidString,
    Unknown,
}
struct Backend {
    linked: AtomicBool,
    code: i32,
    reply: Reply,
    release_code: i32,
    calls: Mutex<Vec<(i32, Vec<Argument>)>>,
    releases: Mutex<usize>,
    text: Mutex<Option<Box<[u16]>>>,
}
impl Backend {
    fn new(reply: Reply) -> Self {
        Self {
            linked: AtomicBool::new(true),
            code: 0,
            reply,
            release_code: 0,
            calls: Mutex::new(Vec::new()),
            releases: Mutex::new(0),
            text: Mutex::new(None),
        }
    }
}
impl ExcelCallBackend for Backend {
    fn link(&self) -> Result<(), ExcelCallError> {
        Ok(())
    }
    fn unlink(&self) {
        self.linked.store(false, Ordering::Relaxed);
    }
    fn is_linked(&self) -> bool {
        self.linked.load(Ordering::Relaxed)
    }
    unsafe fn excel12v_raw(
        &self,
        function: i32,
        result: *mut XLOPER12,
        count: c_int,
        arguments: *mut LPXLOPER12,
    ) -> i32 {
        // SAFETY: test callers uphold the same live-root synchronous backend contract.
        unsafe {
            let args = core::slice::from_raw_parts(arguments, count as usize);
            if function == excel_api_sys::xlFree {
                assert!(result.is_null());
                assert_eq!(count, 1);
                *self.releases.lock().unwrap() += 1;
                let root = &**args.first().unwrap();
                if root.xltype == excel_api_sys::xltypeMulti {
                    let array = root.val.array;
                    drop(Box::from_raw(core::ptr::slice_from_raw_parts_mut(
                        array.lparray,
                        (array.rows * array.columns) as usize,
                    )));
                }
                self.text.lock().unwrap().take();
                return self.release_code;
            }
            let snapshot = args
                .iter()
                .map(|p| {
                    assert_eq!(
                        (**p).xltype & (excel_api_sys::xlbitDLLFree | excel_api_sys::xlbitXLFree),
                        0
                    );
                    let decoded = crate::RawExcelValue::from_callback(&**p).decode().unwrap();
                    match decoded {
                        crate::ExcelValueRef::Reference(r) => {
                            Argument::Reference(OwnedExcelReference::from_borrowed(&r).unwrap())
                        }
                        v => Argument::Value(
                            ExcelValue::from_borrowed_with_limits(v, &ConversionLimits::default())
                                .unwrap(),
                        ),
                    }
                })
                .collect();
            self.calls.lock().unwrap().push((function, snapshot));
            assert!(!result.is_null());
            if self.code != 0 {
                return self.code;
            }
            *result = match self.reply {
                Reply::Boolean(value) => XLOPER12 {
                    val: XLOPER12Value {
                        xbool: i32::from(value),
                    },
                    xltype: excel_api_sys::xltypeBool,
                },
                Reply::Number => XLOPER12 {
                    val: XLOPER12Value { num: 12.5 },
                    xltype: excel_api_sys::xltypeNum,
                },
                Reply::Error => XLOPER12 {
                    val: XLOPER12Value {
                        err: excel_api_sys::xlerrNA,
                    },
                    xltype: excel_api_sys::xltypeErr,
                },
                Reply::Empty => XLOPER12 {
                    val: XLOPER12Value { w: 0 },
                    xltype: excel_api_sys::xltypeNil,
                },
                Reply::Text => {
                    let mut text = counted("Book.xlsx").unwrap();
                    let pointer = text.as_mut_ptr();
                    *self.text.lock().unwrap() = Some(text);
                    XLOPER12 {
                        val: XLOPER12Value { str: pointer },
                        xltype: excel_api_sys::xltypeStr,
                    }
                }
                Reply::Array => {
                    let cells = vec![
                        XLOPER12 {
                            val: XLOPER12Value { num: 7.0 },
                            xltype: excel_api_sys::xltypeNum,
                        },
                        XLOPER12 {
                            val: XLOPER12Value { w: 0 },
                            xltype: excel_api_sys::xltypeNil,
                        },
                    ]
                    .into_boxed_slice();
                    XLOPER12 {
                        val: XLOPER12Value {
                            array: excel_api_sys::XLOPER12Array {
                                lparray: Box::into_raw(cells).cast(),
                                rows: 1,
                                columns: 2,
                            },
                        },
                        xltype: excel_api_sys::xltypeMulti,
                    }
                }
                Reply::InvalidString => XLOPER12 {
                    val: XLOPER12Value {
                        str: ptr::null_mut(),
                    },
                    xltype: excel_api_sys::xltypeStr,
                },
                Reply::Unknown => XLOPER12 {
                    val: XLOPER12Value { w: 0 },
                    xltype: 0x80000000,
                },
            };
            0
        }
    }
}
fn reference() -> OwnedExcelReference {
    OwnedExcelReference::on_sheet(42, vec![ReferenceArea::new(0, 0, 0, 1).unwrap()]).unwrap()
}

#[test]
fn read_copy_limit_releases_the_excel_array_once() {
    let backend = Backend::new(Reply::Array);
    let capability = CallCapability::new(&backend);
    let limits = ConversionLimits {
        max_array_elements: 1,
        ..ConversionLimits::default()
    };
    assert!(matches!(
        crate::WorksheetContext::new(&capability).read_reference(&reference(), &limits),
        Err(ExcelCallError::ResultConversion(_))
    ));
    assert_eq!(*backend.releases.lock().unwrap(), 1);
}

#[test]
fn help_preserves_every_registration_position_including_argument_help() {
    let registration = crate::FunctionRegistration::new(
        "exported",
        "VISIBLE",
        crate::FunctionSignature::new(
            crate::ExcelReturnType::Number,
            &[crate::ExcelArgumentType::Number],
        ),
    )
    .category("Category")
    .description("Function help")
    .arguments(&["input"], &["Argument help"]);
    for help in [None, Some(crate::HelpTopic::new("guide.chm", 42))] {
        let storage =
            CallArguments::registration(&"module.xll".into(), &registration, help).unwrap();
        assert_eq!(storage.roots.len(), 11);
        let expected = [
            "module.xll",
            "exported",
            "BB",
            "VISIBLE",
            "input",
            "",
            "Category",
            "",
            if help.is_some() { "guide.chm!42" } else { "" },
            "Function help",
            "Argument help",
        ];
        for (index, root) in storage.roots.iter().enumerate() {
            // SAFETY: storage owns all roots and their counted strings until this loop ends.
            let value = unsafe { crate::RawExcelValue::from_callback(root) }
                .decode()
                .unwrap();
            let value =
                ExcelValue::from_borrowed_with_limits(value, &ConversionLimits::default()).unwrap();
            assert_eq!(
                value,
                if index == 5 {
                    ExcelValue::Number(1.0)
                } else {
                    ExcelValue::Text(expected[index].into())
                }
            );
        }
    }
}

#[test]
fn macro_sheet_information_permission_does_not_authorize_mutation() {
    let backend = Backend::new(Reply::Text);
    let capability = CallCapability::new(&backend);
    let context = crate::MacroContext::macro_sheet(&capability);
    assert!(matches!(
        context.set_reference(&reference(), ExcelValue::Number(1.0)),
        Err(ExcelCallError::IllegalContext { .. })
    ));
    assert!(matches!(
        context.clear_reference(&reference()),
        Err(ExcelCallError::IllegalContext { .. })
    ));
    assert!(matches!(
        context.calculate_now(),
        Err(ExcelCallError::IllegalContext { .. })
    ));
    assert!(backend.calls.lock().unwrap().is_empty());
    assert!(context.document_info(DocumentInfo::Name, None).is_ok());
}

#[test]
fn reads_preserve_scalar_nil_error_and_array_and_release_once() {
    for reply in [
        Reply::Number,
        Reply::Empty,
        Reply::Error,
        Reply::Text,
        Reply::Array,
    ] {
        let backend = Backend::new(reply);
        let capability = CallCapability::new(&backend);
        let context = crate::WorksheetContext::new(&capability);
        let result = context
            .read_reference(&reference(), &ConversionLimits::default())
            .unwrap();
        match reply {
            Reply::Number => assert_eq!(result, ExcelValue::Number(12.5)),
            Reply::Empty => assert_eq!(result, ExcelValue::Empty),
            Reply::Error => assert_eq!(result, ExcelValue::Error(crate::ExcelError::Na)),
            Reply::Text => assert_eq!(result, ExcelValue::Text("Book.xlsx".into())),
            Reply::Array => {
                let ExcelValue::Array(array) = result else {
                    panic!("array expected")
                };
                assert_eq!((array.rows(), array.columns()), (1, 2));
                assert_eq!(
                    array.values(),
                    &[ExcelValue::Number(7.0), ExcelValue::Empty]
                );
            }
            _ => unreachable!(),
        }
        assert_eq!(
            *backend.calls.lock().unwrap(),
            vec![(16386, vec![Argument::Reference(reference())])]
        );
        assert_eq!(*backend.releases.lock().unwrap(), 1);
    }
}

#[test]
fn read_failures_preserve_codes_and_release_on_conversion_failure() {
    for reply in [Reply::InvalidString, Reply::Unknown] {
        let backend = Backend::new(reply);
        let capability = CallCapability::new(&backend);
        assert!(matches!(
            crate::MacroContext::new(&capability)
                .read_reference(&reference(), &ConversionLimits::default()),
            Err(ExcelCallError::ResultConversion(_))
        ));
        assert_eq!(*backend.releases.lock().unwrap(), 1);
    }
    let mut backend = Backend::new(Reply::Number);
    backend.code = excel_api_sys::xlretUncalced;
    let capability = CallCapability::new(&backend);
    assert!(matches!(
        crate::ThreadSafeContext::new(&capability)
            .read_reference(&reference(), &ConversionLimits::default()),
        Err(ExcelCallError::ExcelFailure {
            code: ExcelReturnCode(excel_api_sys::xlretUncalced),
            ..
        })
    ));
    assert_eq!(*backend.releases.lock().unwrap(), 0);
    let mut backend = Backend::new(Reply::InvalidString);
    backend.release_code = excel_api_sys::xlretFailed;
    let capability = CallCapability::new(&backend);
    assert!(matches!(
        crate::MacroContext::new(&capability)
            .read_reference(&reference(), &ConversionLimits::default()),
        Err(ExcelCallError::ResultConversion(
            crate::ExcelOwnedConversionError::ConversionAndRelease { .. }
        ))
    ));
    assert_eq!(*backend.releases.lock().unwrap(), 1);
}

#[test]
fn write_and_clear_use_exact_roots_with_checked_result_without_transfer() {
    let _allocations = crate::return_alloc::allocation_test_guard();
    let backend = Backend::new(Reply::Boolean(true));
    let capability = CallCapability::new(&backend);
    let context = crate::MacroContext::new(&capability);
    let values = [
        ExcelValue::Number(2.0),
        ExcelValue::Text("=literal".into()),
        ExcelValue::Boolean(true),
        ExcelValue::Error(crate::ExcelError::Na),
        ExcelValue::Array(
            crate::ExcelArray::new(1, 2, vec![ExcelValue::Integer(3), ExcelValue::Empty]).unwrap(),
        ),
    ];
    for value in &values {
        context.set_reference(&reference(), value.clone()).unwrap();
    }
    context.clear_reference(&reference()).unwrap();
    context
        .set_reference(&reference(), ExcelValue::Empty)
        .unwrap();
    let calls = backend.calls.lock().unwrap();
    for (i, value) in values.iter().enumerate() {
        assert_eq!(
            calls[i],
            (
                16387,
                vec![
                    Argument::Reference(reference()),
                    Argument::Value(value.clone())
                ]
            )
        );
    }
    assert_eq!(calls[5], (16387, vec![Argument::Reference(reference())]));
    assert_eq!(calls[6], calls[5]);
    assert_eq!(*backend.releases.lock().unwrap(), 7);
}

#[test]
fn writes_and_clears_reject_false_errors_and_malformed_results() {
    let _allocations = crate::return_alloc::allocation_test_guard();
    for reply in [
        Reply::Boolean(false),
        Reply::Error,
        Reply::Number,
        Reply::Text,
    ] {
        for clear in [false, true] {
            let backend = Backend::new(reply);
            let capability = CallCapability::new(&backend);
            let context = crate::MacroContext::new(&capability);
            let result = if clear {
                context.clear_reference(&reference())
            } else {
                context.set_reference(&reference(), ExcelValue::Integer(1))
            };
            match reply {
                Reply::Boolean(false) => assert_eq!(result, Err(ExcelCallError::CommandRejected)),
                Reply::Error => assert_eq!(
                    result,
                    Err(ExcelCallError::CommandError(crate::ExcelError::Na))
                ),
                _ => assert!(matches!(
                    result,
                    Err(ExcelCallError::MalformedResult {
                        function: "xlSet",
                        ..
                    })
                )),
            }
            assert_eq!(backend.calls.lock().unwrap().len(), 1);
            assert_eq!(*backend.releases.lock().unwrap(), 1);
            assert!(backend.text.lock().unwrap().is_none());
        }
    }
}

#[test]
fn successful_write_does_not_hide_result_release_failure_or_retry() {
    let mut backend = Backend::new(Reply::Boolean(true));
    backend.release_code = excel_api_sys::xlretFailed;
    let capability = CallCapability::new(&backend);
    assert!(matches!(
        crate::MacroContext::new(&capability).clear_reference(&reference()),
        Err(ExcelCallError::ResultConversion(
            crate::ExcelOwnedConversionError::Release(_)
        ))
    ));
    assert_eq!(backend.calls.lock().unwrap().len(), 1);
    assert_eq!(*backend.releases.lock().unwrap(), 1);
}

#[test]
fn invalid_shapes_and_values_do_not_dispatch() {
    let _allocations = crate::return_alloc::allocation_test_guard();
    let backend = Backend::new(Reply::Number);
    let capability = CallCapability::new(&backend);
    let context = crate::MacroContext::new(&capability);
    let multi =
        OwnedExcelReference::on_sheet(1, vec![ReferenceArea::cell(0, 0).unwrap(); 2]).unwrap();
    assert!(
        context
            .read_reference(&multi, &ConversionLimits::default())
            .is_err()
    );
    assert!(
        context
            .set_reference(&multi, ExcelValue::Integer(2))
            .is_err()
    );
    assert!(context.clear_reference(&multi).is_err());
    for value in [
        ExcelValue::Missing,
        ExcelValue::Array(crate::ExcelArray::new(0, 0, vec![]).unwrap()),
        ExcelValue::Array(crate::ExcelArray::new(1, 1, vec![ExcelValue::Missing]).unwrap()),
        ExcelValue::Text("a".repeat(32_768).into()),
    ] {
        assert!(context.set_reference(&reference(), value).is_err());
    }
    assert!(backend.calls.lock().unwrap().is_empty());
}

#[test]
fn typed_document_selectors_validate_positions_results_and_names() {
    for (selector, expected_number) in [
        (DocumentInfo::Name, 1),
        (DocumentInfo::Directory, 2),
        (DocumentInfo::ActiveSheetName, 76),
        (DocumentInfo::ActiveWorkbookName, 88),
        (DocumentInfo::HasChanges, 4),
        (DocumentInfo::ReadOnly, 5),
    ] {
        let boolean = matches!(selector, DocumentInfo::HasChanges | DocumentInfo::ReadOnly);
        let backend = Backend::new(if boolean {
            Reply::Boolean(true)
        } else {
            Reply::Text
        });
        let capability = CallCapability::new(&backend);
        let context = crate::MacroContext::new(&capability);
        let name = if matches!(
            selector,
            DocumentInfo::ActiveSheetName | DocumentInfo::ActiveWorkbookName
        ) {
            None
        } else {
            Some("Book.xlsx")
        };
        let value = context.document_info(selector, name).unwrap();
        assert_eq!(
            value,
            if boolean {
                ExcelValue::Boolean(true)
            } else {
                ExcelValue::Text("Book.xlsx".into())
            }
        );
        let mut args = vec![Argument::Value(ExcelValue::Integer(expected_number))];
        if name.is_some() {
            args.push(Argument::Value(ExcelValue::Text("Book.xlsx".into())));
        }
        assert_eq!(*backend.calls.lock().unwrap(), vec![(188, args)]);
        assert_eq!(*backend.releases.lock().unwrap(), 1);
    }
    let backend = Backend::new(Reply::Error);
    let capability = CallCapability::new(&backend);
    let context = crate::MacroContext::new(&capability);
    assert_eq!(
        context
            .document_info(DocumentInfo::Directory, None)
            .unwrap(),
        ExcelValue::Error(crate::ExcelError::Na)
    );
    for name in ["", "bad\0name"] {
        assert!(
            context
                .document_info(DocumentInfo::Name, Some(name))
                .is_err()
        );
    }
    assert!(
        context
            .document_info(DocumentInfo::ActiveWorkbookName, Some("Book.xlsx"))
            .is_err()
    );
    assert_eq!(backend.calls.lock().unwrap().len(), 1);
    let backend = Backend::new(Reply::Number);
    let capability = CallCapability::new(&backend);
    assert!(matches!(
        crate::MacroContext::new(&capability).document_info(DocumentInfo::Name, None),
        Err(ExcelCallError::MalformedResult { .. })
    ));
    assert_eq!(*backend.releases.lock().unwrap(), 1);
}

#[test]
fn calculate_now_checks_transport_and_result_and_releases_once() {
    for reply in [
        Reply::Boolean(true),
        Reply::Boolean(false),
        Reply::Error,
        Reply::Number,
    ] {
        let backend = Backend::new(reply);
        let capability = CallCapability::new(&backend);
        let result = crate::MacroContext::new(&capability).calculate_now();
        match reply {
            Reply::Boolean(true) => assert!(result.is_ok()),
            Reply::Boolean(false) => assert_eq!(result, Err(ExcelCallError::CommandRejected)),
            Reply::Error => assert_eq!(
                result,
                Err(ExcelCallError::CommandError(crate::ExcelError::Na))
            ),
            _ => assert!(matches!(
                result,
                Err(ExcelCallError::MalformedResult { .. })
            )),
        }
        assert_eq!(*backend.calls.lock().unwrap(), vec![(32799, vec![])]);
        assert_eq!(*backend.releases.lock().unwrap(), 1);
    }
}

#[test]
fn new_commands_reject_illegal_contexts_counts_and_unavailable_backend() {
    let _allocations = crate::return_alloc::allocation_test_guard();
    let backend = Backend::new(Reply::Boolean(true));
    let capability = CallCapability::new(&backend);
    for descriptor in [XL_SET, XLF_GET_DOCUMENT, XLC_CALCULATE_NOW] {
        for permission in [
            CallPermission::Worksheet,
            CallPermission::ThreadSafe,
            CallPermission::Lifecycle,
        ] {
            assert!(matches!(
                capability.call(permission, descriptor, &mut []),
                Err(ExcelCallError::IllegalContext { .. })
            ));
        }
        let mut roots = [ptr::null_mut(); 3];
        assert!(matches!(
            capability.call(CallPermission::Macro, descriptor, &mut roots),
            Err(ExcelCallError::InvalidArgumentCount { .. })
        ));
    }
    assert!(backend.calls.lock().unwrap().is_empty());
    backend.unlink();
    let context = crate::MacroContext::new(&capability);
    assert_eq!(
        context.calculate_now(),
        Err(ExcelCallError::BackendUnavailable)
    );
    assert_eq!(
        context.document_info(DocumentInfo::Name, None),
        Err(ExcelCallError::BackendUnavailable)
    );
    assert_eq!(
        context.set_reference(&reference(), ExcelValue::Integer(1)),
        Err(ExcelCallError::BackendUnavailable)
    );
    assert!(backend.calls.lock().unwrap().is_empty());
}

#[test]
fn backend_failures_are_not_promoted_to_success_or_spurious_releases() {
    let _allocations = crate::return_alloc::allocation_test_guard();
    let mut backend = Backend::new(Reply::Boolean(true));
    backend.code = excel_api_sys::xlretFailed;
    let capability = CallCapability::new(&backend);
    let context = crate::MacroContext::new(&capability);
    assert!(matches!(
        context.calculate_now(),
        Err(ExcelCallError::ExcelFailure { .. })
    ));
    assert!(matches!(
        context.document_info(DocumentInfo::Name, None),
        Err(ExcelCallError::ExcelFailure { .. })
    ));
    assert!(matches!(
        context.set_reference(&reference(), ExcelValue::Number(1.0)),
        Err(ExcelCallError::ExcelFailure { .. })
    ));
    assert!(matches!(
        context.clear_reference(&reference()),
        Err(ExcelCallError::ExcelFailure { .. })
    ));
    assert_eq!(*backend.releases.lock().unwrap(), 0);
}

#[test]
fn return_to_excel_status_survives_every_release_error_wrapper() {
    for code in [
        excel_api_sys::xlretAbort,
        excel_api_sys::xlretUncalced | excel_api_sys::xlretFailed,
        excel_api_sys::xlretFailed,
    ] {
        let release = crate::ExcelReleaseError::ExcelCallFailure { code };
        let errors = [
            ExcelCallError::ExcelFailure {
                function: "fixture",
                code: ExcelReturnCode(code),
            },
            ExcelCallError::RegistrationRelease {
                id_bits: 1.0_f64.to_bits(),
                release: release.clone(),
            },
            ExcelCallError::AcceptedCleanupRelease {
                function: "fixture",
                release: release.clone(),
            },
            ExcelCallError::CleanupRejected {
                function: "fixture",
                actual_type: 4,
                excel_error: None,
                release: Some(release.clone()),
            },
            ExcelCallError::ResultRelease(release.clone()),
            ExcelCallError::ResultConversion(crate::ExcelOwnedConversionError::Release(
                release.clone(),
            )),
            ExcelCallError::ResultConversion(
                crate::ExcelOwnedConversionError::ConversionAndRelease {
                    conversion: crate::ConversionError::UnsupportedReference,
                    release,
                },
            ),
        ];
        for error in errors {
            assert_eq!(
                error.requires_return_to_excel(),
                code != excel_api_sys::xlretFailed,
                "{error:?}"
            );
        }
    }
    assert!(
        !ExcelCallError::ResultRelease(crate::ExcelReleaseError::BackendUnavailable)
            .requires_return_to_excel()
    );
}
