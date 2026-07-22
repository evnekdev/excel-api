# Official Microsoft sample summary

The reference is Microsoft's previous-versions C++ Automation article, "How to automate Excel from C++ without using MFC or #import". The downloaded page is retained outside the repository; source provenance is recorded in `documentation-sources.jsonl`.

The faithful baseline is `CoInitialize(NULL)`, `CLSCTX_LOCAL_SERVER`, `IID_IDispatch`, `LOCALE_USER_DEFAULT` (1024) name lookup, `LOCALE_SYSTEM_DEFAULT` (2048) invocation, and `DISPATCH_PROPERTYGET` for `Workbooks.Add`.
