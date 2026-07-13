#ifdef _WIN32
#include <windows.h>
#include <stddef.h>
#include <stdint.h>
#include "XLCALL.H"
size_t excel_sdk_sizeof_xlref12(void){return sizeof(XLREF12);}
size_t excel_sdk_alignof_xlref12(void){return __alignof(XLREF12);}
size_t excel_sdk_sizeof_fp12_header(void){return offsetof(FP12,array);}
size_t excel_sdk_sizeof_xloper12(void){return sizeof(XLOPER12);}
size_t excel_sdk_alignof_xloper12(void){return __alignof(XLOPER12);}
size_t excel_sdk_offsetof_xloper12_xltype(void){return offsetof(XLOPER12,xltype);}
uint32_t excel_sdk_xltype_ref(void){return xltypeRef;}
uint32_t excel_sdk_xltype_int(void){return xltypeInt;}
uint32_t excel_sdk_xlbit_xlfree(void){return xlbitXLFree;}
uint32_t excel_sdk_xlbit_dllfree(void){return xlbitDLLFree;}
#endif
