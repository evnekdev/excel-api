/*
**  File:           SAMPLE\EXAMPLES\EXAMPLE.C
**  Description:    XLL SDK Example
**
**
**  This file uses the framework library
**  (frmwrk32.lib).
**
*/

#include <ctype.h>
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>

/*
** rgFuncs
**
** This is a table of all the functions exported by this module.
** These functions are all registered (in xlAutoOpen) when you
** open the XLL. Before every string, leave a space for the
** byte count. The format of this table is the same as
** the last seven arguments to the REGISTER function.
** rgFuncsRows define the number of rows in the table. The
** dimension [3] represents the number of columns in the table.
*/
#define rgFuncsRows 29

static LPWSTR rgFuncs[rgFuncsRows][7] = {
	{L"CallerExample",				L"I", L"CallerExample"},
	{L"debugPrintfExample",			L"I", L"debugPrintfExample"},
	{L"EvaluateExample",			L"I", L"EvaluateExample"},
	{L"Excel12fExample",			L"I", L"Excel12fExample"},
	{L"Excel12Example",				L"I", L"Excel12Example"},
	{L"InitFrameworkExample",		L"I", L"InitFrameworkExample"},
	{L"TempActiveCellExample",		L"I", L"TempActiveCellExample"},
	{L"TempActiveColumnExample",	L"I", L"TempActiveColumnExample"},
	{L"TempActiveRefExample",		L"I", L"TempActiveRefExample"},
	{L"TempActiveRowExample",		L"I", L"TempActiveRowExample"},
	{L"TempBoolExample",			L"I", L"TempBoolExample"},
	{L"TempErrExample",				L"P", L"TempErrExample"},
	{L"TempIntExample",				L"I", L"TempIntExample"},
	{L"TempMissingExample",			L"I", L"TempMissingExample"},
	{L"TempNumExample",				L"I", L"TempNumExample"},
	{L"TempStrExample",				L"I", L"TempStrExample"},
	{L"fArray",						L"Q", L"fArray"},
	{L"xlCoerceExample",			L"II",L"xlCoerceExample"},
	{L"xlFreeExample",				L"I", L"xlFreeExample"},
	{L"xlGetInstExample",			L"I", L"xlGetInstExample"},
	{L"xlGetInstPtrExample",		L"I", L"xlGetInstPtrExample"},
	{L"xlGetNameExample",			L"I", L"xlGetNameExample"},
	{L"xlSetExample",				L"II",L"xlSetExample"},
	{L"xlSheetIdExample",			L"I", L"xlSheetIdExample"},
	{L"xlSheetNmExample",			L"I", L"xlSheetNmExample"},
	{L"xlStackExample",				L"I", L"xlStackExample"},
	{L"xlUDFExample",				L"I", L"xlUDFExample"},
	{L"InternationalExample",		L"I", L"InternationalExample"},
	{L"CalcCircum",					L"BB", L"CalcCircum"}
};

/*
** DllMain
**
** This function is called by LibEntry which is called
** by Windows when the DLL is first loaded. LibEntry initializes the
** DLL's heap if a HEAPSIZE is specified in the DLL's .DEF file, and
** then calls DllMain. The following DllMain function satisfies that
** call. The DllMain function should perform additional initialization
** tasks required by the DLL. In this example, we byte-count all the strings
** in the preceding table. DllMain will be called will be called once per 
** process. In Win32 DllMain replaces both the LibMain and WEP functions.
**
** Arguments:
**
**      HANDLE hInstance			Instance handle
**      ULONG ul_reason_for_call	Reason DllMain was called
**      LPVOID lpReserved			Reserved
**
** Returns:
**
**      int                 1 if initialization is successful.
*/
BOOL WINAPI DllMain(HANDLE hInstance, ULONG ul_reason_for_call,
		   LPVOID lpReserved)
{

	/* Deprecated due to MSVC handling of statically allocated strings */

	/*
	** In the following for loops, the strings in rgFuncs[] are byte-counted
	** so that they won't need to be byte-counted later.
	*/

	/*if(ul_reason_for_call == DLL_PROCESS_ATTACH)
	{
		int i,j;

		for (i = 0; i < rgFuncsRows; i++) 
		{
			for (j = 0; j < 7; j++) 
			{
				rgFuncs[i][j][0] = (BYTE)lstrlen (rgFuncs[i][j]+1);
			}
		}
	}*/

	return 1;
}


/*
** xlAutoOpen
**
** xlAutoOpen is how Microsoft Excel loads XLL files.
** When you open an XLL, Microsoft Excel calls the xlAutoOpen
** function, and nothing more.
**
** More specifically, xlAutoOpen is called by Microsoft Excel:
**
**  - when you open this XLL file from the File menu,
**  - when this XLL is in the XLSTART directory, and is
**		automatically opened when Microsoft Excel starts,
**  - when Microsoft Excel opens this XLL for any other reason, or
**  - when a macro calls REGISTER(), with only one argument, which is the
**		name of this XLL.
**
** xlAutoOpen is also called by the Add-in Manager when you add this XLL
** as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
** REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.
**
** xlAutoOpen should:
**
**  - register all the functions you want to make available while this
**		XLL is open,
**
**  - add any menus or menu items that this XLL supports,
**
**  - perform any other initialization you need, and
**
**  - return 1 if successful, or return 0 if your XLL cannot be opened.
*/
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{

	static XLOPER12 xDLL;	/* name of this DLL */
	int i;					/* Loop index */

	/*
	** In the following block of code the name of the XLL is obtained by
	** calling xlGetName. This name is used as the first argument to the
	** REGISTER function to specify the name of the XLL. Next, the XLL loops
	** through the rgFuncs[] table, registering each function in the table using
	** xlfRegister. Functions must be registered before you can add a menu
	** item.
	*/

	Excel12f(xlGetName, &xDLL, 0);

        for (i=0;i<rgFuncsRows;i++) 
		{
			Excel12f(xlfRegister, 0, 4,
				(LPXLOPER12)&xDLL,
				(LPXLOPER12)TempStr12(rgFuncs[i][0]),
				(LPXLOPER12)TempStr12(rgFuncs[i][1]),
				(LPXLOPER12)TempStr12(rgFuncs[i][2]));
		}

	/* Free the XLL filename */
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

	return 1;
}

/*
** xlAutoClose
**
** xlAutoClose is called by Microsoft Excel:
**
**  - when you quit Microsoft Excel, or
**  - when a macro sheet calls UNREGISTER(), giving a string argument
**		which is the name of this XLL.
**
** xlAutoClose is called by the Add-in Manager when you remove this XLL from
** the list of loaded add-ins. The Add-in Manager first calls xlAutoRemove,
** then calls UNREGISTER("EXAMPLE.XLL"), which in turn calls xlAutoClose.
**
**
** xlAutoClose should:
**
**  - Remove any menus or menu items that were added in xlAutoOpen,
**
**  - do any necessary global cleanup, and
**
**  - delete any names that were added (names of exported functions, and
**		so on). Remember that registering functions may cause names to be created.
**
** xlAutoClose does NOT have to unregister the functions that were registered
** in xlAutoOpen. This is done automatically by Microsoft Excel after
** xlAutoClose returns.
**
** xlAutoClose should return 1.
*/
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
	int i;

	/*
	** This block first deletes all names added by xlAutoOpen or by
	** xlAutoRegister.
	*/

	for (i = 0; i < rgFuncsRows; i++)
		Excel12f(xlfSetName, 0, 1, TempStr12(rgFuncs[i][2]));

	return 1;
}

///***************************************************************************
// lpwstricmp()
//
// Purpose: Compares a pascal string and a null-terminated C-string to see
// if they are equal.  Method is case insensitive
//
// Parameters:
//
//      LPWSTR s     First string (null-terminated)
//      LPWSTR t     Second string (byte counted)
//
// Returns: 
//
//      int         0 if they are equal
//                  Nonzero otherwise
//
// Comments:
//
//      Unlike the usual string functions, lpwstricmp
//      doesn't care about collating sequence.
//
// History:  Date       Author        Reason
///***************************************************************************

int lpwstricmp(LPWSTR s, LPWSTR t)
{
	int i;

	if (wcslen(s) != *t)
		return 1;

	for (i = 1; i <= s[0]; i++)
	{
		if (towlower(s[i-1]) != towlower(t[i]))
			return 1;
	}
	return 0;
}

/*
** xlAutoRegister
**
** This function is called by Microsoft Excel if a macro sheet tries to
** register a function without specifying the type_text argument. If that
** happens, Microsoft Excel calls xlAutoRegister, passing the name of the
** function that the user tried to register. xlAutoRegister should use the
** normal REGISTER function to register the function, but this time it must
** specify the type_text argument. If xlAutoRegister does not recognize the
** function name, it should return a #VALUE! error. Otherwise, it should
** return whatever REGISTER returned.
**
** Arguments:
**
**	    LPXLOPER12 pxName   xltypeStr containing the
**                          name of the function
**                          to be registered. This is not
**                          case sensitive, because
**                          Microsoft Excel uses Pascal calling
**                          convention.
**
** Returns:
**
**      LPXLOPER12          xltypeNum containing the result
**                          of registering the function,
**                          or xltypeErr containing #VALUE!
**                          if the function could not be
**                          registered.
*/
__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xDLL, xRegId;
	int i;

	/*
	** This block initializes xRegId to a #VALUE! error first. This is done in
	** case a function is not found to register. Next, the code loops through the
	** functions in rgFuncs[] and uses lpstricmp to determine if the current
	** row in rgFuncs[] represents the function that needs to be registered.
	** When it finds the proper row, the function is registered and the
	** register ID is returned to Microsoft Excel. If no matching function is
	** found, an xRegId is returned containing a #VALUE! error.
	*/

	xRegId.xltype = xltypeErr;
	xRegId.val.err = xlerrValue;

	for (i = 0; i < rgFuncsRows; i++) 
	{
		if (!lpwstricmp(rgFuncs[i][0], pxName->val.str)) 
		{
			Excel12f(xlGetName, &xDLL, 0);

			Excel12f(xlfRegister, 0, 4,
		  		(LPXLOPER12)&xDLL,
		  		(LPXLOPER12)TempStr12(rgFuncs[i][0]),
		  		(LPXLOPER12)TempStr12(rgFuncs[i][1]),
		  		(LPXLOPER12)TempStr12(rgFuncs[i][2]));

			/* Free the XLL filename */
			Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

			return (LPXLOPER12)&xRegId;
		}
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms

	return (LPXLOPER12)&xRegId;
}

/*
** xlAutoAdd
**
** This function is called by the Add-in Manager only. When you add a
** DLL to the list of active add-ins, the Add-in Manager calls xlAutoAdd()
** and then opens the XLL, which in turn calls xlAutoOpen.
**
*/
__declspec(dllexport) int WINAPI xlAutoAdd(void)
{
	XCHAR szBuf[255];

	wsprintfW((LPWSTR)szBuf, L"Thank you for adding Example.XLL\n build date %hs, time %hs",__DATE__, __TIME__);

	/* Display a dialog box indicating that the XLL was successfully added */
	Excel12f(xlcAlert, 0, 2, TempStr12(szBuf), TempInt12(2));
	return 1;
}

/*
** xlAutoRemove
**
** This function is called by the Add-in Manager only. When you remove
** an XLL from the list of active add-ins, the Add-in Manager calls
** xlAutoRemove() and then UNREGISTER("EXAMPLE.XLL").
**
** You can use this function to perform any special tasks that need to be
** performed when you remove the XLL from the Add-in Manager's list
** of active add-ins. For example, you may want to delete an
** initialization file when the XLL is removed from the list.
*/
__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
	/* Display a dialog box indicating that the XLL was successfully removed */
	Excel12f(xlcAlert, 0, 2, TempStr12(L"Thank you for removing Example.XLL!"), TempInt12(2));
	return 1;
}

/* xlAddInManagerInfo12
**
**
** This function is called by the Add-in Manager to find the long name
** of the add-in. If xAction = 1, this function should return a string
** containing the long name of this XLL, which the Add-in Manager will use
** to describe this XLL. If xAction = 2 or 3, this function should return
** #VALUE!.
**
** Arguments
**
**      LPXLOPER12 xAction    The information you want; either
**                          1 = the long name of the
**                              add in, or
**                          2 = reserved
**                          3 = reserved
**
** Return value
**
**      LPXLOPER12            The long name or #VALUE!.
**
*/
__declspec(dllexport) LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction;

	/*
	** This code coerces the passed-in value to an integer. This is how the
	** code determines what is being requested. If it receives a 1, it returns a
	** string representing the long name. If it receives anything else, it
	** returns a #VALUE! error.
	*/

	Excel12f(xlCoerce, &xIntAction, 2, xAction, TempInt12(xltypeInt));

	if(xIntAction.val.w == 1) 
	{
		xInfo.xltype = xltypeStr;
		xInfo.val.str = L"\026Example Standalone DLL";
	}
	else 
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms

	return (LPXLOPER12)&xInfo;
}

/*
** CallerExample
**
** This function illustrates the xlfCaller function, which
** selects the cell from which it was called.
**
*/
__declspec(dllexport) short WINAPI CallerExample(void)
{
	XLOPER12 xRes;

	Excel12(xlfCaller, &xRes, 0);
	Excel12(xlcSelect, 0, 1, (LPXLOPER12)&xRes);
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}


/*
** debugPrintfExample
**
** This function prints a string to show that control was passed to the
** function.  The _DEBUG flag must be defined before compiling or else this
** function does nothing.
*/
__declspec(dllexport) short WINAPI debugPrintfExample(void)
{

#ifdef _DEBUG
	debugPrintf("Made it!\r");
#endif

	return 1;
}


/*
** EvaluateExample
**
** This function demonstrates using evaluate to look up a cell reference. It
** uses xlfEvaluate to coerce the text "!B38" to the contents of cell B38. 
**
*/
__declspec(dllexport) short WINAPI EvaluateExample(void)
{
	XLOPER12 xFormulaText, xRes, xRes2, xInt;

	xFormulaText.xltype = xltypeStr;
	xFormulaText.val.str = L"\004!B38";
	Excel12(xlfEvaluate, &xRes, 1, (LPXLOPER12)&xFormulaText);

	xInt.xltype = xltypeInt;
	xInt.val.w = 2;
	Excel12(xlcAlert, &xRes2, 2, (LPXLOPER12)&xRes, (LPXLOPER12)&xInt);
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes);
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes2);

	return 1;
}

/*
** Excel12fExample
**
** This function demonstrates the debugging capability of Excel12f().
** It passes a bad argument to the Excel12f() function, which outputs
** a debug message on the terminal attached to COM1.
**
**

*/
__declspec(dllexport) short WINAPI Excel12fExample(void)
{
	Excel12f(xlcDisplay, 0, 1, 0);
	return 1;
}

/*
** Excel12Example
**
** This function illustrates the Excel12() function. It
** uses Excel12() to select the cell from which it was called.
**
*/
__declspec(dllexport) short WINAPI Excel12Example(void)
{
	XLOPER12 xRes;

	Excel12(xlfCaller, &xRes, 0);
	Excel12(xlcSelect, 0, 1, (LPXLOPER12)&xRes);
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes);

	return 1;
}

/*
** InitFrameworkExample
**
** This function uses InitFramework() to free all temporary memory.
*/
__declspec(dllexport) short WINAPI InitFrameworkExample(void)
{
	InitFramework();
	return 1;
}

/*
** TempActiveCellExample
**
** This function uses TempActiveCell() to display the contents of
** cell B94 on the active sheet.
**
*/
__declspec(dllexport) short WINAPI TempActiveCellExample(void)
{
	Excel12f(xlcAlert, 0, 1, TempActiveCell12(93,1));
	return 1;
}

/*
** TempActiveColumnExample
**
** This function uses TempActiveColumn() to select an entire column.
**
*/
__declspec(dllexport) short WINAPI TempActiveColumnExample(void)
{
	Excel12f(xlcSelect, 0, 1, TempActiveColumn12(1));
	return 1;
}

/*
** TempActiveRefExample
**
** This function uses TempActiveRef() to select A105:C110.
**
*/
__declspec(dllexport) short WINAPI TempActiveRefExample(void)
{
	Excel12f(xlcSelect, 0, 1, TempActiveRef12(104, 109, 0, 2));
	return 1;
}

/*
** TempActiveRowExample
**
** This function uses TempActiveRow() to select an entire row.
**
*/
__declspec(dllexport) short WINAPI TempActiveRowExample(void)
{
	Excel12f(xlcSelect, 0, 1, TempActiveRow12(112));
	return 1;
}

/*
** TempBoolExample
**
** This function uses TempBool to clear the status bar. Temporary
** memory is freed when Excel12f() is called.
**
*/
__declspec(dllexport) short WINAPI TempBoolExample(void)
{
	Excel12f(xlcMessage, 0, 1, TempBool12(0));
	return 1;
}

/*
** TempErrExample
**
** This function uses TempErr to return a #VALUE! error to Microsoft Excel.
**
** This function can cause a memory leak as temp memory allocated by TempErr12
** may never be freed. 
**
*/
__declspec(dllexport) LPXLOPER12 WINAPI TempErrExample(void)
{
	return TempErr12(xlerrValue);
}

/*
** TempIntExample
**
** This function uses TempInt() to pass an argument to xlfGetWorkspace.
*/
__declspec(dllexport) short WINAPI TempIntExample(void)
{
	XLOPER12 xRes;

	Excel12f(xlfGetWorkspace, &xRes, 1, TempInt12(44));
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** TempMissingExample
**
** This example uses TempMissing12 to provide 3 missing arguments to
** xlcWorkspace followed by a Boolean FALSE to suppress the display of
** worksheet scroll-bars. The first 3 arguments correspond to other 
** workspace settings which are unaffected
**
*/
__declspec(dllexport) short WINAPI TempMissingExample(void)
{
	XLOPER12 xBool;

	xBool.xltype = xltypeBool;
	xBool.val.xbool = 0;
	Excel12f(xlcWorkspace, 0, 4, TempMissing12(), TempMissing12(),
		TempMissing12(), (LPXLOPER12)&xBool);
	return 1;
}

/*
** TempNumExample
**
** This function uses TempNum() to pass an argument to GET.WORKSPACE.
*/
__declspec(dllexport) short WINAPI TempNumExample(void)
{
	XLOPER12 xRes;

	Excel12f(xlfGetWorkspace, &xRes, 1, TempNum12(44));
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** TempStrExample
**
** This function uses TempStrConst() to create a string for an alert message box.
**
*/
__declspec(dllexport) short WINAPI TempStrExample(void)
{
	Excel12f(xlcAlert, 0, 1, TempStr12(L"Made it!"));
	return 1;
}

/*
** fArray
**
** This example consists of two routines: fArray and xlAutoFree().
** This function creates an xltypeMulti containing eight values. It returns
** this array to Microsoft Excel with the xlbitDLLFree bit set. When
** Microsoft Excel is done with the values, it calls xlAutoFree(), which
** frees the memory that fArray() allocated.
*/

HANDLE hArray;

__declspec(dllexport) LPXLOPER12 WINAPI fArray(void)
{
    LPXLOPER12 pxArray;
    static XLOPER12 xMulti;
    int i;
	int rwcol;

	xMulti.xltype = xltypeMulti | xlbitDLLFree;
	xMulti.val.array.columns = 1;
	xMulti.val.array.rows = 8;

	//For large values of rows and columns, this would overflow
	//use __int64 in that case and return an error if rwcol
	//contains a number that won't fit in sizeof(int) bytes

	rwcol = xMulti.val.array.columns * xMulti.val.array.rows; 

	pxArray = (LPXLOPER12)GlobalLock(hArray = GlobalAlloc(GMEM_ZEROINIT, rwcol * sizeof(XLOPER12)));

	xMulti.val.array.lparray = pxArray;

	for(i = 0; i < rwcol; i++) 
	{
		pxArray[i].xltype = xltypeInt;
		pxArray[i].val.w = i;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms

	return (LPXLOPER12)&xMulti;
}

/*
** xlAutoFree
**
** Demonstrates the xlAutoFree callback. Frees the memory allocated by fArray as noted
** in the comment above.
**
*/

__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
	GlobalUnlock(hArray);
	GlobalFree(hArray);
	return;
}

/*
** xlCoerceExample
**
** This function takes a value and converts it to a string.
** Then, the function displays the string in an alert box.
**
*/
__declspec(dllexport) short WINAPI xlCoerceExample(short iVal)
{
	XLOPER12 xStr, xInt, xDestType;

	xInt.xltype = xltypeInt;
	xInt.val.w = iVal;

	xDestType.xltype = xltypeInt;
	xDestType.val.w = xltypeStr;

	Excel12f(xlCoerce, &xStr, 2, (LPXLOPER12)&xInt, (LPXLOPER12)&xDestType);

	Excel12f(xlcAlert, 0, 1, (LPXLOPER12)&xStr);
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xStr);

	return 1;
}

/*
** xlFreeExample
**
** This example calls GET.WORKSPACE(1) to return (as a string) the platform
** on which Microsoft Excel is currently running. The code copies
** this returned string into a buffer for later use. The standard
** strcpy() function is not used to copy the string to the buffer because
** strcpy() expects a null-terminated string and the returned value is a
** byte-counted string. The code places the buffer back into the
** XLOPER12 for later use with Excel12f(). Finally, the code displays
** the string in an alert box.
**
*/
__declspec(dllexport) short WINAPI xlFreeExample(void)
{

	XLOPER12 xRes, xInt;
    XCHAR buffer[cchMaxStz];
    int i,len;

	// Create an XLOPER12 for the argument to Getworkspace
	xInt.xltype = xltypeInt;
	xInt.val.w = 1;
	// Call GetWorkspace
	Excel12f(xlfGetWorkspace, &xRes, 1, (LPXLOPER12)&xInt);
	
	// Get the length of the returned string
	len = (int)xRes.val.str[0];
	if (len > cchMaxStz - 2) //Take into account 1st char, which contains the length
        len = cchMaxStz - 2; //and the null terminator. Truncate if necessary to fit
							 //buffer
	// Copy to buffer
	for(i = 1; i <= len; i++)
		buffer[i] = xRes.val.str[i];

	// Null terminate, Not necessary but a good idea
	buffer[len] = '\0';
    buffer[0] = len;

	// Free the string returned from Excel
	Excel12f(xlFree, 0, 1, &xRes);

	// Create a new string XLOPER12 for the alert
	xRes.xltype = xltypeStr;
	xRes.val.str = buffer;

	// Show the alert
	Excel12f(xlcAlert, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** xlGetInstExample
**
** The following example compares the last instance of Microsoft
** Excel to the current instance of Microsoft Excel. If they are
** the same, the code returns 1; if they are not the same, it returns 0;
** if the function fails, it returns -1.
*/
__declspec(dllexport) short WINAPI xlGetInstExample(void)
{
	XLOPER12 xRes;
	static HANDLE hOld = 0;
	short iRet;

	if (Excel12f(xlGetInst, &xRes, 0) != xlretSuccess)
		iRet = -1;
	else
		{
		HANDLE hNew;

		hNew = (HANDLE)xRes.val.w;
		if (hNew != hOld)
			iRet = 0;
		else
			iRet = 1;
		hOld = hNew;
		}

	return iRet;
}

/*
** xlGetInstPtrExample
**
** The following example compares the last instance of Microsoft
** Excel to the current instance of Microsoft Excel. If they are
** the same, the code returns 1; if they are not the same, it returns 0;
** if the function fails, it returns -1.
*/
__declspec(dllexport) short WINAPI xlGetInstPtrExample(void)
{
	XLOPER12 xRes;
	static HANDLE hOld = 0;
	short iRet;

	if (Excel12f(xlGetInstPtr, &xRes, 0) != xlretSuccess)
		iRet = -1;
	else
		{
		HANDLE hNew;

		hNew = xRes.val.bigdata.h.hdata;
        if (hNew != hOld)
			iRet = 0;
		else
			iRet = 1;
		hOld = hNew;
		}

	return iRet;
}

/*
** xlGetNameExample
**
** The following code retrieves the name of the DLL that contains the code,
** and then displays it in an alert box.
**
*/
__declspec(dllexport) short WINAPI xlGetNameExample(void)
{
	XLOPER12 xRes;

	Excel12f(xlGetName, &xRes, 0);
	Excel12f(xlcAlert, 0, 1, (LPXLOPER12)&xRes);
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** xlSetExample
**
** The following code places the passed-in value into B205:B206.
*/
__declspec(dllexport) short WINAPI xlSetExample(short iVal)
{
	XLOPER12 xRef, xValue;

	xRef.xltype = xltypeSRef;
	xRef.val.sref.count = 1;
	xRef.val.sref.ref.rwFirst = 204;
	xRef.val.sref.ref.rwLast = 205;
	xRef.val.sref.ref.colFirst = 1;
	xRef.val.sref.ref.colLast = 1;
	xValue.xltype = xltypeInt;
	xValue.val.w = iVal;
	Excel12(xlSet, 0, 2, (LPXLOPER12)&xRef, (LPXLOPER12)&xValue);
	return 1;
}

/*
** xlSheetIdExample
**
** This function gets the sheet ID for Sheet1 in workbook BOOK1.XLSX.
**
*/
__declspec(dllexport) short WINAPI xlSheetIdExample(void)
{       
	XLOPER12 xSheetName, xRes;
	XCHAR szBuf[sizeof(IDSHEET)*2+2+1];

	xSheetName.xltype = xltypeStr;
	xSheetName.val.str = L"\022[BOOK1.XLSX]Sheet1";
	Excel12(xlSheetId, &xRes, 1, (LPXLOPER12)&xSheetName);
	wsprintfW((LPWSTR)szBuf, L"0x%p",xRes.val.mref.idSheet);
	Excel12f(xlcAlert, 0, 1, TempStr12(szBuf));
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** xlSheetNmExample
**
** This function displays a message containing the name of the sheet
** that called it.
**
*/
__declspec(dllexport) short WINAPI xlSheetNmExample(void)
{
	XLOPER12 xRes, xSheetName;

	Excel12(xlfCaller, &xRes, 0);
	Excel12(xlSheetNm, &xSheetName, 1, (LPXLOPER12)&xRes);
	Excel12(xlcAlert, 0, 1, (LPXLOPER12)&xSheetName);
	Excel12(xlFree, 0, 1, (LPXLOPER12)&xSheetName);
	return 1;
}

/*
** xlStackExample
**
** This function displays a message that indicates the amount of stack
** space remaining.
**
*/
__declspec(dllexport) short WINAPI xlStackExample(void)
{
	XLOPER12 xRes;

	Excel12(xlStack, &xRes, 0);
	Excel12(xlcAlert, 0, 1, (LPXLOPER12)&xRes);
	return 1;
}

/*
** xlUDFExample
**
** This function demonstrates how to run a macro from an XLL. It runs
** a macro called TestMacro in a workbook called BOOK1.XLSX
*/
__declspec(dllexport) short WINAPI xlUDFExample(void)
{       
	XLOPER12 xMacroName, xMacroRef, xRes;

	xMacroName.xltype = xltypeStr;
	xMacroName.val.str = L"\034[BOOK1.XLSX]Macro1!TestMacro";
	Excel12(xlfEvaluate, &xMacroRef, 1, (LPXLOPER12)&xMacroName);
	Excel12(xlUDF, &xRes, 1, (LPXLOPER12)&xMacroRef);
	return 1;
}

/*
** InternationalExample
**
** This routine demonstrates how to set the xlIntl bit to allow function names to
** be intepreted correclty on international versions of Excel.
**
*/
__declspec(dllexport) int WINAPI InternationalExample()
{
	XLOPER12 xResult, xSum;

	xSum.xltype = xltypeStr;
	xSum.val.str = L"\011=SUM(1,2)";

	return Excel12f(xlcFormula | xlIntl, &xResult, 2, (LPXLOPER12) &xSum, TempActiveRef12(237,237,1,1));
}

/*
** CalcCircum
**
** Calculates the circumference of a circle given the radius
**
*/

__declspec(dllexport) double WINAPI CalcCircum(double pdRadius)
{
	return pdRadius * 6.283185308;
}
