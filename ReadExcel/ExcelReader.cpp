#define _CRT_SECURE_NO_WARNINGS

#include <iostream>
#include "ExcelReader.h"

using namespace std;


// AutoWrap() - Automation helper function...
HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...) {
	// Begin variable-argument list...
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) {
		//MessageBox(NULL, L"NULL IDispatch passed to AutoWrap()", L"Error", 0x10010);
		cout << "NULL IDispatch passed to AutoWrap()" << endl;
		_exit(0);
	}

	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	char buf[200];
	char szName[256];


	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
		//MessageBox(NULL, buf, L"AutoWrap()", 0x10010);
		cout << buf << endl;
		_exit(0);
		return hr;
	}

	// Allocate memory for arguments...
	VARIANT* pArgs = new VARIANT[cArgs + 1];
	// Extract arguments...
	for (int i = 0; i < cArgs; i++) {
		pArgs[i] = va_arg(marker, VARIANT);
	}

	// Build DISPPARAMS
	dp.cArgs = cArgs;
	dp.rgvarg = pArgs;

	// Handle special-case for property-puts!
	if (autoType & DISPATCH_PROPERTYPUT) {
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
	}

	// Make the call!
	hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, NULL, NULL);
	if (FAILED(hr)) {
		sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
		//MessageBox(NULL, buf, L"AutoWrap()", 0x10010);
		cout << buf << endl;
		_exit(0);
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete[] pArgs;

	return hr;
}

int readExcel(CString xlFilename, ExcelData &xlData, int colSize)
{
	CoInitialize(NULL);

	// Get the CLSID of EXCEL 
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

	if (FAILED(hr)) {
		cout << "CLSIDFromProgID() function call failed!" << endl;
		return -1;
	}

	// Create an instance
	IDispatch *pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pXlApp);
	if (FAILED(hr)) {
		cout << "Please check if EXCEL has been installed!";
		return -1;
	}

	//// Display, set the Application.Visible property to 1
	//VARIANT x;
	//x.vt = VT_I4;
	//x.lVal = 0;
	//AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlApp, L"Visible", 1, x);
	
	// Get the Workbooks collection
	IDispatch *pXlBooks;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
		pXlBooks = result.pdispVal;
	}

	// Array to store information
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	SAFEARRAYBOUND sab[2];
	//	sab[0].lLbound = 1; sab[0].cElements = 40; 
	//	sab[1].lLbound = 1; sab[1].cElements = 16; 
	sab[0].lLbound = 1; sab[0].cElements = 1000;
	sab[1].lLbound = 1; sab[1].cElements = 30;
	arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	//int tableNum;

	// Call the Workbooks.Open() method to open an existing Workbook
	IDispatch *pXlBook;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(xlFilename); //xlFilename.AllocSysString();
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Open", 1, parm);
		pXlBook = result.pdispVal;
	}

	IDispatch *pXlSheet;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
		pXlSheet = result.pdispVal;
	}

	// Select a Range
	IDispatch *pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(L"A1:Z1000");

		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
		VariantClear(&parm);

		pXlRange = result.pdispVal;
	}

	// int colSize = 30;
	// Read data within this Range
	AutoWrap(DISPATCH_PROPERTYGET, &arr, pXlRange, L"Value", 0);

	RowData row_data;	// get row data
	VARIANT eleVar;		// get cell data
	CString strVal;		// get CString type in the cell data 
	LONGLONG dblVal;	// get int type in the cell data
	long indices[2];

	for (int i=1; i <= 1000; i++)
	{
		for (int j = 1; j <= colSize; j++)
		{
			// strVal = "";
			// VARIANT eleVar;
			
			indices[0] = i;
			indices[1] = j;
			SafeArrayGetElement(arr.parray, indices, (void *)&eleVar);
			if (eleVar.vt == VT_BSTR)
			{
				strVal = eleVar.bstrVal;
			}
			else if (eleVar.vt == VT_R8)
			{
				dblVal = eleVar.dblVal;
				strVal.Format(L"%lld", dblVal);
			}
			else if (eleVar.vt == VT_NULL)
			{
				strVal = "";
			}
			else
			{
				strVal = "";
			}
			
			// last line
			if (j == 1 && strVal.IsEmpty())
			{
				goto end;
			}
			// first row, last column
			if (i == 1 && strVal.IsEmpty())
			{
				colSize = j;
				break;
			}

			row_data.push_back(strVal);
			//std::wcout<< strVal.GetString() << " ";
		}
		xlData.push_back(row_data); 
		row_data.clear();
		//std::cout << std::endl;
	}
	end:
		AutoWrap(DISPATCH_METHOD, NULL, pXlBook, L"Close", 0);
		VariantClear(&arr);
		pXlRange->Release();
		pXlSheet->Release();
		pXlBook->Release();
	
	// Release all interfaces and variables
	AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);
	pXlBooks->Release();
	pXlApp->Release();

	// Unregister the COM library
	CoUninitialize();
	return 0;
}
