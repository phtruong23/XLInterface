#include "XLIO.h"
//#include <iostream>
//using namespace std;

XLIO::XLIO()
{
	rowSize = 1000;
	colSize = 30;
	rangeStr = L"A1:Z1000";

	Init();
}

XLIO::~XLIO()
{
	// Unregister the COM library
	CoUninitialize();
}

// Close current workbook
void XLIO::Close()
{
	AutoWrap(DISPATCH_METHOD, NULL, pXlBook, L"Close", 0);
	pXlRange->Release();
	pXlSheet->Release();
	pXlBook->Release();
}

void XLIO::Quit()
{
	// Release all interfaces and variables
	AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Quit", 0);
	pXlBooks->Release();
	pXlApp->Release();
}


BOOL XLIO::Init()
{
	CoInitialize(NULL);

	// Get the CLSID of EXCEL 
	CLSID clsid;
	HRESULT hr = CLSIDFromProgID(L"Excel.Application", &clsid);

	if (FAILED(hr)) {
		//cout << "CLSIDFromProgID() function call failed!" << endl;
		MessageBox(NULL, L"CLSIDFromProgID() function call failed!", L"XLIO", MB_ICONERROR);
		return FALSE;
	}

	// Create an instance
	//IDispatch* pXlApp;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pXlApp);
	if (FAILED(hr)) {
		//cout << "Please check if EXCEL has been installed!";
		return FALSE;
	}

	// Get the Workbooks collection
	//IDispatch* pXlBooks;
	{
		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Workbooks", 0);
		pXlBooks = result.pdispVal;
	}
	return TRUE;
}

BOOL XLIO::Open(CString filename)
{
	// Call the Workbooks.Open() method to open an existing Workbook
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(filename); //xlFilename.AllocSysString();
		VARIANT result;
		VariantInit(&result);
		hr = AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Open", 1, parm);
		pXlBook = result.pdispVal;
		
		VariantClear(&parm);
		if (FAILED(hr))
			return FALSE;
	}

	//IDispatch* pXlSheet;
	{
		VARIANT result;
		VariantInit(&result);
		hr = AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
		pXlSheet = result.pdispVal;
		
		if (FAILED(hr))
			return FALSE;
	}

	return TRUE;
}

BOOL XLIO::New()
{
	// Call Workbooks.Open() to get a new workbook...
	{
		VARIANT result;
		VariantInit(&result);
		hr = AutoWrap(DISPATCH_PROPERTYGET, &result, pXlBooks, L"Add", 0);
		pXlBook = result.pdispVal;
		
		if (FAILED(hr))
			return FALSE;
	}

	// Get ActiveSheet object
	//IDispatch* pXlSheet;
	{
		VARIANT result;
		VariantInit(&result);
		hr = AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"ActiveSheet", 0);
		pXlSheet = result.pdispVal;
		if (FAILED(hr))
			return FALSE;
	}
	return TRUE;
}

BOOL XLIO::Save(CString filename)
{
	// pXlBook->Save
	// Set .Saved property of workbook to TRUE so we aren't prompted
	// to save when we tell Excel to quit...
	{
		VARIANT x;
		x.vt = VT_I4;
		x.lVal = 1;
		/*x.vt = VT_BSTR;
		x.bstrVal = ::SysAllocString(filename);*/

		hr = AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlBook, L"Saved", 1, x);
		//hr = AutoWrap(DISPATCH_METHOD, NULL, pXlApp, L"Save", 0);

		VariantClear(&x);
		if (FAILED(hr))
			return FALSE;
	}
	return TRUE;
}

BOOL XLIO::SaveAs(CString filename)
{
	// pXlBook->Save
	{
		// Convert the NULL-terminated string to BSTR.
		VARIANT vtFileName;
		vtFileName.vt = VT_BSTR;
		vtFileName.bstrVal = SysAllocString(filename);

		VARIANT vtFormat;
		vtFormat.vt = VT_I4;
		vtFormat.lVal = 6;     // XlFileFormat::xlCSV

		// If there are more than 1 parameters passed, they MUST be pass in 
		// reversed order. Otherwise, you may get the error 0x80020009.
		hr = AutoWrap(DISPATCH_METHOD, NULL, pXlBook, L"SaveAs", 2, vtFormat, vtFileName);

		VariantClear(&vtFileName);
		if (FAILED(hr))
			return FALSE;
	}
	return TRUE;
}

/***
* Get data of the index sheet 
* index = 0: Get current sheet data
***/
BOOL XLIO::GetSheet(uint8_t index, ExcelData& xlData)
{
	//Changing the ActiveSheet    
	if (index)
	{
		{
			VARIANT itemn;
			itemn.vt = VT_I4;
			itemn.lVal = index;

			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_PROPERTYGET, &result, pXlApp, L"Worksheets", 1, itemn);
			pXlSheet = result.pdispVal;
		}

		{
			VARIANT result;
			VariantInit(&result);
			AutoWrap(DISPATCH_METHOD, &result, pXlSheet, L"Activate", 0);
		}
	}
	GetData(xlData);
	return TRUE;
}
BOOL XLIO::GetRange(int rowNum, int colNum, ExcelData& xlData)
{
	rowSize = rowNum;
	colSize = colNum;

	GetData(xlData);
	return TRUE;
}
BOOL XLIO::GetColumns(int colNum, ExcelData& xlData)
{
	colSize = colNum;
	GetData(xlData);
	return TRUE;
}
BOOL XLIO::GetRows(int rowNum, ExcelData& xlData)
{
	rowSize = rowNum;
	GetData(xlData);
	return TRUE;
}

BOOL XLIO::GetData(ExcelData& xlData)
{
	// Array to store information
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	SAFEARRAYBOUND sab[2];
	//	sab[0].lLbound = 1; sab[0].cElements = 40; 
	//	sab[1].lLbound = 1; sab[1].cElements = 16; 
	sab[0].lLbound = 1; sab[0].cElements = 1000;
	sab[1].lLbound = 1; sab[1].cElements = 30;
	arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	 
	// Select a Range
	//IDispatch* pXlRange;
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(rangeStr);

		VARIANT result;
		VariantInit(&result);
		hr = AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
		pXlRange = result.pdispVal;

		VariantClear(&parm);
	}

	// Read data within this Range
	AutoWrap(DISPATCH_PROPERTYGET, &arr, pXlRange, L"Value", 0);
	if (FAILED(hr))
		return FALSE;

	RowData row_data;	// get row data
	VARIANT eleVar;		// get cell data
	CString strVal;		// get CString type in the cell data 
	LONGLONG dblVal;	// get int type in the cell data
	long indices[2];

	for (int i = 1; i <= rowSize; i++)
	{
		for (int j = 1; j <= colSize; j++)
		{
			// strVal = "";
			// VARIANT eleVar;

			indices[0] = i;
			indices[1] = j;
			SafeArrayGetElement(arr.parray, indices, (void*)&eleVar);
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
				goto out;
			}
			// first row, last column
			if (i == 1 && strVal.IsEmpty())
			{
				colSize = j;
				break;
			}

			row_data.push_back(strVal);
		}
		xlData.push_back(row_data);
		row_data.clear();
	}
	out:
	VariantClear(&arr);
	return TRUE;
}


// range: L"A1:O15"
BOOL XLIO::SetRange(CString range, ExcelData& xlData)
{
	rangeStr = range;
	SetData(xlData);
	return TRUE;
}

BOOL XLIO::SetData(ExcelData & xlData)
{
	// Create a 15x15 safearray of variants...
	VARIANT arr;
	arr.vt = VT_ARRAY | VT_VARIANT;
	{
		SAFEARRAYBOUND sab[2];
		sab[0].lLbound = 1; sab[0].cElements = 1000;
		sab[1].lLbound = 1; sab[1].cElements = 30;
		arr.parray = SafeArrayCreate(VT_VARIANT, 2, sab);
	}

	// Fill safearray with some values...
	for (int i = 1; i <= xlData.size(); i++) {
		for (int j = 1; j <= xlData[0].size(); j++) {
			// Create entry value for (i,j)
			VARIANT tmp;
			tmp.vt = VT_BSTR;
			tmp.bstrVal = SysAllocString(xlData.at(i-1).at(j-1));
			// Add to safearray...
			long indices[] = { i,j };
			SafeArrayPutElement(arr.parray, indices, (void*)&tmp);
		}
	}

	// Get Range object for the Range A1:O15...
	{
		VARIANT parm;
		parm.vt = VT_BSTR;
		parm.bstrVal = ::SysAllocString(rangeStr); //L"A1:O15"

		VARIANT result;
		VariantInit(&result);
		AutoWrap(DISPATCH_PROPERTYGET, &result, pXlSheet, L"Range", 1, parm);
		VariantClear(&parm);

		pXlRange = result.pdispVal;
	}

	// Set range with our safearray...
	AutoWrap(DISPATCH_PROPERTYPUT, NULL, pXlRange, L"Value", 1, arr);

	return TRUE;
}

// AutoWrap() - Automation helper function...
HRESULT XLIO::AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...) {
	// Begin variable-argument list...
	va_list marker;
	va_start(marker, cArgs);

	if (!pDisp) {
		MessageBox(NULL, L"NULL IDispatch passed to AutoWrap()", L"Error", 0x10010);
		//cout << "NULL IDispatch passed to AutoWrap()" << endl;
		//_exit(0);
		return -1;
	}

	// Variables used...
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	DISPID dispidNamed = DISPID_PROPERTYPUT;
	DISPID dispID;
	HRESULT hr;
	//char buf[200];
	char szName[256];
	
	// Convert down to ANSI
	WideCharToMultiByte(CP_ACP, 0, ptName, -1, szName, 256, NULL, NULL);

	// Get DISPID for name passed...
	hr = pDisp->GetIDsOfNames(IID_NULL, &ptName, 1, LOCALE_USER_DEFAULT, &dispID);
	if (FAILED(hr)) {
		//sprintf(buf, "IDispatch::GetIDsOfNames(\"%s\") failed w/err 0x%08lx", szName, hr);
		//cout << buf << endl;
		//_exit(0);
		inform.Format(L"IDispatch::GetIDsOfNames(%s) failed w/err 0x%08lx", ptName, hr);
		MessageBox(NULL, inform, L"AutoWrap", 0x10010);
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
		//sprintf(buf, "IDispatch::Invoke(\"%s\"=%08lx) failed w/err 0x%08lx", szName, dispID, hr);
		//cout << buf << endl;
		//_exit(0);
		inform.Format(L"IDispatch::Invoke(%s=%08lx) failed w/err 0x%08lx", ptName, dispID, hr);
		MessageBox(NULL, inform, L"AutoWrap()", 0x10010);
		return hr;
	}
	// End variable-argument section...
	va_end(marker);

	delete[] pArgs;

	return hr;
}
