#pragma once
#define _CRT_SECURE_NO_WARNINGS

#include <atlstr.h>
#include <vector>
#include <string>
#include <ole2.h>


typedef std::vector<CString> RowData;
typedef std::vector<RowData> ExcelData;

class XLIO
{
public:
	XLIO();
	virtual ~XLIO();

	BOOL Init();
	BOOL Open(CString filename);
	BOOL New();
	BOOL Save(CString filename);
	BOOL SaveAs(CString filename);
	void Close();
	void Quit();

	BOOL GetSheet(uint8_t index, ExcelData& xlData);
	BOOL GetRange(int rowNum, int colNum, ExcelData& xlData);
	BOOL GetColumns(int colNum, ExcelData& xlData);
	BOOL GetRows(int rowNum, ExcelData& xlData);
	
	BOOL SetRange(CString range, ExcelData& xlData);

protected:
	IDispatch* pXlApp;
	IDispatch* pXlBooks;
	IDispatch* pXlBook;
	IDispatch* pXlSheet;
	IDispatch* pXlRange;

private:
	HRESULT hr;
	CString inform;
	int rowSize;
	int colSize;
	CString rangeStr;
	
	BOOL GetData(ExcelData& xlData);
	BOOL SetData(ExcelData& xlData);
	HRESULT AutoWrap(int autoType, VARIANT* pvResult, IDispatch* pDisp, LPOLESTR ptName, int cArgs...);
};