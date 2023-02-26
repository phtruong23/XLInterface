#ifndef EXCEL_READER_H
#define EXCEL_READER_H

#include <atlstr.h>
#include <vector>
#include <string>
#include <ole2.h>

typedef std::vector<CString> RowData;
typedef std::vector<RowData> ExcelData;

int readExcel(CString xlFilename, ExcelData &xlData, int colSize=30);

#endif 