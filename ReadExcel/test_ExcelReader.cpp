#include "ExcelReader.h"
#include <string.h>
#include <iostream>

int main()
{
    CString filename = L"D:/Workspace/Workspace_Cpp/Excel/ReadExcelUsingCOM/smc_monitor_20230223.xlsx";
    ExcelData xlData;

    std::cout << filename.GetString() << std::endl;
    MessageBox(NULL, filename, L"ExcelReader", MB_OK);

    readExcel(filename, xlData, 7);

    CString tmp;
    for (auto row : xlData)
    {
        for (auto val : row)
        {
            //std::wcout << val.GetString() << " ";
            tmp += val + L" ";
        }
        //std::wcout << std::endl;
        tmp += L"\n";
    }
    MessageBox(NULL, tmp, L"ExcelReader", MB_OK);
    return 0;
}