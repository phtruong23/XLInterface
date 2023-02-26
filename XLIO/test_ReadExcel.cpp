#include <iostream>
#include "XLIO.h"

int test_read_excel()
{
    XLIO XLReader;
    CString filename = L"D:/Workspace/Workspace_Cpp/Excel/ReadExcelUsingCOM/monitor.xlsx";
    ExcelData xlData;
    bool result;

    result = XLReader.Open(filename);
    if (!result)
    {
        std::cout << "Can't open file" << std::endl;
        return -1;
    }
    result = XLReader.GetColumns(7, xlData);
    if (!result)
    {
        std::cout << "Can't read file" << std::endl;
        return -1;
    }

    XLReader.Close();
    XLReader.Quit();

    CString tmp;

    for (auto row : xlData)
    {
        for (auto val : row)
        {
            std::wcout << val.GetString() << " ";
            tmp += val + L" ";
        }
        std::wcout << std::endl;
        tmp += L"\n";
    }
    MessageBox(NULL, tmp, L"ExcelReader", MB_OK);
    return 0;
}

int test_write_excel()
{
    XLIO XLReader;
    CString filename = L"D:/Workspace/Workspace_Cpp/Excel/ReadExcelUsingCOM/monitor2.xlsx";
    ExcelData xlData;
    RowData rDt;
    bool result;

    rDt.push_back("Begin");
    rDt.push_back("End");
    rDt.push_back("Room_code");
    rDt.push_back("Bed_code");
    rDt.push_back("1st Check");
    rDt.push_back("2nd Check");
    rDt.push_back("Prediction");
    xlData.push_back(rDt);

    result = XLReader.New();

    result = XLReader.SetRange(L"A1:H2", xlData);

    //XLReader.Save(filename);
    XLReader.SaveAs(filename);
    XLReader.Close();
    XLReader.Quit();

    return 0;
}

int main()
{
    //test_read_excel();
    test_write_excel();

    return 0;
}