# ExcelWrapperCpp-CLI, Last Update: 2/2/17
Flexible Wrapper for Excel using C++/CLI.  The classes expose the underlying API, so you can use these methods and the original API where holes exist in this API. 

Documentation at: Main Page->ExcelWrapper->html

Interfaces for:

Excel.Application, destructor to Quit() when the program finishes

Workbook

Worksheet

Range, and Value Functions

Cells, and Value Functions

Worksheet.UsedRange.Rows.Count


There is also a portion written in native C++.  This is for easy conversion from Excel Range values to std::string or double.  You will need to include ExcelWrapper->Debug->ExcelWrapper.lib for the C++ functionality and the corresponding include file Native.h


Author: Graduate Student Chad K. Crowe
