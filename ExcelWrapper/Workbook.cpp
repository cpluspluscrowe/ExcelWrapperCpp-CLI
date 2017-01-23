#include "stdafx.h"
#include "Workbook.h"
#include "ExcelApplication.h"

WorkbookWrapper::Workbook::Workbook(String^ filePath)
{

}

WorkbookWrapper::Workbook^ ExcelApplicationWrapper::ExcelApplication::Workbooks::Open(String^ filePath){
	WorkbookWrapper::Workbook^ wb = gcnew WorkbookWrapper::Workbook(filePath);
	return wb;
}