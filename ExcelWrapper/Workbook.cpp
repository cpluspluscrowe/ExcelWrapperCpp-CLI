#include "stdafx.h"
#include "Workbook.h"
#include "ExcelApplication.h"
#include "Worksheet.h"

///Workbook functions
WorkbookWrapper::Workbook::Workbook(Excel::Application^ xl,System::String^ filePath)
{
	this->wrappedWorkbook = xl->Workbooks->Open(filePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	this->Sheets = gcnew WorkbookWrapper::WorkbookSheetsWrapper(this->wrappedWorkbook);
}

WorkbookWrapper::Workbook::!Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}

WorkbookWrapper::Workbook::~Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}
Excel::Workbook^ WorkbookWrapper::Workbook::GetWrappedWorkbook(){
	return this->wrappedWorkbook;
}


///Sheets Wrapper functions
WorkbookWrapper::WorkbookSheetsWrapper::WorkbookSheetsWrapper(Excel::Workbook^ workbook){
	this->wrappedWorkbook = workbook;
}

WorksheetWrapper::Worksheet^ WorkbookWrapper::WorkbookSheetsWrapper::operator [](String^ worksheetName){
	return static_cast<WorksheetWrapper::Worksheet^>(gcnew WorksheetWrapper::Worksheet(static_cast<Excel::Worksheet^>(this->wrappedWorkbook->Sheets[worksheetName])));
}
WorksheetWrapper::Worksheet^ WorkbookWrapper::WorkbookSheetsWrapper::operator [](int worksheetNumber){
	return static_cast<WorksheetWrapper::Worksheet^>(gcnew WorksheetWrapper::Worksheet(static_cast<Excel::Worksheet^>(this->wrappedWorkbook->Sheets[worksheetNumber])));
}
