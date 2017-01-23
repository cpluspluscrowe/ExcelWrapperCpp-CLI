#include "stdafx.h"
#include "Workbook.h"
#include "ExcelApplication.h"
#include "Worksheet.h"

///Workbook functions
ExcelApplicationWrapper::Workbook::Workbook(Excel::Application^ xl, System::String^ filePath)
{
	this->wrappedWorkbook = xl->Workbooks->Open(filePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	this->Sheets = gcnew ExcelApplicationWrapper::WorkbookSheetsWrapper(this->wrappedWorkbook);
}

ExcelApplicationWrapper::Workbook::!Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}

ExcelApplicationWrapper::Workbook::~Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}
Excel::Workbook^ ExcelApplicationWrapper::Workbook::GetWrappedWorkbook(){
	return this->wrappedWorkbook;
}


///Sheets Wrapper functions
ExcelApplicationWrapper::WorkbookSheetsWrapper::WorkbookSheetsWrapper(Excel::Workbook^ workbook){
	this->wrappedWorkbook = workbook;
}

ExcelApplicationWrapper::Worksheet^ ExcelApplicationWrapper::WorkbookSheetsWrapper::operator [](String^ worksheetName){
	return static_cast<ExcelApplicationWrapper::Worksheet^>(gcnew ExcelApplicationWrapper::Worksheet(static_cast<Excel::Worksheet^>(this->wrappedWorkbook->Sheets[worksheetName])));
}
ExcelApplicationWrapper::Worksheet^ ExcelApplicationWrapper::WorkbookSheetsWrapper::operator [](int worksheetNumber){
	return static_cast<ExcelApplicationWrapper::Worksheet^>(gcnew ExcelApplicationWrapper::Worksheet(static_cast<Excel::Worksheet^>(this->wrappedWorkbook->Sheets[worksheetNumber])));
}
