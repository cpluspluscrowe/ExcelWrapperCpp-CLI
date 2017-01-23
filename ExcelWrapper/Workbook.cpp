#include "stdafx.h"
#include "Workbook.h"
#include "ExcelApplication.h"
#include "Worksheet.h"

//Workbook
WorkbookWrapper::Workbook::Workbook(Excel::Application^ xl,System::String^ filePath)
{
	this->wrappedWorkbook = xl->Workbooks->Open(filePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	for each(Excel::Worksheet^ worksheet in this->wrappedWorkbook->Sheets){
		this->worksheetDict.Add(worksheet->Name, gcnew WorksheetWrapper::Worksheet(worksheet));
	}
}

WorkbookWrapper::Workbook::!Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}

WorkbookWrapper::Workbook::~Workbook(){
	this->wrappedWorkbook->Close((System::Object^)false, Type::Missing, Type::Missing);
}

//Sheets
//WorkbookWrapper::Workbook::Sheets::Sheets(System::String^ sheetName){
//	WorksheetWrapper::Worksheet^ ws = gcnew WorksheetWrapper::Worksheet();
//}