// This is the main DLL file.

#include "stdafx.h"

#include "ExcelApplication.h"

//ExcelApplication
ExcelApplicationWrapper::ExcelApplication::ExcelApplication(){
	Excel::Application^ xl = gcnew Excel::Application();
	xl->Visible = true;
	this->xl = xl;
}
ExcelApplicationWrapper::ExcelApplication::~ExcelApplication(){
	this->xl->Quit();
}
ExcelApplicationWrapper::ExcelApplication::!ExcelApplication(){
	this->xl->Quit();
}
ExcelApplicationWrapper::ExcelApplication::Workbooks::Workbooks(Excel::Application^ xl){
	this->xl = xl;
}
Excel::Application^ ExcelApplicationWrapper::ExcelApplication::GetWrappedExcelApplication(){
	return this->xl;
}

//Workbooks
WorkbookWrapper::Workbook^ ExcelApplicationWrapper::ExcelApplication::Workbooks::Open(String^ filePath){
	WorkbookWrapper::Workbook^ wb = gcnew WorkbookWrapper::Workbook(this->xl, filePath);
	return wb;
}