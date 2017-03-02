// This is the main DLL file.

#include "stdafx.h"

#include "ExcelApplication.h"

//ExcelApplication
ExcelApplicationWrapper::ExcelApplication::ExcelApplication(){
	Excel::Application^ xl = gcnew Application();
	xl->Visible = true;
	this->xl = xl;
	this->Workbooks = gcnew ExcelApplicationWrapper::Workbooks(this->xl);
}
void ExcelApplicationWrapper::ExcelApplication::SetAlerts(bool showAlerts){
	this->xl->DisplayAlerts = showAlerts;
}
void ExcelApplicationWrapper::ExcelApplication::SetVisibility(bool isVisible){
	this->xl->Visible = isVisible;
}
void ExcelApplicationWrapper::ExcelApplication::Quit(){
	this->xl->Quit();
}
ExcelApplicationWrapper::ExcelApplication::~ExcelApplication(){
	this->xl->Quit();
}
ExcelApplicationWrapper::ExcelApplication::!ExcelApplication(){
	this->xl->Quit();
}
Excel::Application^ ExcelApplicationWrapper::ExcelApplication::GetWrappedExcelApplication(){
	return this->xl;
}

//Workbooks
ExcelApplicationWrapper::Workbook^ ExcelApplicationWrapper::Workbooks::Open(String^ filePath){
	ExcelApplicationWrapper::Workbook^ wb = gcnew ExcelApplicationWrapper::Workbook(this->xl, filePath);
	return wb;
}
ExcelApplicationWrapper::Workbook^ ExcelApplicationWrapper::Workbooks::Open(std::string filePath){
	String^ convertedFilePath = gcnew String(filePath.c_str());
	ExcelApplicationWrapper::Workbook^ wb = gcnew ExcelApplicationWrapper::Workbook(this->xl, convertedFilePath);
	return wb;
}
ExcelApplicationWrapper::Workbooks::Workbooks(Excel::Application^ xl){
	this->xl = xl;
}