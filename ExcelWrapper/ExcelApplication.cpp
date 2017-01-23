// This is the main DLL file.

#include "stdafx.h"

#include "ExcelApplication.h"

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