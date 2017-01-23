// ExcelWrapper.h

#pragma once
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
using namespace System;
#include "Workbook.h"

namespace ExcelApplicationWrapper{
	/*! will return a WorkbookWrapper::Workbook*/
	public ref class Workbooks{
	public:
		Excel::Application^ xl;
		Workbooks(Excel::Application^ xl);
		ExcelApplicationWrapper::Workbook^ Open(System::String^ filePath);
		ExcelApplicationWrapper::Workbook^ Open(std::string filePath);
	};

	public ref class ExcelApplication
	{
	public:
		Excel::Application^ xl;
		ExcelApplication();
		!ExcelApplication();
		~ExcelApplication();

		Excel::Application^ GetWrappedExcelApplication();

		ExcelApplicationWrapper::Workbooks^ Workbooks;
	};
}