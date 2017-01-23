// ExcelWrapper.h

#pragma once
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
using namespace System;
#include "Workbook.h"

namespace ExcelApplicationWrapper{
	public ref class ExcelApplication
	{
	public:
		Excel::Application^ xl;
		ExcelApplication();
		!ExcelApplication();
		~ExcelApplication();

		Excel::Application^ GetWrappedExcelApplication();

		/*! will return a WorkbookWrapper::Workbook*/
		ref class Workbooks{
		public:
			Excel::Application^ xl;
			Workbooks(Excel::Application^ xl);
			WorkbookWrapper::Workbook^ Open(System::String^ filePath);
		};
	};
}