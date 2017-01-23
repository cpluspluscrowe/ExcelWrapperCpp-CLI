// ExcelWrapper.h

#pragma once
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
using namespace System;
#include "Workbook.h"

namespace ExcelApplicationWrapper{
	public ref class ExcelApplication
	{
		Excel::Application^ xl;
		ExcelApplication();
		!ExcelApplication();
		~ExcelApplication();

		ref class Workbooks{
			WorkbookWrapper::Workbook^ Open(String^ filePath);
		};
	};
}