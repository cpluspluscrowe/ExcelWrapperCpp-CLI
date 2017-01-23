#pragma once
using namespace Microsoft::Office::Interop::Excel;
using namespace System::Collections::Generic;
using namespace System;
namespace Excel = Microsoft::Office::Interop::Excel;
#include <map>
#include <memory>
#include "Worksheet.h"

namespace WorkbookWrapper{
	public ref class Workbook
	{
	public:
		Workbook(Excel::Application^ xl, System::String^ filePath);
		!Workbook();
		~Workbook();
		ref class Sheets{

		};
	private:
		Excel::Workbook^ wrappedWorkbook;
		Dictionary<String^, WorksheetWrapper::Worksheet^> worksheetDict;
	};
}
