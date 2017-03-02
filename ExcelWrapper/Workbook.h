#pragma once
using namespace Microsoft::Office::Interop::Excel;
using namespace System::Collections::Generic;
using namespace System;
namespace Excel = Microsoft::Office::Interop::Excel;
#include <map>
#include <memory>
#include "Worksheet.h"

namespace ExcelApplicationWrapper{
	///Workbook.Sheets Wrapper
	public ref class WorkbookSheetsWrapper{
	public:
		WorkbookSheetsWrapper(Excel::Workbook^ workbook);
		ExcelApplicationWrapper::Worksheet^ ExcelApplicationWrapper::WorkbookSheetsWrapper::operator [](String^ worksheetName);
		ExcelApplicationWrapper::Worksheet^ ExcelApplicationWrapper::WorkbookSheetsWrapper::operator [](int worksheetNumber);

		ExcelApplicationWrapper::Worksheet^ Add();
	private:
		Excel::Workbook^ wrappedWorkbook;
	};

	///Workbook Wrapper
	public ref class Workbook
	{
	public:
		Workbook(Excel::Application^ xl, System::String^ filePath);
		Workbook(Excel::Workbook^ wb);
		!Workbook();
		~Workbook();
		
		Excel::Workbook^ GetWrappedWorkbook();
		bool Save();
		bool Close(bool saveIt);

		ExcelApplicationWrapper::WorkbookSheetsWrapper^ Sheets;
		String^ Name;
	private:
		Excel::Workbook^ wrappedWorkbook;
		Dictionary<String^, ExcelApplicationWrapper::Worksheet^> worksheetNameDict;
		Dictionary<int, ExcelApplicationWrapper::Worksheet^> worksheetIntDict;
	};
}
