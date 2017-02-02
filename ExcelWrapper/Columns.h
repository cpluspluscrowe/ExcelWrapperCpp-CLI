#pragma once
using namespace System;
using namespace System::Collections::Generic;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
#include "Worksheet.h"

namespace ExcelApplicationWrapper{
	ref class Columns
	{
	public:
		Columns(ExcelApplicationWrapper::Worksheet^ currentSheet);

		bool IsStringInColumn(int columnNumber, String^ stringLooking4);
		bool IsStringInColumn(String^ columnLetter, String^ stringLooking4);
		int GetLastUsedRow(int columnNumber);
		int GetLastUsedRow(String^ columnLetter);
		List<ExcelApplicationWrapper::Range^>^ FindInColumn(String^ looking4InColumn);
		void SetColumnIndex(int columnIndex);
		void SetColumnIndexByLetter(String^ columnLetter);
	private:
		int columnIndex;
		ExcelApplicationWrapper::Worksheet^ currentSheet;
	};
}