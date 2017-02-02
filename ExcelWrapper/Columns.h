#pragma once
#include "Range.h"
using namespace System;
using namespace System::Collections::Generic;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	ref class Worksheet;
	public ref class WorksheetColumnsWrapper
	{
	public:
		WorksheetColumnsWrapper(ExcelApplicationWrapper::Worksheet^ currentSheet);

		bool IsStringInColumn(String^ stringLooking4);
		int GetLastUsedRow();
		Queue<ExcelApplicationWrapper::Range^>^ FindInColumn(String^ looking4InColumn);
		void SetColumnIndex(int columnIndex);
		void SetColumnIndexByLetter(String^ columnLetter);
	private:
		int columnIndex;
		ExcelApplicationWrapper::Worksheet^ currentSheet;
	};
}