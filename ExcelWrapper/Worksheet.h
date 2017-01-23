#pragma once
#include "Range.h"
using namespace System;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	///Worksheet.Range wrapper
	public ref class WorksheetRangeWrapper{
	public:
		WorksheetRangeWrapper();
		WorksheetRangeWrapper(Excel::Worksheet^ worksheet);

		ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ range1);
		ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1, String^ rangeString2);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet.Cells Wrapper
	public ref class WorksheetCellsWrapper{
	public:
		WorksheetCellsWrapper();
		WorksheetCellsWrapper(Excel::Worksheet^ worksheet);

		ExcelApplicationWrapper::Range^ operator()(int row, int column);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet Wrapper
	public ref class Worksheet
	{
	public:
		Worksheet(Excel::Worksheet^ worksheet);

		Excel::Worksheet^ GetWrappedWorksheet();

		WorksheetRangeWrapper^ Range;
		WorksheetCellsWrapper^ Cells;
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};
}
