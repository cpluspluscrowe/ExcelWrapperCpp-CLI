#pragma once
#include "Range.h"
#include "Columns.h"
using namespace System;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{

	public ref class RowsWrapper sealed{
	public:
		RowsWrapper(Excel::Worksheet^ worksheet);

		int Count;
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	public ref class WorksheetUsedRangeWrapper sealed{
	public:
		WorksheetUsedRangeWrapper(Excel::Worksheet^ worksheet);

		RowsWrapper^ Rows;
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet.Range wrapper
	public ref class WorksheetRangeWrapper{
	public:
		WorksheetRangeWrapper(Excel::Worksheet^ worksheet);

		ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ range1);
		ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1, String^ rangeString2);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet.Cells Wrapper
	public ref class WorksheetCellsWrapper{
	public:
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
		ExcelApplicationWrapper::WorksheetColumnsWrapper^ Columns(int i);
		ExcelApplicationWrapper::WorksheetColumnsWrapper^ Columns(String^ columnLetter);

		WorksheetRangeWrapper^ Range;
		WorksheetCellsWrapper^ Cells;
		WorksheetUsedRangeWrapper^ UsedRange;
	private:
		Excel::Worksheet^ wrappedWorksheet;
		ExcelApplicationWrapper::WorksheetColumnsWrapper^ currentColumn;
	};
}
