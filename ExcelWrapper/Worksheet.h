#pragma once
#include "Range.h"
using namespace System;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
namespace WorksheetWrapper{
	///Worksheet.Range wrapper
	public ref class WorksheetRangeWrapper{
	public:
		WorksheetRangeWrapper();
		WorksheetRangeWrapper(Excel::Worksheet^ worksheet);

		RangeWrapper::Range^ WorksheetWrapper::WorksheetRangeWrapper::operator()(String^ range1);
		RangeWrapper::Range^ WorksheetWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1, String^ rangeString2);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet.Cells Wrapper
	public ref class WorksheetCellsWrapper{
	public:
		WorksheetCellsWrapper();
		WorksheetCellsWrapper(Excel::Worksheet^ worksheet);

		RangeWrapper::Range^ operator()(int row,int column);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};

	///Worksheet Wrapper
	ref class Worksheet
	{
	public:
		Worksheet(Excel::Worksheet^ worksheet);

		Excel::Worksheet^ GetWrappedWorksheet();

		WorksheetRangeWrapper Range;
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};
}
