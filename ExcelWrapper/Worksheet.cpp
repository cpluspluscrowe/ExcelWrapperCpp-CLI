#include "stdafx.h"
#include "Worksheet.h"
#include "Range.h"

//Worksheet Wrapper
ExcelApplicationWrapper::Worksheet::Worksheet(Excel::Worksheet^ worksheet)
{
	this->wrappedWorksheet = worksheet;
	this->Range = gcnew WorksheetRangeWrapper(worksheet);
	this->Cells = gcnew WorksheetCellsWrapper(worksheet);
}

Excel::Worksheet^ ExcelApplicationWrapper::Worksheet::GetWrappedWorksheet(){
	return this->wrappedWorksheet;
}


//Worksheet.Range Wrapper
ExcelApplicationWrapper::WorksheetRangeWrapper::WorksheetRangeWrapper(){

}

ExcelApplicationWrapper::WorksheetRangeWrapper::WorksheetRangeWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString){
	return gcnew ExcelApplicationWrapper::Range(this->wrappedWorksheet->Range[rangeString, Type::Missing]);
}
ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1, String^ rangeString2){
	return gcnew ExcelApplicationWrapper::Range(this->wrappedWorksheet->Range[rangeString1, rangeString2]);
}

//Worksheet.Cells Wrapper
ExcelApplicationWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(){

}
ExcelApplicationWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}
ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetCellsWrapper::operator()(int row, int column){
	return gcnew ExcelApplicationWrapper::Range(static_cast<Excel::Range^>(this->wrappedWorksheet->Cells[row, column]));
}
