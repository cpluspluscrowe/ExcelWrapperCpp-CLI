#include "stdafx.h"
#include "Worksheet.h"
#include "Range.h"

//Worksheet Wrapper
WorksheetWrapper::Worksheet::Worksheet(Excel::Worksheet^ worksheet)
{
	this->wrappedWorksheet = worksheet;
}

Excel::Worksheet^ WorksheetWrapper::Worksheet::GetWrappedWorksheet(){
	return this->wrappedWorksheet;
}


//Worksheet.Range Wrapper
WorksheetWrapper::WorksheetRangeWrapper::WorksheetRangeWrapper(){

}

WorksheetWrapper::WorksheetRangeWrapper::WorksheetRangeWrapper(Excel::Worksheet^ worksheet){

}

RangeWrapper::Range^ WorksheetWrapper::WorksheetRangeWrapper::operator()(String^ rangeString){
	return gcnew RangeWrapper::Range(this->wrappedWorksheet->Range[rangeString, Type::Missing]);
}
RangeWrapper::Range^ WorksheetWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1,String^ rangeString2){
	return gcnew RangeWrapper::Range(this->wrappedWorksheet->Range[rangeString1, rangeString2]);
}

//Worksheet.Cells Wrapper
WorksheetWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(){

}
WorksheetWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}
RangeWrapper::Range^ WorksheetWrapper::WorksheetCellsWrapper::operator()(int row, int column){
	return gcnew RangeWrapper::Range(static_cast<Excel::Range^>(this->wrappedWorksheet->Cells[row, column]));
}
