#include "stdafx.h"
#include "Worksheet.h"
#include "Range.h"

//Worksheet Wrapper
ExcelApplicationWrapper::Worksheet::Worksheet(Excel::Worksheet^ worksheet)
{
	this->wrappedWorksheet = worksheet;
	this->Range = gcnew WorksheetRangeWrapper(worksheet);
	this->Cells = gcnew WorksheetCellsWrapper(worksheet);
	this->UsedRange = gcnew WorksheetUsedRangeWrapper(worksheet);
}

Excel::Worksheet^ ExcelApplicationWrapper::Worksheet::GetWrappedWorksheet(){
	return this->wrappedWorksheet;
}
//Worksheet Functions
int ExcelApplicationWrapper::Worksheet::GetLastRowInColumn(int columnNumber){
	int lastRow = this->UsedRange->Rows->Count;
	while (lastRow > 1 && this->Cells(lastRow, columnNumber)->IsNull()){
		lastRow -= 1;
	}
	return lastRow;
}
int ExcelApplicationWrapper::Worksheet::GetLastRowInColumn(String^ columnLetter){
	int lastRow = this->UsedRange->Rows->Count;
	while (lastRow > 1 && this->Range(columnLetter + lastRow.ToString())->IsNull()){
		lastRow -= 1;
	}
	return lastRow;
}
bool ExcelApplicationWrapper::Worksheet::IsStringInColumn(int columnNumber,String^ stringLooking4){
	for (int i = 1; i <= this->UsedRange->Rows->Count; i++){
		if (!this->Cells(i, columnNumber)->IsNull()){
			if (this->Cells(i, columnNumber)->GetString() == stringLooking4){
				return true;
			}
		}
	}
	return false;
}
bool ExcelApplicationWrapper::Worksheet::IsStringInColumn(String^ columnLetter, String^ stringLooking4){
	for (int i = 1; i <= this->UsedRange->Rows->Count; i++){
		if (!this->Range(columnLetter + i.ToString())->IsNull()){
			if (this->Range(columnLetter + i.ToString())->GetString() == stringLooking4){
				return true;
			}
		}
	}
	return false;
}


//Worksheet.Range Wrapper

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
ExcelApplicationWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetCellsWrapper::operator()(int row, int column){
	return gcnew ExcelApplicationWrapper::Range(static_cast<Excel::Range^>(this->wrappedWorksheet->Cells[row, column]));
}
//UsedRangeWrapper
ExcelApplicationWrapper::WorksheetUsedRangeWrapper::WorksheetUsedRangeWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
	this->Rows = gcnew RowsWrapper(worksheet);
	this->Rows->Count = this->wrappedWorksheet->UsedRange->Rows->Count;
}
//RowsWrapper
ExcelApplicationWrapper::RowsWrapper::RowsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

