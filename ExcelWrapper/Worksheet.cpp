#include "stdafx.h"
#include "Worksheet.h"
#include "Range.h"

///Worksheet Wrapper
ExcelApplicationWrapper::Worksheet::Worksheet(Excel::Worksheet^ worksheet)
{
	this->wrappedWorksheet = worksheet;
	this->Range = gcnew WorksheetRangeWrapper(worksheet);
	this->Cells = gcnew WorksheetCellsWrapper(worksheet);
	this->UsedRange = gcnew WorksheetUsedRangeWrapper(worksheet);
	this->currentColumn = gcnew ExcelApplicationWrapper::WorksheetColumnsWrapper(this);
	this->Name = this->wrappedWorksheet->Name;
}
///Worksheet Functions
Excel::Worksheet^ ExcelApplicationWrapper::Worksheet::GetWrappedWorksheet(){
	return this->wrappedWorksheet;
}
ExcelApplicationWrapper::WorksheetColumnsWrapper^ ExcelApplicationWrapper::Worksheet::Columns(int i){
	this->currentColumn->SetColumnIndex(i);
	return this->currentColumn;
}

ExcelApplicationWrapper::WorksheetColumnsWrapper^ ExcelApplicationWrapper::Worksheet::Columns(String^ columnLetter){
	this->currentColumn->SetColumnIndexByLetter(columnLetter);
	return this->currentColumn;
}
void ExcelApplicationWrapper::Worksheet::Hide(bool hide){
	if (hide){
		this->wrappedWorksheet->Visible = Microsoft::Office::Interop::Excel::XlSheetVisibility::xlSheetHidden;
	}
	else{
		this->wrappedWorksheet->Visible = Microsoft::Office::Interop::Excel::XlSheetVisibility::xlSheetVisible;
	}
	
}
///Worksheet.Range Wrapper
ExcelApplicationWrapper::WorksheetRangeWrapper::WorksheetRangeWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString){
	return gcnew ExcelApplicationWrapper::Range(this->wrappedWorksheet->Range[rangeString, Type::Missing]);
}
ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetRangeWrapper::operator()(String^ rangeString1, String^ rangeString2){
	return gcnew ExcelApplicationWrapper::Range(this->wrappedWorksheet->Range[rangeString1, rangeString2]);
}

///Worksheet.Cells Wrapper
ExcelApplicationWrapper::WorksheetCellsWrapper::WorksheetCellsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

ExcelApplicationWrapper::Range^ ExcelApplicationWrapper::WorksheetCellsWrapper::operator()(int row, int column){
	return gcnew ExcelApplicationWrapper::Range(static_cast<Excel::Range^>(this->wrappedWorksheet->Cells[row, column]));
}
///UsedRangeWrapper
ExcelApplicationWrapper::WorksheetUsedRangeWrapper::WorksheetUsedRangeWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
	this->Rows = gcnew RowsWrapper(worksheet);
	this->Rows->Count = this->wrappedWorksheet->UsedRange->Rows->Count;
}
///RowsWrapper
ExcelApplicationWrapper::RowsWrapper::RowsWrapper(Excel::Worksheet^ worksheet){
	this->wrappedWorksheet = worksheet;
}

