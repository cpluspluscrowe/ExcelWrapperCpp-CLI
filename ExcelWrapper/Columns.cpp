#include "stdafx.h"
#include "Columns.h"
#include "Worksheet.h"


ExcelApplicationWrapper::WorksheetColumnsWrapper::WorksheetColumnsWrapper(ExcelApplicationWrapper::Worksheet^ currentSheet)
{
	this->currentSheet = currentSheet;
}

bool ExcelApplicationWrapper::WorksheetColumnsWrapper::IsStringInColumn(String^ stringLooking4){
	for (int i = 1; i <= this->currentSheet->UsedRange->Rows->Count; i++){
		if (!this->currentSheet->Cells(i, this->columnIndex)->IsNull()){
			if (this->currentSheet->Cells(i, this->columnIndex)->GetString() == stringLooking4){
				return true;
			}
		}
	}
	return false;
}

int ExcelApplicationWrapper::WorksheetColumnsWrapper::GetLastUsedRow(){
	int lastRow = this->currentSheet->UsedRange->Rows->Count;
	while (lastRow > 1 && this->currentSheet->Cells(lastRow, this->columnIndex)->IsNull()){
		lastRow -= 1;
	}
	return lastRow;
}

Queue<ExcelApplicationWrapper::Range^>^ ExcelApplicationWrapper::WorksheetColumnsWrapper::FindInColumn(String^ looking4InColumn){
	Queue<ExcelApplicationWrapper::Range^>^ rangeVector = gcnew Queue<ExcelApplicationWrapper::Range^>();
	for (int i = 1; i <= this->currentSheet->UsedRange->Rows->Count; i++){
		if (!this->currentSheet->Cells(i, this->columnIndex)->IsNull()){
			if (this->currentSheet->Cells(i, this->columnIndex)->GetString() == looking4InColumn){
				rangeVector->Enqueue(currentSheet->Cells(i, this->columnIndex));
			}
		}
	}
	return rangeVector;
}

void ExcelApplicationWrapper::WorksheetColumnsWrapper::SetColumnIndex(int columnIndex){
	this->columnIndex = columnIndex;
}

void ExcelApplicationWrapper::WorksheetColumnsWrapper::SetColumnIndexByLetter(String^ columnLetter){
	this->columnIndex = this->currentSheet->Range(columnLetter + "1")->GetWrappedRange()->Column;
}