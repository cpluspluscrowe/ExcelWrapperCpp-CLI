#include "stdafx.h"
#include "Columns.h"


ExcelApplicationWrapper::Columns::Columns(ExcelApplicationWrapper::Worksheet^ currentSheet)
{
	this->currentSheet = currentSheet;
}

bool ExcelApplicationWrapper::Columns::IsStringInColumn(int columnNumber, String^ stringLooking4){
	for (int i = 1; i <= this->currentSheet->UsedRange->Rows->Count; i++){
		if (!this->currentSheet->Cells(i, columnNumber)->IsNull()){
			if (this->currentSheet->Cells(i, columnNumber)->GetString() == stringLooking4){
				return true;
			}
		}
	}
	return false;
}
bool ExcelApplicationWrapper::Columns::IsStringInColumn(String^ columnLetter, String^ stringLooking4){
	for (int i = 1; i <= this->currentSheet->UsedRange->Rows->Count; i++){
		if (!this->currentSheet->Range(columnLetter + i.ToString())->IsNull()){
			if (this->currentSheet->Range(columnLetter + i.ToString())->GetString() == stringLooking4){
				return true;
			}
		}
	}
	return false;
}

int ExcelApplicationWrapper::Columns::GetLastUsedRow(int columnNumber){
	int lastRow = this->currentSheet->UsedRange->Rows->Count;
	while (lastRow > 1 && this->currentSheet->Cells(lastRow, columnNumber)->IsNull()){
		lastRow -= 1;
	}
	return lastRow;
}
int ExcelApplicationWrapper::Columns::GetLastUsedRow(String^ columnLetter){
	int lastRow = this->currentSheet->UsedRange->Rows->Count;
	while (lastRow > 1 && this->currentSheet->Range(columnLetter + lastRow.ToString())->IsNull()){
		lastRow -= 1;
	}
	return lastRow;
}