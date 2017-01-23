#include "stdafx.h"
#include "Worksheet.h"


WorksheetWrapper::Worksheet::Worksheet(Excel::Worksheet^ worksheet)
{
	this->wrappedWorksheet = worksheet;
}
