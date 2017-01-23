#include "stdafx.h"
#include "stdafx.h"
#include "Range.h"

///Range Wrapper Code
ExcelApplicationWrapper::Range::Range(Excel::Range^ rng){
	this->wrappedRange = rng;
	auto rngVal = rng->Value2;
	if (rngVal == nullptr){
		this->Value2 = "nullptr";
	}
	else{
		this->Value2 = rngVal->ToString();
	}
}
