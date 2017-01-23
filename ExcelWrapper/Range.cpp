#include "stdafx.h"
#include "Range.h"

///Range Wrapper Code
RangeWrapper::Range::Range(Excel::Range^ rng){
	this->wrappedRange = rng;
}

///Cells Wrapper Code
RangeWrapper::Cells::Cells(Excel::Range^ rng){
	this->wrappedRange = rng;
}
