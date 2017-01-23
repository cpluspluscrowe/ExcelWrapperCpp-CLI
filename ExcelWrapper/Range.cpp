#include "stdafx.h"
#include "stdafx.h"
#include "Range.h"

///Range Wrapper Code
RangeWrapper::Range::Range(Excel::Range^ rng){
	this->wrappedRange = rng;
}
