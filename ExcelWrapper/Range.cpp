#include "stdafx.h"
#include "Range.h"
#include <msclr/marshal_cppstd.h>

///Range Wrapper Code
ExcelApplicationWrapper::Range::Range(Excel::Range^ rng){
	this->cppValue2 = new CppValue(nullptr);
	this->wrappedRange = rng;
	auto rngVal = rng->Value2;
	if (rngVal == nullptr){
		this->Value2 = nullptr;
	}
	else{
		this->Value2 = rngVal->ToString();
		msclr::interop::marshal_context context;
		std::string cppValue2 = context.marshal_as<std::string>(this->Value2);
		this->cppValue2 = new CppValue(&cppValue2);
	}

}

CppValue* ExcelApplicationWrapper::Range::GetCppValue(){
	return this->cppValue2;
}

/*

*/