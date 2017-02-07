#include "stdafx.h"
#include "Range.h"
#include <msclr/marshal_cppstd.h>

///Range Wrapper Code
ExcelApplicationWrapper::Range::Range(Excel::Range^ rng){
	this->wrappedRange = rng;
	double doubleValue;
	if (rng->Text != nullptr && rng->Text != ""){
		if (rng->Count == 1){
			auto isDouble = System::Double::TryParse(rng->Value2->ToString(), doubleValue);
			if (isDouble){
				this->dValue = doubleValue;
			}
			else{
				this->sValue = rng->Value2->ToString();
				msclr::interop::marshal_context context;
				std::string cellValue2 = context.marshal_as<std::string>(this->sValue);
			}
		}
	}
	else{
		this->sValue = nullptr;
		this->dValue = nullptr;
	}
}


System::String^ ExcelApplicationWrapper::Range::GetString(){
	return this->sValue;
}
double^ ExcelApplicationWrapper::Range::GetDouble(){
	return this->dValue;
}

bool ExcelApplicationWrapper::Range::IsNull(){
	if (this->dValue == nullptr && this->sValue == nullptr){
		return true;
	}
	else{
		return false;
	}
}

Excel::Range^ ExcelApplicationWrapper::Range::GetWrappedRange(){
	return this->wrappedRange;
}

void ExcelApplicationWrapper::Range::SetValue(int value2PutInCell){
	this->wrappedRange->Value2 = value2PutInCell;
}
void ExcelApplicationWrapper::Range::SetValue(double value2PutInCell){
	this->wrappedRange->Value2 = value2PutInCell;
}
void ExcelApplicationWrapper::Range::SetValue(System::String^ value2PutInCell){
	this->wrappedRange->Value2 = value2PutInCell;
}
