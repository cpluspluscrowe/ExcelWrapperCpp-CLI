#include "stdafx.h"
#include "Range.h"
#include <msclr/marshal_cppstd.h>

///Range Wrapper Code

ExcelApplicationWrapper::Range::Range(Excel::Range^ rng){
	this->wrappedRange = rng;
}

System::String^ ExcelApplicationWrapper::Range::GetFormula(){
	System::Object^ formula = this->wrappedRange->Formula;
	if (formula != nullptr){
		return this->ToString();
	}
	else{
		return nullptr;
	}
}

System::String^ ExcelApplicationWrapper::Range::GetValueString(){
	System::Object^ value2 = this->GetValue2();
	if (value2 != nullptr){
		return value2->ToString();
	}
	else{
		return nullptr;
	}
}
System::String^ ExcelApplicationWrapper::Range::GetText(){
	return this->wrappedRange->Text->ToString();
}
System::Object^ ExcelApplicationWrapper::Range::GetValue2(){
	return this->wrappedRange->Value2;
}

bool ExcelApplicationWrapper::Range::IsNull(){
	auto text = this->GetText();
	if (text == nullptr || text == ""){
		return true;
	}
	else{
		return false;
	}
}

bool ExcelApplicationWrapper::Range::HasFormula(){
	if (this->GetFormula() == nullptr){
		return false;
	}
	else{
		return true;
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
double ExcelApplicationWrapper::Range::GetDouble(){
	double result;
	System::Double::TryParse(this->GetText(),result);
	return result;
}

