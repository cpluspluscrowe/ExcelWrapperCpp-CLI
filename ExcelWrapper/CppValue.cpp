#include "stdafx.h"
#include "CppValue.h"

std::shared_ptr<std::string> CppValue::GetCppSharedPString(){
	return this->cppStringSharedP;
}

CppValue::CppValue(std::string* cppValue2){
	if (cppValue2 == nullptr){
		//do nothing, leave as nullptr
	}
	else{
		this->cppStringSharedP = std::make_shared<std::string>(*cppValue2);
	}
}