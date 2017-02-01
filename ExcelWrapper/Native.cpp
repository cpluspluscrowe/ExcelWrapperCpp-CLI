#include "stdafx.h"
#include "Native.h"

Native::Native(double nativeValue){
	this->dValue = std::make_shared<double>(nativeValue);
}
Native::Native(std::string nativeValue){
	this->sValue = std::make_shared<std::string>(nativeValue);
}

Native::~Native(){

}