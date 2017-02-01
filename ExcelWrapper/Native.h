#pragma once
#include <string>
#include <memory>
#include <algorithm>
#include <locale>

#ifdef COMPILE_NATIVE_LIB
#define MY_API __declspec(dllexport)
#else
#define MY_API __declspec(dllimport)
#endif

class MY_API Native{
public:	
	Native(std::string cellValue);
	Native(double cellValue);
	~Native();
	template <typename T>
	T GetValue(){
		if (this->dValue != nullptr){
			return dValue;
		}
		else{
			return sValue;
		}
	}
private:
	std::shared_ptr<double> dValue;
	std::shared_ptr<std::string> sValue;
};