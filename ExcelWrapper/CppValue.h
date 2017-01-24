#pragma once
#include <string>
#include <memory>

#ifdef COMPILE_NATIVE_LIB
#define MY_API __declspec(dllexport)
#else
#define MY_API __declspec(dllimport)
#endif

class MY_API CppValue{
public:
	CppValue(std::string* cppValue2);

	std::shared_ptr<std::string> GetCppSharedPString();
private:
	std::shared_ptr<std::string> cppStringSharedP;
};