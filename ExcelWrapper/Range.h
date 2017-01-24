#pragma once
#include <string>
#include <memory>
#include "CppValue.h"
#pragma make_public(CppValue)
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);

		CppValue* GetCppValue();

		System::String^ Value2;	
		CppValue* cppValue2;
	private:
		Excel::Range^ wrappedRange; 
	};
}
