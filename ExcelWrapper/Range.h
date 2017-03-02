#pragma once
#include <string>
#include <memory>
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);
		bool IsNull();
		double GetDouble();
		System::String^ GetText();
		System::String^ GetFormula();
		System::Object^ GetValue2();
		System::String^ GetValueString();
		void SetValue(int value2PutInCell);
		void SetValue(double value2PutInCell);
		void SetValue(System::String^ value2PutInCell);
		bool HasFormula();
		Excel::Range^ ExcelApplicationWrapper::Range::GetWrappedRange();
	private:
		Excel::Range^ wrappedRange;
		System::String^ sValue;
		double^ dValue;
		bool isNull;
	};
}

