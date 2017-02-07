#pragma once
#include <string>
#include <memory>
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	///Native
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);
		bool IsNull();
		System::String^ GetString();
		double^ GetDouble();
		void SetValue(int value2PutInCell);
		void SetValue(double value2PutInCell);
		void SetValue(System::String^ value2PutInCell);
		Excel::Range^ ExcelApplicationWrapper::Range::GetWrappedRange();
	private:
		Excel::Range^ wrappedRange;
		System::String^ sValue;
		double^ dValue;
	};
}

