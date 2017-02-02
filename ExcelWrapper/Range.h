#pragma once
#include <string>
#include <memory>
#include "Native.h"
#pragma make_public(Native)
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;

namespace ExcelApplicationWrapper{
	///Native
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);

		Native* GetNative();
		bool IsNull();
		System::String^ GetString();
		double^ GetDouble();
		void SetValue(int value2PutInCell);
		void SetValue(double value2PutInCell);
		void SetValue(System::String^ value2PutInCell);
		Excel::Range^ ExcelApplicationWrapper::Range::GetWrappedRange();
	private:
		Excel::Range^ wrappedRange;
		Native* native;
		System::String^ sValue;
		double^ dValue;
	};
}
