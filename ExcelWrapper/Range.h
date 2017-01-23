#pragma once
using namespace System;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
namespace ExcelApplicationWrapper{
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);
		String^ Value2;
	private:
		Excel::Range^ wrappedRange;
	};
}
