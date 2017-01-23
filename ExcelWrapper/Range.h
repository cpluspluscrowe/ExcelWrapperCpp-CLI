#pragma once
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
namespace RangeWrapper{
	public ref class Range
	{
	public:
		Range(Excel::Range^ rng);
	private:
		Excel::Range^ wrappedRange;
	};

	public ref class Cells
	{
	public:
		Cells(Excel::Range^ rng);
	private:
		Excel::Range^ wrappedRange;
	};
}
