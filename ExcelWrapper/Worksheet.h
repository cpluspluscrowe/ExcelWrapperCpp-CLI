#pragma once
using namespace System;
using namespace Microsoft::Office::Interop::Excel;
namespace Excel = Microsoft::Office::Interop::Excel;
namespace WorksheetWrapper{
	ref class Worksheet
	{
	public:
		Worksheet(Excel::Worksheet^ worksheet);
	private:
		Excel::Worksheet^ wrappedWorksheet;
	};
}
