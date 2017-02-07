#include <msclr/marshal_cppstd.h>
#include <memory>
#include <string>

namespace ExcelApplicationWrapper{
	class Convert{
	public:
		static std::shared_ptr<std::string> GetSPNativeString(System::String^ sValue){
			if (sValue != nullptr){
				return std::make_shared<std::string>(msclr::interop::marshal_as<std::string>(sValue));
			}
			else{
				return nullptr;
			}
		}
		static std::shared_ptr<double> GetSPNativeDouble(double^ dValue){
			if (dValue != nullptr){
				double val = (double)dValue;
				return std::make_shared<double>(val);
			}
			else{
				return nullptr;
			}
		}
	};
}
