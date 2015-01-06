// dllmain.h : 模块类的声明。

class CWordAddinModule : public ATL::CAtlDllModuleT< CWordAddinModule >
{
public :
	DECLARE_LIBID(LIBID_WordAddinLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_WORDADDIN, "{58AD4C76-0B87-4D9E-9166-735595B601DC}")
};

extern class CWordAddinModule _AtlModule;
