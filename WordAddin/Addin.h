// Addin.h : CAddin 的声明

#pragma once
#include "resource.h"       // 主符号



#include "WordAddin_i.h"




#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Windows CE 平台(如不提供完全 DCOM 支持的 Windows Mobile 平台)上无法正确支持单线程 COM 对象。定义 _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA 可强制 ATL 支持创建单线程 COM 对象实现并允许使用其单线程 COM 对象实现。rgs 文件中的线程模型已被设置为“Free”，原因是该模型是非 DCOM Windows CE 平台支持的唯一线程模型。"
#endif

using namespace ATL;


// CAddin

_ATL_FUNC_INFO DocumentOpenInfo = {CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH | VT_BYREF}};
_ATL_FUNC_INFO DocumentBeforeCloseInfo = {CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH | VT_BYREF,VT_BYREF | VT_BOOL}};
_ATL_FUNC_INFO StartupInfo = {CC_STDCALL,VT_EMPTY,0};
_ATL_FUNC_INFO QuitInfo = {CC_STDCALL,VT_EMPTY,0};

class ATL_NO_VTABLE CAddin :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAddin, &CLSID_Addin>,
	public ISupportErrorInfo,
	public IDispatchImpl<IAddin, &IID_IAddin, &LIBID_WordAddinLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &LIBID_AddInDesignerObjects, /* wMajor = */ 1>,
	public IDispEventSimpleImpl<10 ,CAddin, &__uuidof(Word::ApplicationEvents2)>,
	public IDispEventSimpleImpl<11 ,CAddin, &__uuidof(Word::ApplicationEvents2)>,
	public IDispEventSimpleImpl<12 ,CAddin, &__uuidof(Word::ApplicationEvents2)>,
	public IDispEventSimpleImpl<13 ,CAddin, &__uuidof(Word::ApplicationEvents2)>
{
public:
	CAddin()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)


	BEGIN_COM_MAP(CAddin)
		COM_INTERFACE_ENTRY(IAddin)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(ISupportErrorInfo)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()

	// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		OutputDebugString(_T("OnConnection"));
		DocumentOpenEvent::DispEventAdvise(Application, &__uuidof(Word::ApplicationEvents2));
		DocumentBeforeCloseEvent::DispEventAdvise(Application, &__uuidof(Word::ApplicationEvents2));
		StartupEvent::DispEventAdvise(Application, &__uuidof(Word::ApplicationEvents2));
		QuitEvent::DispEventAdvise(Application, &__uuidof(Word::ApplicationEvents2));
		return E_NOTIMPL;
	}
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		OutputDebugString(_T("OnDisconnection"));
		return E_NOTIMPL;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		OutputDebugString(_T("OnAddInsUpdate"));
		return E_NOTIMPL;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		OutputDebugString(_T("OnStartupComplete"));
		return E_NOTIMPL;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		OutputDebugString(_T("OnBeginShutdown"));
		return E_NOTIMPL;
	}

public:
	typedef IDispEventSimpleImpl<10 ,CAddin, &__uuidof(Word::ApplicationEvents2)> DocumentOpenEvent;  // 打开一个文档的事件
	typedef IDispEventSimpleImpl<11 ,CAddin, &__uuidof(Word::ApplicationEvents2)> DocumentBeforeCloseEvent; // 关闭文件的事件
	typedef IDispEventSimpleImpl<12 ,CAddin, &__uuidof(Word::ApplicationEvents2)> StartupEvent;  // 打开Word事件
	typedef IDispEventSimpleImpl<13 ,CAddin, &__uuidof(Word::ApplicationEvents2)> QuitEvent; // 关闭Word事件

	void __stdcall DocumentOpen (/*[in]*/ struct _Document * Doc )
	{
		OutputDebugString(_T("DocumentOpen"));
	}
	void __stdcall DocumentBeforeClose (/*[in]*/ struct _Document * Doc,/*[in]*/ VARIANT_BOOL * Cancel )
	{
		OutputDebugString(_T("DocumentBeforeClose"));
	}
	void __stdcall Startup ()
	{
		OutputDebugString(_T("Startup"));
	}
	void __stdcall Quit ()
	{
		OutputDebugString(_T("Quit"));
	}

	BEGIN_SINK_MAP(CAddin)
		SINK_ENTRY_INFO(10,__uuidof(Word::ApplicationEvents2),/*dispinterface*/0x04,DocumentOpen, &DocumentOpenInfo)
		SINK_ENTRY_INFO(11,__uuidof(Word::ApplicationEvents2),/*dispinterface*/0x06,DocumentBeforeClose, &DocumentBeforeCloseInfo)
		SINK_ENTRY_INFO(12,__uuidof(Word::ApplicationEvents2),/*dispinterface*/0x01,Startup, &StartupInfo)
		SINK_ENTRY_INFO(13,__uuidof(Word::ApplicationEvents2),/*dispinterface*/0x02,Quit, &QuitInfo)
	END_SINK_MAP()

};

OBJECT_ENTRY_AUTO(__uuidof(Addin), CAddin)
