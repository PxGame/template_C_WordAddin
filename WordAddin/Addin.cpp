// Addin.cpp : CAddin µÄÊµÏÖ

#include "stdafx.h"
#include "Addin.h"


// CAddin

STDMETHODIMP CAddin::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* const arr[] = 
	{
		&IID_IAddin
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}
