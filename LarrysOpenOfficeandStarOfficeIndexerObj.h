// LarrysOpenOfficeandStarOfficeIndexerObj.h : Declaration of the CLarrysOpenOfficeandStarOfficeIndexerObj

#pragma once
#include "resource.h"       // main symbols

#include "LarrysOpenOfficeandStarOfficeIndexer.h"

// CLarrysOpenOfficeandStarOfficeIndexerObj

class ATL_NO_VTABLE CLarrysOpenOfficeandStarOfficeIndexerObj : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CLarrysOpenOfficeandStarOfficeIndexerObj, &CLSID_LarrysOpenOfficeandStarOfficeIndexerObj>,
	public ISupportErrorInfo,
	public IDispatchImpl<ILarrysOpenOfficeandStarOfficeIndexerObj, &IID_ILarrysOpenOfficeandStarOfficeIndexerObj, &LIBID_LarrysOpenOfficeandStarOfficeIndexerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CLarrysOpenOfficeandStarOfficeIndexerObj()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_LARRYSOPENOFFICEANDSTAROFFICEINDEXEROBJ)


BEGIN_COM_MAP(CLarrysOpenOfficeandStarOfficeIndexerObj)
	COM_INTERFACE_ENTRY(ILarrysOpenOfficeandStarOfficeIndexerObj)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
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

	STDMETHOD(HandleFile)(BSTR RawFullPath, IDispatch* IFactory);
};

OBJECT_ENTRY_AUTO(__uuidof(LarrysOpenOfficeandStarOfficeIndexerObj), CLarrysOpenOfficeandStarOfficeIndexerObj)
