// LarrysOpenOfficeandStarOfficeIndexer.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"
#include "LarrysOpenOfficeandStarOfficeIndexer.h"

class CLarrysOpenOfficeandStarOfficeIndexerModule : public CAtlDllModuleT< CLarrysOpenOfficeandStarOfficeIndexerModule >
{
public :
	DECLARE_LIBID(LIBID_LarrysOpenOfficeandStarOfficeIndexerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_LARRYSOPENOFFICEANDSTAROFFICEINDEXER, "{069CD576-C879-4744-B367-E0E1603453E9}")
};

CLarrysOpenOfficeandStarOfficeIndexerModule _AtlModule;


// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
	if (SUCCEEDED(hr)) {
		LarGDSPlugin GDSPlugin(PluginClassID);

		vector<string> Extensions;

		// OpenOffice 1.x files
		Extensions.push_back("sxc");
		Extensions.push_back("stc");
		Extensions.push_back("sxd");
		Extensions.push_back("std");
		Extensions.push_back("sxi");
		Extensions.push_back("sti");
		Extensions.push_back("sxw");
		Extensions.push_back("sxg");
		Extensions.push_back("stw");
		Extensions.push_back("stm");

		// OpenDocument files
		Extensions.push_back("odt");
		Extensions.push_back("ott");
		Extensions.push_back("odg");
		Extensions.push_back("otg");
		Extensions.push_back("odp");
		Extensions.push_back("otp");
		Extensions.push_back("ods");
		Extensions.push_back("ots");
		Extensions.push_back("odf");

		GDSPlugin.RegisterPluginWithExtensions("Larry's OpenOffice and StarOffice Indexer 1.01", "Indexes all OpenOffice.org/StarOffice6+ files", "", Extensions);
	}

	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();

	if (SUCCEEDED(hr)) {
		LarGDSPlugin GDSPlugin(PluginClassID);
		GDSPlugin.UnregisterPlugin();
	}
	
	return hr;
}