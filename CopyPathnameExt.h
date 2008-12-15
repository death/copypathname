// CopyPathnameExt.h : Declaration of the CCopyPathnameExt

#pragma once
#include "resource.h"       // main symbols

#include <shlobj.h>
#include <comdef.h>
#include <string>
#include <vector>

// CCopyPathnameExt

[
	coclass,
	threading(apartment),
	aggregatable(never),
	vi_progid("CopyPathname.CopyPathnameExt"),
	progid("CopyPathname.CopyPathnameExt.1"),
	version(1.0),
	uuid("2EECB91C-3E6A-4D08-BDCA-05FB48CCE9F3"),
	helpstring("CopyPathnameExt Class")
]
class ATL_NO_VTABLE CCopyPathnameExt
    : public IShellExtInit
    , public IContextMenu
{
public:
	CCopyPathnameExt()
	{
	}

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

    // IShellExtInit
    STDMETHODIMP Initialize(LPCITEMIDLIST folder, LPDATAOBJECT dataObj, HKEY progID);

    // IContextMenu
    STDMETHODIMP GetCommandString(UINT, UINT, UINT *, LPSTR, UINT);
    STDMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO);
    STDMETHODIMP QueryContextMenu(HMENU, UINT, UINT, UINT, UINT);


private:
    std::vector<std::string> m_pathnames;
};

