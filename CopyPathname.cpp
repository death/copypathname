// CopyPathname.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "resource.h"

// The module attribute causes DllMain, DllRegisterServer and DllUnregisterServer to be automatically implemented for you
[ module(dll, uuid = "{269825B6-049A-4C89-BD68-8616196560EC}", 
		 name = "CopyPathname", 
		 helpstring = "CopyPathname 1.0 Type Library",
		 resource_name = "IDR_COPYPATHNAME") ]
class CCopyPathnameModule
{
public:
    HRESULT DllRegisterServer(BOOL typelib = TRUE)
    {
        if (!(GetVersion() & 0x80000000UL)) {
            CRegKey reg;
            LONG lRet = reg.Open(HKEY_LOCAL_MACHINE,
                                 _T("Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"),
                                 KEY_SET_VALUE);

            if (lRet != ERROR_SUCCESS)
                return E_ACCESSDENIED;

            lRet = reg.SetValue(_T("CopyPathnameExt"), 
                                _T("{2EECB91C-3E6A-4D08-BDCA-05FB48CCE9F3}"));

            if (lRet != ERROR_SUCCESS)
                return E_ACCESSDENIED;
        }

        {
            CRegKey reg;
            LONG ret = reg.Open(HKEY_CLASSES_ROOT,
                                _T("*\\shellex\\ContextMenuHandlers"),
                                KEY_SET_VALUE | KEY_CREATE_SUB_KEY);
            if (ret != ERROR_SUCCESS)
                return E_ACCESSDENIED;

            ret = reg.SetKeyValue(_T("CopyPathnameExt"),
                                  _T("{2EECB91C-3E6A-4D08-BDCA-05FB48CCE9F3}"));
            if (ret != ERROR_SUCCESS)
                return E_ACCESSDENIED;
        }

        {
            CRegKey reg;
            LONG ret = reg.Open(HKEY_CLASSES_ROOT,
                                _T("Directory\\shellex\\ContextMenuHandlers"),
                                KEY_SET_VALUE | KEY_CREATE_SUB_KEY);
            if (ret != ERROR_SUCCESS)
                return E_ACCESSDENIED;

            ret = reg.SetKeyValue(_T("CopyPathnameExt"),
                                  _T("{2EECB91C-3E6A-4D08-BDCA-05FB48CCE9F3}"));
            if (ret != ERROR_SUCCESS)
                return E_ACCESSDENIED;
        }

        return CAtlDllModuleT<CCopyPathnameModule>::DllRegisterServer(typelib);
    }

    HRESULT DllUnregisterServer(BOOL typelib = TRUE)
    {
        if (!(GetVersion() & 0x80000000UL)) {
            CRegKey reg;
            LONG lRet = reg.Open(HKEY_LOCAL_MACHINE,
                                 _T("Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved"),
                                 KEY_SET_VALUE);

            if (lRet == ERROR_SUCCESS)
                lRet = reg.DeleteValue(_T("{2EECB91C-3E6A-4D08-BDCA-05FB48CCE9F3}"));
        }

        {
            CRegKey reg;
            LONG ret = reg.Open(HKEY_CLASSES_ROOT,
                                _T("*\\shellex\\ContextMenuHandlers"),
                                KEY_SET_VALUE);
            if (ret == ERROR_SUCCESS)
                reg.DeleteSubKey(_T("CopyPathnameExt"));
        }

        {
            CRegKey reg;
            LONG ret = reg.Open(HKEY_CLASSES_ROOT,
                                _T("Directory\\shellex\\ContextMenuHandlers"),
                                KEY_SET_VALUE);
            if (ret == ERROR_SUCCESS)
                reg.DeleteSubKey(_T("CopyPathnameExt"));
        }

        return CAtlDllModuleT<CCopyPathnameModule>::DllUnregisterServer(typelib);
    }
};
