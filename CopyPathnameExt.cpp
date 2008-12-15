// CopyPathnameExt.cpp : Implementation of CCopyPathnameExt

#include "stdafx.h"
#include "CopyPathnameExt.h"

#include <atlconv.h>

STDMETHODIMP CCopyPathnameExt::Initialize(LPCITEMIDLIST folder, LPDATAOBJECT dataObj, HKEY progID)
{
    FORMATETC fmt = {CF_HDROP, 0, DVASPECT_CONTENT, -1, TYMED_HGLOBAL};
    STGMEDIUM stg = {TYMED_HGLOBAL};
    if (FAILED(dataObj->GetData(&fmt, &stg)))
        return E_INVALIDARG;

    HDROP drop = (HDROP)GlobalLock(stg.hGlobal);
    if (!drop)
        return E_INVALIDARG;

    HRESULT hr = S_OK;
    UINT nfiles = DragQueryFile(drop, 0xFFFFFFFF, 0, 0);
    if (nfiles == 1) {
        if (DragQueryFileA(drop, 0, m_pathname, MAX_PATH)) {
            // We return S_OK, so Explorer will call QueryInterface again
            // and get a pointer to the IContextMenu interface.
        } else {
            hr = E_INVALIDARG;
        }
    } else {
        hr = E_INVALIDARG;
    }

    GlobalUnlock(stg.hGlobal);
    ReleaseStgMedium(&stg);
    return hr;

}

STDMETHODIMP CCopyPathnameExt::GetCommandString(UINT command,
                                                UINT flags,
                                                UINT *,
                                                LPSTR name,
                                                UINT maxChars)
{
    USES_CONVERSION;

    if (command != 0)
        return E_INVALIDARG;

    if (!(flags & GCS_HELPTEXT))
        return E_INVALIDARG;

    LPCTSTR help = _T("Copy pathname to clipboard");

    if (flags & GCS_UNICODE)
        lstrcpynW((LPWSTR)name, T2CW(help), maxChars);
    else
        lstrcpynA(name, T2CA(help), maxChars);

    return S_OK;
}

namespace
{
    bool CopyTextToClipboard(HWND hwnd, LPCSTR text)
    {
        bool copied = false;
        
        if (OpenClipboard(hwnd)) {
            if (EmptyClipboard()) {
                if (text[0]) {
                    HGLOBAL hBuf = GlobalAlloc(GHND, lstrlenA(text) + 1);
                    if (hBuf) {
                        char *buf = static_cast<char *>(GlobalLock(hBuf));
                        if (buf) {
                            lstrcpyA(buf, text);
                            if (GlobalUnlock(hBuf) == 0 && GetLastError() == NO_ERROR) {
                                if (SetClipboardData(CF_TEXT, hBuf))
                                    copied = true;
                            }
                        }
                        if (copied == false)
                            GlobalFree(hBuf);
                    }
                } else {
                    copied = true;
                }
            }
            CloseClipboard();
        }

        return copied;
    }
}

STDMETHODIMP CCopyPathnameExt::InvokeCommand(LPCMINVOKECOMMANDINFO cmdInfo)
{
    if (HIWORD(cmdInfo->lpVerb) != 0)
        return E_INVALIDARG;

    switch (LOWORD(cmdInfo->lpVerb)) {
        case 0:
            {
                CopyTextToClipboard(cmdInfo->hwnd, m_pathname);
            }
            return S_OK;
        default:
            return E_INVALIDARG;
    }
}

STDMETHODIMP CCopyPathnameExt::QueryContextMenu(HMENU menu, 
                                                UINT index, 
                                                UINT firstCommand, 
                                                UINT lastCommand,
                                                UINT flags)
{
    if (flags & CMF_DEFAULTONLY)
        return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 0);

    InsertMenu(menu, index, MF_STRING | MF_BYPOSITION, firstCommand, _T("Copy pathname"));
    
    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, 1);
}

