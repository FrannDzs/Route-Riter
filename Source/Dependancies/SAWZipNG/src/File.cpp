// File.cpp : Implementation of CFile
#include "stdafx.h"
#include "SAWZipNG.h"
#include "File.h"

/////////////////////////////////////////////////////////////////////////////
// CFile


STDMETHODIMP CFile::get_Name(BSTR *pVal)
{
	CComBSTR(m_fileinfo.GetFileName()).CopyTo(pVal);
	return S_OK;
}

