// FileInfo.cpp : Implementation of CFileInfo
#include "stdafx.h"
#include "SAWZipNG.h"
#include "FileInfo.h"

/////////////////////////////////////////////////////////////////////////////
// CFileInfo


STDMETHODIMP CFileInfo::get_Filename(BSTR *pVal)
{
	CComBSTR(m_fileHdr.GetFileName()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CFileInfo::put_Filename(BSTR newVal)
{
	USES_CONVERSION;
    m_fileHdr.SetFileName(OLE2T(newVal));
	return S_OK;
}

STDMETHODIMP CFileInfo::get_Comment(BSTR *pVal)
{
	CComBSTR(m_fileHdr.GetComment()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CFileInfo::put_Comment(BSTR newVal)
{
	USES_CONVERSION;
    m_fileHdr.SetComment(OLE2T(newVal));
	return S_OK;
}

STDMETHODIMP CFileInfo::get_CompressionRatio(float *pVal)
{
	*pVal = m_fileHdr.GetCompressionRatio();
	return S_OK;
}

STDMETHODIMP CFileInfo::get_CompressionSize(long *pVal)
{
    *pVal = m_fileHdr.GetEffComprSize();
	return S_OK;
}

STDMETHODIMP CFileInfo::get_Attributes(long *pVal)
{
    *pVal = m_fileHdr.GetOriginalAttributes();
	return S_OK;
}

STDMETHODIMP CFileInfo::get_Encrypted(VARIANT_BOOL *pVal)
{
    *pVal = m_fileHdr.IsEncrypted() ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CFileInfo::get_UncompressedSize(long *pVal)
{
    *pVal = m_fileHdr.m_uUncomprSize;
	return S_OK;
}


STDMETHODIMP CFileInfo::get_ModificationDate(DATE *pVal)
{
	DosDateTimeToVariantTime(m_fileHdr.m_uModDate, m_fileHdr.m_uModTime, pVal);
	return S_OK;
}

STDMETHODIMP CFileInfo::put_ModificationDate(DATE newVal)
{
	VariantTimeToDosDateTime(newVal, &m_fileHdr.m_uModDate, &m_fileHdr.m_uModTime);
	return S_OK;
}

STDMETHODIMP CFileInfo::get_Crc32(long *pVal)
{
    *pVal = m_fileHdr.m_uCrc32;
	return S_OK;
}

STDMETHODIMP CFileInfo::get_IsDirectory(VARIANT_BOOL *pVal)
{
    *pVal = m_fileHdr.IsDirectory() ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
