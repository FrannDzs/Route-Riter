// FileInfo.h : Declaration of the CFileInfo

#ifndef __FILEINFO_H_
#define __FILEINFO_H_

#include "resource.h"       // main symbols
#include "ZipArchive.h"

/////////////////////////////////////////////////////////////////////////////
// CFileInfo
class ATL_NO_VTABLE CFileInfo : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFileInfo, &CLSID_FileInfo>,
	public IDispatchImpl<IFileInfo, &IID_IFileInfo, &LIBID_SAWZipNG>
{
public:
	CFileInfo()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_FILEINFO)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFileInfo)
	COM_INTERFACE_ENTRY(IFileInfo)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// IFileInfo
public:
	STDMETHOD(get_IsDirectory)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_Crc32)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_ModificationDate)(/*[out, retval]*/ DATE *pVal);
	STDMETHOD(put_ModificationDate)(/*[in]*/ DATE newVal);
	STDMETHOD(get_UncompressedSize)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_Encrypted)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_Attributes)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_CompressionSize)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_CompressionRatio)(/*[out, retval]*/ float *pVal);
	STDMETHOD(get_Comment)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Comment)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_Filename)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Filename)(/*[in]*/ BSTR newVal);

    void SetFileHeader(const CZipFileHeader &hdr) { m_fileHdr = hdr; }

private:
    CZipFileHeader m_fileHdr;
};

#endif //__FILEINFO_H_
