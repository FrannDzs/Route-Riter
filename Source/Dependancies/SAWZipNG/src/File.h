// File.h : Declaration of the CFile

#ifndef __FILE_H_
#define __FILE_H_

#include "resource.h"       // main symbols
#include "ZipArchive.h"

/////////////////////////////////////////////////////////////////////////////
// CFile
class ATL_NO_VTABLE CFile : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CFile, &CLSID_File>,
	public IDispatchImpl<IFile, &IID_IFile, &LIBID_SAWZIPNGLib>
{
public:
	CFile()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_FILE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CFile)
	COM_INTERFACE_ENTRY(IFile)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

// IFile
public:
	STDMETHOD(get_Name)(/*[out, retval]*/ BSTR *pVal);

public:

    void SetFileHeader(const CZipFileHeader &hdr) { m_fileinfo = hdr; }
private:

    CZipFileHeader m_fileinfo;
};

#endif //__FILE_H_
