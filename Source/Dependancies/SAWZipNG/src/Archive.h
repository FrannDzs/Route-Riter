// Archive.h : Declaration of the CArchive

#ifndef __ARCHIVE_H_
#define __ARCHIVE_H_

/* BACKUP events, due to the variant_bool, out bug:
	HRESULT Fire_OnDiskNeeded(LONG disk, VARIANT_BOOL * Cancel)
	{
		CComVariant varResult;
		T* pT = static_cast<T*>(this);
		int nConnectionIndex;
		CComVariant* pvars = new CComVariant[2];
		int nConnections = m_vec.GetSize();

		for (nConnectionIndex = 0; nConnectionIndex < nConnections; nConnectionIndex++)
		{
			pT->Lock();
			CComPtr<IUnknown> sp = m_vec.GetAt(nConnectionIndex);
			pT->Unlock();
			IDispatch* pDispatch = reinterpret_cast<IDispatch*>(sp.p);
			if (pDispatch != NULL)
			{
				VariantClear(&varResult);
				pvars[1] = disk;
//				pvars[0] = Cancel;
                pvars[0].vt = VT_BOOL | VT_BYREF;
                pvars[0].byref = Cancel;
				DISPPARAMS disp = { pvars, NULL, 2, 0 };
				pDispatch->Invoke(0x5, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, &varResult, NULL, NULL);
			}
		}
		delete[] pvars;
		return varResult.scode;
	
	}
*/

#include <string>
#ifdef _UNICODE
 typedef std::wstring cstring;
#else
 typedef std::string cstring;
#endif
#include <vector>

#include "resource.h"       // main symbols

#include "ZipArchive.h"
#include "SAWZipNGCP.h"

/////////////////////////////////////////////////////////////////////////////
// CArchive
class ATL_NO_VTABLE CArchive : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CArchive, &CLSID_Archive>,
	public ISupportErrorInfo,
	public IConnectionPointContainerImpl<CArchive>,
	public IDispatchImpl<IArchive, &IID_IArchive, &LIBID_SAWZipNG>,
	public CProxy_IArchiveEvents< CArchive >
{
public:

    CArchive() : m_ignoreCrc(VARIANT_FALSE), 
                 m_enableFindFast(VARIANT_FALSE), 
                 m_case(VARIANT_FALSE)
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_ARCHIVE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CArchive)
	COM_INTERFACE_ENTRY(IArchive)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
	COM_INTERFACE_ENTRY(IConnectionPointContainer)
	COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
END_COM_MAP()

BEGIN_CONNECTION_POINT_MAP(CArchive)
CONNECTION_POINT_ENTRY(DIID__IArchiveEvents)
END_CONNECTION_POINT_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IArchive
public:
	STDMETHOD(get_CaseSensitivity)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_CaseSensitivity)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(RenameFile)(/*[in]*/ long index, /*[in]*/ BSTR newFilename, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(get_EnableFindFast)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_EnableFindFast)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_SpanMode)(/*[out, retval]*/ tagSpanMode *pVal);
	STDMETHOD(get_CurrentDisk)(/*[out, retval]*/ short *pVal);
	STDMETHOD(WillBeDuplicated)(/*[in]*/ BSTR FilePath, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL FullPath, /*[in, optional, defaultvalue(0)]*/ VARIANT_BOOL FileNameOnly, /*[out, retval]*/ long *index);
	STDMETHOD(PredictFileNameInZip)(/*[in]*/ BSTR FileName, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL FullPath, /*[in, optional, defaultvalue(0)]*/ VARIANT_BOOL Exact, /*[out, retval]*/ BSTR *FileNameInZip);
	STDMETHOD(get_TempPath)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_TempPath)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_SystemCompatibility)(/*[out, retval]*/ tagZipPlatform *pVal);
	STDMETHOD(put_SystemCompatibility)(/*[in]*/ tagZipPlatform newVal);
	STDMETHOD(SetFileComment)(/*[in]*/ long index, /*[in]*/ BSTR comment);
	STDMETHOD(get_ArchivePath)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(GetIndexes)(/*[in]*/ VARIANT filenames, /*[out, retval]*/ VARIANT *indexes);
	STDMETHOD(FindFile)(/*[in]*/ BSTR filename, /*[in, optional, defaultvalue(FF_DEFAULT)]*/ tagFFCaseSensitivity caseSensitive, /*[in, optional, defaultvalue(0)]*/ VARIANT_BOOL filenameOnly, /*[out, retval]*/ long *index);
	STDMETHOD(get_IgnoreCrc)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_IgnoreCrc)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(FindFiles)(/*[in]*/ BSTR Pattern, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL Fullpath, /*[out, retval]*/ VARIANT *indexes);
	STDMETHOD(DeleteFiles)(/*[in*/ VARIANT indexes);
	STDMETHOD(DeleteFile)(/*[in]*/ long index);
    STDMETHOD(AddFolderWithWildcard)(/*[in]*/ BSTR foldername, /*[in]*/ BSTR wildcard, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL includeSubDirs, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL fullPath, /*[in, optional, defaultvalue(-1)]*/ short level, /*[in, optional, defaultvalue(0)]*/ tagSmartness smartlevel, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(AddFolder)(/*[in]*/ BSTR foldername, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL includeSubDirs, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL fullpath, /*[in, optional, defaultvalue(-1)]*/ short level, /*[in, optional, defaultvalue(0)]*/ tagSmartness smartLevel, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(Flush)();
	STDMETHOD(get_AutoFlush)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_AutoFlush)(/*[in]*/ VARIANT_BOOL newVal);
	STDMETHOD(get_Comment)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Comment)(/*[in]*/ BSTR newVal);
	STDMETHOD(get_SearchBufferSize)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_SearchBufferSize)(/*[in]*/ long newVal);
	STDMETHOD(get_GeneralBufferSize)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_GeneralBufferSize)(/*[in]*/ long newVal);
	STDMETHOD(get_WriteBufferSize)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_WriteBufferSize)(/*[in]*/ long newVal);
	STDMETHOD(TestFile)(/*[in]*/ long index, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(PredictExtractedFileName)(/*[in]*/ BSTR FileNameInZip, /*[in]*/ BSTR Path, /*[in]*/ VARIANT_BOOL FullPath, /*[in]*/ BSTR NewName, /*[out, retval]*/ BSTR* result);
	STDMETHOD(get_Password)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Password)(/*[in]*/ BSTR newVal);
	STDMETHOD(ExtractAs)(/*[in]*/ long index, /*[in]*/ BSTR path, /*[in]*/ BSTR newName, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL fullpath, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(Extract)(/*[in]*/ long index, /*[in]*/ BSTR path, /*[in, optional, defaultvalue(1)]*/ VARIANT_BOOL fullpath, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(AddFileAs)(/*[in]*/ BSTR filename, /*[in]*/ BSTR nameInZip, /*[in, optional, defaultvalue(-1)]*/ short level, /*[in, optional, defaultvalue(0)]*/ tagSmartness smartLevel, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(AddFile)(/*[in]*/ BSTR filename, /*[in, optional, defaultvalue(true)]*/ VARIANT_BOOL fullpath, /*[in, optional, defaultvalue(-1)]*/ short level, /*[in, optional, defaultvalue(0)]*/ tagSmartness smartLevel, /*[in, optional, defaultvalue(65536)]*/ long bufferSize, /*[out, retval]*/ VARIANT_BOOL *result);
	STDMETHOD(get_RootPath)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_RootPath)(/*[in]*/ BSTR newVal);
	STDMETHOD(Create)(/*[in]*/ BSTR filename, /*[in, optional, defaultvalue(CM_CREATE)]*/ tagCreateMode createMode, /*[in, optional, defaultvalue(0)]*/ int volumeSize);
	STDMETHOD(GetFileInfo)(/*[in]*/ long index, /*[out, retval]*/ IFileInfo **fi);
	STDMETHOD(get_Closed)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(Close)();
	STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_DirCount)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_FileCount)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_ReadOnly)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(Open)(/*[in]*/ BSTR filename, /*[in, optional, defaultvalue(OM_OPEN)]*/ tagOpenMode openMode, /*[in, optional, defaultvalue(0)]*/ int volumeSize);

    void Fire(CZipActionCallback *cb);

    void FinalRelease()
	{
        if ( ! m_archive.IsClosed() )
		    m_archive.Close();
	}

private:
    CZipArchive m_archive;

    static const int callbackType;

    void LoadAllFiles(const cstring &path, const cstring &wildcards, bool recurse, std::vector<cstring> &list);
    VARIANT_BOOL m_ignoreCrc;
    VARIANT_BOOL m_enableFindFast;
    VARIANT_BOOL m_case;
};

class ArchiveActionCallback : public CZipActionCallback
{
public:
    ArchiveActionCallback(CArchive *archive) : m_archive(archive)
    {
    }

    virtual ~ArchiveActionCallback()
    {
    }

    inline bool Callback(int iProgress)
    {
        BSTR filename;
        CComBSTR(m_szFileInZip).CopyTo(&filename);
        VARIANT_BOOL cancel;

        switch(m_iType)
        {
        case CZipArchive::cbAdd:
            {
                m_archive->Fire_OnAdd(filename, m_uTotalSoFar, m_uTotalToDo, &cancel);
            }
        case CZipArchive::cbExtract:
            {
                m_archive->Fire_OnExtract(filename, m_uTotalSoFar, m_uTotalToDo, &cancel);
            }
        case CZipArchive::cbDelete:
            {
                m_archive->Fire_OnDelete(filename, m_uTotalSoFar, m_uTotalToDo, &cancel);
            }
        case CZipArchive::cbAddStore:
            {
                m_archive->Fire_OnStore(filename, m_uTotalSoFar, m_uTotalToDo, &cancel);
            }
        }
        if ( cancel == VARIANT_TRUE )
            return false;

        return true;
    }
private:
    CArchive *m_archive;
};

class ArchiveSpanCallback : public CZipSpanCallback
{
public:
    ArchiveSpanCallback(CArchive *archive) : m_archive(archive)
    {
    }

    virtual ~ArchiveSpanCallback()
    {
    }

    inline bool Callback(int iProgress)
    {
        switch(iProgress)
        {
        case -1:
            {
                VARIANT_BOOL cancel;
                m_archive->Fire_OnDiskNeeded(m_uDiskNeeded, &cancel);
                if ( cancel == VARIANT_TRUE )
                    return false;
                break;
            }
        case -2:
            break;
        case -3:
            break;
        case -4:
            break;
        default:
            ;
        }
        return true;
    }
private:
    CArchive *m_archive;
};

#endif //__ARCHIVE_H_
