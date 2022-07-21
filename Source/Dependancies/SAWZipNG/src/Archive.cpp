// Archive.cpp : Implementation of CArchive
#include "stdafx.h"
#include "SAWZipNG.h"
#include "Archive.h"
#include "FileInfo.h"

#include "ZipString.h"
#include "ZipPlatform.h"

#include "comvector.h"

/////////////////////////////////////////////////////////////////////////////
// CArchive

const int CArchive::callbackType =   CZipArchive::cbAdd 
                                   | CZipArchive::cbExtract
                                   | CZipArchive::cbDelete
                                   | CZipArchive::cbAddStore;

STDMETHODIMP CArchive::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IArchive
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CArchive::Open(BSTR filename, tagOpenMode openMode, int volumeSize)
{
	USES_CONVERSION;
 
    m_archive.SetCallback(new ArchiveActionCallback(this), callbackType);
    m_archive.SetSpanCallback(new ArchiveSpanCallback(this));

	// Check for ReadOnly
    try
    {
        m_archive.Open(OLE2T(filename), openMode == OM_READONLY ? CZipArchive::zipOpenReadOnly : CZipArchive::zipOpen, volumeSize);
    }
	catch (CZipException e)
    { 
		return Error(e.GetErrorDescription());
    }
    return S_OK;
}

// Returns true when the archive is readonly
STDMETHODIMP CArchive::get_ReadOnly(VARIANT_BOOL *pVal)
{
    *pVal = m_archive.IsReadOnly() ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

// Returns the number of files in the archive
STDMETHODIMP CArchive::get_FileCount(long *pVal)
{
    *pVal = m_archive.GetCount(true);
	return S_OK;
}

// Returns the number of directories in the archive
STDMETHODIMP CArchive::get_DirCount(long *pVal)
{
    *pVal = m_archive.GetCount() - m_archive.GetCount(true);
	return S_OK; 
}

// Returns the number of files and directories in the archive
STDMETHODIMP CArchive::get_Count(long *pVal)
{
    *pVal = m_archive.GetCount();
	return S_OK;
}

// Close the archive file
STDMETHODIMP CArchive::Close()
{
    m_archive.Close();
	return S_OK;
}

STDMETHODIMP CArchive::get_Closed(VARIANT_BOOL *pVal)
{
    *pVal = m_archive.IsClosed();
	return S_OK;
}

STDMETHODIMP CArchive::GetFileInfo(long index, IFileInfo **fi)
{
    CZipFileHeader header;

    try
    {
        m_archive.GetFileInfo(header, index);

        CComObject<CFileInfo> *pfile;
		CComObject<CFileInfo>::CreateInstance(&pfile);
        pfile->SetFileHeader(header);
		return pfile->QueryInterface(__uuidof(IFileInfo), (LPVOID *)fi);
    }
    catch(CZipException e)
    {
    }

	return S_OK;
}

STDMETHODIMP CArchive::Create(BSTR filename, tagCreateMode mode, int volumeSize)
{
    if ( ! m_archive.IsClosed() )
        m_archive.Close();

    m_archive.SetCallback(new ArchiveActionCallback(this), callbackType);
    m_archive.SetSpanCallback(new ArchiveSpanCallback(this));

    USES_CONVERSION;
    try
    {
        if (    mode == CM_CREATE_SPAN
             && volumeSize == 0 )
        {
            // Avoid exception which result in an access violation (don't know why)
            if ( ! ZipPlatform::IsDriveRemovable(OLE2T(filename)) )
            {
                return Error("A PKZip-span can only be created on a removable disk");
            }
        }
        m_archive.Open(OLE2T(filename), mode == CM_CREATE_SPAN ? CZipArchive::zipCreateSpan : CZipArchive::zipCreate, volumeSize);
    }
	catch (CZipException e)
    { 
        m_archive.Close(CZipArchive::afAfterException);
		return Error(e.GetErrorDescription());
    }
	return S_OK;
}

STDMETHODIMP CArchive::get_RootPath(BSTR *pVal)
{
	CComBSTR(m_archive.GetRootPath()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CArchive::put_RootPath(BSTR newVal)
{
	USES_CONVERSION;
    cstring path = OLE2T(newVal);
    if ( path.length() == 0 )
        m_archive.SetRootPath(NULL);
    else
        m_archive.SetRootPath(path.c_str());
	return S_OK;
}

STDMETHODIMP CArchive::AddFile(BSTR filename, VARIANT_BOOL fullpath, short level, tagSmartness smartLevel, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;
    try
    {
        *result = m_archive.AddNewFile(OLE2T(filename),
                                       level,
                                       fullpath == VARIANT_TRUE,
                                       smartLevel,
                                       bufferSize) ? VARIANT_TRUE : VARIANT_FALSE;
    }
    catch(CZipException e)
	{
		return Error(e.GetErrorDescription());
    }
                         
	return S_OK;
}

STDMETHODIMP CArchive::AddFileAs(BSTR filename, BSTR nameInZip, short level, tagSmartness smartLevel, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;
    try
    {
        *result = m_archive.AddNewFile(OLE2T(filename),
                                       OLE2T(nameInZip),
                                       level,
                                       smartLevel,
                                       bufferSize) ? VARIANT_TRUE : VARIANT_FALSE;
    }
    catch(CZipException e)
    {
		return Error(e.GetErrorDescription());
    }
                         
	return S_OK;
}

STDMETHODIMP CArchive::Extract(long index, BSTR path, VARIANT_BOOL fullpath, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;

	try
	{
		*result = m_archive.ExtractFile(index, OLE2T(path), fullpath == VARIANT_TRUE, NULL, bufferSize) ? VARIANT_TRUE : VARIANT_FALSE;
	}
	catch (CZipException e)
    { 
		return Error(e.GetErrorDescription());
    }
	return S_OK;
}

STDMETHODIMP CArchive::ExtractAs(long index, BSTR path, BSTR newName, VARIANT_BOOL fullpath, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;
    try
    {
	    *result = m_archive.ExtractFile(index, OLE2T(path), fullpath == VARIANT_TRUE, OLE2T(newName), bufferSize) ? VARIANT_TRUE : VARIANT_FALSE;
    }
	catch (CZipException e)
    { 
		return Error(e.GetErrorDescription());
    }
	return S_OK;
}

STDMETHODIMP CArchive::get_Password(BSTR *pVal)
{
	CComBSTR(m_archive.GetPassword()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CArchive::put_Password(BSTR newVal)
{
	USES_CONVERSION;
    cstring psw = OLE2T(newVal);
    if ( psw.length() == 0 )
        m_archive.SetPassword(NULL);
    else
        m_archive.SetPassword(psw.c_str());
	return S_OK;
}

STDMETHODIMP CArchive::PredictExtractedFileName(BSTR FileNameInZip, BSTR Path, VARIANT_BOOL FullPath, BSTR NewName, BSTR *result)
{
    USES_CONVERSION;

    LPCSTR TNewName = OLE2T(NewName);
    CZipString fileName = m_archive.PredictExtractedFileName(OLE2T(FileNameInZip),
                                                             OLE2T(Path),
                                                             FullPath == VARIANT_TRUE,
                                                             _tcslen(TNewName) == 0 ? NULL : TNewName);
    CComBSTR(fileName).CopyTo(result);
	return S_OK;
}

STDMETHODIMP CArchive::TestFile(long index, long bufferSize, VARIANT_BOOL *result)
{
    try
    {
	    *result = m_archive.TestFile(index, bufferSize) ? VARIANT_TRUE : VARIANT_FALSE;
	}
	catch (CZipException e)
    { 
		return Error(e.GetErrorDescription());
    }
	return S_OK;
}

STDMETHODIMP CArchive::get_WriteBufferSize(long *pVal)
{
    int size;
    m_archive.GetAdvanced(&size, NULL, NULL);
    *pVal = size;
	return S_OK;
}

STDMETHODIMP CArchive::put_WriteBufferSize(long newVal)
{
    m_archive.SetAdvanced(newVal);
	return S_OK;
}

STDMETHODIMP CArchive::get_GeneralBufferSize(long *pVal)
{
    int size;
    m_archive.GetAdvanced(NULL, &size, NULL);
    *pVal = size;
	return S_OK;
}

STDMETHODIMP CArchive::put_GeneralBufferSize(long newVal)
{
    int writeSize;
    int searchSize;
    m_archive.GetAdvanced(&writeSize, NULL, &searchSize);
    m_archive.SetAdvanced(writeSize, newVal, searchSize);
	return S_OK;
}

STDMETHODIMP CArchive::get_SearchBufferSize(long *pVal)
{
    int size;
    m_archive.GetAdvanced(NULL, NULL, &size);
    *pVal = size;
	return S_OK;
}

STDMETHODIMP CArchive::put_SearchBufferSize(long newVal)
{
    int writeSize;
    int generalSize;
    m_archive.GetAdvanced(&writeSize, &generalSize, NULL);
    m_archive.SetAdvanced(writeSize, generalSize, newVal);
	return S_OK;
}

STDMETHODIMP CArchive::get_Comment(BSTR *pVal)
{
	CComBSTR(m_archive.GetGlobalComment()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CArchive::put_Comment(BSTR newVal)
{
	USES_CONVERSION;
    m_archive.SetGlobalComment(OLE2T(newVal));
	return S_OK;
}

STDMETHODIMP CArchive::get_AutoFlush(VARIANT_BOOL *pVal)
{
    *pVal = m_archive.GetAutoFlush() ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CArchive::put_AutoFlush(VARIANT_BOOL newVal)
{
    m_archive.SetAutoFlush(newVal == VARIANT_TRUE);
	return S_OK;
}

STDMETHODIMP CArchive::Flush()
{
	m_archive.Flush();
	return S_OK;
}

STDMETHODIMP CArchive::AddFolder(BSTR foldername, VARIANT_BOOL includeSubDirs, VARIANT_BOOL fullpath, short level, tagSmartness smartLevel, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;
    cstring path(OLE2T(foldername));

	if ( *path.rbegin() != _T('\\') )
		path += _T('\\');

    std::vector<cstring> filelist;
    LoadAllFiles(path, "*.*", includeSubDirs == VARIANT_TRUE, filelist);

    *result = VARIANT_TRUE;
    for(std::vector<cstring>::iterator it = filelist.begin(); it != filelist.end(); it++)
    {
		try
		{
			if ( ! m_archive.AddNewFile(it->c_str(),
										level,
										fullpath == VARIANT_TRUE,
										smartLevel,
										bufferSize) )
			{
				*result = VARIANT_FALSE;
				break;
			}
		}
		catch (CZipException e)
		{ 
			return Error(e.GetErrorDescription());
		}
   }

	return S_OK;
}

void CArchive::LoadAllFiles(const cstring &path, const cstring &wildcard, bool recurse, std::vector<cstring> &list)
{
    WIN32_FIND_DATA findFileData;
    cstring find = path + wildcard;
    HANDLE handleFind = FindFirstFile(find.c_str() , &findFileData);
    if ( handleFind == INVALID_HANDLE_VALUE )
    {
         if ( recurse )
         {
            // No files found with the wildcard. If recurse is true try to find them in the
            // subdirectories
            find = path + "*.*";
            handleFind = FindFirstFile(find.c_str(), &findFileData);
            if ( handleFind == INVALID_HANDLE_VALUE )
                return;
        
            do
            {
	            if (   _tcscmp(findFileData.cFileName, _T(".")) == 0
		            || _tcscmp(findFileData.cFileName, _T("..")) == 0 )
		            continue;

		        if ( (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY )
		        {
			        LoadAllFiles(path + findFileData.cFileName + _T('\\'), wildcard, true, list);
			        continue;
		        }
            }
    	    while( FindNextFile(handleFind, &findFileData) );
        
            FindClose(handleFind);
         }
         return;
    }

    do
    {
        // Skip dots
	    if (   _tcscmp(findFileData.cFileName, _T(".")) == 0
		    || _tcscmp(findFileData.cFileName, _T("..")) == 0 )
		    continue;

        // Skip system files
		if ( (findFileData.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) == FILE_ATTRIBUTE_SYSTEM )
			continue;

		if ( 
                recurse
             && (findFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY 
           )
		{
			LoadAllFiles(path + findFileData.cFileName + _T('\\'), wildcard, recurse, list);
			continue;
		}

        list.push_back(path + findFileData.cFileName);
    }
	while( FindNextFile(handleFind, &findFileData) );
	
	FindClose(handleFind);
}

STDMETHODIMP CArchive::AddFolderWithWildcard(BSTR foldername, BSTR wildcard, VARIANT_BOOL includeSubDirs, VARIANT_BOOL fullPath, short level, tagSmartness smartlevel, long bufferSize, VARIANT_BOOL *result)
{
    USES_CONVERSION;
    cstring path(OLE2T(foldername));

	if ( *path.rbegin() != _T('\\') )
		path += _T('\\');

    std::vector<cstring> filelist;
    LoadAllFiles(path, OLE2T(wildcard), includeSubDirs == VARIANT_TRUE, filelist);

    *result = VARIANT_TRUE;
    for(std::vector<cstring>::iterator it = filelist.begin(); it != filelist.end(); it++)
    {
		try
		{
			if ( ! m_archive.AddNewFile(it->c_str(),
										level,
										fullPath == VARIANT_TRUE,
										smartlevel,
										bufferSize) )
			{
				*result = VARIANT_FALSE;
				break;
			}
		}
		catch (CZipException e)
		{ 
			return Error(e.GetErrorDescription());
		}
   }

	return S_OK;
}

STDMETHODIMP CArchive::DeleteFile(long index)
{
    m_archive.DeleteFile(index);
	return S_OK;
}

STDMETHODIMP CArchive::DeleteFiles(VARIANT indexes)
{
    if ( indexes.vt != (VT_ARRAY | VT_VARIANT) )
        return DISP_E_TYPEMISMATCH;

    CComVectorData<VARIANT> vIndexes(indexes.parray);
    if( !vIndexes )
        return E_UNEXPECTED;

    CZipWordArray array;
    for( int i = 0; i < vIndexes.Length(); ++i )
    {
        if ( vIndexes[i].vt == VT_I4 )
            array.Add(vIndexes[i].lVal);
        else if ( vIndexes[i].vt == VT_I2 )
            array.Add(vIndexes[i].iVal);
        else
            return DISP_E_TYPEMISMATCH;
    }
	
	try
	{
	    m_archive.DeleteFiles(array);
	}
	catch (CZipException e)
    { 
		return Error(e.GetErrorDescription());
    }

	return S_OK;
}

STDMETHODIMP CArchive::FindFiles(BSTR Pattern, VARIANT_BOOL Fullpath, VARIANT *indexes)
{
    if ( ! indexes )
        return E_POINTER;

    USES_CONVERSION;

    CZipWordArray array;
    m_archive.FindMatches(OLE2T(Pattern), array, Fullpath == VARIANT_TRUE);
    
    CComVector<VARIANT> vIndexes(array.GetSize());
    if( !vIndexes )
        return E_OUTOFMEMORY;

    CComVectorData<VARIANT> vData(vIndexes);
    if( !vData )
        return E_UNEXPECTED;

    for(int i = 0; i < array.GetSize(); i++)
        vData[i] = CComVariant(array[i]);

    vIndexes.DetachTo(indexes);
	return S_OK;
}

STDMETHODIMP CArchive::get_IgnoreCrc(VARIANT_BOOL *pVal)
{
    *pVal = m_ignoreCrc;
	return S_OK;
}

STDMETHODIMP CArchive::put_IgnoreCrc(VARIANT_BOOL newVal)
{
	m_ignoreCrc = newVal;
    m_archive.SetIgnoreCRC(newVal == VARIANT_TRUE);

	return S_OK;
}

STDMETHODIMP CArchive::FindFile(BSTR filename, tagFFCaseSensitivity caseSensitive, VARIANT_BOOL filenameOnly, long *index)
{
    CZipArchive::FFCaseSens c;
    switch(caseSensitive)
    {
    case FF_SENSITIVE:
        c = CZipArchive::ffCaseSens;
        break;
    case FF_NON_SENSITIVE:
        c = CZipArchive::ffNoCaseSens;
        break;
    default:
        c = CZipArchive::ffDefault;
        break;
    }

    USES_CONVERSION;
    *index = m_archive.FindFile(OLE2T(filename), c, filenameOnly == VARIANT_TRUE);
    return S_OK;
}

STDMETHODIMP CArchive::GetIndexes(VARIANT filenames, VARIANT *indexes)
{
    USES_CONVERSION;

    if ( ! indexes )
        return E_POINTER;

    if ( filenames.vt != (VT_ARRAY | VT_VARIANT) )
        return DISP_E_TYPEMISMATCH;

    CComVectorData<VARIANT> vFilenames(filenames.parray);
    if( !vFilenames )
        return E_UNEXPECTED;

    CZipStringArray array;
    for( int i = 0; i < vFilenames.Length(); ++i )
    {
        if ( vFilenames[i].vt == VT_BSTR )
            array.Add(OLE2T(vFilenames[i].bstrVal));
        else
            return DISP_E_TYPEMISMATCH;
    }

    CZipWordArray indexArray;
    m_archive.GetIndexes(array, indexArray);
        
    CComVector<VARIANT> vIndexes(indexArray.GetSize());
    if( !vIndexes )
        return E_OUTOFMEMORY;

    CComVectorData<VARIANT> vData(vIndexes);
    if( !vData )
        return E_UNEXPECTED;

    for(i = 0; i < indexArray.GetSize(); i++)
        vData[i] = CComVariant(indexArray[i]);

    vIndexes.DetachTo(indexes);
 
    return S_OK;
}

STDMETHODIMP CArchive::get_ArchivePath(BSTR *pVal)
{
	CComBSTR(m_archive.GetArchivePath()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CArchive::SetFileComment(long index, BSTR comment)
{
	USES_CONVERSION;

    m_archive.SetFileComment(index, OLE2T(comment));
	return S_OK;
}

STDMETHODIMP CArchive::get_SystemCompatibility(tagZipPlatform *pVal)
{
    *pVal = (Platform) m_archive.GetSystemCompatibility();
	return S_OK;
}

STDMETHODIMP CArchive::put_SystemCompatibility(tagZipPlatform newVal)
{
    m_archive.SetSystemCompatibility(newVal);
	return S_OK;
}

STDMETHODIMP CArchive::get_TempPath(BSTR *pVal)
{
	CComBSTR(m_archive.GetTempPath()).CopyTo(pVal);
	return S_OK;
}

STDMETHODIMP CArchive::put_TempPath(BSTR newVal)
{
	USES_CONVERSION;

    m_archive.SetTempPath(OLE2T(newVal));
	return S_OK;
}

STDMETHODIMP CArchive::PredictFileNameInZip(BSTR FileName, VARIANT_BOOL FullPath, VARIANT_BOOL Exact, BSTR *FileNameInZip)
{
    USES_CONVERSION;

    CComBSTR(m_archive.PredictFileNameInZip(OLE2T(FileName), 
                                            FullPath == VARIANT_TRUE,
                                            CZipArchive::prAuto, 
                                            Exact == VARIANT_TRUE)).CopyTo(FileNameInZip);
	return S_OK;
}

STDMETHODIMP CArchive::WillBeDuplicated(BSTR FilePath, VARIANT_BOOL FullPath, VARIANT_BOOL FileNameOnly, long *index)
{
    USES_CONVERSION;

    *index = m_archive.WillBeDuplicated(OLE2T(FilePath),
                                        FullPath == VARIANT_TRUE,
                                        FileNameOnly == VARIANT_TRUE,
                                        CZipArchive::prAuto);

	return S_OK;
}

STDMETHODIMP CArchive::get_CurrentDisk(short *pVal)
{
    *pVal = m_archive.GetCurrentDisk();
	return S_OK;
}

STDMETHODIMP CArchive::get_SpanMode(tagSpanMode *pVal)
{
	*pVal = (tagSpanMode) m_archive.GetSpanMode();
	return S_OK;
}

STDMETHODIMP CArchive::get_EnableFindFast(VARIANT_BOOL *pVal)
{
    *pVal = m_enableFindFast;
	return S_OK;
}

STDMETHODIMP CArchive::put_EnableFindFast(VARIANT_BOOL newVal)
{
    m_enableFindFast = newVal;
    m_archive.EnableFindFast(newVal == VARIANT_TRUE);
	return S_OK;
}

STDMETHODIMP CArchive::RenameFile(long index, BSTR newFilename, VARIANT_BOOL *result)
{
    USES_CONVERSION;

    *result = m_archive.RenameFile(index, OLE2T(newFilename));

	return S_OK;
}

void CArchive::Fire(CZipActionCallback *cb)
{
}

STDMETHODIMP CArchive::get_CaseSensitivity(VARIANT_BOOL *pVal)
{
    *pVal = m_case;
	return S_OK;
}

STDMETHODIMP CArchive::put_CaseSensitivity(VARIANT_BOOL newVal)
{
	// TODO: Add your implementation code here
    m_case = newVal;
    m_archive.SetCaseSensitivity(newVal == VARIANT_TRUE);
	return S_OK;
}
