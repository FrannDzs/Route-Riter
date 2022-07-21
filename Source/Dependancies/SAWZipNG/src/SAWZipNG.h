/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Jul 13 13:29:47 2003
 */
/* Compiler settings for C:\SAWZipNG\src\SAWZipNG.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __SAWZipNG_h__
#define __SAWZipNG_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IArchive_FWD_DEFINED__
#define __IArchive_FWD_DEFINED__
typedef interface IArchive IArchive;
#endif 	/* __IArchive_FWD_DEFINED__ */


#ifndef ___IArchiveEvents_FWD_DEFINED__
#define ___IArchiveEvents_FWD_DEFINED__
typedef interface _IArchiveEvents _IArchiveEvents;
#endif 	/* ___IArchiveEvents_FWD_DEFINED__ */


#ifndef __IFileInfo_FWD_DEFINED__
#define __IFileInfo_FWD_DEFINED__
typedef interface IFileInfo IFileInfo;
#endif 	/* __IFileInfo_FWD_DEFINED__ */


#ifndef __Archive_FWD_DEFINED__
#define __Archive_FWD_DEFINED__

#ifdef __cplusplus
typedef class Archive Archive;
#else
typedef struct Archive Archive;
#endif /* __cplusplus */

#endif 	/* __Archive_FWD_DEFINED__ */


#ifndef __FileInfo_FWD_DEFINED__
#define __FileInfo_FWD_DEFINED__

#ifdef __cplusplus
typedef class FileInfo FileInfo;
#else
typedef struct FileInfo FileInfo;
#endif /* __cplusplus */

#endif 	/* __FileInfo_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

/* interface __MIDL_itf_SAWZipNG_0000 */
/* [local] */ 

typedef /* [v1_enum] */ 
enum OpenMode
    {	OM_OPEN	= 0,
	OM_READONLY	= 1
    }	tagOpenMode;

typedef /* [v1_enum] */ 
enum CreateMode
    {	CM_CREATE	= 0,
	CM_CREATE_SPAN	= 1
    }	tagCreateMode;

typedef /* [v1_enum] */ 
enum FFCaseSensitivity
    {	FF_DEFAULT	= 0,
	FF_SENSITIVE	= 1,
	FF_NON_SENSITIVE	= 2
    }	tagFFCaseSensitivity;

typedef /* [v1_enum] */ 
enum Smartness
    {	SM_LAZY	= 0,
	SM_CPASSDIR	= 0x1,
	SM_CPFILE	= 0x2,
	SM_NOT_COMP_SMALL	= 0x4,
	SM_CHECK_FOR_EFF	= 0x8,
	SM_MEMORY_FLAG	= 0x10,
	SM_CHECK_FOR_EFF_IN_MEM	= SM_MEMORY_FLAG | SM_CHECK_FOR_EFF,
	SM_SMART_PASS	= SM_CPASSDIR | SM_CPFILE,
	SM_SMART_ADD	= SM_NOT_COMP_SMALL | SM_CHECK_FOR_EFF,
	SM_SMART_SAFE	= SM_SMART_PASS | SM_NOT_COMP_SMALL,
	SM_SMART_TEST	= SM_SMART_PASS | SM_SMART_ADD
    }	tagSmartness;

typedef /* [v1_enum] */ 
enum Platform
    {	ZP_DOS_FAT	= 0,
	ZP_AMIGA	= ZP_DOS_FAT + 1,
	ZP_VAX_VMS	= ZP_AMIGA + 1,
	ZP_UNIX	= ZP_VAX_VMS + 1,
	ZP_VM_CMS	= ZP_UNIX + 1,
	ZP_ATARI	= ZP_VM_CMS + 1,
	ZP_OS2_HPFS	= ZP_ATARI + 1,
	ZP_MAC	= ZP_OS2_HPFS + 1,
	ZP_Z_SYSTEM	= ZP_MAC + 1,
	ZP_CP_M	= ZP_Z_SYSTEM + 1,
	ZP_NTFS	= ZP_CP_M + 1
    }	tagZipPlatform;

typedef /* [v1_enum] */ 
enum SpanMode
    {	SPM_EXIST_TD	= -2,
	SPM_EXIST_PKZIP	= -1,
	SPM_NO	= 0,
	SPM_CREATE_PKZIP	= 1,
	SPM_CREATE_TD	= 2
    }	tagSpanMode;




extern RPC_IF_HANDLE __MIDL_itf_SAWZipNG_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_SAWZipNG_0000_v0_0_s_ifspec;

#ifndef __IArchive_INTERFACE_DEFINED__
#define __IArchive_INTERFACE_DEFINED__

/* interface IArchive */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IArchive;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D4C11DAF-8B64-11D7-923B-000000000000")
    IArchive : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Open( 
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagOpenMode openMode = OM_OPEN,
            /* [defaultvalue][optional][in] */ int volumeSize = 0) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ReadOnly( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_FileCount( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_DirCount( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Count( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Close( void) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Closed( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetFileInfo( 
            /* [in] */ long index,
            /* [retval][out] */ IFileInfo __RPC_FAR *__RPC_FAR *fi) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Create( 
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagCreateMode createMode = CM_CREATE,
            /* [defaultvalue][optional][in] */ int volumeSize = 0) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_RootPath( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_RootPath( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddFile( 
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddFileAs( 
            /* [in] */ BSTR filename,
            /* [in] */ BSTR nameInZip,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Extract( 
            /* [in] */ long index,
            /* [in] */ BSTR path,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ExtractAs( 
            /* [in] */ long index,
            /* [in] */ BSTR path,
            /* [in] */ BSTR newName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Password( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Password( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PredictExtractedFileName( 
            /* [in] */ BSTR FileNameInZip,
            /* [in] */ BSTR Path,
            /* [in] */ VARIANT_BOOL FullPath,
            /* [in] */ BSTR NewName,
            /* [retval][out] */ BSTR __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE TestFile( 
            /* [in] */ long index,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_WriteBufferSize( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_WriteBufferSize( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_GeneralBufferSize( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_GeneralBufferSize( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SearchBufferSize( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_SearchBufferSize( 
            /* [in] */ long newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Comment( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FindFile( 
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagFFCaseSensitivity caseSensitive,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL filenameOnly,
            /* [retval][out] */ long __RPC_FAR *index) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Comment( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_AutoFlush( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_AutoFlush( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Flush( void) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddFolder( 
            /* [in] */ BSTR foldername,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE AddFolderWithWildcard( 
            /* [in] */ BSTR foldername,
            /* [in] */ BSTR wildcard,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullPath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DeleteFile( 
            /* [in] */ long index) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE DeleteFiles( 
            /* [in] */ VARIANT indexes) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FindFiles( 
            /* [in] */ BSTR Pattern,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL Fullpath,
            /* [retval][out] */ VARIANT __RPC_FAR *indexes) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_IgnoreCrc( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_IgnoreCrc( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetIndexes( 
            /* [in] */ VARIANT filenames,
            /* [retval][out] */ VARIANT __RPC_FAR *indexes) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ArchivePath( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetFileComment( 
            /* [in] */ long index,
            /* [in] */ BSTR comment) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SystemCompatibility( 
            /* [retval][out] */ tagZipPlatform __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_SystemCompatibility( 
            /* [in] */ tagZipPlatform newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_TempPath( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_TempPath( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE PredictFileNameInZip( 
            /* [in] */ BSTR FileName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL Exact,
            /* [retval][out] */ BSTR __RPC_FAR *FileNameInZip) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE WillBeDuplicated( 
            /* [in] */ BSTR FilePath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FileNameOnly,
            /* [retval][out] */ long __RPC_FAR *index) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CurrentDisk( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SpanMode( 
            /* [retval][out] */ tagSpanMode __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_EnableFindFast( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_EnableFindFast( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE RenameFile( 
            /* [in] */ long index,
            /* [in] */ BSTR newFilename,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CaseSensitivity( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_CaseSensitivity( 
            /* [in] */ VARIANT_BOOL newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IArchiveVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IArchive __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IArchive __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IArchive __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IArchive __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IArchive __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IArchive __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IArchive __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Open )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagOpenMode openMode,
            /* [defaultvalue][optional][in] */ int volumeSize);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ReadOnly )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_FileCount )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_DirCount )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Count )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Close )( 
            IArchive __RPC_FAR * This);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Closed )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetFileInfo )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [retval][out] */ IFileInfo __RPC_FAR *__RPC_FAR *fi);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Create )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagCreateMode createMode,
            /* [defaultvalue][optional][in] */ int volumeSize);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_RootPath )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_RootPath )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddFile )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddFileAs )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR filename,
            /* [in] */ BSTR nameInZip,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Extract )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [in] */ BSTR path,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *ExtractAs )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [in] */ BSTR path,
            /* [in] */ BSTR newName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Password )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Password )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PredictExtractedFileName )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR FileNameInZip,
            /* [in] */ BSTR Path,
            /* [in] */ VARIANT_BOOL FullPath,
            /* [in] */ BSTR NewName,
            /* [retval][out] */ BSTR __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *TestFile )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_WriteBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_WriteBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_GeneralBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_GeneralBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SearchBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_SearchBufferSize )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Comment )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindFile )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR filename,
            /* [defaultvalue][optional][in] */ tagFFCaseSensitivity caseSensitive,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL filenameOnly,
            /* [retval][out] */ long __RPC_FAR *index);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Comment )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_AutoFlush )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_AutoFlush )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Flush )( 
            IArchive __RPC_FAR * This);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddFolder )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR foldername,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *AddFolderWithWildcard )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR foldername,
            /* [in] */ BSTR wildcard,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL fullPath,
            /* [defaultvalue][optional][in] */ short level,
            /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
            /* [defaultvalue][optional][in] */ long bufferSize,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DeleteFile )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DeleteFiles )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT indexes);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FindFiles )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR Pattern,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL Fullpath,
            /* [retval][out] */ VARIANT __RPC_FAR *indexes);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_IgnoreCrc )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_IgnoreCrc )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIndexes )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT filenames,
            /* [retval][out] */ VARIANT __RPC_FAR *indexes);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ArchivePath )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetFileComment )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [in] */ BSTR comment);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SystemCompatibility )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ tagZipPlatform __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_SystemCompatibility )( 
            IArchive __RPC_FAR * This,
            /* [in] */ tagZipPlatform newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_TempPath )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_TempPath )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *PredictFileNameInZip )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR FileName,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL Exact,
            /* [retval][out] */ BSTR __RPC_FAR *FileNameInZip);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *WillBeDuplicated )( 
            IArchive __RPC_FAR * This,
            /* [in] */ BSTR FilePath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
            /* [defaultvalue][optional][in] */ VARIANT_BOOL FileNameOnly,
            /* [retval][out] */ long __RPC_FAR *index);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CurrentDisk )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SpanMode )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ tagSpanMode __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_EnableFindFast )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_EnableFindFast )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *RenameFile )( 
            IArchive __RPC_FAR * This,
            /* [in] */ long index,
            /* [in] */ BSTR newFilename,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CaseSensitivity )( 
            IArchive __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_CaseSensitivity )( 
            IArchive __RPC_FAR * This,
            /* [in] */ VARIANT_BOOL newVal);
        
        END_INTERFACE
    } IArchiveVtbl;

    interface IArchive
    {
        CONST_VTBL struct IArchiveVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IArchive_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IArchive_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IArchive_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IArchive_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IArchive_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IArchive_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IArchive_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IArchive_Open(This,filename,openMode,volumeSize)	\
    (This)->lpVtbl -> Open(This,filename,openMode,volumeSize)

#define IArchive_get_ReadOnly(This,pVal)	\
    (This)->lpVtbl -> get_ReadOnly(This,pVal)

#define IArchive_get_FileCount(This,pVal)	\
    (This)->lpVtbl -> get_FileCount(This,pVal)

#define IArchive_get_DirCount(This,pVal)	\
    (This)->lpVtbl -> get_DirCount(This,pVal)

#define IArchive_get_Count(This,pVal)	\
    (This)->lpVtbl -> get_Count(This,pVal)

#define IArchive_Close(This)	\
    (This)->lpVtbl -> Close(This)

#define IArchive_get_Closed(This,pVal)	\
    (This)->lpVtbl -> get_Closed(This,pVal)

#define IArchive_GetFileInfo(This,index,fi)	\
    (This)->lpVtbl -> GetFileInfo(This,index,fi)

#define IArchive_Create(This,filename,createMode,volumeSize)	\
    (This)->lpVtbl -> Create(This,filename,createMode,volumeSize)

#define IArchive_get_RootPath(This,pVal)	\
    (This)->lpVtbl -> get_RootPath(This,pVal)

#define IArchive_put_RootPath(This,newVal)	\
    (This)->lpVtbl -> put_RootPath(This,newVal)

#define IArchive_AddFile(This,filename,fullpath,level,smartLevel,bufferSize,result)	\
    (This)->lpVtbl -> AddFile(This,filename,fullpath,level,smartLevel,bufferSize,result)

#define IArchive_AddFileAs(This,filename,nameInZip,level,smartLevel,bufferSize,result)	\
    (This)->lpVtbl -> AddFileAs(This,filename,nameInZip,level,smartLevel,bufferSize,result)

#define IArchive_Extract(This,index,path,fullpath,bufferSize,result)	\
    (This)->lpVtbl -> Extract(This,index,path,fullpath,bufferSize,result)

#define IArchive_ExtractAs(This,index,path,newName,fullpath,bufferSize,result)	\
    (This)->lpVtbl -> ExtractAs(This,index,path,newName,fullpath,bufferSize,result)

#define IArchive_get_Password(This,pVal)	\
    (This)->lpVtbl -> get_Password(This,pVal)

#define IArchive_put_Password(This,newVal)	\
    (This)->lpVtbl -> put_Password(This,newVal)

#define IArchive_PredictExtractedFileName(This,FileNameInZip,Path,FullPath,NewName,result)	\
    (This)->lpVtbl -> PredictExtractedFileName(This,FileNameInZip,Path,FullPath,NewName,result)

#define IArchive_TestFile(This,index,bufferSize,result)	\
    (This)->lpVtbl -> TestFile(This,index,bufferSize,result)

#define IArchive_get_WriteBufferSize(This,pVal)	\
    (This)->lpVtbl -> get_WriteBufferSize(This,pVal)

#define IArchive_put_WriteBufferSize(This,newVal)	\
    (This)->lpVtbl -> put_WriteBufferSize(This,newVal)

#define IArchive_get_GeneralBufferSize(This,pVal)	\
    (This)->lpVtbl -> get_GeneralBufferSize(This,pVal)

#define IArchive_put_GeneralBufferSize(This,newVal)	\
    (This)->lpVtbl -> put_GeneralBufferSize(This,newVal)

#define IArchive_get_SearchBufferSize(This,pVal)	\
    (This)->lpVtbl -> get_SearchBufferSize(This,pVal)

#define IArchive_put_SearchBufferSize(This,newVal)	\
    (This)->lpVtbl -> put_SearchBufferSize(This,newVal)

#define IArchive_get_Comment(This,pVal)	\
    (This)->lpVtbl -> get_Comment(This,pVal)

#define IArchive_FindFile(This,filename,caseSensitive,filenameOnly,index)	\
    (This)->lpVtbl -> FindFile(This,filename,caseSensitive,filenameOnly,index)

#define IArchive_put_Comment(This,newVal)	\
    (This)->lpVtbl -> put_Comment(This,newVal)

#define IArchive_get_AutoFlush(This,pVal)	\
    (This)->lpVtbl -> get_AutoFlush(This,pVal)

#define IArchive_put_AutoFlush(This,newVal)	\
    (This)->lpVtbl -> put_AutoFlush(This,newVal)

#define IArchive_Flush(This)	\
    (This)->lpVtbl -> Flush(This)

#define IArchive_AddFolder(This,foldername,includeSubDirs,fullpath,level,smartLevel,bufferSize,result)	\
    (This)->lpVtbl -> AddFolder(This,foldername,includeSubDirs,fullpath,level,smartLevel,bufferSize,result)

#define IArchive_AddFolderWithWildcard(This,foldername,wildcard,includeSubDirs,fullPath,level,smartLevel,bufferSize,result)	\
    (This)->lpVtbl -> AddFolderWithWildcard(This,foldername,wildcard,includeSubDirs,fullPath,level,smartLevel,bufferSize,result)

#define IArchive_DeleteFile(This,index)	\
    (This)->lpVtbl -> DeleteFile(This,index)

#define IArchive_DeleteFiles(This,indexes)	\
    (This)->lpVtbl -> DeleteFiles(This,indexes)

#define IArchive_FindFiles(This,Pattern,Fullpath,indexes)	\
    (This)->lpVtbl -> FindFiles(This,Pattern,Fullpath,indexes)

#define IArchive_get_IgnoreCrc(This,pVal)	\
    (This)->lpVtbl -> get_IgnoreCrc(This,pVal)

#define IArchive_put_IgnoreCrc(This,newVal)	\
    (This)->lpVtbl -> put_IgnoreCrc(This,newVal)

#define IArchive_GetIndexes(This,filenames,indexes)	\
    (This)->lpVtbl -> GetIndexes(This,filenames,indexes)

#define IArchive_get_ArchivePath(This,pVal)	\
    (This)->lpVtbl -> get_ArchivePath(This,pVal)

#define IArchive_SetFileComment(This,index,comment)	\
    (This)->lpVtbl -> SetFileComment(This,index,comment)

#define IArchive_get_SystemCompatibility(This,pVal)	\
    (This)->lpVtbl -> get_SystemCompatibility(This,pVal)

#define IArchive_put_SystemCompatibility(This,newVal)	\
    (This)->lpVtbl -> put_SystemCompatibility(This,newVal)

#define IArchive_get_TempPath(This,pVal)	\
    (This)->lpVtbl -> get_TempPath(This,pVal)

#define IArchive_put_TempPath(This,newVal)	\
    (This)->lpVtbl -> put_TempPath(This,newVal)

#define IArchive_PredictFileNameInZip(This,FileName,FullPath,Exact,FileNameInZip)	\
    (This)->lpVtbl -> PredictFileNameInZip(This,FileName,FullPath,Exact,FileNameInZip)

#define IArchive_WillBeDuplicated(This,FilePath,FullPath,FileNameOnly,index)	\
    (This)->lpVtbl -> WillBeDuplicated(This,FilePath,FullPath,FileNameOnly,index)

#define IArchive_get_CurrentDisk(This,pVal)	\
    (This)->lpVtbl -> get_CurrentDisk(This,pVal)

#define IArchive_get_SpanMode(This,pVal)	\
    (This)->lpVtbl -> get_SpanMode(This,pVal)

#define IArchive_get_EnableFindFast(This,pVal)	\
    (This)->lpVtbl -> get_EnableFindFast(This,pVal)

#define IArchive_put_EnableFindFast(This,newVal)	\
    (This)->lpVtbl -> put_EnableFindFast(This,newVal)

#define IArchive_RenameFile(This,index,newFilename,result)	\
    (This)->lpVtbl -> RenameFile(This,index,newFilename,result)

#define IArchive_get_CaseSensitivity(This,pVal)	\
    (This)->lpVtbl -> get_CaseSensitivity(This,pVal)

#define IArchive_put_CaseSensitivity(This,newVal)	\
    (This)->lpVtbl -> put_CaseSensitivity(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_Open_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR filename,
    /* [defaultvalue][optional][in] */ tagOpenMode openMode,
    /* [defaultvalue][optional][in] */ int volumeSize);


void __RPC_STUB IArchive_Open_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_ReadOnly_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_ReadOnly_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_FileCount_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_FileCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_DirCount_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_DirCount_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_Count_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_Count_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_Close_Proxy( 
    IArchive __RPC_FAR * This);


void __RPC_STUB IArchive_Close_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_Closed_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_Closed_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_GetFileInfo_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [retval][out] */ IFileInfo __RPC_FAR *__RPC_FAR *fi);


void __RPC_STUB IArchive_GetFileInfo_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_Create_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR filename,
    /* [defaultvalue][optional][in] */ tagCreateMode createMode,
    /* [defaultvalue][optional][in] */ int volumeSize);


void __RPC_STUB IArchive_Create_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_RootPath_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_RootPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_RootPath_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArchive_put_RootPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_AddFile_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR filename,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
    /* [defaultvalue][optional][in] */ short level,
    /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_AddFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_AddFileAs_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR filename,
    /* [in] */ BSTR nameInZip,
    /* [defaultvalue][optional][in] */ short level,
    /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_AddFileAs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_Extract_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [in] */ BSTR path,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_Extract_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_ExtractAs_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [in] */ BSTR path,
    /* [in] */ BSTR newName,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_ExtractAs_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_Password_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_Password_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArchive_put_Password_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_PredictExtractedFileName_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR FileNameInZip,
    /* [in] */ BSTR Path,
    /* [in] */ VARIANT_BOOL FullPath,
    /* [in] */ BSTR NewName,
    /* [retval][out] */ BSTR __RPC_FAR *result);


void __RPC_STUB IArchive_PredictExtractedFileName_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_TestFile_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_TestFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_WriteBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_WriteBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_WriteBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IArchive_put_WriteBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_GeneralBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_GeneralBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_GeneralBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IArchive_put_GeneralBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_SearchBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_SearchBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_SearchBufferSize_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long newVal);


void __RPC_STUB IArchive_put_SearchBufferSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_Comment_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_FindFile_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR filename,
    /* [defaultvalue][optional][in] */ tagFFCaseSensitivity caseSensitive,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL filenameOnly,
    /* [retval][out] */ long __RPC_FAR *index);


void __RPC_STUB IArchive_FindFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_Comment_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArchive_put_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_AutoFlush_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_AutoFlush_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_AutoFlush_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IArchive_put_AutoFlush_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_Flush_Proxy( 
    IArchive __RPC_FAR * This);


void __RPC_STUB IArchive_Flush_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_AddFolder_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR foldername,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL fullpath,
    /* [defaultvalue][optional][in] */ short level,
    /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_AddFolder_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_AddFolderWithWildcard_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR foldername,
    /* [in] */ BSTR wildcard,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL includeSubDirs,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL fullPath,
    /* [defaultvalue][optional][in] */ short level,
    /* [defaultvalue][optional][in] */ tagSmartness smartLevel,
    /* [defaultvalue][optional][in] */ long bufferSize,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_AddFolderWithWildcard_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_DeleteFile_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index);


void __RPC_STUB IArchive_DeleteFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_DeleteFiles_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT indexes);


void __RPC_STUB IArchive_DeleteFiles_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_FindFiles_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR Pattern,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL Fullpath,
    /* [retval][out] */ VARIANT __RPC_FAR *indexes);


void __RPC_STUB IArchive_FindFiles_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_IgnoreCrc_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_IgnoreCrc_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_IgnoreCrc_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IArchive_put_IgnoreCrc_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_GetIndexes_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT filenames,
    /* [retval][out] */ VARIANT __RPC_FAR *indexes);


void __RPC_STUB IArchive_GetIndexes_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_ArchivePath_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_ArchivePath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_SetFileComment_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [in] */ BSTR comment);


void __RPC_STUB IArchive_SetFileComment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_SystemCompatibility_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ tagZipPlatform __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_SystemCompatibility_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_SystemCompatibility_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ tagZipPlatform newVal);


void __RPC_STUB IArchive_put_SystemCompatibility_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_TempPath_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_TempPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_TempPath_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IArchive_put_TempPath_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_PredictFileNameInZip_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR FileName,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL Exact,
    /* [retval][out] */ BSTR __RPC_FAR *FileNameInZip);


void __RPC_STUB IArchive_PredictFileNameInZip_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_WillBeDuplicated_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ BSTR FilePath,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL FullPath,
    /* [defaultvalue][optional][in] */ VARIANT_BOOL FileNameOnly,
    /* [retval][out] */ long __RPC_FAR *index);


void __RPC_STUB IArchive_WillBeDuplicated_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_CurrentDisk_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_CurrentDisk_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_SpanMode_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ tagSpanMode __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_SpanMode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_EnableFindFast_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_EnableFindFast_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_EnableFindFast_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IArchive_put_EnableFindFast_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IArchive_RenameFile_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ long index,
    /* [in] */ BSTR newFilename,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *result);


void __RPC_STUB IArchive_RenameFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IArchive_get_CaseSensitivity_Proxy( 
    IArchive __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IArchive_get_CaseSensitivity_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IArchive_put_CaseSensitivity_Proxy( 
    IArchive __RPC_FAR * This,
    /* [in] */ VARIANT_BOOL newVal);


void __RPC_STUB IArchive_put_CaseSensitivity_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IArchive_INTERFACE_DEFINED__ */



#ifndef __SAWZipNG_LIBRARY_DEFINED__
#define __SAWZipNG_LIBRARY_DEFINED__

/* library SAWZipNG */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_SAWZipNG;

#ifndef ___IArchiveEvents_DISPINTERFACE_DEFINED__
#define ___IArchiveEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IArchiveEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__IArchiveEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D4C11DB1-8B64-11D7-923B-000000000000")
    _IArchiveEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IArchiveEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            _IArchiveEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            _IArchiveEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            _IArchiveEvents __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            _IArchiveEvents __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            _IArchiveEvents __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            _IArchiveEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            _IArchiveEvents __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        END_INTERFACE
    } _IArchiveEventsVtbl;

    interface _IArchiveEvents
    {
        CONST_VTBL struct _IArchiveEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IArchiveEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _IArchiveEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _IArchiveEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _IArchiveEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _IArchiveEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _IArchiveEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _IArchiveEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IArchiveEvents_DISPINTERFACE_DEFINED__ */


#ifndef __IFileInfo_INTERFACE_DEFINED__
#define __IFileInfo_INTERFACE_DEFINED__

/* interface IFileInfo */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IFileInfo;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D4C11DB4-8B64-11D7-923B-000000000000")
    IFileInfo : public IDispatch
    {
    public:
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Filename( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Filename( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Comment( 
            /* [retval][out] */ BSTR __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_Comment( 
            /* [in] */ BSTR newVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CompressionRatio( 
            /* [retval][out] */ float __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_CompressionSize( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Attributes( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Encrypted( 
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_UncompressedSize( 
            /* [retval][out] */ long __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ModificationDate( 
            /* [retval][out] */ DATE __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE put_ModificationDate( 
            /* [in] */ DATE newVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IFileInfoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IFileInfo __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IFileInfo __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IFileInfo __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Filename )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Filename )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Comment )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ BSTR __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_Comment )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ BSTR newVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CompressionRatio )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ float __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_CompressionSize )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Attributes )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Encrypted )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_UncompressedSize )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ long __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ModificationDate )( 
            IFileInfo __RPC_FAR * This,
            /* [retval][out] */ DATE __RPC_FAR *pVal);
        
        /* [helpstring][id][propput] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *put_ModificationDate )( 
            IFileInfo __RPC_FAR * This,
            /* [in] */ DATE newVal);
        
        END_INTERFACE
    } IFileInfoVtbl;

    interface IFileInfo
    {
        CONST_VTBL struct IFileInfoVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IFileInfo_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IFileInfo_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IFileInfo_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IFileInfo_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IFileInfo_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IFileInfo_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IFileInfo_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IFileInfo_get_Filename(This,pVal)	\
    (This)->lpVtbl -> get_Filename(This,pVal)

#define IFileInfo_put_Filename(This,newVal)	\
    (This)->lpVtbl -> put_Filename(This,newVal)

#define IFileInfo_get_Comment(This,pVal)	\
    (This)->lpVtbl -> get_Comment(This,pVal)

#define IFileInfo_put_Comment(This,newVal)	\
    (This)->lpVtbl -> put_Comment(This,newVal)

#define IFileInfo_get_CompressionRatio(This,pVal)	\
    (This)->lpVtbl -> get_CompressionRatio(This,pVal)

#define IFileInfo_get_CompressionSize(This,pVal)	\
    (This)->lpVtbl -> get_CompressionSize(This,pVal)

#define IFileInfo_get_Attributes(This,pVal)	\
    (This)->lpVtbl -> get_Attributes(This,pVal)

#define IFileInfo_get_Encrypted(This,pVal)	\
    (This)->lpVtbl -> get_Encrypted(This,pVal)

#define IFileInfo_get_UncompressedSize(This,pVal)	\
    (This)->lpVtbl -> get_UncompressedSize(This,pVal)

#define IFileInfo_get_ModificationDate(This,pVal)	\
    (This)->lpVtbl -> get_ModificationDate(This,pVal)

#define IFileInfo_put_ModificationDate(This,newVal)	\
    (This)->lpVtbl -> put_ModificationDate(This,newVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_Filename_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_Filename_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFileInfo_put_Filename_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFileInfo_put_Filename_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_Comment_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ BSTR __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFileInfo_put_Comment_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [in] */ BSTR newVal);


void __RPC_STUB IFileInfo_put_Comment_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_CompressionRatio_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ float __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_CompressionRatio_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_CompressionSize_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_CompressionSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_Attributes_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_Attributes_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_Encrypted_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ VARIANT_BOOL __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_Encrypted_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_UncompressedSize_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ long __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_UncompressedSize_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IFileInfo_get_ModificationDate_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [retval][out] */ DATE __RPC_FAR *pVal);


void __RPC_STUB IFileInfo_get_ModificationDate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propput] */ HRESULT STDMETHODCALLTYPE IFileInfo_put_ModificationDate_Proxy( 
    IFileInfo __RPC_FAR * This,
    /* [in] */ DATE newVal);


void __RPC_STUB IFileInfo_put_ModificationDate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IFileInfo_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Archive;

#ifdef __cplusplus

class DECLSPEC_UUID("D4C11DB0-8B64-11D7-923B-000000000000")
Archive;
#endif

EXTERN_C const CLSID CLSID_FileInfo;

#ifdef __cplusplus

class DECLSPEC_UUID("D4C11DB5-8B64-11D7-923B-000000000000")
FileInfo;
#endif
#endif /* __SAWZipNG_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

unsigned long             __RPC_USER  VARIANT_UserSize(     unsigned long __RPC_FAR *, unsigned long            , VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  VARIANT_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, VARIANT __RPC_FAR * ); 
void                      __RPC_USER  VARIANT_UserFree(     unsigned long __RPC_FAR *, VARIANT __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
