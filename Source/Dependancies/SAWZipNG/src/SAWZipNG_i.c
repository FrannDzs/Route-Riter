/* this file contains the actual definitions of */
/* the IIDs and CLSIDs */

/* link this file in with the server and any clients */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Jul 13 13:29:47 2003
 */
/* Compiler settings for C:\SAWZipNG\src\SAWZipNG.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )
#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

const IID IID_IArchive = {0xD4C11DAF,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


const IID LIBID_SAWZipNG = {0xD4C11DA3,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


const IID DIID__IArchiveEvents = {0xD4C11DB1,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


const IID IID_IFileInfo = {0xD4C11DB4,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


const CLSID CLSID_Archive = {0xD4C11DB0,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


const CLSID CLSID_FileInfo = {0xD4C11DB5,0x8B64,0x11D7,{0x92,0x3B,0x00,0x00,0x00,0x00,0x00,0x00}};


#ifdef __cplusplus
}
#endif

