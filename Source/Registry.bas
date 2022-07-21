Attribute VB_Name = "Registry"
'***************************************************************************
'*
'* File: Registry.Bas
'*
'* Type: Visual Module Subroutine
'*
'* Name: Basic Registry Read/Write/Delete
'*
'* Additional Notes:
'*
'* $History: Registry.bas $
'*
'*      *****************  Version 2  *****************
'*      User: Qzdcg8       Date: 9/10/00    Time: 10:02
'*      Updated in $/NCAS ORT-GI3/References/ORT QT
'*      Fixed process thread retention problem and added Quick Query Toolbar
'*      capability
'*
'*      *****************  Version 1  *****************
'*      User: Qzdcg8       Date: 15/02/00   Time: 11:36
'*      Created in $/NCAS ORT-GI3/References/ORT QT
'*      Components of ORTQT/ISQL
'*
'*      *****************  Version 9  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:36
'*      Updated in $/NCAS ORT/Reference/ORT MM
'*      06/08/98 - was using incorrect version of ORTHTMLDocument class, also
'*      missing dcaSecurity.dll reference
'*
'*      *****************  Version 8  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:06
'*      Updated in $/NCAS ORT/Reference/ORT DH
'*      06/08/98 - was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 7  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:04
'*      Updated in $/NCAS ORT/Reference/ORT FS
'*      06/08/98 - Was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 6  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:03
'*      Updated in $/NCAS ORT/Reference/ORT QA
'*      06/08/98 - Was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 5  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:01
'*      Updated in $/NCAS ORT/Reference/ORT RDE
'*      06/08/98 - Was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 4  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 12:00
'*      Updated in $/NCAS ORT/Reference/ORT RR
'*      06/08/98 - Was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 3  *****************
'*      User: Mzrh69       Date: 6/08/98    Time: 11:55
'*      Updated in $/NCAS ORT/Reference/ORT RV
'*      06/08/98 - Was using incorrect version of ORTHTMLDocument class
'*
'*      *****************  Version 2  *****************
'*      User: Mzrh69       Date: 5/08/98    Time: 15:34
'*      Updated in $/NCAS ORT/Reference/ORT MM
'*      Missing class ORTSearchField added
'*
'*      *****************  Version 1  *****************
'*      User: Mzrh69       Date: 3/08/98    Time: 16:19
'*      Created in $/NCAS ORT/Reference/Registry
'*
'*      *****************  Version 1  *****************
'*      User: Lz5xjh       Date: 6/07/98    Time: 9:55
'*      Created in $/NCAS RDM/Sources/UI Controllers/dcaRDMExplorer
'*
'*      *****************  Version 21  *****************
'*      User: Poynting     Date: 4/27/98    Time: 4:09p
'*      Updated in $/Clarify/Sources/Client/ObligationServerDLL
'*      CSR 17, CAA0147
'*      Corrected the permissions requested on registry keys to the minimum set
'*      required.
'*
'*      *****************  Version 20  *****************
'*      User: Amos         Date: 18/04/98   Time: 16:19
'*      Updated in $/NCAS MC/Sources/Debug Message Services/dcaMessageLogger
'*
'*      *****************  Version 19  *****************
'*      User: Amos         Date: 18/04/98   Time: 9:43
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*
'*      *****************  Version 18  *****************
'*      User: Amos         Date: 18/04/98   Time: 9:40
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*
'*      *****************  Version 16  *****************
'*      User: Poynting     Date: 3/20/98    Time: 4:27p
'*      Updated in $/Clarify/Sources/Client/ObligationServerDLL
'*      CAR007
'*      Corrected the access permissions used to read registry keys.
'*
'*      *****************  Version 15  *****************
'*      User: Amos         Date: 11/02/98   Time: 14:47
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*
'*      *****************  Version 14  *****************
'*      User: Amos         Date: 11/02/98   Time: 14:24
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*      Corrected GetNextId Routine to correct Stored Procedure Calling
'*      Mechanim
'*
'*      *****************  Version 13  *****************
'*      User: Amos         Date: 11/02/98   Time: 14:02
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*
'*      *****************  Version 12  *****************
'*      User: Cholerton    Date: 5-02-98    Time: 6:12p
'*      Updated in $/NCAS MC/Sources/Debug Message Services/dcaMessageLogger
'*      Disable/comment-out all debug code (and error handling code that causes
'*      the debug code to be called).  This is because this module is now
'*      called by the debug routines themselves.
'*
'*      *****************  Version 11  *****************
'*      User: Amos         Date: 5/02/98    Time: 17:35
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*
'*      *****************  Version 10  *****************
'*      User: Amos         Date: 5/02/98    Time: 12:01
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*
'*      *****************  Version 9  *****************
'*      User: Amos         Date: 5/02/98    Time: 10:18
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbManager
'*      Debug removed from critical section of AllocateConnection, Doevents
'*      removed from all code.  GetNextId stored procedures implemented. Round
'*      Robin connection allocation implemented.  Connection handle set to long
'*      in all code.
'*
'*      *****************  Version 8  *****************
'*      User: Szc2t2       Date: 12/01/98   Time: 13:54
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*      Changed ReportErr to ReRaiseErr
'*
'*      *****************  Version 7  *****************
'*      User: Szc2t2       Date: 8/01/98    Time: 14:50
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*      Changed order: logged errors before raising them
'*
'*      *****************  Version 6  *****************
'*      User: Szc2t2       Date: 7/01/98    Time: 12:02
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*      Added debug messages and error handling routines
'*
'*      *****************  Version 5  *****************
'*      User: Fzs56l       Date: 12/17/97   Time: 10:55a
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*      Added error processing to GetKeyValue
'*
'*      *****************  Version 3  *****************
'*      User: Fzs56l       Date: 12/16/97   Time: 10:35a
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'*      Added the functions to the module to allow for easy used.
'*
'*      *****************  Version 2  *****************
'*      User: Fzs56l       Date: 12/16/97   Time: 9:51a
'*      Updated in $/Technical Direction/Sources/Database Services/dcaDbCommon
'****************************************************************************

'----------------------------------------------------------------------------
' Compiler Directives
'----------------------------------------------------------------------------
Option Explicit


'----------------------------------------------------------------------------
' Constant Declarations
'----------------------------------------------------------------------------

' Unit Name


'----------------------------------------------------------------------------
' Public Data Declarations
'----------------------------------------------------------------------------

'Type Constants
Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

'Primary Hive Constants

Global Const HKEY_CURRENT_USER = &H80000001
'Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

'Permission Level for Key Access
Global Const KEY_QUERY_VALUE = &H1
Global Const KEY_SET_VALUE = &H2
Global Const KEY_CREATE_SUB_KEY = &H4
Global Const KEY_WRITE = (KEY_CREATE_SUB_KEY Or KEY_SET_VALUE)
'Global Const KEY_ALL_ACCESS = &H3F
Global Const KEY_ENUMERATE_SUB_KEYS = &H8
Global Const REG_OPTION_NON_VOLATILE = 0

'Registry API Return Constants
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259
Const KEY_READ = &H20019
'Custom Data Types
'Public Type FILETIME
'        dwLowDateTime As Long
'        dwHighDateTime As Long
'End Type


Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3

Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
'----------------------------------------------------------------------------
' Private Data Declarations
'----------------------------------------------------------------------------

' Object key for debug routines


'----------------------------------------------------------------------------
' Registry API Declarations
'----------------------------------------------------------------------------

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long
    
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long

Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As Long _
    , lpcbData As Long) As Long

Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String _
    , ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long
        
Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long

Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    lpcbName As Long, lpReserved As Long, ByVal lpClass As String, _
    lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long

Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
        ByVal cbName As Long) As Long

Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, lpReserved As Long, lpType As Long, _
    lpData As Byte, lpcbData As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long
'----------------------------------------------------------------------------
'Public methods
'----------------------------------------------------------------------------

'---------------------------------------------------------------------------
' Name : SetValueEx
'
' Description : sets the value of a registry entry
'
' Parameters :
'       hkey [in] - one of the primary hive constants e.g
'                       HKEY_LOCAL_MACHINE
'       sValueName [in] - full path name of the value to be set
'       lType      [in] - one of the type constants - type of
'                         data to be written
'       vvalue     [in] - new value
'
' Return Value : error code or 0 (if successful)
'
' Additional Notes:
'
'---------------------------------------------------------------------------

Public Function SetValueEx( _
            ByVal hKey As Long, _
            ByVal sValueName As String, _
            ByVal lType As Long, _
            ByVal vValue As Variant) _
            As Long

   
    
    On Error GoTo Handler

    Dim lValue As Long
    Dim sValue As String

    Select Case lType

        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, _
                                  lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
    
    Exit Function
    
Handler:
    
'    ' Report the error
'
'    LogErrHandler "Registry.SetValueEx"
'
'    ReRaiseErr
    
End Function


'---------------------------------------------------------------------------
' Name : CreateNewKey
'
' Description : Creates a new key in the registry
'
' Parameters :
'   sNewKeyName [in] - full path name for the new key
'   lPredefinedKey [in] - root hive macro (e.g HKEY_LOCAL_MACHINE)
'
' Additional Notes:
'
'---------------------------------------------------------------------------

Public Sub CreateNewKey( _
            ByVal sNewKeyName As String, _
            ByVal lPredefinedKey As Long)

    
    
    Dim hNewKey As Long
    Dim lretval As Long

    On Error GoTo Handler

    lretval = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
        vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        0&, hNewKey, lretval)
    RegCloseKey (hNewKey)
  
    Exit Sub
    
Handler:
    
'    ' Report the error
'    LogErrHandler "Registry.CreateNewKey"
'
'    ReRaiseErr
       
End Sub

       
'---------------------------------------------------------------------------
' Name : SetKeyValue
'
' Description : sets the value for a key
'
' Parameters :
'   hHivekey [in] - root hive key macro
'   sKeyname [in] - full path name of the key to be modified
'   sValueName [in] - name of the value to be set within the key
'   vValueSetting [in] - new value for the key
'   lValueType[in] - type of the new value
'
' Additional Notes:
'
'-------------------------------------------------------------------------

Public Sub SetKeyValue( _
                ByVal hHiveKey As Long, _
                ByVal sKeyName As String, _
                ByVal sValueName As String, _
                ByVal vValueSetting As Variant, _
                ByVal lValueType As Long)
      
    
    
    Dim lretval As Long
    Dim hKey As Long

    On Error GoTo Handler

    'lretval = RegOpenKeyEx(hHiveKey, sKeyName, 0, KEY_WRITE, hKey)
    lretval = RegOpenKeyEx(hHiveKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lretval = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
    
    Exit Sub
    
Handler:

'    ' Report the error
'    LogErrHandler "Registry.SetKeyValue"
'
'    ReRaiseErr
    
End Sub


'---------------------------------------------------------------------------
' Name : GetKeyValue
'
' Description : Retreive the value for a key
'
' Parameters :
'   hHiveKey [in] - root hive macro
'   sKey     [in] - full path name of the key
'   sValue   [in] - name of the value to be retrieved
'
' Return Value :
'   Contents of the specified key value
'--------------------------------------------------------------------------

Function GetKeyValue( _
            ByVal hHiveKey As Long, _
            ByVal sKey As String, _
            ByVal sValue As String) _
            As String
    
    
    Dim lretval As Long
    Dim vValue As Variant
    
    Dim hKey As Long

    On Error GoTo Handler:

    lretval = RegOpenKeyEx(hHiveKey, sKey, 0, KEY_QUERY_VALUE, hKey)
    lretval = QueryValueEx(hKey, sValue, vValue)
    If lretval <> 0 Then
        'key is not registered
        GetKeyValue = Empty
        Exit Function
    End If
    
    RegCloseKey (hKey)

    If vValue = vbNullString Then
        GetKeyValue = vbNullString
        Exit Function
    End If
    
    vValue = Left$(vValue, InStr(vValue, vbNullChar) - 1)

    GetKeyValue = CStr(vValue)
    
    Exit Function
    
Handler:
    
'    ' Report the error
'    LogErrHandler "Registry.GetKeyValue"
'
'    ReRaiseErr
       
End Function


'---------------------------------------------------------------------------
' Name : QueryValueEx
'
' Description :
'
' Parameters :
'   lhkey [in] - root hive macro
'   szValueName [in] - name of value to be retrieved
'   vvalue [out] - value retreived
'
' Return Value : 0 if successful
'
' Additional Notes :
'---------------------------------------------------------------------------

Public Function QueryValueEx( _
            ByVal lhKey As Long, _
            ByVal szValueName As String, _
            ByRef vValue As Variant) _
            As Long

   

    Dim cch As Long
    Dim lRC As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError
    lRC = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lRC <> ERROR_NONE Then Error 5

    Select Case lType
        Case REG_SZ:
            sValue = String(cch, 0)
            lRC = RegQueryValueExString(lhKey, szValueName, _
                                        0&, lType, sValue, cch)
            If lRC = ERROR_NONE Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        Case REG_DWORD:
            lRC = RegQueryValueExLong(lhKey, szValueName, _
                                      0&, lType, lValue, cch)
            If lRC = ERROR_NONE Then vValue = lValue
        Case Else
            lRC = -1
    End Select
    
QueryValueExExit:

    QueryValueEx = lRC
    Exit Function

QueryValueExError:

    Resume QueryValueExExit
    

End Function


'---------------------------------------------------------------------------
' Name : DeleteKey
'
' Description : deletes a key and all its values & subkeys
'
' Parameters :
'   hStartKey [in] - root key macro
'   skeyName  [in] - full key name
'
' Return Value : 0 if successful
'---------------------------------------------------------------------------

Public Function DeleteKey( _
                ByVal hStartKey As Long, _
                ByVal sKeyName As String) _
                As Long

    

    Dim lwRtn, lwSubKeyLength As Long
    
    Dim szSubkey As String * 256
    Dim hKey As Long
    

    On Error GoTo Handler

    lwRtn = RegOpenKeyEx(hStartKey, sKeyName, 0, (KEY_WRITE Or KEY_ENUMERATE_SUB_KEYS), hKey)
     
    If lwRtn = ERROR_NONE Then

        Do While (lwRtn = ERROR_NONE)
            lwSubKeyLength = 256
            lwRtn = RegEnumKey(hKey, 0, szSubkey, lwSubKeyLength)
            If lwRtn = ERROR_NO_MORE_ITEMS Then
                lwRtn = RegDeleteKey(hStartKey, sKeyName)
                Exit Function
            ElseIf lwRtn = ERROR_NONE Then
                lwRtn = DeleteKey(hKey, szSubkey)
            End If
        Loop
        RegCloseKey (hKey)
    Else
        lwRtn = ERROR_BADKEY
    End If

    DeleteKey = lwRtn
    
    Exit Function
    
Handler:
    
'    ' Report the error
'    LogErrHandler "Registry.DeleteKey"
'
'    ReRaiseErr
    
End Function


'--------------------------------------------------------------------------
' Name : GetSubKeys
'
' Description : Retrieves all the subkeys for a key
'
' Parameters :
'       hHiveKey [in] - root hive macro
'       sKeyName [in] - name of key
'       cSubKeys [out] - collection of subkeys
'--------------------------------------------------------------------------

Public Sub GetSubKeys( _
            ByVal hHiveKey As Long, _
            ByVal sKeyName As String, _
            ByRef cSubKeys As Collection)
    
   
    
    Dim lretval As Long
    Dim hKey As Long
    Dim sSubKey As String * 256 'Fixed length string holding return value
    Dim sSubKeyTrim As String 'Clean string version
    Dim lKeylength As Long 'Length of return value
    Dim lIndex As Long 'Current index position

    On Error GoTo Handler

    'Open Parent key for enumeration
    lretval = RegOpenKeyEx(hHiveKey, sKeyName, _
                           0, KEY_ENUMERATE_SUB_KEYS, hKey)
    'if key does not exist then return an empty collection
    If lretval <> ERROR_NONE Then
        RegCloseKey (hHiveKey)
        RegCloseKey (hKey)
        Exit Sub
    End If
    
    lIndex = 0
    lKeylength = 256

    'Enumerate Subkeys and add keys to collection
    Do While RegEnumKey(hKey, lIndex, sSubKey, lKeylength) <> ERROR_NO_MORE_ITEMS
        sSubKeyTrim = Left$(sSubKey, InStr(sSubKey, vbNullChar) - 1)
        cSubKeys.Add sSubKeyTrim
        lIndex = lIndex + 1
    Loop

    RegCloseKey (hHiveKey)
    RegCloseKey (hKey)
    
    Exit Sub

Handler:
    
'    ' Report the error
'    LogErrHandler "Registry.GetSubKeys"
'
'    ReRaiseErr
    
End Sub

Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim retval As Long
    Dim valuetype As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
        Exit Function
    End If
    
    ' prepare a 1K receiving resBinary
    length = 1024
    ReDim resBinary(0 To length - 1) As Byte
    
    ' read the registry key
    retval = RegQueryValueEx(handle, ValueName, 0, valuetype, resBinary(0), _
        length)
    ' if resBinary was too small, try again
    If retval = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To length - 1) As Byte
        retval = RegQueryValueEx(handle, ValueName, 0, valuetype, resBinary(0), _
            length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valuetype
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(length - 1)
            CopyMemory ByVal resString, resBinary(0), length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey handle
            Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' close the registry key
    RegCloseKey handle
End Function






