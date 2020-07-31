Attribute VB_Name = "modRegistry"
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - modRegistry
' *********************************************************************

'-------------------------------------------------------------------------------
' This file contains definitions of constants and functions for Registry access
'
'-------------------------------------------------------------------------------
Option Explicit
'-------------------------------------------------------------------------------
' Constants as defined in WINNT.H file
'-------------------------------------------------------------------------------
' Registry Data types
'-------------------------------------------------------------------------------
Public Const REG_NONE As Long = 0                       ' No value type
Public Const REG_SZ As Long = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ As Long = 2                  ' Unicode nul terminated string
                                                        ' (with environment variable references)
Public Const REG_BINARY As Long = 3                     ' Free form binary
Public Const REG_DWORD As Long = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN As Long = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN As Long = 5           ' 32-bit number
Public Const REG_LINK As Long = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ As Long = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST As Long = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10
'-------------------------------------------------------------------------------
' Main Key vaues
'-------------------------------------------------------------------------------
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
'-------------------------------------------------------------------------------
' Return codes from Registry functions.
'-------------------------------------------------------------------------------
Public Const ERROR_SUCCESS As Long = 0
Public Const ERROR_BADDB As Long = 1009
Public Const ERROR_BADKEY As Long = 1010
Public Const ERROR_CANTOPEN As Long = 1011
Public Const ERROR_CANTREAD As Long = 1012
Public Const ERROR_CANTWRITE As Long = 1013
Public Const ERROR_OUTOFMEMORY As Long = 14
Public Const ERROR_INVALID_PARAMETER As Long = 87
Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_NO_MORE_ITEMS As Long = 259
Public Const ERROR_MORE_DATA As Long = 234
'-------------------------------------------------------------------------------
' Read/Write permissions:
'-------------------------------------------------------------------------------
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_CREATE_LINK As Long = &H20
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const WRITE_OWNER = &H80000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Public Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Public Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Public Const KEY_EXECUTE = KEY_READ
'-------------------------------------------------------------------------------
' Open/Create Options
'-------------------------------------------------------------------------------
Public Const REG_OPTION_RESERVED As Long = &H0       ' Parameter is reserved
Public Const REG_OPTION_NON_VOLATILE As Long = &H0   ' Key is preserved
                                                     ' when system is rebooted
Public Const REG_OPTION_VOLATILE As Long = &H1       ' Key is not preserved
                                                     ' when system is rebooted
Public Const REG_OPTION_CREATE_LINK As Long = &H2    ' Created key is a
                                                     ' symbolic link
Public Const REG_OPTION_BACKUP_RESTORE As Long = &H4 ' open for backup or restore
                                                     ' special access rules
                                                     ' privilege required
Public Const REG_OPTION_OPEN_LINK  As Long = &H8     ' Open symbolic link
'-------------------------------------------------------------------------------
' Key creation/open disposition
'-------------------------------------------------------------------------------
Public Const REG_CREATED_NEW_KEY As Long = &H1       ' New Registry Key created
Public Const REG_OPENED_EXISTING_KEY As Long = &H2   ' Existing Key opened
'-------------------------------------------------------------------------------
' Windows version constants
'-------------------------------------------------------------------------------
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
'-------------------------------------------------------------------------------
' FILETIME type is needed for RegEnumKey and
'   RegQueryInfoKey
'-------------------------------------------------------------------------------
Type FILETIME
    lLowDateTime    As Long
    lHighDateTime   As Long
End Type
'-------------------------------------------------------------------------------
' OsVersionInfo type is needed for GetVersionEx
'-------------------------------------------------------------------------------
Type OSVERSIONINFO
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatform As Long
    szCSDVersion As String * 128
End Type
'-------------------------------------------------------------------------------
' Note: we need to check the Windows Version because in 95/98 the "RegDeleteKey"
' function will delete the key & all subkeys, while in NT/2000 we need to call
' the alternative function "SHDeleteKey" (may not be available in 95/98)
' (if Internet Explorer 4.0 or later is installed - it will work for both!)
'-------------------------------------------------------------------------------
'
' Function declarations
'
Public Declare Function RegConnectRegistry Lib "advapi32.dll" _
      (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
'-------------------------------------------------------------------------------
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
      (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
      (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
      (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
      (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function SHDeleteKey Lib "Shlwapi.dll" Alias "SHDeleteKeyA" _
      (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function SHDeleteEmptyKey Lib "Shlwapi.dll" Alias "SHDeleteEmptyKeyA" _
      (ByVal hKey As Long, ByVal lpSubKey As String) As Long
'-------------------------------------------------------------------------------
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
      (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) _
      As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
      (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
       lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, _
       lpcbData As Long) As Long
Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
      (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, _
       ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
       lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, _
       lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
       lpftLastWriteTime As FILETIME) As Long
'-------------------------------------------------------------------------------
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
      (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
       ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
      (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
       lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
      (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
       ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
      (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
       ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
       lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
      (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, _
       ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
       lpftLastWriteTime As FILETIME) As Long
'-------------------------------------------------------------------------------
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" _
    (lpStruct As OSVERSIONINFO) As Long


'-------------------------------------------------------------------------------
'   CheckValExists - Checks if the specified value exists at the given path
'
'   hKey      - Specifies the tree that contains the strPath path.
'   szPath    - Contains the key with the value to be deleted.
'   szValueName - The value to remove
'
'   Returns True if value exists, False if not
'-------------------------------------------------------------------------------
Public Function CheckValExists(ByVal hKey As Long, ByVal szPath As String, _
                               ByVal szValueName As String) As Boolean
Dim lRegResult As Long
Dim lValueType As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long
    CheckValExists = False
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    If (lRegResult = ERROR_SUCCESS) Then _
        lRegResult = RegQueryValueEx(hCurKey, szValueName, 0&, lValueType, _
                                     ByVal 0&, lDataBufferSize)
    If (lRegResult = ERROR_SUCCESS) Then CheckValExists = True
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   CopyRegistryByte - Copies a Byte array value to a specified new value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szOldValName - The source value name
'   szNewValName - The destination value name
'   bCopyAnyway  - Copy even if the destination value already exists
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function CopyRegistryByte(ByVal hKey As Long, ByVal szPath As String, _
                                 ByVal szOldValName As String, ByVal szNewValName As String, _
                                 ByVal bCopyAnyway As Boolean) As Boolean
Dim lRegResult As Long
Dim hCurKey As Long
Dim lDisp As Long
Dim byData() As Byte
Dim lValueType As Long
Dim lDataBufferSize As Long
'
    CopyRegistryByte = False
    If (Not bCopyAnyway) Then       ' Check if the new value exists
        If (CheckValExists(hKey, szPath, szNewValName)) Then Exit Function
    End If
' Open the key and get number of bytes
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, szOldValName, 0&, lValueType, ByVal 0&, lDataBufferSize)
' Get the previous value
    If (lRegResult = ERROR_SUCCESS) Then
        If (lValueType = REG_BINARY) Then
' Initialise buffers and retrieve value
            If (lDataBufferSize > 0) Then
                ReDim byData(lDataBufferSize - 1) As Byte
                lRegResult = RegQueryValueEx(hCurKey, szOldValName, 0&, lValueType, byData(0), lDataBufferSize)
            Else
                ReDim byData(0) As Byte
            End If
            CopyRegistryByte = True ' Indicate sucess in the first part
        End If
    End If
' Write the copy
    If (CopyRegistryByte = True) Then
        CopyRegistryByte = False    ' Invalidate again
        lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                    KEY_WRITE, ByVal 0, hCurKey, lDisp)
' Pass the first array element and length of array
        If (lRegResult = ERROR_SUCCESS) Then _
           lRegResult = RegSetValueEx(hCurKey, szNewValName, 0&, REG_BINARY, byData(0), lDataBufferSize)
        If (lRegResult = ERROR_SUCCESS) Then
            CopyRegistryByte = True ' Indicate complete success
        End If
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
'
End Function

'-------------------------------------------------------------------------------
'   CopyRegistryLong - Copies a Long value to a specified new value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szOldValName - The source value name
'   szNewValName - The destination value name
'   bCopyAnyway  - Copy even if the destination value already exists
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function CopyRegistryLong(ByVal hKey As Long, ByVal szPath As String, _
                                 ByVal szOldValName As String, ByVal szNewValName As String, _
                                 ByVal bCopyAnyway As Boolean) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lDisp As Long
Dim lData As Long
Dim lValueType As Long
Dim lDataBufferSize As Long
'
    CopyRegistryLong = False
    If (Not bCopyAnyway) Then       ' Check if the new value exists
        If (CheckValExists(hKey, szPath, szNewValName)) Then Exit Function
    End If
' Get the previous value
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lDataBufferSize = 4       ' 4 bytes = 32 bits = long
    lRegResult = RegQueryValueEx(hCurKey, szOldValName, 0&, lValueType, lData, lDataBufferSize)
    If (lRegResult = ERROR_SUCCESS) And (lValueType = REG_DWORD) Then
        CopyRegistryLong = True     ' Indicate sucess in the first part
    End If
' Write the copy
    If (CopyRegistryLong = True) Then
        CopyRegistryLong = False    ' Invalidate again
        lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                    KEY_WRITE, ByVal 0, hCurKey, lDisp)
        If (lRegResult = ERROR_SUCCESS) Then _
            lRegResult = RegSetValueEx(hCurKey, szNewValName, 0&, REG_DWORD, lData, 4)
        If (lRegResult = ERROR_SUCCESS) Then
            CopyRegistryLong = True ' Indicate complete success
        End If
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   CopyRegistryString - Copies a String  value to a specified new value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szOldValName - The source value name
'   szNewValName - The destination value name
'   bCopyAnyway  - Copy even if the destination value already exists
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function CopyRegistryString(ByVal hKey As Long, ByVal szPath As String, _
                            ByVal szOldValName As String, ByVal szNewValName As String, _
                            ByVal bCopyAnyway As Boolean) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lDisp As Long
Dim szData As String
Dim lValueType As Long
Dim lDataBufferSize As Long
'
    CopyRegistryString = False
    If (Not bCopyAnyway) Then       ' Check if the new value exists
        If (CheckValExists(hKey, szPath, szNewValName)) Then Exit Function
    End If
' Open the key and get length of string
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, szOldValName, ByVal 0&, lValueType, ByVal 0&, lDataBufferSize)
' Get the previous value
    If (lRegResult = ERROR_SUCCESS) Then
        If lValueType = REG_SZ Then
' Initialise string buffer and retrieve string
            szData = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, szOldValName, 0&, 0&, ByVal szData, lDataBufferSize)
            CopyRegistryString = True   ' Indicate sucess in the first part
        End If
    End If
' Write the copy
    If (CopyRegistryString = True) Then
        CopyRegistryString = False      ' Invalidate again
        lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                    KEY_WRITE, ByVal 0, hCurKey, lDisp)
        If (lRegResult = ERROR_SUCCESS) Then _
            lRegResult = RegSetValueEx(hCurKey, szNewValName, 0, REG_SZ, ByVal szData, Len(szData))
        If (lRegResult = ERROR_SUCCESS) Then
            CopyRegistryString = True   ' Indicate complete success
        End If
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   CreateKey - Creates a new Registry Key
'
'   hKey      - Specifies the tree that contains the strPath path.
'   szPath    - Contains the key to be created.
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function CreateKey(hKey As Long, szPath As String) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lDisp As Long
'
    CreateKey = False
    lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                KEY_WRITE, ByVal 0, hCurKey, lDisp)
    If (lRegResult = ERROR_SUCCESS) Then
        CreateKey = True
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
'
End Function

'-------------------------------------------------------------------------------
'   DeleteKey - Deletes a registry key
'
'   hKey      - Specifies the tree that contains the strPath path.
'   szPath    - Contains the key to be deleted.
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function DeleteKey(ByVal hKey As Long, ByVal szPath As String) As Boolean
Dim lRegResult As Long
'
    DeleteKey = False
'    lRegResult = RegDeleteKey(hKey, szPath)
    lRegResult = SHDeleteKey(hKey, szPath)
    If (lRegResult = ERROR_SUCCESS) Then
        DeleteKey = True
    End If
'
End Function

'-------------------------------------------------------------------------------
'   DeleteValue - Removes a named value from the specified registry key.
'
'   hKey      - Specifies the tree that contains the strPath path.
'   szPath    - Contains the key with the value to be deleted.
'   szValueName - The value to remove
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function DeleteValue(ByVal hKey As Long, _
                            ByVal szPath As String, _
                            ByVal szValueName As String) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
'
    DeleteValue = False
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_WRITE, hCurKey)
    If (lRegResult = ERROR_SUCCESS) Then
        lRegResult = RegDeleteValue(hCurKey, szValueName)
        If (lRegResult = ERROR_SUCCESS) Then
            DeleteValue = True
        End If
        lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
    End If
'
End Function

'-------------------------------------------------------------------------------
'   GetAllKeys - Gets all the keys in the specified path
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all keys
'
'   Returns: an array of strings in a variant
'-------------------------------------------------------------------------------
Public Function GetAllKeys(ByVal hKey As Long, ByVal szPath As String) As Variant
'
Dim lRegResult As Long
Dim lCounter As Long
Dim hCurKey As Long
Dim szBuffer As String
Dim lDataBufferSize As Long
Dim szNames() As String
Dim intZeroPos As Integer
'
    lCounter = 0
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
'
    Do
' Initialise buffers (longest possible length=255)
        lDataBufferSize = 255
        szBuffer = String(lDataBufferSize, " ")
        lRegResult = RegEnumKey(hCurKey, lCounter, szBuffer, lDataBufferSize)
        If (lRegResult = ERROR_SUCCESS) Then
' Tidy up string and save it
            ReDim Preserve szNames(lCounter) As String
            intZeroPos = InStr(szBuffer, Chr$(0))
            If intZeroPos > 0 Then
                szNames(UBound(szNames)) = left$(szBuffer, intZeroPos - 1)
            Else
                szNames(UBound(szNames)) = szBuffer
            End If
        lCounter = lCounter + 1
        Else
            Exit Do
        End If
    Loop
    If (lCounter > 0) Then
        GetAllKeys = szNames
    End If
'
End Function

'-------------------------------------------------------------------------------
' GetAllValues - Enumerates the values for the specified open registry key
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'
'   Returns: a 2D array. - (x,0) is value name
'                        - (x,1) is value type (see constants)
'-------------------------------------------------------------------------------
Public Function GetAllValues(ByVal hKey As Long, ByVal szPath As String) As Variant
Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim szValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim szNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer
Dim Finisheddata() As Variant
'
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
'
    Do
' Initialise bufffers
        lValueNameSize = 255
        szValueName = String$(lValueNameSize, " ")
        lDataBufferSize = 4000
'
        lRegResult = RegEnumValue(hCurKey, lCounter, szValueName, lValueNameSize, ByVal 0&, lValueType, byDataBuffer(0), lDataBufferSize)
        If (lRegResult = ERROR_SUCCESS) Then
' Save the type
            ReDim Preserve szNames(lCounter) As String
            ReDim Preserve lTypes(lCounter) As Long
            lTypes(UBound(lTypes)) = lValueType
'Tidy up string and save it
            intZeroPos = InStr(szValueName, Chr$(0))
            If intZeroPos > 0 Then
                szNames(UBound(szNames)) = left$(szValueName, intZeroPos - 1)
            Else
                szNames(UBound(szNames)) = szValueName
            End If
            lCounter = lCounter + 1
        Else
            Exit Do
        End If
    Loop
'Move data into array
    If (lCounter > 0) Then
        ReDim Finisheddata(UBound(szNames), 0 To 1) As Variant
'
        For lCounter = 0 To UBound(szNames)
            Finisheddata(lCounter, 0) = szNames(lCounter)
            Finisheddata(lCounter, 1) = lTypes(lCounter)
        Next
        GetAllValues = Finisheddata
    End If
End Function

'-------------------------------------------------------------------------------
'   GetMainKey - Used to convert main key strings to their values
'
'   szKeyName - The string description of the main key
'
'   Returns a long value of the Main Key
'-------------------------------------------------------------------------------
Function GetMainKey(ByVal szKeyName As String) As Long
    Select Case szKeyName
        Case "HKEY_CLASSES_ROOT"
            GetMainKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetMainKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetMainKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetMainKey = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA"
            GetMainKey = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG"
            GetMainKey = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            GetMainKey = HKEY_DYN_DATA
        Case Else
            GetMainKey = 0
    End Select
End Function

'-------------------------------------------------------------------------------
'   GetRegistryByte - Retrieves a Byte type value from the registry
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   Default - Optional default value to be returned if error
'
'   Returns a Variant containing the buffer of Byte values read
'-------------------------------------------------------------------------------
Public Function GetRegistryByte(ByVal hKey As Long, ByVal szPath As String, _
                                ByVal szValueName As String, _
                                Optional ByVal Default As Variant) As Variant
Dim lValueType As Long
Dim byBuffer() As Byte
Dim lDataBufferSize As Long
Dim lRegResult As Long
Dim hCurKey As Long
' Set the default value
    If Not IsEmpty(Default) Then
        If VarType(Default) = vbArray + vbByte Then
            GetRegistryByte = Default
        Else
            GetRegistryByte = 0
        End If
    Else
        GetRegistryByte = 0
    End If
' Open the key and get number of bytes
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, szValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If (lRegResult = ERROR_SUCCESS) Then
        If (lValueType = REG_BINARY) And (lDataBufferSize > 0) Then
' Initialise buffers and retrieve value
            ReDim byBuffer(lDataBufferSize - 1) As Byte
            lRegResult = RegQueryValueEx(hCurKey, szValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
            GetRegistryByte = byBuffer
        Else        ' The value is not BYTE
        End If
    Else            ' Error ???
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   GetRegistryLong - Retrieves a Long type value from the registry
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   Default - Optional default value to be returned if error
'
'   Returns a Long containing the value read
'-------------------------------------------------------------------------------
Public Function GetRegistryLong(ByVal hKey As Long, ByVal szPath As String, _
                                ByVal szValueName As String, _
                                Optional ByVal Default As Long) As Long
Dim lRegResult As Long
Dim lValueType As Long
Dim lBuffer As Long
Dim lDataBufferSize As Long
Dim hCurKey As Long
' Set the default value
    If Not IsEmpty(Default) Then
        GetRegistryLong = Default
    Else
        GetRegistryLong = 0
    End If
'
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lDataBufferSize = 4       ' 4 bytes = 32 bits = long
    lRegResult = RegQueryValueEx(hCurKey, szValueName, 0&, lValueType, lBuffer, lDataBufferSize)
    If (lRegResult = ERROR_SUCCESS) Then
        If lValueType = REG_DWORD Then
            GetRegistryLong = lBuffer
        Else        ' The Value is not DWORD
        End If
    Else            ' Error ???
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   GetRegistryString - Retrieves a String type value from the registry
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   Default - Optional default value to be returned if error
'
'   Returns a String containing the value read
'-------------------------------------------------------------------------------
Public Function GetRegistryString(ByVal hKey As Long, ByVal szPath As String, _
                                  ByVal szValueName As String, _
                                  Optional ByVal Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim szBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long

On Error Resume Next

' Set the default value
    If Not IsEmpty(Default) Then
        GetRegistryString = Default
    Else
        GetRegistryString = ""
    End If
' Open the key and get length of string
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, szValueName, ByVal 0&, lValueType, ByVal 0&, lDataBufferSize)
    If (lRegResult = ERROR_SUCCESS) Then
        If lValueType = REG_SZ Then
' Initialise string buffer and retrieve string
            szBuffer = String(lDataBufferSize, " ")
            lRegResult = RegQueryValueEx(hCurKey, szValueName, 0&, 0&, ByVal szBuffer, lDataBufferSize)
' Format the string
            intZeroPos = InStr(szBuffer, Chr$(0))
            If intZeroPos > 0 Then
                GetRegistryString = left$(szBuffer, intZeroPos - 1)
            Else
                GetRegistryString = szBuffer
            End If
        Else    ' The Value is not String
        End If
    Else        ' Error ???
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   GetTypeDescr - Used to convert the type values to their description strings
'
'   nType - The type value
'
'   Returns a String description
'-------------------------------------------------------------------------------
Function GetTypeDescr(ByVal nType As Long) As String
    Select Case nType
    Case REG_NONE
        GetTypeDescr = "REG_NONE"
    Case REG_SZ
        GetTypeDescr = "REG_SZ"
    Case REG_EXPAND_SZ
        GetTypeDescr = "REG_EXPAND_SZ"
    Case REG_BINARY
        GetTypeDescr = "REG_BINARY"
    Case REG_DWORD
        GetTypeDescr = "REG_DWORD"
    Case REG_DWORD_LITTLE_ENDIAN
        GetTypeDescr = "REG_DWORD_LITTLE_ENDIAN"
    Case REG_DWORD_BIG_ENDIAN
        GetTypeDescr = "REG_DWORD_BIG_ENDIAN"
    Case REG_LINK
        GetTypeDescr = "REG_LINK"
    Case REG_MULTI_SZ
        GetTypeDescr = "REG_MULTI_SZ"
    Case REG_RESOURCE_LIST
        GetTypeDescr = "REG_RESOURCE_LIST"
    Case REG_FULL_RESOURCE_DESCRIPTOR
        GetTypeDescr = "REG_FULL_RESOURCE_DESCRIPTOR"
    Case REG_RESOURCE_REQUIREMENTS_LIST
        GetTypeDescr = "REG_RESOURCE_REQUIREMENTS_LIST"
    Case Else
        GetTypeDescr = "Unknown"
    End Select
End Function

'-------------------------------------------------------------------------------
'   GetTypeDescr - Used to convert the type description strings to their values
'
'   szType - The type description string
'
'   Returns a Long value
'-------------------------------------------------------------------------------
Function GetTypeValue(ByVal szType As String) As Long
    Select Case szType
    Case "REG_NONE"
        GetTypeValue = REG_NONE
    Case "REG_SZ"
        GetTypeValue = REG_SZ
    Case "REG_EXPAND_SZ"
        GetTypeValue = REG_EXPAND_SZ
    Case "REG_BINARY"
        GetTypeValue = REG_BINARY
    Case "REG_DWORD"
        GetTypeValue = REG_DWORD
    Case "REG_DWORD_LITTLE_ENDIAN"
        GetTypeValue = REG_DWORD_LITTLE_ENDIAN
    Case "REG_DWORD_BIG_ENDIAN"
        GetTypeValue = REG_DWORD_BIG_ENDIAN
    Case "REG_LINK"
        GetTypeValue = REG_LINK
    Case "REG_MULTI_SZ"
        GetTypeValue = REG_MULTI_SZ
    Case "REG_RESOURCE_LIST"
        GetTypeValue = REG_RESOURCE_LIST
    Case "REG_FULL_RESOURCE_DESCRIPTOR"
        GetTypeValue = REG_FULL_RESOURCE_DESCRIPTOR
    Case "REG_RESOURCE_REQUIREMENTS_LIST"
        GetTypeValue = REG_RESOURCE_REQUIREMENTS_LIST
    Case Else
        GetTypeValue = -1
    End Select
End Function

'-------------------------------------------------------------------------------
'   GetWinVersion - Retrieves the Windows version information
'
'   Returns the version ID (see constants) if success, -1 if failure
'-------------------------------------------------------------------------------
Public Function GetWinVersion() As Long
Dim nStat As Long
Dim OsVers As OSVERSIONINFO
    GetWinVersion = -1
    OsVers.dwVersionInfoSize = 148&
    nStat = GetVersionEx(OsVers)
    If (nStat > 0) Then GetWinVersion = OsVers.dwPlatform
End Function

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Public Function RenameRegistryKey(ByVal hKey As Long, ByVal szPath As String, _
                                  ByVal szOldKey As String, ByVal szNewKey As String) _
                                  As Boolean
Dim hCurKey As Long
Dim lRegResult As Long

Dim nSubK As Long
Dim nSubV As Long
Dim nDummy(6) As Long
Dim szNewPath As String
Dim tFileT As FILETIME
'
    RenameRegistryKey = False
' Check if the key is empty
    nDummy(6) = 256
    szNewPath = Space(nDummy(6))
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    If (lRegResult = ERROR_SUCCESS) Then _
        lRegResult = RegQueryInfoKey(hCurKey, szNewPath, nDummy(6), nDummy(0), _
                                     nSubK, nDummy(1), nDummy(2), nSubV, _
                                     nDummy(3), nDummy(4), nDummy(5), tFileT)
    RegCloseKey (hCurKey)  ' Never forget to close the Key
    If (lRegResult <> ERROR_SUCCESS) Then Exit Function
    If ((nSubK > 0) Or (nSubV)) Then Exit Function
'
    szNewPath = Replace(szPath, "\" & szOldKey, "\" & szNewKey, , , vbTextCompare)
    lRegResult = SHDeleteEmptyKey(hKey, szPath)
    If (lRegResult = ERROR_SUCCESS) Then
        RenameRegistryKey = CreateKey(hKey, szNewPath)
    End If
End Function

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Public Function RenameRegistryValue(ByVal hKey As Long, ByVal szPath As String, _
                                    ByVal szOldVal As String, ByVal szNewVal As String) _
                                    As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lValueType As Long
Dim lDataBufferSize As Long
Dim bStat As Boolean
' Get the valur information
    RenameRegistryValue = False
    lRegResult = RegOpenKeyEx(hKey, szPath, ByVal 0, KEY_READ, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, szOldVal, 0&, lValueType, ByVal 0&, lDataBufferSize)
    Call RegCloseKey(hCurKey)   ' Never forget to close the Key
' Start Renaming (actually read the old value and create new, delete the old)
    If (lRegResult = ERROR_SUCCESS) Then
        Select Case lValueType
        Case REG_NONE
            bStat = CopyRegistryByte(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_SZ
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_EXPAND_SZ
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_BINARY
            bStat = CopyRegistryByte(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_DWORD
            bStat = CopyRegistryLong(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_DWORD_LITTLE_ENDIAN
            bStat = CopyRegistryLong(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_DWORD_BIG_ENDIAN
            bStat = CopyRegistryLong(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_LINK
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_MULTI_SZ
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_RESOURCE_LIST
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_FULL_RESOURCE_DESCRIPTOR
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        Case REG_RESOURCE_REQUIREMENTS_LIST
            bStat = CopyRegistryString(hKey, szPath, szOldVal, szNewVal, False)
        End Select
        If bStat Then Call DeleteValue(hKey, szPath, szOldVal)
    Else
        bStat = False
    End If
    RenameRegistryValue = bStat
End Function

'-------------------------------------------------------------------------------
'   SaveRegistryByte - Saves a Byte array to a specified value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   byData    - The byte array to save (Make sure that the array starts with
'               element 0 otherwise it will not be saved!)
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function SaveRegistryByte(ByVal hKey As Long, ByVal szPath As String, _
                            ByVal szValueName As String, byData() As Byte) As Boolean
Dim lRegResult As Long
Dim hCurKey As Long
Dim lDisp As Long
'
    SaveRegistryByte = False
    lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                KEY_WRITE, ByVal 0, hCurKey, lDisp)
' Pass the first array element and length of array
    If (lRegResult = ERROR_SUCCESS) Then _
       lRegResult = RegSetValueEx(hCurKey, szValueName, 0&, REG_BINARY, byData(0), UBound(byData()) + 1)
    If (lRegResult = ERROR_SUCCESS) Then
        SaveRegistryByte = True
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
'
End Function

'-------------------------------------------------------------------------------
'   SaveRegistryLong - Saves a Long to a specified value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   lData     - The Long value to be saved
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function SaveRegistryLong(ByVal hKey As Long, ByVal szPath As String, _
                                 ByVal szValueName As String, ByVal lData As Long) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lDisp As Long
'
    SaveRegistryLong = False
    lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                KEY_WRITE, ByVal 0, hCurKey, lDisp)
    If (lRegResult = ERROR_SUCCESS) Then _
        lRegResult = RegSetValueEx(hCurKey, szValueName, 0&, REG_DWORD, lData, 4)
    If (lRegResult = ERROR_SUCCESS) Then
        SaveRegistryLong = True
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function

'-------------------------------------------------------------------------------
'   SaveRegistryString - Saves a String to a specified value name
'
'   hKey      - Specifies the tree that contains the szPath path.
'   szPath    - The path to scan for all values
'   szValueName - The value name
'   szData    - The String to be saved
'
'   Returns True if success, False if failure
'-------------------------------------------------------------------------------
Public Function SaveRegistryString(ByVal hKey As Long, ByVal szPath As String, _
                                   ByVal szValueName As String, szData As String) As Boolean
Dim hCurKey As Long
Dim lRegResult As Long
Dim lDisp As Long
'
    SaveRegistryString = False
    lRegResult = RegCreateKeyEx(hKey, szPath, ByVal 0, ByVal 0, REG_OPTION_NON_VOLATILE, _
                                KEY_WRITE, ByVal 0, hCurKey, lDisp)
    If (lRegResult = ERROR_SUCCESS) Then _
        lRegResult = RegSetValueEx(hCurKey, szValueName, 0, REG_SZ, ByVal szData, Len(szData))
    If (lRegResult = ERROR_SUCCESS) Then
        SaveRegistryString = True
    End If
    lRegResult = RegCloseKey(hCurKey)   ' Never forget to close the Key
End Function



