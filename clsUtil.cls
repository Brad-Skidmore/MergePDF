VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - clsUtil
' *********************************************************************

Option Explicit

Private Const REP_CHAR160 = "ñÑ¶"
Private Const BAD_SEED As Long = -9999
Private Const MIN_SEED As Long = 300000
Private Const MAX_SEED As Long = 999999
Private Const SECURITY_SEED As Long = 934618
Private Const SZ_SECURITY_SEED As String = " x961118 7g t46111872i f36111801j c56111823w x761118 5g b86111876t e07111828h w66111844e"
'==============================================================================================================
'WIN API
'==============================================================================================================
'Window OS System Info

'Used By Private Declare Function GetVersionEx
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

'used by utGetTaskbarHeight and
Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const WM_CLOSE = &H10 'BGS used for Close window

Private Const NO_DB_VERSION_TABLE_ERROR As Integer = 3078
Private Const DB_VERSION_DUPLICATE_VALUE_ERROR As Integer = 3022
Private Const DB_VERSION_ITEM_NOT_FOUND_ERROR As Integer = 3265
Private Const SW_RESTORE = 9
Private Const SPI_GETWORKAREA = 48
Private Const SE_ERR_NOASSOC = 31


'WIN API Methods
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'WIN API Routines
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'=======================================================================================================================
'END WIN API
'=======================================================================================================================

'Member Vars used for Passing Global Vars
Private msAPPEXEName As String                  'Application Name
Private msMainAppEXEName As String              'Main Application name
Private msCommandString As String               'Command String passed to Application
Private msInstallDir As String                  'Installed directory
Private mlSECURITY_SEED As Long
Private msSECURITY_SEED As String
'Flags
Private mbUCText As Boolean

'Objects
Private mFSO As Scripting.FileSystemObject

'Error handling
Private mlErrNum As Long
Private msErrSrc As String
Private msErrDesc As String

Private Property Get ClassName() As String
    ClassName = App.EXEName & ".clsUtil"
End Property

Public Property Get FSO() As Scripting.FileSystemObject
    Set FSO = mFSO
End Property

Public Property Let MySeed(ByVal plSeed As Long)
    mlSECURITY_SEED = plSeed
End Property
Public Property Get MySeed() As Long
    MySeed = mlSECURITY_SEED
End Property

Public Property Let gsAppEXEName(psName As String)
    msAPPEXEName = psName
End Property
Public Property Get gsAppEXEName() As String
    gsAppEXEName = msAPPEXEName
End Property

Public Property Let gsMainAppEXEName(psName As String)
    msMainAppEXEName = psName
End Property
Public Property Get gsMainAppEXEName() As String
    gsMainAppEXEName = msMainAppEXEName
End Property

Public Property Let gsCommandString(psName As String)
    msCommandString = psName
End Property
Public Property Get gsCommandString() As String
    gsCommandString = msCommandString
End Property

Public Property Let gsInstallDir(psDir As String)
    msInstallDir = psDir
End Property
Public Property Get gsInstallDir() As String
    gsInstallDir = msInstallDir
End Property

Public Property Let bUCText(ByVal pbFlag As Boolean)
    mbUCText = pbFlag
End Property
Public Property Get bUCText() As Boolean
    bUCText = mbUCText
End Property

Public Sub utAlwaysOnTop(pForm As Object, pbSetOnTop As Boolean, Optional plHwnd As Long = 0)
    On Error GoTo EH
    Dim lFlag As Long
    Dim lHwnd As Long
    
    If plHwnd > 0 Then
        lHwnd = plHwnd
    Else
        lHwnd = pForm.hWnd
    End If
    
    If pbSetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    
    SetWindowPos lHwnd, lFlag, _
    pForm.left / Screen.TwipsPerPixelX, _
    pForm.top / Screen.TwipsPerPixelY, _
    pForm.Width / Screen.TwipsPerPixelX, _
    pForm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utAlwaysOnTop: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

Public Sub utSelText(pTextBox As Object)
'Purpose: Highlights All Text

'Parameters : TextBox

'Returns: Just Highlights TExt in the pTextBox

'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
    On Error GoTo EH
    With pTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    Exit Sub
EH:
    Err.Clear
End Sub

Public Sub utUCText(pControl As Object)
    On Error GoTo EH
    Dim iSelpos As Integer
    
    If Not mbUCText Then
        mbUCText = True
        iSelpos = pControl.SelStart
        With pControl
            .Text = UCase(.Text)
            .SelStart = iSelpos
            .SelLength = 0
        End With
        mbUCText = False
    End If
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utUCText: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

Public Function utFileExists(strFile As String, Optional pbDirOnly As Boolean) As Boolean
    '10.24.2002 Use the File System Object since it is Superior to Dir function.
    On Error GoTo EH
    
    If strFile <> vbNullString Then
        If Not pbDirOnly Then
            utFileExists = mFSO.FileExists(strFile)
        Else
            utFileExists = mFSO.FolderExists(strFile)
        End If
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFileExists: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetFileData(psFilePath As String, Optional pbLock As Boolean = False) As String
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    
    iFFile = FreeFile
    'piFFile = iFFile
    If pbLock Then
        Open psFilePath For Binary Access Read Lock Read As #iFFile
    Else
        Open psFilePath For Binary Access Read As #iFFile
    End If
    lMyFileLen = FileLen(psFilePath) + 2
    utGetFileData = Input(lMyFileLen, #iFFile)
    If Not pbLock Then
        Close #iFFile
    End If
    
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetFileData: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Close #iFFile
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetFilePath(ByVal psFullQualPath As String) As String
    On Error GoTo EH
    
    If InStr(1, psFullQualPath, "\", vbBinaryCompare) > 0 Then
        utGetFilePath = left(psFullQualPath, InStrRev(psFullQualPath, "\", , vbBinaryCompare))
    Else
        utGetFilePath = vbNullString
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFileExists: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function
Public Function utGetFileName(ByVal psFullQualPath As String)
    On Error GoTo EH
    
    utGetFileName = mFSO.GetFileName(psFullQualPath)
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFileExists: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Sub utSaveFileData(psFilePath As String, psFileData As String, Optional psDelimeter As String, Optional pbLock As Boolean = False)
    On Error GoTo EH
    Dim lMyFileLen As Long
    Dim iFFile As Integer
    
    iFFile = FreeFile
    'piFFile = iFFile
    Open psFilePath For Binary Access Write As #iFFile
    Put #iFFile, 1, psFileData & psDelimeter
    If Not pbLock Then
        Close #iFFile
    End If
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utSaveFileData: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Close #iFFile
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

Public Function utFindSetForm(goForms As Object, psFormName As String, poForm As Object) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    
    For iCount = 0 To goForms.Count - 1
        If UCase(goForms(iCount).Name) = UCase(psFormName) Then
            Set poForm = goForms(iCount)
            utFindSetForm = True
            Exit For
        End If
    Next
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFindSetForm: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utFormExists(goForms As Object, psFormName As String) As Boolean
    On Error GoTo EH
    Dim iCount As Integer
    
    For iCount = 0 To goForms.Count - 1
        If UCase(goForms(iCount).Name) = UCase(psFormName) Then
            utFormExists = True
            Exit For
        End If
    Next
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFormExists: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetCaption(ByVal lHwnd As Long) As String
     On Error GoTo EH
    Dim sInput As String
    Dim lLen As Long
    
    lLen = GetWindowTextLength(lHwnd)
    sInput = String(lLen&, 0)
    Call GetWindowText(lHwnd&, sInput, lLen + 1)
    utGetCaption = sInput
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetCaption: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utFindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    On Error GoTo EH
    'calls Win API
    'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    utFindWindow = FindWindow(lpClassName, lpWindowName)
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFindWindow: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utFindNextWindow(psAppCaption As String, plhWndStart As Long, plhWndFound As Long) As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim lRetVal As Long
    Dim sFoundCaption As String
    Dim sDoNotClose As String
    
    lHwnd = plhWndStart
LOOK_FOR_WINDOW:
    Do Until lHwnd = 0
        lHwnd = GetNextWindow(lHwnd, 2)
        sFoundCaption = utGetCaption(lHwnd)
        If InStr(1, sFoundCaption, psAppCaption, vbTextCompare) > 0 Then
            If InStr(1, sDoNotClose, sFoundCaption, vbTextCompare) = 0 Then
                plhWndFound = lHwnd
                utFindNextWindow = True
                Exit Do
            End If
        End If
    Loop
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFindNextWindow: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utFindWindowPartial(AppTitle As String, _
   Optional Method As FindWindowPartialTypes = FwpStartsWith, _
   Optional CaseSensitive As Boolean = False, _
   Optional MustBeVisible As Boolean = False) As Long
   On Error GoTo EH
   
   utFindWindowPartial = FindWindowPartial(AppTitle, Method, CaseSensitive, MustBeVisible)
   
   Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFindWindowPartial: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utLookForWindow(psWindowCaption As String, Optional plTimeOutSeconds As Long = 60) As Boolean
    On Error GoTo EH
    Dim bFound As Boolean
    Dim lSleep As Long
    
    bFound = True
    Do Until utWindowFound(psWindowCaption)
        lSleep = lSleep + 1
        If lSleep > plTimeOutSeconds Then
            bFound = False
            Exit Do
        End If
    Loop
    
    utLookForWindow = bFound
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utLookForWindow: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utAppActivatePartial(AppTitle As String, Optional Method As FindWindowPartialTypes = FwpStartsWith, Optional CaseSensitive As Boolean = False) As Long
    On Error GoTo EH
    Dim hWndApp As Long
    '
    ' Retrieve window handle for first top-level window
    ' that starts with or contains the passed string.
    '
    hWndApp = utFindWindowPartial(AppTitle, Method, CaseSensitive, True)
    
    If hWndApp Then
       '
       ' Switch to it, restoring if need be.
       '
       If IsIconic(hWndApp) Then
          Call ShowWindow(hWndApp, SW_RESTORE)
       End If
       
       Call SetForegroundWindow(hWndApp)
       
       utAppActivatePartial = hWndApp
       
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utAppActivatePartial: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetSetting(psAPP As String, psSection As String, psKey As String, _
                                   Optional pvDefault As String = vbNullString) As Variant
    On Error GoTo EH
    
    utGetSetting = GetRegSetting(psSection & "\" & psAPP, psKey, pvDefault)
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetSetting: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

'BGS 3.2.2006 This function Gets registry setting value
Private Function GetRegSetting(psAppPath As String, _
                                    psKey As String, _
                                    Optional psDefaultSetting As String) As String
    On Error GoTo EH
    Dim sSetting As String
                                            
    sSetting = modRegistry.GetRegistryString(HKEY_LOCAL_MACHINE, psAppPath, psKey, psDefaultSetting)
    
    If sSetting <> vbNullString Then
        GetRegSetting = sSetting
    Else
        GetRegSetting = psDefaultSetting
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = vbCrLf & ClassName & ": " & "Error In Private Function GetRegSetting:" & " Desc: " & Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Sub utSaveSetting(psAPP As String, psSection As String, psKey As String, psSetting As String)
    On Error GoTo EH
    
    SaveRegSetting psSection & "\" & psAPP, psKey, psSetting
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utSaveSetting: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

'BGS 3.2.2006 This function saves registry setting value
Private Sub SaveRegSetting(psAppPath As String, psKey As String, psSetting As String)
    On Error GoTo EH
    Dim sSetting As String
    
    sSetting = psSetting
    
    modRegistry.SaveRegistryString HKEY_LOCAL_MACHINE, psAppPath, psKey, sSetting
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = vbCrLf & ClassName & ": " & "Error In Private Sub SaveRegSetting:" & " Desc: " & Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

Public Function utGetWinOSVersion() As String
    On Error GoTo EH
    Dim WinOS As OSVERSIONINFO
    Dim sVSInfo As String
    
    WinOS.dwOSVersionInfoSize = Len(WinOS)
    GetVersionEx WinOS
    
    sVSInfo = "Build Number: " & WinOS.dwBuildNumber & vbCrLf
    sVSInfo = sVSInfo & "Major Version: " & WinOS.dwMajorVersion & vbCrLf
    sVSInfo = sVSInfo & "Minor Verison: " & WinOS.dwMinorVersion & vbCrLf
    sVSInfo = sVSInfo & "OS VS Info Size: " & WinOS.dwOSVersionInfoSize & vbCrLf
    sVSInfo = sVSInfo & "Platform ID: " & WinOS.dwPlatformId & vbCrLf
    sVSInfo = sVSInfo & "CSD Version: " & WinOS.szCSDVersion & vbCrLf
    
    utGetWinOSVersion = sVSInfo
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetWinOSVersion: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetMyComputerName() As String
    On Error GoTo EH
    Dim lSize As Long
    Dim lRet As Long
    
    lSize = 255
    utGetMyComputerName = Space(lSize)
    lRet = GetComputerName(utGetMyComputerName, lSize)
    utGetMyComputerName = left(utGetMyComputerName, lSize)
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetMyComputerName: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utGetMyUserName() As String
   On Error GoTo EH
    Dim lSize As Long
    Dim lRet As Long
    
    lSize = 255
    utGetMyUserName = Space(lSize)
    lRet = GetUserName(utGetMyUserName, lSize)
    utGetMyUserName = left(utGetMyUserName, lSize)
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetMyUserName: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utCleanSQLString(pvText As Variant) As Variant
    On Error GoTo EH
    
    'BGS 11.21.2001 add some more cleaning
    If VarType(pvText) = vbString Then
        utCleanSQLString = Replace(pvText, "'", "''")
    Else
        utCleanSQLString = pvText
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utCleanSQLString: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utCleanValString(psValText As String) As String
    'Val function Bug in VB6
    'http://msdn.microsoft.com/vbasic/productinfo/previous/vb6/tips/01pasttips.asp
    'Need to parse out both % and !  because these trailing equate to Double and Single
    'and Val() bugs because it can't convert Double or single into integer
    On Error GoTo EH
    Dim lCount As Long
    Dim sTemp As String
    
    sTemp = psValText
    
    sTemp = Trim(sTemp)
    For lCount = 1 To Len(sTemp)
        If IsNumeric(Mid(sTemp, lCount, 1)) Then
            Exit For
        End If
    Next
    sTemp = Mid(sTemp, lCount)
    
    sTemp = Replace(sTemp, ",", vbNullString)
    sTemp = Replace(sTemp, "%", vbNullString)
    sTemp = Replace(sTemp, "!", vbNullString)
    utCleanValString = sTemp
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utCleanValString: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Sub utCleanValTextBox(pvText As Variant)
    On Error GoTo EH
    Dim lPos As Long
    Dim sTemp As String
    
    'Clean out any Special chars \/:*?"<>|
    If IsObject(pvText) Then
        lPos = pvText.SelStart
        sTemp = pvText.Text
    Else
        sTemp = CStr(pvText)
    End If
    
    sTemp = Replace(sTemp, "%", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "!", vbNullString, , , vbBinaryCompare)
    
    If IsObject(pvText) Then
        pvText.Text = sTemp
        pvText.SelStart = lPos
    Else
        pvText = sTemp
    End If
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utCleanValTextBox: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub


Public Function utGetTickCount() As String
    On Error GoTo EH
    
    'Uses WIN API Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    utGetTickCount = CStr(timeGetTime)
                         
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetTickCount: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Sub utSuffixLabels(plblArray As Object, Optional plLen As Long = 25)
    On Error Resume Next
    Dim lCount As Long
    Dim oLabel As Label

    For Each oLabel In plblArray
        If Len(oLabel.Caption) < plLen Then
            oLabel.Caption = oLabel.Caption & String(plLen - Len(oLabel.Caption), ".")
        End If
    Next
    
End Sub

Public Function utShowFormIgnoreModality(poForm As Object, Optional pbShow As Boolean = True) As Boolean
    On Error GoTo EH
    
    If Not poForm Is Nothing Then
        If pbShow Then
            If poForm.left < 0 Then
                poForm.left = poForm.left + 20000
                poForm.Refresh
            End If
        Else
            If poForm.left > 0 Then
                poForm.left = poForm.left - 20000
                poForm.Refresh
            End If
        End If
    End If
    
    utShowFormIgnoreModality = True
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utShowFormIgnoreModality: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utCleanDirectory(ByRef poTextBox As Object) As String
    Dim sDir As String
    
    utCleanFileFolderName poTextBox, True
    
    sDir = poTextBox.Text
    
    utCleanDirectory = sDir
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utCleanFileFolderName: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Sub utCleanFileFolderName(ByRef pvText As Variant, Optional ByVal pbFilePath As Boolean)
    On Error GoTo EH
    Dim lPos As Long
    Dim sTemp As String
    'Clean out any Special chars \/:*?"<>|
    If IsObject(pvText) Then
        lPos = pvText.SelStart
        sTemp = pvText.Text
    Else
        sTemp = CStr(pvText)
    End If
    If Not pbFilePath Then
        sTemp = Replace(sTemp, "\", vbNullString, , , vbBinaryCompare)
    End If
    sTemp = Replace(sTemp, "/", vbNullString, , , vbBinaryCompare)
    If Not pbFilePath Then
        sTemp = Replace(sTemp, ":", vbNullString, , , vbBinaryCompare)
    End If
    sTemp = Replace(sTemp, "*", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "?", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, """", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "<", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, ">", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, "|", vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, vbCr, vbNullString, , , vbBinaryCompare)
    sTemp = Replace(sTemp, vbLf, vbNullString, , , vbBinaryCompare)
    If IsObject(pvText) Then
        pvText.Text = sTemp
        pvText.SelStart = lPos
    Else
        pvText = sTemp
    End If
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Sub utCleanFileFolderName: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub

Public Sub utBubbleSort(ByRef pvArray As Variant, _
                        Optional ByVal pb2D As Boolean = False, _
                        Optional ByVal plSortBy2dPos = 0)
    On Error GoTo EH
    Dim lMainLoop As Long
    Dim lSubLoop As Long
    Dim lPos As Long
    Dim lCompLen As Long
    Dim lCompLen2 As Long
    
    For lMainLoop = UBound(pvArray, 1) To LBound(pvArray, 1) Step -1
        For lSubLoop = LBound(pvArray) + 1 To lMainLoop
            'Only sort non nullstrings
            If pb2D Then
                lCompLen = Len(pvArray(lSubLoop - 1, plSortBy2dPos))
                lCompLen2 = Len(pvArray(lSubLoop, plSortBy2dPos))
                If lCompLen > lCompLen2 Then
                    lCompLen = lCompLen2
                End If
                If left(pvArray(lSubLoop - 1, plSortBy2dPos), lCompLen) > left(pvArray(lSubLoop, plSortBy2dPos), lCompLen) Then
                    'Sort all the elements of the 2 dimension
                    For lPos = LBound(pvArray, 2) To UBound(pvArray, 2)
                        Call SwitchPlace(pvArray(lSubLoop - 1, lPos), pvArray(lSubLoop, lPos))
                    Next
                End If
            Else
                If pvArray(lSubLoop) <> vbNullString Then
                    lCompLen = Len(pvArray(lSubLoop - 1))
                    lCompLen2 = Len(pvArray(lSubLoop))
                    If lCompLen > lCompLen2 Then
                        lCompLen = lCompLen2
                    End If
                    If left(pvArray(lSubLoop - 1), lCompLen) > left(pvArray(lSubLoop), lCompLen) Then
                        Call SwitchPlace(pvArray(lSubLoop - 1), pvArray(lSubLoop))
                    End If
                End If
            End If
        Next
    Next
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = vbCrLf & ClassName & ": " & "Error In Friend Sub utBubbleSort:" & " Desc: " & Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Sub




Private Sub SwitchPlace(a As Variant, b As Variant)
    'used by Bubble sort
    Dim C As Variant
    C = a
    a = b
    b = C
End Sub

Public Function utValidDate(psDate As String) As String
    On Error GoTo EH
    'will retunrn "12:00:00 AM" if invalid
    Dim sDate As String
    
    sDate = psDate
    
    If IsDate(sDate) Then
        'BGS 5.8.2002 Check to see if they entered like 12:00AM
        'it will default to year 1899 All dates enterd should be at least
        '1900 and above. Unless we have over 100 year old claims.. Hmm you think ?
        If Format(sDate, "YYYY") > 1899 Then
            sDate = Format(sDate, "MM/DD/YY")
        Else
            sDate = NULL_DATE
        End If
    Else
        sDate = NULL_DATE
    End If
    
    utValidDate = sDate
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utValidDate: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utValidSSN(psSSN As String) As String
    '
End Function

'============================================================================================================
'This function will look for the Application window by Title caption (First partial match)
'If it does not find it then it will loop for specified time
'============================================================================================================
Public Function utWindowFound(ByVal psName As String, Optional plSleepSeconds As Long = 1) As Long
    On Error GoTo EH
    Dim lSleep As Long
    Dim lRet As Long
    
    'BGS 1.28.2002 Sleep and look for the window
    For lSleep = 1 To plSleepSeconds * 10
        DoEvents
        Sleep 100
        lRet = utAppActivatePartial(psName)
        If lRet Then
            utWindowFound = lRet
            Exit Function
        End If
    Next
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utWindowFound: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utDeleteFile(psFilePath As String) As String
    On Error GoTo EH
    
    mFSO.DeleteFile psFilePath, True
    
    Exit Function
EH:
    utDeleteFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utCopyFile(psSourceFilePath As String, psDestFilePath As String) As String
    On Error GoTo EH
    
    mFSO.CopyFile psSourceFilePath, psDestFilePath, True
    
    Exit Function
EH:
    utCopyFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utMoveFile(psSourceFilePath As String, psDestFilePath As String) As String
    On Error GoTo EH
    
    mFSO.MoveFile psSourceFilePath, psDestFilePath
    
    Exit Function
EH:
    utMoveFile = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utMoveFolder(psSourceFolderPath As String, psDestFolderPath As String) As String
    On Error GoTo EH
    
    mFSO.MoveFolder psSourceFolderPath, psDestFolderPath
    
    Exit Function
EH:
    utMoveFolder = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utDynamicArraySet(pVarArray As Variant) As Boolean
    'Purpose: To see if a Dynamic array has been set
    'Parameters : pVarArray As Variant: Send in any Dynamic array data type
    'Returns: True if has been set, false if not
    'Author : BGS-3/24/2000
    'Revision History:  SMR     Initials    Date    Description
    
    On Error GoTo NOT_SET
    Dim iRet As Integer
    
    If IsArray(pVarArray) Then
        iRet = LBound(pVarArray, 1)
        'if the Lbound call to the first dimension of
        'pVarArray does not error then the dynamic array must
        'be set so...
        utDynamicArraySet = True
        Exit Function
    End If
    
NOT_SET:
    utDynamicArraySet = False
End Function

Public Function utDeleteDir(psDirPath As String) As String
    On Error GoTo EH

    mFSO.DeleteFolder psDirPath, True

    Exit Function
EH:
    utDeleteDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utMakeDir(psDirPath As String) As String
    On Error GoTo EH
    
    mFSO.CreateFolder psDirPath
    
    Exit Function
EH:
    utMakeDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utCopyDir(psSourceDirPath As String, psDestDirPath As String) As String
    On Error GoTo EH
    
    mFSO.CopyFolder psSourceDirPath, psDestDirPath, True
    
    Exit Function
EH:
    utCopyDir = Err.Number & vbCrLf & vbCrLf & Err.Description
End Function

Public Function utGetSystemDir() As String
    On Error GoTo EH
    Dim sDir As String
    Dim lRet As Long
    
    sDir = Space(260)
    lRet = GetSystemDirectory(sDir, Len(sDir))
    sDir = left(sDir, lRet)
    
    utGetSystemDir = sDir
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utGetSystemDir: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utFindCBOItem(psSearchText As String, pCBO As Object, piPos As Integer) As String
    On Error GoTo EH
    Dim iCount As Integer
    Dim oCbo As Object
    
    Set oCbo = pCBO
    
    If oCbo.ListCount > 0 Then
        For iCount = 0 To oCbo.ListCount - 1
            If InStr(1, left(oCbo.List(iCount), piPos), psSearchText, vbTextCompare) > 0 Then
                utFindCBOItem = oCbo.List(iCount)
                Exit Function
            End If
        Next
        'BGS if we did not find the search item then pass it back what was sent in
        utFindCBOItem = psSearchText
    End If
    
    Set oCbo = Nothing
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utFindCBOItem: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utNumInTextSortFormat(psText As String) As String
    'This function will search Text string and format any numerics
    'with 15 digit 0 place holder so numbers can be sorted within a string.
    On Error GoTo EH
    Dim sText As String
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lCount As Long
    Dim lCount2 As Long
    Dim varyText As Variant
    
    sText = psText
    
    varyText = Split(sText, Chr(32))
    
    If IsArray(varyText) Then
        For lCount = LBound(varyText, 1) To UBound(varyText, 1)
            sText = varyText(lCount)
            If IsNumeric(sText) Then
                sText = Format(sText, "000000000000000.000000000000000000000000000000000000000000000")
                varyText(lCount) = sText
            Else
                lCount2 = 1
                sTemp = vbNullString
                sTemp2 = vbNullString
                'If there is a Numeric + Alpha I.E. 1ST street or 2ND etc
                'Need to find the pos where the number stops and format it as well!
                'Need to format a number in the middle of text.
                Do Until lCount2 > Len(sText)
                    If Not IsNumeric(Mid(sText, lCount2, 1)) Then
                        sTemp = sTemp & Mid(sText, lCount2, 1)
                        lCount2 = lCount2 + 1
                    Else
                        sTemp2 = vbNullString
                        Do Until Not IsNumeric(Mid(sText, lCount2, 1))
                            sTemp2 = sTemp2 & Mid(sText, lCount2, 1)
                            lCount2 = lCount2 + 1
                        Loop
                        sTemp = sTemp & Format(sTemp2, "000000000000000.000000000000000000000000000000000000000000000")
                    End If
                Loop
                varyText(lCount) = sTemp
            End If
        Next
        utNumInTextSortFormat = Join(varyText, Chr(32))
    Else
        utNumInTextSortFormat = vbNullString
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utNumInTextSortFormat: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utScrubAlphaNumeric(psScrubMe As String, Optional pbAllowSpaces As Boolean = False) As String
    On Error GoTo EH
    Dim sSCrub As String
    Dim sClean As String
    Dim lPos As Long
    Dim sLetter As String
    
    sSCrub = psScrubMe
    sClean = psScrubMe
    
    For lPos = 1 To Len(sSCrub)
        sLetter = Mid(sSCrub, lPos, 1)
        If Not utIsAlphaNumeric(sLetter, pbAllowSpaces) Then
            sClean = Replace(sClean, sLetter, vbNullString, , , vbBinaryCompare)
        End If
    Next
    
    utScrubAlphaNumeric = sClean
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utScrubAlphaNumeric: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utIsAlphaNumeric(psLetter As String, Optional pbAllowSpaces As Boolean = False) As Boolean
    On Error GoTo EH
    
    If pbAllowSpaces Then
        utIsAlphaNumeric = psLetter Like "[A-Za-z0-9_- ]"
    Else
        utIsAlphaNumeric = psLetter Like "[A-Za-z0-9]"
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utIsAlphaNumeric: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utScrubNumeric(psScrubMe As String, _
                            Optional pbAllowDecimal As Boolean = False, _
                            Optional pbAllowPercent As Boolean = False) As String
    On Error GoTo EH
    Dim sSCrub As String
    Dim sClean As String
    Dim lPos As Long
    Dim sLetter As String
    Dim lFoundDecPos As Long
    Dim bScrubAgain As Boolean
    
    sSCrub = psScrubMe
    sClean = psScrubMe
    
    If pbAllowPercent Then
        'If allowing a Percent then it must have a percent at the very end
        sClean = Trim(sClean)
        If InStr(1, sClean, "%", vbTextCompare) = 0 And sClean <> vbNullString Then
            sClean = sClean & "%"
        End If
        
        If sClean <> vbNullString Then
            sClean = Replace(sClean, "%", vbNullString, , , vbBinaryCompare)
            sClean = sClean & "%"
        End If
         
    End If
    
    sSCrub = sClean
SCRUB_AGAIN:
    lFoundDecPos = 0
    For lPos = 1 To Len(sSCrub)
        sLetter = Mid(sSCrub, lPos, 1)
        If pbAllowDecimal Then
            If StrComp(sLetter, ".", vbBinaryCompare) = 0 Then
                If lFoundDecPos = 0 Then
                    lFoundDecPos = lPos
                Else
                    'If this is a repeating decimal then get rid of it subsequent ones
                    sLetter = "#" ' "#" will be cleaned below
                    Mid(sSCrub, lPos, 1) = sLetter
                End If
            End If
        End If
        If Not utIsNumeric(sLetter, pbAllowDecimal, pbAllowPercent) Then
            sClean = Replace(sClean, sLetter, vbNullString, , , vbBinaryCompare)
        End If
    Next
    
    If InStr(1, sSCrub, "#", vbBinaryCompare) > 0 And Not bScrubAgain Then
        sClean = sSCrub
        bScrubAgain = True
        GoTo SCRUB_AGAIN
    End If
    
    If sClean = "%" Then
        sClean = vbNullString
    End If
    
    utScrubNumeric = sClean
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utScrubNumeric: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utIsNumeric(psLetter As String, _
                            Optional pbAllowDecimal As Boolean = False, _
                            Optional pbAllowPercent As Boolean = False) As Boolean
    On Error GoTo EH
    
    If pbAllowDecimal And Not pbAllowPercent Then
        utIsNumeric = psLetter Like "[0-9_.]"
    ElseIf Not pbAllowDecimal And pbAllowPercent Then
        utIsNumeric = psLetter Like "[0-9_%]"
    ElseIf pbAllowDecimal And pbAllowPercent Then
        utIsNumeric = psLetter Like "[0-9_._%]"
    Else
        utIsNumeric = psLetter Like "[0-9]"
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utIsNumeric: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function


Public Function utIsControlArray(MyForm As Object, MyControl As Object) As Boolean
    
    'BGS 8/1/1999 Added this function to determin if a Control is part of
    'a control array or not. I had to do this because VB does not have a
    'function that figures this out. (IsArray does not work on Control Arrays)
    On Error GoTo EH
    Dim MyCount As Integer
    Dim CheckMyControl As Object
    
    
    
    For Each CheckMyControl In MyForm.Controls
        If CheckMyControl.Name = MyControl.Name Then
            MyCount = MyCount + 1
            If MyCount > 1 Then
                Exit For
            End If
        End If
    Next
    
    utIsControlArray = MyCount - 1
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utIsControlArray: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function utActiveFile(psActiveFilePath As String, plCurrentTry As Long) As Boolean
    '3.3.2004 Active file will return true IF
    'The file is is actively being written to
    On Error GoTo EH
    Dim iFFile As Integer
    Dim sActiveFile As String
    
    sActiveFile = psActiveFilePath
    plCurrentTry = plCurrentTry + 1
    iFFile = FreeFile
    Open sActiveFile For Binary Access Read Lock Read As #iFFile
    Close #iFFile
    utActiveFile = False
    Exit Function
EH:
    Close #iFFile
    utActiveFile = True
End Function

Public Function RemoveItemFromDictionary(pDictionary As Scripting.Dictionary, psKey As String) As Boolean
    On Error GoTo EH
    
    If Not pDictionary Is Nothing Then
        If pDictionary.Exists(psKey) Then
            pDictionary.Remove psKey
            RemoveItemFromDictionary = True
        Else
            RemoveItemFromDictionary = False
        End If
    Else
        RemoveItemFromDictionary = False
    End If
    
    Exit Function
EH:
    Err.Clear
    RemoveItemFromDictionary = False
End Function

Public Function GetItemFromDictionary(pDictionary As Scripting.Dictionary, psKey As String) As String
    On Error GoTo EH
    
    If Not pDictionary Is Nothing Then
        If pDictionary.Exists(psKey) Then
            GetItemFromDictionary = pDictionary(psKey)
        Else
            GetItemFromDictionary = vbNullString
        End If
    Else
        GetItemFromDictionary = vbNullString
    End If
    
    Exit Function
EH:
    Err.Clear
    GetItemFromDictionary = vbNullString
End Function

Public Function AddItemToCollection(pCol As Collection, pvItem As Variant, psKey As String) As Boolean
    On Error GoTo EH
    
    If pCol Is Nothing Then
        Set pCol = New Collection
    End If
    
    pCol.Add psKey, pvItem
    
    AddItemToCollection = True
    
    Exit Function
EH:
    Err.Clear
    AddItemToCollection = False
End Function

Public Function RemoveItemFromCollection(pCol As Collection, psKey As String) As Boolean
    On Error GoTo EH
    
    If Not pCol Is Nothing Then
        pCol.Remove psKey
    Else
        RemoveItemFromCollection = False
    End If
    
    RemoveItemFromCollection = True
    
    Exit Function
EH:
    Err.Clear
    RemoveItemFromCollection = False
End Function

Public Function GetItemFromCollection(pCol As Collection, psKey As String) As Variant
    On Error GoTo EH
    
    If Not pCol Is Nothing Then
        GetItemFromCollection = pCol(psKey)
    Else
        GetItemFromCollection = vbNullString
    End If
    
    Exit Function
EH:
    Err.Clear
    GetItemFromCollection = vbNullString
End Function

Public Function utShellExecute(Optional plHwnd As Long = -1, _
                                Optional pslpOperation As String = "OPEN", _
                                Optional pslpFile As String, _
                                Optional pslpParameters As String = vbNullString, _
                                Optional pslpDirectory As String = "App.Path", _
                                Optional plnShowCmd As VBA.VbAppWinStyle = vbNormalFocus, _
                                Optional pbCreateTempFile As Boolean = False, _
                                Optional pbUseTimeStampFileName As Boolean = False, _
                                Optional pbShowMessage As Boolean = False, _
                                Optional psTempFileCaption As String) As Boolean
    On Error GoTo EH
    Dim lHwnd As Long
    Dim slpOperation As String
    Dim slpFile As String
    Dim slpParameters As String
    Dim slpDirectory As String
    Dim lnShowCmd As VBA.VbAppWinStyle
    Dim sErrorMess As String
    Dim sTmpExt As String
    Dim sTmpFile As String
    Dim lRet As Long
    Dim sDir As String
    
    utShellExecute = False
    
    'Get info from Params
    If plHwnd = -1 Then
        lHwnd = GetDesktopWindow
    End If
    slpOperation = pslpOperation
    If pslpFile = vbNullString Then
        Exit Function
    Else
        slpFile = pslpFile
    End If
     
    slpParameters = pslpParameters
    If pslpDirectory = "App.Path" Then
        If gsInstallDir <> vbNullString Then
            slpDirectory = gsInstallDir
        Else
            slpDirectory = App.Path
        End If
    Else
        slpDirectory = pslpDirectory
    End If
    
    lnShowCmd = plnShowCmd
    
    If pbCreateTempFile Then
        'Be sure the Temp dir exixts
        If Not CreateTempDir(sErrorMess) Then
            Err.Raise -999, "CreateTempDir", sErrorMess
        End If
        'See if need to create a Temp file name
        If pbUseTimeStampFileName Then
            sTmpExt = Mid(slpFile, InStrRev(slpFile, ".", , vbBinaryCompare))
            sTmpFile = "C:\Temp\Temp_" & psTempFileCaption & "_" & CStr(timeGetTime()) & sTmpExt
        Else
            sTmpFile = "C:\Temp\Temp_" & psTempFileCaption & "_" & Mid(slpFile, InStrRev(slpFile, "\", , vbBinaryCompare) + 1)
        End If
        sErrorMess = utCopyFile(slpFile, sTmpFile)
    Else
        sTmpFile = slpFile
    End If
    
    lRet = ShellExecute(lHwnd, slpOperation, sTmpFile, slpParameters, slpDirectory, lnShowCmd)
        'Check to see if the Associated Application opened the file
        'If not that means there is no application associated with the file
        'in that case open Explorer to give the user a chance to open the file
    If lRet = SE_ERR_NOASSOC Then
        sDir = Space(260)
        lRet = GetSystemDirectory(sDir, Len(sDir))
        sDir = left(sDir, lRet)
        lRet = ShellExecute(lHwnd, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & sTmpFile, sDir, lnShowCmd)
    End If
   
    utShellExecute = True
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function utShellExecute: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function CLEANUP() As Boolean
    On Error GoTo EH
    
    'Delete any files that were created under Temp Directory
    DelTempDirFiles
    
    CLEANUP = True
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function CLEANUP: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function CreateTempDir(Optional psErrorMess As String) As Boolean
    On Error GoTo EH
    Dim sError As String
    
    CreateTempDir = False
    
    If Not utFileExists("C:\Temp", True) Then
        'Create C:\Temp dir
        sError = utMakeDir("C:\Temp")
        If sError <> vbNullString Then
            psErrorMess = sError
            Exit Function
        End If
    End If
    
    CreateTempDir = True
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function CreateTempDir: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function DelTempDirFiles(Optional psTempDir As String = "C:\Temp", _
                                Optional psTempFile As String = "Temp_*") As Boolean
    On Error GoTo EH
    Dim sTempDir As String
    Dim sTempFile As String
    Dim sError As String
    
    DelTempDirFiles = False
    
    sTempDir = psTempDir
    sTempFile = psTempFile
    
    If utFileExists(sTempDir, True) Then
        'Clear any temp files created with "Temp_*"
        If Dir(sTempDir & "\" & sTempFile, vbNormal) <> vbNullString Then
            sError = utDeleteFile(sTempDir & "\" & sTempFile)
        End If
    End If
    
    DelTempDirFiles = True
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function DelTempDirFiles: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function ConvertTwipsToPixels(plTwips As Long) As Long
    On Error GoTo EH
    
    ConvertTwipsToPixels = plTwips / 15
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function ConvertTwipsToPixels: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function ConvertPixelsToTwips(plPixels As Long) As Long
    On Error GoTo EH
    
    ConvertPixelsToTwips = plPixels * 15
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function ConvertPixelsToTwips: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function GetFlagText(pbFlag As Boolean) As String
    On Error GoTo EH
    Dim sFlagText As String
    
    If pbFlag Then
        sFlagText = Chr(160)
    Else
        sFlagText = Chr(32)
    End If
    
    GetFlagText = sFlagText
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function GetFlagText: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function GetFlagFromText(psFlagText As String) As Boolean
    On Error GoTo EH
    Dim bFlag As Boolean
    
    If psFlagText = Chr(160) Then
        bFlag = True
    Else
        bFlag = False
    End If
    
    GetFlagFromText = bFlag
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function GetFlagFromText: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Public Function GetvbVarTypeFromText(psvbVarTypeText As String) As VBA.VbVarType
    On Error GoTo EH
    Dim sDataType As String
    Dim iDataType As VBA.VbVarType
    
    sDataType = psvbVarTypeText
    
    If StrComp(sDataType, "vbArray", vbTextCompare) = 0 Then
        iDataType = vbArray
    ElseIf StrComp(sDataType, "vbBoolean", vbTextCompare) = 0 Then
        iDataType = vbBoolean
    ElseIf StrComp(sDataType, "vbByte", vbTextCompare) = 0 Then
        iDataType = vbByte
    ElseIf StrComp(sDataType, "vbCurrency", vbTextCompare) = 0 Then
        iDataType = vbCurrency
    ElseIf StrComp(sDataType, "vbDataObject", vbTextCompare) = 0 Then
        iDataType = vbDataObject
    ElseIf StrComp(sDataType, "vbDate", vbTextCompare) = 0 Then
       iDataType = vbDate
    ElseIf StrComp(sDataType, "vbDecimal", vbTextCompare) = 0 Then
       iDataType = vbDecimal
    ElseIf StrComp(sDataType, "vbDouble", vbTextCompare) = 0 Then
       iDataType = vbDouble
    ElseIf StrComp(sDataType, "vbEmpty", vbTextCompare) = 0 Then
       iDataType = vbEmpty
    ElseIf StrComp(sDataType, "vbError", vbTextCompare) = 0 Then
       iDataType = vbError
    ElseIf StrComp(sDataType, "vbInteger", vbTextCompare) = 0 Then
       iDataType = vbInteger
    ElseIf StrComp(sDataType, "vbLong", vbTextCompare) = 0 Then
       iDataType = vbLong
    ElseIf StrComp(sDataType, "vbNull", vbTextCompare) = 0 Then
       iDataType = vbNull
    ElseIf StrComp(sDataType, "vbObject", vbTextCompare) = 0 Then
       iDataType = vbObject
    ElseIf StrComp(sDataType, "vbSingle", vbTextCompare) = 0 Then
       iDataType = vbSingle
    ElseIf StrComp(sDataType, "vbString", vbTextCompare) = 0 Then
        iDataType = vbString
    ElseIf StrComp(sDataType, "vbUserDefinedType", vbTextCompare) = 0 Then
        iDataType = vbUserDefinedType
    ElseIf StrComp(sDataType, "vbVariant", vbTextCompare) = 0 Then
        iDataType = vbVariant
    Else
        iDataType = vbVariant
    End If
    
    GetvbVarTypeFromText = iDataType
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = vbCrLf & ClassName & ": " & "Error in Public Function GetvbVarTypeFromText: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function
Public Function utFormWinRegPos(psAppEXEName As String, poMyForm As Object, Optional pbSave As Boolean, _
                              Optional pfrmOffset As Object, Optional pctrlOffset As Object, _
                              Optional pbUseFullCaption As Boolean = True, Optional pbUseFrmName As Boolean) As Boolean
    utFormWinRegPos = FormWinRegPos(psAppEXEName, poMyForm, pbSave, pfrmOffset, pctrlOffset, pbUseFullCaption, pbUseFrmName)
                              
End Function

Public Sub utHideAllForms(goForms As Object, Optional psSkipForms As String)
    On Error GoTo EH
    Dim MyForm As Form
    
    For Each MyForm In goForms
        If InStr(1, psSkipForms, MyForm.Name, vbTextCompare) = 0 Then
            MyForm.Visible = False
        End If
    Next
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub utHideAllForms"
End Sub

Public Sub utShowAllForms(goForms As Object, oPassMindTray As Object, psNavPos As String, Optional psSkipForms As String)
    On Error GoTo EH
    Dim MyForm As Form
    Dim bCheckModalFlag As Boolean
    
    On Error GoTo EH
    
    For Each MyForm In goForms
        If InStr(1, psSkipForms, MyForm.Name, vbTextCompare) = 0 Then
            If MyForm.Visible Then
                If MyForm.WindowState <> vbMaximized Then
                    MyForm.WindowState = vbNormal
                    'Check the Pos of the open forms to see if they
                    'are hidden behind the Nav Screen. If they are, Move the
                    'Left to make it fully visible
                    If psNavPos = "LEFT" Then
                        If MyForm.left < oPassMindTray.Width Then
                            MyForm.left = oPassMindTray.left + oPassMindTray.Width
                        End If
                    Else
                        If MyForm.left + MyForm.Width > oPassMindTray.left Then
                            MyForm.left = oPassMindTray.left - MyForm.Width
                        End If
                    End If
                End If
            End If
            On Error Resume Next
            MyForm.Show
            If Err.Number > 0 Then
                Err.Clear
            End If
            On Error GoTo EH
        End If
    Next
    
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Sub utShowAllForms"
End Sub

Public Function utGetTaskbarHeight() As Long
    utGetTaskbarHeight = GetTaskbarHeight
End Function

Public Function utGetPath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
                         
    utGetPath = GetPath(psAppEXEName, psName, psMess, psFileMess, psDeFaultPath, plHwnd, psFilter, psSelFile, plFlags, pbCenterForm, pbShowOpen)
                         
End Function

Public Function utGetUserAccounts() As Variant
    On Error GoTo EH
    
    Dim sKeySet As String
    
    sKeySet = GetSetting(App.EXEName, "USERS", "KeySet", "False")
    If sKeySet <> "False" Then
        utGetUserAccounts = Split(sKeySet, "|")
    Else
        SaveSetting App.EXEName, "USERS", "KeySet", vbNullString
        utGetUserAccounts = Split(vbNullString, "|")
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function utGetUserAccounts"
End Function

Public Function SaveUserAccounts(ByRef pvaryUserAccounts As Variant) As Boolean
    On Error GoTo EH
    
    Dim sKeySet As String
    
    'Sort the Array
    utBubbleSort pvaryUserAccounts
    
    sKeySet = Join(pvaryUserAccounts, "|")
    
    SaveSetting App.EXEName, "USERS", "KeySet", sKeySet
    
    SaveUserAccounts = True
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function SaveUserAccounts"
End Function

Public Function utRemoveUserAccount(ByVal psUserName As String) As Boolean
    On Error GoTo EH
    
    Dim varyUsers As Variant
    Dim sTemp As String
    Dim lPos As Long
    
    If utDoesThisUserAccountExist(psUserName, varyUsers, lPos) Then
        Delete1FromArray varyUsers, lPos
        If SaveUserAccounts(varyUsers) Then
            DeleteSetting App.EXEName, psUserName & "_SECURITY"
            DeleteSetting App.EXEName, psUserName & "_FORM_POSN"
            DeleteSetting App.EXEName, psUserName & "_GENERAL"
'            SaveSetting App.EXEName, "MSG", "COMMAND", "SHUT_DOWN_Application"
            utRemoveUserAccount = True
        End If
    Else
        utRemoveUserAccount = False
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function utRemoveUserAccount"
End Function

Private Sub Delete1FromArray(ByRef pvArrayToChange As Variant, ByVal plDelPos As Long)
    On Error GoTo EH
    
    Dim vTmpArray As Variant
    Dim nTmpArrayCounter As Long
    Dim nTmpNewArrayCounter As Long
    Dim lPos As Long
    
    If Not IsArray(pvArrayToChange) Then
        Exit Sub
    End If
    
    vTmpArray = pvArrayToChange
    
    If UBound(vTmpArray, 1) > 0 Then
        lPos = UBound(vTmpArray) - 1
    Else
        lPos = 0
    End If
    
    ReDim pvArrayToChange(0 To lPos)
    
    If lPos = 0 Then
        If UBound(vTmpArray, 1) = 1 Then
            If plDelPos = 0 Then
                pvArrayToChange(lPos) = vTmpArray(1)
            Else
                pvArrayToChange(lPos) = vTmpArray(0)
            End If
        Else
            pvArrayToChange(lPos) = vbNullString
        End If
    Else
        For nTmpArrayCounter = 0 To lPos + 1
            If Not (nTmpArrayCounter = plDelPos) Then
                pvArrayToChange(nTmpNewArrayCounter) = vTmpArray(nTmpArrayCounter)
                nTmpNewArrayCounter = nTmpNewArrayCounter + 1
            End If
        Next nTmpArrayCounter
    End If
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Delete1FromArray"
End Sub

Private Sub Add1ToArray(ByRef pvArrayToChange As Variant, ByVal pvItemVal As Variant)
    On Error GoTo EH
    Dim lCurUbound As Long
    Dim vMyArray As Variant
    
    If IsArray(pvArrayToChange) Then
        vMyArray = pvArrayToChange
        lCurUbound = UBound(vMyArray, 1)
    Else
        lCurUbound = -1
    End If
    
    ReDim Preserve vMyArray(0 To lCurUbound + 1)
    vMyArray(lCurUbound + 1) = pvItemVal
    
    pvArrayToChange = vMyArray
    Erase vMyArray
    Exit Sub
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Private Sub Add1ToArray"
End Sub

Public Function utAddUserAccount(ByVal psUserName As String) As Boolean
    On Error GoTo EH
    
    Dim varyUsers As Variant
    Dim sTemp As String
    Dim lPos As Long
    
    If Not utDoesThisUserAccountExist(psUserName, varyUsers) Then
        Add1ToArray varyUsers, psUserName
        utAddUserAccount = SaveUserAccounts(varyUsers)
    Else
        utAddUserAccount = False
    End If
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function utAddUserAccount"
End Function

Public Function utDoesThisUserAccountExist(ByVal psUserName As String, _
                                            Optional ByRef pvAryUsers As Variant, _
                                            Optional ByRef plPos As Variant) As Boolean
    On Error GoTo EH
    
    Dim sTemp As String
    Dim varyUsers As Variant
    Dim lPos As Long
    
    varyUsers = utGetUserAccounts
    
    If Not IsMissing(pvAryUsers) Then
        pvAryUsers = varyUsers
    End If
    
    For lPos = LBound(varyUsers, 1) To UBound(varyUsers, 1)
        sTemp = varyUsers(lPos)
        If StrComp(sTemp, psUserName, vbTextCompare) = 0 Then
            utDoesThisUserAccountExist = True
            If Not IsMissing(plPos) Then
                plPos = lPos
            End If
            Exit For
        End If
    Next
    
    Exit Function
EH:
    mlErrNum = Err.Number
    msErrSrc = Err.Source
    msErrDesc = Err.Description
    Err.Raise mlErrNum, msErrSrc, msErrDesc & vbCrLf & ClassName & vbCrLf & "Public Function utDoesThisUserAccountExist"
End Function


Public Sub utErrorLog(poErr As ErrObject, psAppName As String, psClassName As String, psProcName As String, Optional pbShowError As Boolean = False)
    Dim sErrMess As String
    Dim sFileName As String
    Dim sErrorNum As String
    Dim sErrorDescp As String
    Dim sErrorLogDir As String
    Dim sSilentError As String
    
    
    'Set these before the Error handler for this Proc
    sErrorNum = poErr.Number
    sErrorDescp = poErr.Description
        
    'Need to set this here so we won't clear the first Error
    On Error GoTo EH
    sErrMess = "<<<<<<<<<< BEGIN ERROR MESSAGE >>>>>>>>>>" & vbCrLf
    sErrMess = sErrMess & Now() & vbCrLf
    sErrMess = sErrMess & "AppName: " & psAppName & vbCrLf
    sErrMess = sErrMess & "ClassName: " & psClassName & vbCrLf
    sErrMess = sErrMess & "ProcName: " & psProcName & vbCrLf
    sErrMess = sErrMess & "ERROR # " & sErrorNum & " " & Now & vbCrLf
    sErrMess = sErrMess & sErrorDescp & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<<  END ERROR MESSAGE  >>>>>>>>>>" & vbCrLf & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<< BEGIN WIN OS INFO >>>>>>>>>>" & vbCrLf
    sErrMess = sErrMess & goUtil.utGetWinOSVersion & vbCrLf
    sErrMess = sErrMess & "<<<<<<<<<<  END WIN OS INFO  >>>>>>>>>>" & vbCrLf
    
    'Get the Error Log Install Dir
    sErrorLogDir = App.Path
    
    'Check to be sure the Error Log Folder is there
    If Not goUtil.utFileExists(sErrorLogDir & "\ErrorLog", True) Then
        goUtil.utMakeDir sErrorLogDir & "\ErrorLog"
    End If
        
    sFileName = sErrorLogDir & "\ErrorLog\" & App.EXEName & "_" & Format(Now(), "YYMMDDHHMMSS") & "_Error.Log"
    goUtil.utSaveFileData sFileName, sErrMess
    
    If pbShowError Then
        MsgBox sErrMess, vbCritical + vbOKOnly, "Error"
    End If
    
    Exit Sub
EH:
    MsgBox App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Sub ErrorLog" & vbCrLf & "Error # " & Err.Number & vbCrLf & Err.Description & vbCrLf
End Sub

Private Sub Class_Initialize()
    On Error Resume Next
    Set mFSO = New Scripting.FileSystemObject
End Sub

Private Sub Class_Terminate()
    Set mFSO = Nothing
End Sub









