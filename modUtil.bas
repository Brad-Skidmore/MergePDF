Attribute VB_Name = "modUtil"
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - modUtil
' *********************************************************************

Option Explicit

Public Const NULL_DATE As String = "12:00:00 AM"
Public Const L_DLM_TAG As String = "|"
Public Const L_DLM_ITEM As String = "@"
Public Const FILLER_STR As String = "ήρ"
Public Const FILLER_STR_ADD As String = "A"
Public Const SEC_MAX_LEN As Long = 20
'Printer Types
Public Type ATPrinter
    PRINTER_NAME As String
    PRINTER_PORT As String
    PRINTER_DRIVER As String
End Type

'Constants used by FindWindowPartial to call
'Private Declare Function SystemParametersInfo
Public Enum FindWindowPartialTypes
   FwpStartsWith = 0
   FwpContains = 1
   FwpMatches = 2
End Enum

'Used Reports Print Options
Public Enum PrintFormat
    RawText = 0
    Translated
End Enum

'If Data in a collection changed
Public Enum IsDirty
    NoChange = 0
    AddMe
    DeleteMe
End Enum

'Used for Export PDF
Public Enum ExportType
    ARExcel = 0
    ARPdf
    ARRtf
    ARHtml
    ARText
End Enum

'Used to Pass various misc Paramers
Public Type udtParameter
    ParamName As String
    ParamValue As Variant
End Type

'Used for Moving Items in ListView Up or Down
Public Enum MoveListItem
    MoveUp = -1
    MoveDown = 1
End Enum

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Const SPI_GETWORKAREA = 48
Public Const SE_ERR_NOASSOC = 31
Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

'
' Private variables needed to support enumeration
'
Private m_hWnd As Long
Private m_Method As FindWindowPartialTypes
Private m_CaseSens As Boolean
Private m_Visible As Boolean
Private m_AppTitle As String


'Error handling
Private mlErrNum As Long
Private msErrSrc As String
Private msErrDesc As String

'====================================================================================================
'WIN API
'====================================================================================================
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Property Get ClassName() As String
    ClassName = "modUtil"
End Property

Public Sub ShowError(ByRef psData As String, _
                    ByVal plErrNum As Long, _
                    ByVal psErrSrc As String, _
                    ByVal psErrDesc As String, _
                    ByVal psLogPath As String)
    On Error GoTo EH
    Dim sMess As String
    Dim oUtil As clsUtil
    
    Set oUtil = New clsUtil
    
    sMess = "Error#" & plErrNum & vbCrLf & "Err Src: " & psErrSrc & vbCrLf & "Err Desc: " & psErrDesc
    
    MsgBox sMess, vbCritical + vbOKOnly, "Error"
    
    psData = psData & sMess
    
    'Save the Log
    oUtil.utSaveFileData psLogPath, psData
    
    'Show the Import Log
    oUtil.utShellExecute , , psLogPath, , App.Path
    Set oUtil = Nothing
    
    Exit Sub
EH:
    MsgBox "Error handler Error" & vbCrLf & CStr(Err.Number) & vbCrLf & Err.Source & vbCrLf & Err.Description
End Sub

'************************Begin Note 1.28.2002 **************************

'*The following Block of code was obtained from http://www.mvps.org/vb/*
'The source code is free to use within any application as long as the actual
'uncompiled source code is not sold or distributed to other programmers.
' *********************************************************************
'  Copyright ©1995-2000 Karl E. Peterson, All Rights Reserved
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************

Public Function FindWindowPartial(AppTitle As String, _
   Optional Method As FindWindowPartialTypes = FwpStartsWith, _
   Optional CaseSensitive As Boolean = False, _
   Optional MustBeVisible As Boolean = False) As Long
   On Error GoTo EH
   
   m_hWnd = 0
   m_Method = Method
   m_CaseSens = CaseSensitive
   m_AppTitle = AppTitle
   
   '
   ' Upper-case search string if case-insensitive.
   '
   If m_CaseSens = False Then
      m_AppTitle = UCase$(m_AppTitle)
   End If
   '
   ' Fire off enumeration, and return m_hWnd when done.
   '
   Call EnumWindows(AddressOf EnumWindowsProc, MustBeVisible)
   
   FindWindowPartial = m_hWnd
   
   Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = ClassName & ": " & "Error in Public Function FindWindowPartial: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    On Error GoTo EH
   Static WindowText As String
   Static nRet As Long
   '
   ' Make sure we meet visibility requirements.
   '
   If lParam Then 'window must be visible
      If IsWindowVisible(hWnd) = False Then
         EnumWindowsProc = True
         Exit Function
      End If
   End If
   '
   ' Retrieve windowtext (caption)
   '
   WindowText = Space$(256)
   nRet = GetWindowText(hWnd, WindowText, Len(WindowText))
   If nRet Then
      '
      ' Clean up window text and prepare for comparison.
      '
      WindowText = left$(WindowText, nRet)
      If m_CaseSens = False Then
         WindowText = UCase$(WindowText)
      End If
      '
      ' Use appropriate method to determine if
      ' current window's caption either starts
      ' with, contains, or matches passed string.
      '
      Select Case m_Method
         Case FwpStartsWith
            If InStr(WindowText, m_AppTitle) = 1 Then
               m_hWnd = hWnd
            End If
         Case FwpContains
            If InStr(WindowText, m_AppTitle) <> 0 Then
               m_hWnd = hWnd
            End If
         Case FwpMatches
            If WindowText = m_AppTitle Then
               m_hWnd = hWnd
            End If
      End Select
   End If
   '
   ' Return True to continue enumeration if we haven't
   ' found what we're looking for.
   '
   EnumWindowsProc = (m_hWnd = 0)
   Exit Function
EH:
    mlErrNum = Err.Number
    msErrDesc = ClassName & ": " & "Error in Private Function EnumWindowsProc: " & " Desc: " & Err.Description
    msErrSrc = Err.Source
    Err.Raise mlErrNum, msErrSrc, msErrDesc
End Function
'************************End Note 1.28.2002 **************************

Public Function FormWinRegPos(psAppEXEName As String, pMyForm As Form, Optional pbSave As Boolean, _
                              Optional pfrmOffset As Form, Optional pctrlOffset As Control, _
                              Optional pbUseFullCaption As Boolean = True, Optional pbUseFrmName As Boolean) As Boolean
'Purpose: This Procedure can be used by AnyForm to Get or Save the Form Position
'         from the Windows Registry using Save Setting and GetSetting :)

'Parameters : pMyForm As Form, Optional pbSave As Boolean

'Returns: FormWinRegPos Returns True Only if  Retrieving and Finds Stored  Values
'         FormWinRegPos Returns False if Retreieving and does not find Stored Values
'         FormWinRegPos Returns False When Saveing IE pbSave is Set to True.


'Author : BGS - 3/10/2000

'Revision History:  SMR     Initials    Date        Description
'                     1     BGS         3/21/2000   Added Optional pfrOffset incase you want to Offset the posn
'                                                   in realation to another form.
'                     2     BGS         10/23/2001  Changed the SECTION to enter ALL forms under FORM_POSN
'                                                   Also check for Borderstyle do Change width or height on non sizable windows :)
                    
    'This Procedure will Either Retrieve or Save Form Posn values
    'Best used on Form Load and Unload or QueryUnLoad
    Dim sCap As String
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    On Error GoTo EH
    With pMyForm
        If Not pbUseFullCaption Then
            If pbUseFrmName Then
                sCap = left(.Name, InStr(1, .Caption, " "))
            Else
                sCap = left(.Caption, InStr(1, .Caption, " "))
            End If
            
        Else
            If pbUseFrmName Then
                sCap = .Name & " "
            Else
                sCap = .Caption & " "
            End If
            
        End If
        If pbSave Then
            'If Saving then do this...
            'If Form was minimized or Maximized then Closed Need to Save Windowstate
            'THEN... set Back to Normal Or previous non Max or Min State then Save
            'Posn Parameters
            
            SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_WindowState", .WindowState
            
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .Visible = True
                .WindowState = vbNormal
            End If
            
            'Save AppName...FrmName...KeyName...Value
            If pfrmOffset Is Nothing Then
                'Check to be sure windows didn't screw up and get set to something way off in la la land
                If .top < 0 Then .top = 0
                If .top > Screen.Height Then .top = Screen.Height - 100
                If .left < 0 Then .left = 0
                If .left > Screen.Width Then .left = Screen.Width - 100
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Top", .top
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Left", .left
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Width", .Width
            Else
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Top", .top - pfrmOffset.top - pctrlOffset.top
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Left", .left - pfrmOffset.left
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Height", .Height
                SaveSetting psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Width", .Width
            End If
        Else
            'If Not Saveing Must Be Getting ..
            'Need to ref AppName...FrmName...KeyName
            '(If nothing Stored Use The Exisiting Form value)
            If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
                .WindowState = vbNormal
            End If
            If pfrmOffset Is Nothing Then
                .top = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Top", .top)
                .left = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Left", .left)
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            Else
                .top = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Top", .top) + pfrmOffset.top + pctrlOffset.top
                .left = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Left", .left) + pfrmOffset.left
                If .BorderStyle = vbSizable Or .BorderStyle = vbSizableToolWindow Then
                    .Height = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Height", .Height)
                    .Width = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_Width", .Width)
                End If
                'Be Sure WindowState is set last (Can't Change POSN if vbMinimized Or Maximized)
                .WindowState = GetSetting(psAppEXEName, gfrmMergePDF.Caption & "_FORM_POSN", .Name & sCap & "_WindowState", .WindowState)
            End If
        End If
    End With

    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Function FormWinRegPos" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function


Public Function GetTaskbarHeight() As Long
    On Error GoTo EH
    Dim lRes As Long
    Dim rectVal As RECT
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
    Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Function GetTaskbarHeight" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

Public Function GetPath(psAppEXEName As String, psName As String, psMess As String, psFileMess As String, psDeFaultPath As String, plHwnd As Long, _
                         Optional psFilter As String = vbNullString, _
                         Optional psSelFile As String, _
                         Optional plFlags As Long, _
                         Optional pbCenterForm As Boolean = True, _
                         Optional pbShowOpen As Boolean = True) As String
    On Error GoTo EH
    Dim sOpen As SelectedFile
    Dim sDir As String
    Dim sMyFilter As String
    Dim sInitDir As String
    Dim sNewCatName As String
    Dim MyFileDialog As OPENFILENAME
    Dim lErrNum As Long
    Dim sErrDesc As String
    
    FileDialog = MyFileDialog
    sMyFilter = psFilter
    FileDialog.sFilter = sMyFilter
    
    ' See Standard CommonDialog Flags for all options
    If plFlags > 0 Then
        FileDialog.flags = plFlags
    Else
        FileDialog.flags = OFN_HIDEREADONLY Or OFN_NOVALIDATE
    End If
    FileDialog.sDlgTitle = psMess
    FileDialog.sFile = psFileMess
    If goUtil.utFileExists(psDeFaultPath, False) Or goUtil.utFileExists(psDeFaultPath, True) Then
        sInitDir = psDeFaultPath
    Else
        sInitDir = GetSetting(psAppEXEName, "Dir", psName, "Error")
    End If
    
    If sInitDir <> "Error" And sInitDir <> vbNullString Then
        FileDialog.sInitDir = sInitDir
    Else
        FileDialog.sInitDir = "C:\"
    End If
    
    If pbShowOpen Then
        sOpen = ShowOpen(plHwnd, pbCenterForm)
    Else
        sOpen = ShowSave(plHwnd, pbCenterForm)
    End If
    If Err.Number <> 32755 And sOpen.bCanceled = False Then
        sDir = left(sOpen.sLastDirectory, InStrRev(sOpen.sLastDirectory, "\"))
        If sDir = vbNullString And sOpen.nFilesSelected = 1 Then
            sDir = left(sOpen.sFiles(1), InStrRev(sOpen.sFiles(1), "\"))
        End If
        'Set the selected file
        psSelFile = Replace(sOpen.sFiles(1), sDir, vbNullString)
        SaveSetting psAppEXEName, "Dir", psName, sDir
        GetPath = sDir
    Else
        GetPath = psDeFaultPath
    End If
    
   Exit Function
EH:
    lErrNum = Err.Number
    sErrDesc = Err.Description
    Err.Raise lErrNum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Function GetPath" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
End Function

'Public Sub EnterUserPass(psSection As String, psUserName As String, pvPass As Variant _
'                         , Optional poForm As Object _
'                         , Optional psUpdateOldUserName As String _
'                         , Optional pbDoSave As Boolean = False _
'                         , Optional psSaveUserName As String _
'                         , Optional psSavePassword As String _
'                         , Optional ByRef poUseMyIcon As Object)
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'
'    Dim sCryptUserName As String
'    Dim sCryptPass As String
'    Dim sRet As String
'    Dim sSectionOld As String
'    Dim sSectionNew As String
'    Dim vAry As Variant
'    Dim lPos As Long
'    Dim lSec As Long
'    Dim sKeyName As String
'    Dim sKeyData As String
'
'    If pbDoSave Then
'        If psUpdateOldUserName <> vbNullString Then
'            For lSec = 1 To 3
'                Select Case lSec
'                    Case 1
'                        sSectionOld = psUpdateOldUserName & "_SECURITY"
'                        sSectionNew = psUserName & "_SECURITY"
'                    Case 2
'                        sSectionOld = psUpdateOldUserName & "_GENERAL"
'                        sSectionNew = psUserName & "_GENERAL"
'                    Case 3
'                         sSectionOld = psUpdateOldUserName & "_FORM_POSN"
'                         sSectionNew = psUserName & "_FORM_POSN"
'                End Select
'
'                vAry = GetAllSettings(App.EXEName, sSectionOld)
'
'                For lPos = LBound(vAry, 1) To UBound(vAry, 1)
'                    'Copy Settings over to New Name
'                    sKeyName = vAry(lPos, 0)
'                    sKeyData = vAry(lPos, 1)
'                    SaveSetting App.EXEName, sSectionNew, sKeyName, sKeyData
'
'                    'DeleteSetting from Old name
'                    DeleteSetting App.EXEName, sSectionOld, sKeyName
'                Next lPos
'
'                'Delete Main Section Key is handled inside:
'                'Remove User Name from USERS (goUtil.utRemoveUserAccount)
'
'            Next lSec
'            'Remove User Name from USERS
'            goUtil.utRemoveUserAccount psUpdateOldUserName
'        End If
'    End If
'
'    'If pass is control that has text then we are
'    'prompting for confirmation of the old password first if it exists
'    'and then re enter the new password
'    If IsObject(pvPass) And Not pbDoSave Then
'        'security
'        goUtil.MySeed = Abs(CLng(goUtil.utGetTickCount)) * -1
'        sCryptPass = GetCryptSetting(App.EXEName, psSection, "PASSWORD")
'        If sCryptPass <> vbNullString Then
'            If Not poForm Is Nothing Then
'                'PUt input box top left of form
'                sRet = MyInputBox("Please enter old password.", "OLD PASSWORD", , poForm.left, poForm.top, "*", poUseMyIcon)
'                If sRet = vbNullString Or Len(sRet) > 50 Then
'                    If Len(sRet) > 50 Then
'                        MsgBox "The password you entered exceeds the maximum number of characters allowed(50).", vbOKOnly + vbExclamation, "INVALID PASSWORD"
'                    End If
'                    'they clicked on cancel
'                    GoTo CLEANUP
'                End If
'            Else
'                'Put input box default windows pos
'                sRet = MyInputBox("Please enter old password.", "OLD PASSWORD", , , , "*", poUseMyIcon)
'                If sRet = vbNullString Or Len(sRet) > 50 Then
'                    If Len(sRet) > 50 Then
'                        MsgBox "The password you entered exceeds the maximum number of characters allowed(50).", vbOKOnly + vbExclamation, "INVALID PASSWORD"
'                    End If
'                    'they clicked on cancel
'                    GoTo CLEANUP
'                End If
'            End If
'            'check sret against the old password
'            If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
'                MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
'                GoTo CLEANUP
'            Else
'                GoTo ENTER_NEWPASS
'            End If
'
'        Else
'            'if there is no password saved yet then just ask for
'            'Password
'ENTER_NEWPASS:
'            sRet = MyInputBox("Please enter a new password.", "ENTER NEW PASSWORD", , poForm.left, poForm.top, "*", poUseMyIcon)
'            If sRet = vbNullString Then
'                GoTo CLEANUP
'            Else
'                sCryptPass = sRet
'                'Ask them to double check the password they just entered
'                sRet = MyInputBox("Please enter the same password again.", "ENTER PASSWORD AGAIN", , poForm.left, poForm.top, "*", poUseMyIcon)
'                If sRet = vbNullString Or Len(sRet) > 50 Then
'                    If Len(sRet) > 50 Then
'                        MsgBox "The password you entered exceeds the maximum number of characters allowed(50).", vbOKOnly + vbExclamation, "INVALID PASSWORD"
'                    End If
'                    GoTo CLEANUP
'                Else
'                    If StrComp(sCryptPass, sRet, vbBinaryCompare) <> 0 Then
'                        MsgBox "The password you entered does not match.", vbOKOnly + vbExclamation, "INCORRECT PASSWORD"
'                        GoTo CLEANUP
'                    End If
'                End If
'            End If
'
'            'Disaplay password
'            pvPass.Text = sCryptPass
'        End If
'    Else
'        sCryptPass = psSavePassword
'    End If
'    'set the user Name
'    sCryptUserName = psSaveUserName
'
'    If sCryptPass <> vbNullString Then
'        If pbDoSave Then
'            'Security
'            goUtil.MySeed = Abs(CLng(goUtil.utGetTickCount)) * -1
'            SaveCryptSetting App.EXEName, psSection, "PASSWORD", sCryptPass
'        End If
'    End If
'    If sCryptUserName <> vbNullString Then
'        If pbDoSave Then
'            'Security
'            goUtil.MySeed = Abs(CLng(goUtil.utGetTickCount)) * -1
'            SaveCryptSetting App.EXEName, psSection, "USER_NAME", sCryptUserName
'        End If
'    End If
'
'    If gfrmMergePDF.Caption = vbNullString Or StrComp(gfrmMergePDF.Caption, psUserName, vbTextCompare) <> 0 Then
'        If pbDoSave Then
'            goUtil.utAddUserAccount psUserName
'        End If
'    End If
'    If pbDoSave Then
'        gfrmMergePDF.Caption = psSaveUserName
'        goUtil.MySeed = Abs(CLng(goUtil.utGetTickCount)) * -1
'        gfrmMergePDF.Password = goUtil.Encode(psSavePassword)
'    End If
'CLEANUP:
'
'    Exit Sub
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lErrNum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Sub EnterPass" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
'End Sub

'Public Function MyInputBox(ByVal psPrompt As String _
'                            , ByVal psTitle As String _
'                            , Optional ByVal psDefault As String = vbNullString _
'                            , Optional ByVal pXPos As Long = 0 _
'                            , Optional ByVal pYPos As Long = 0 _
'                            , Optional ByVal psPassWordChar As String = vbNullString _
'                            , Optional ByRef poUseMyIcon As Object) As String
'    On Error GoTo EH
'    Dim lErrNum As Long
'    Dim sErrDesc As String
'    Dim ofrmInputBox As frmInputBox
'
'
'    Set ofrmInputBox = New frmInputBox
'
'    Load ofrmInputBox
'    With ofrmInputBox
'        .Caption = psTitle
'        .lblPrompt.Caption = psPrompt
'        .txtInput.Text = psDefault
'        .txtInput.PasswordChar = psPassWordChar
'        If pXPos = 0 And pYPos = 0 Then
'            .left = Screen.Width / 2 - ofrmInputBox.Width / 2
'            .top = Screen.Height / 2 - ofrmInputBox.Height / 2
'        Else
'            .left = pXPos
'            .top = pYPos
'        End If
'        If Not poUseMyIcon Is Nothing Then
'            If TypeOf poUseMyIcon Is Form Then
'                .Icon = poUseMyIcon.Icon
'            ElseIf TypeOf poUseMyIcon Is CommandButton Then
'                .Icon = poUseMyIcon.Picture
'            End If
'        End If
'    End With
'    ofrmInputBox.Show vbModal
'
'    MyInputBox = ofrmInputBox.txtInput.Text
'
'    Unload ofrmInputBox
'    Set ofrmInputBox = Nothing
'
'    Exit Function
'EH:
'    lErrNum = Err.Number
'    sErrDesc = Err.Description
'    Err.Raise lErrNum, , App.EXEName & vbCrLf & ClassName & vbCrLf & "Public Function MyInputBox" & vbCrLf & "Error # " & lErrNum & vbCrLf & sErrDesc & vbCrLf
'End Function

