Attribute VB_Name = "modStartMain"
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - modStartMain
' *********************************************************************
Option Explicit

'Password Special Chars List
Public Const PASS_SPCL_CHARS = "!@#$%^&*()_-+=<,>.?/:;{[}]"

'Pass this one Global Object between Apps
Public goUtil As clsUtil
Public gfrmMergePDF As frmMergePDF

Private Property Get msClassName() As String
    msClassName = "modStartMain"
End Property

Public Sub Main()
    On Error GoTo EH
    Dim vAryCommand As Variant
    Dim sRawPDFFilesDir As String
    Dim sSinglePDFOutputDir As String
    Dim sSinglePDFOutputName As String
    Dim bRemovePdfExtFromBookMark As Boolean
    Dim bCaseSensitiveSort As Boolean
    
    Dim sMess As String
    'Set Public Objects Here
    Set goUtil = New clsUtil
    
    'Split up the Command string if one is sent.  MergePDF is being shelled to execute without GUI.
    If InStr(1, Command$, "RawPDFFilesDir", vbTextCompare) > 0 Then
        vAryCommand = Split(Command$, "|")
        If IsArray(vAryCommand) Then
            sRawPDFFilesDir = vAryCommand(1)
            If InStr(1, Command$, "SinglePDFOutputDir", vbTextCompare) > 0 Then
                sSinglePDFOutputDir = vAryCommand(3)
            End If
            If InStr(1, Command$, "SinglePDFOutputName", vbTextCompare) > 0 Then
                sSinglePDFOutputName = vAryCommand(5)
            End If
            'Default removing the .pdf from the bookmark names
            bRemovePdfExtFromBookMark = True
            If InStr(1, Command$, "RemovePdfExtFromBookMark", vbTextCompare) > 0 Then
                bRemovePdfExtFromBookMark = CBool(vAryCommand(7))
            End If
            'Default Sort to NOT be Case Sensitive
            bCaseSensitiveSort = False
            If InStr(1, Command$, "CaseSensitiveSort", vbTextCompare) > 0 Then
                bCaseSensitiveSort = CBool(vAryCommand(9))
            End If
            
            'Build the batch file code
            
            'Merge the many PDFs to a single PDF using each individual pdf file name as a bookmark name.
            modMergePDF.MergePDFFiles sRawPDFFilesDir, sSinglePDFOutputDir, sSinglePDFOutputName, bRemovePdfExtFromBookMark, bCaseSensitiveSort
        Else
            End
            Exit Sub
        End If
    Else
        'If no command string then run the GUI
        Set gfrmMergePDF = New frmMergePDF
        Load gfrmMergePDF
        gfrmMergePDF.Show vbModeless
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Sub Main"
End Sub

