VERSION 5.00
Begin VB.Form frmMergePDF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge PDF Files"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   Icon            =   "frmMergePDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCaseSensitiveSort 
      Caption         =   "Case Sensitive Sort"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton cmdBuildBatchCodeOnly 
      Caption         =   "&Build Batch Code Only"
      Height          =   495
      Left            =   10440
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtBatchFileCode 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2760
      Width           =   12375
   End
   Begin VB.CheckBox chkRemovePDFextFromBookMarkNames 
      Caption         =   "Remove .pdf From Book Mark Names"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "&Merge"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdMergedPDFFileNamePath 
      Height          =   330
      Left            =   120
      Picture         =   "frmMergePDF.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Browse to Merged PDF File Name Path"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtMergedPDFFileNamePath 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   12135
   End
   Begin VB.TextBox txtPDFFilesToBeMergedDirectory 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   12135
   End
   Begin VB.CommandButton cmdPDFFilesToBeMergedDirectory 
      Height          =   330
      Left            =   120
      Picture         =   "frmMergePDF.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Browse to PDF Files To Be Merged Directory"
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblMergedPDFFileNamePath 
      Caption         =   "Merged PDF File Name Path"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Label lblPDFFilesToBeMergedDirectory 
      Caption         =   "PDF Files To Be Merged Directory"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "frmMergePDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - frmMergePDF
' *********************************************************************

Option Explicit

Private Property Get msClassName() As String
    msClassName = Me.Name
End Property

Private Sub cmdBuildBatchCodeOnly_Click()
    DoMerge True
End Sub

Private Sub cmdMerge_Click()
    DoMerge
End Sub

Private Sub DoMerge(Optional pBuildBatchFileCodeOnly As Boolean = False)
    On Error GoTo EH
    Dim sRawPDFFilesDir As String: sRawPDFFilesDir = Trim(txtPDFFilesToBeMergedDirectory.Text)
    Dim sSinglePDFOutputName As String: sSinglePDFOutputName = Trim(txtMergedPDFFileNamePath.Text)
    Dim sSinglePDFOutputDir As String: sSinglePDFOutputDir = sSinglePDFOutputName
    Dim bRemovePdfExtFromBookMark As Boolean: bRemovePdfExtFromBookMark = (chkRemovePDFextFromBookMarkNames.Value = VBRUN.CheckBoxConstants.vbChecked)
    Dim bCaseSensitiveSort As Boolean: bCaseSensitiveSort = (chkCaseSensitiveSort.Value = VBRUN.CheckBoxConstants.vbChecked)
    
    sSinglePDFOutputName = goUtil.utGetFileName(sSinglePDFOutputDir)
    sSinglePDFOutputDir = goUtil.utGetFilePath(sSinglePDFOutputDir)
    
    If Not goUtil.utFileExists(sRawPDFFilesDir, True) Then
        MsgBox App.EXEName & vbCrLf & msClassName & vbCrLf & "PDF Files To Be Merged Directory Is INVALID!", vbCritical
        Exit Sub
    End If
    
    If Not goUtil.utFileExists(sSinglePDFOutputDir, True) Then
        MsgBox App.EXEName & vbCrLf & msClassName & vbCrLf & "Merged PDF File Name Path Is INVALID!", vbCritical
        Exit Sub
    End If
    
    If Not pBuildBatchFileCodeOnly Then
        modMergePDF.MergePDFFiles sRawPDFFilesDir, sSinglePDFOutputDir, sSinglePDFOutputName, bRemovePdfExtFromBookMark, bCaseSensitiveSort, True
        
        'Show the directory
        goUtil.utShellExecute Me.hWnd, , sSinglePDFOutputDir & sSinglePDFOutputName
    End If
    
    'Build the Batch File Code
    txtBatchFileCode.Text = modMergePDF.BuildBatchFileCode(sRawPDFFilesDir, sSinglePDFOutputDir, sSinglePDFOutputName, bRemovePdfExtFromBookMark, bCaseSensitiveSort)
    
Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub DoMerge"
End Sub

Private Sub cmdPDFFilesToBeMergedDirectory_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    
    sMyFilter = sMyFilter & "PDF (*.pdf)" & SD & "*.pdf" & SD
   
    sPath = goUtil.utGetPath(App.EXEName, "PDFFilePath", "Browse to the PDF Files Directory to be Merged and select the top pdf file.", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If goUtil.utFileExists(sPath & sSelFile) Then
        txtPDFFilesToBeMergedDirectory.Text = sPath '& sSelFile
    Else
        txtPDFFilesToBeMergedDirectory.Text = vbNullString
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdPDFFilesToBeMergedDirectory_Click"
End Sub

Private Sub cmdMergedPDFFileNamePath_Click()
    On Error GoTo EH
    Dim sPath As String
    Dim sSelFile As String
    Dim lFileSize As Long
    Dim sMyFilter As String
    
    sMyFilter = sMyFilter & "All (*.*)" & SD & "*.*" & SD
    
    sPath = goUtil.utGetPath(App.EXEName, "PDFFilePath", "Browse to the directory you want to merge pdfs and create the single PDF file name.", "", sPath, Me.hWnd, sMyFilter, sSelFile)
    
    If goUtil.utFileExists(sPath & sSelFile) Then
        txtMergedPDFFileNamePath.Text = sPath & sSelFile
    Else
        txtMergedPDFFileNamePath.Text = vbNullString
    End If
    
    Exit Sub
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Private Sub cmdMergedPDFFileNamePath_Click"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    goUtil.CLEANUP
End Sub

Private Sub txtBatchFileCode_GotFocus()
    goUtil.utSelText txtBatchFileCode
End Sub

Private Sub txtMergedPDFFileNamePath_GotFocus()
    goUtil.utSelText txtMergedPDFFileNamePath
End Sub

Private Sub txtPDFFilesToBeMergedDirectory_GotFocus()
    goUtil.utSelText txtPDFFilesToBeMergedDirectory
End Sub
