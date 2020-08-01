Attribute VB_Name = "modMergePDF"
'  http://www.xlsure.com 2020.07.30
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
'  Merge PDF Files - modMergePDF
' *********************************************************************

Option Explicit
'PDF documents must be declared in general declaration space and not local!
Private moMainDoc As Acrobat.AcroPDDoc
Private moTempDoc As Acrobat.AcroPDDoc

Private Property Get msClassName() As String
    msClassName = "modMergePDF"
End Property

Public Function MergePDFFiles(psRawPDFFilesDir As String, _
                                psSinglePDFOutputDir As String, _
                                psSinglePDFOutputName As String, _
                                Optional ByVal pbRemovePdfExtFromBookMark As Boolean = True, _
                                Optional pbCaseSensitiveSort As Boolean = False, _
                                Optional ByVal pbShowError As Boolean = False) As Boolean
    On Error GoTo EH
    
    Dim bFirstDoc As Boolean
    Dim sRawPDFFilesDir As String
    Dim sSinglePDFOutputDir As String
    Dim sSinglePDFOutputName As String
    Dim saryFileSort() As String
    Dim sBMName As String
    'Track pos of things
    Dim lBMPageNo As Long
    Dim lPos As Long
    Dim lFile As Long
    Dim lInsertPageAfter As Long
    Dim lNumPages As Long
    Dim lRet As Long
    'Need to use Adobe internal Java Object
    'in order to Add Book marks
    Dim oJSO As Object 'JavaScript Object
    Dim oBookMarkRoot As Object
    'File I/O
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim oFSO As Scripting.FileSystemObject
    
    
    sRawPDFFilesDir = psRawPDFFilesDir
    sSinglePDFOutputDir = psSinglePDFOutputDir
    sSinglePDFOutputName = psSinglePDFOutputName
    
    Set oFSO = New Scripting.FileSystemObject
    
    Set oFolder = oFSO.GetFolder(sRawPDFFilesDir)
    
    bFirstDoc = True

    If oFolder.Files.Count = 0 Then
        Exit Function
    End If
    
    'Because the FSO folder files collection does not allow for
    'Native sorting, need to plug all the files into an array and sort that motha
    ReDim saryFileSort(1 To oFolder.Files.Count)
    lFile = 0
    For Each oFile In oFolder.Files
        lFile = lFile + 1
        saryFileSort(lFile) = oFile.Name
    Next
    
    'Once they is all in der sor the array
    'Sort is Case Sensitive
    If pbCaseSensitiveSort Then
        goUtil.utBubbleSort saryFileSort
    End If
    
    For lFile = 1 To UBound(saryFileSort, 1)
        If LCase(Right(saryFileSort(lFile), 4)) = ".pdf" Then
            If bFirstDoc Then
                bFirstDoc = False
                Set moMainDoc = CreateObject("AcroExch.PDDoc") 'New AcroPDDoc
                lRet = moMainDoc.Open(sRawPDFFilesDir & saryFileSort(lFile))
                Set oJSO = moMainDoc.GetJSObject
                Set oBookMarkRoot = oJSO.BookMarkRoot
                sBMName = saryFileSort(lFile)
                lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                If lPos > 0 Then
                    sBMName = left(sBMName, lPos - 1) & ".pdf"
                End If
                If pbRemovePdfExtFromBookMark Then
                    sBMName = Replace(sBMName, ".pdf", vbNullString, , , vbTextCompare)
                End If
                lRet = oBookMarkRoot.CreateChild(sBMName, "this.pageNum =0", lFile - 1)
            Else
                Set moTempDoc = CreateObject("AcroExch.PDDoc") 'New AcroPDDoc
                lRet = moTempDoc.Open(sRawPDFFilesDir & saryFileSort(lFile))
                'get the Book mark page number before the actual instert of new pages
                lBMPageNo = moMainDoc.GetNumPages
                lInsertPageAfter = lBMPageNo - 1
                lNumPages = moTempDoc.GetNumPages
                lRet = moMainDoc.InsertPages(lInsertPageAfter, moTempDoc, 0, lNumPages, 0)
                moTempDoc.Close
                If lRet = 0 Then
                    sBMName = saryFileSort(lFile)
                    lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                    If lPos > 0 Then
                        sBMName = left(sBMName, lPos - 1) & ".pdf"
                    End If
                    'Need to copy the errored document over to be included in the enitre document
                    goUtil.utCopyFile sRawPDFFilesDir & saryFileSort(lFile), sSinglePDFOutputDir & "\" & sBMName
                    sBMName = "PDF Insert Page Error_" & sBMName
                Else
                    sBMName = saryFileSort(lFile)
                    lPos = InStr(1, sBMName, "_{", vbBinaryCompare)
                    If lPos > 0 Then
                        sBMName = left(sBMName, lPos - 1) & ".pdf"
                    End If
                End If
                If pbRemovePdfExtFromBookMark Then
                    sBMName = Replace(sBMName, ".pdf", vbNullString, , , vbTextCompare)
                End If
                lRet = oBookMarkRoot.CreateChild(sBMName, "this.pageNum =" & lBMPageNo, lFile - 1)
            End If
        End If
    Next
    
    lRet = moMainDoc.Save(1, sSinglePDFOutputDir & "\" & sSinglePDFOutputName)
    moMainDoc.Close
    
    MergePDFFiles = True
    
CLEAN_UP:
    Set oFolder = Nothing
    Set oFile = Nothing
    Set oFSO = Nothing
    Set oBookMarkRoot = Nothing
    Set oJSO = Nothing
    Set moMainDoc = Nothing
    Set moTempDoc = Nothing
    
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function MergePDFFiles", pbShowError
End Function

Public Function BuildBatchFileCode(psRawPDFFilesDir As String, _
                                    psSinglePDFOutputDir As String, _
                                    psSinglePDFOutputName As String, _
                                    pbRemovePdfExtFromBookMark As Boolean, _
                                    pbCaseSensitiveSort As Boolean) As String
    
    On Error GoTo EH
    
    Dim sRawPDFFilesDir As String: sRawPDFFilesDir = psRawPDFFilesDir
    Dim sSinglePDFOutputDir As String: sSinglePDFOutputDir = psSinglePDFOutputDir
    Dim sSinglePDFOutputName As String: sSinglePDFOutputName = psSinglePDFOutputName
    Dim bRemovePdfExtFromBookMark As Boolean: bRemovePdfExtFromBookMark = pbRemovePdfExtFromBookMark
    
    Dim sCommandLine As String
    
    sCommandLine = "RawPDFFilesDir|" & sRawPDFFilesDir _
                    & "|SinglePDFOutputDir|" & sSinglePDFOutputDir _
                    & "|SinglePDFOutputName|" & sSinglePDFOutputName _
                    & "|RemovePdfExtFromBookMark|" & CStr(bRemovePdfExtFromBookMark) _
                    & "|CaseSensitiveSort|" & CStr(pbCaseSensitiveSort)

    BuildBatchFileCode = """" & App.Path & "\" & App.EXEName & ".exe"" """ & sCommandLine
    Exit Function
EH:
    goUtil.utErrorLog Err, App.EXEName, msClassName, "Public Function BuildBatchFileCode"
End Function

