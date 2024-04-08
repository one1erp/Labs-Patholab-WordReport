VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const csHeBrEw As String = "iso-8859-8" ' Hebrew character set
Private Const wdDoNotSaveChanges = 0
Private Const wdFormatDocument = 0
Private Const wdPrintAllDocument = 0
Private Const wdPrintDocumentContent = 0
Private Const wdPrintAllPages = 0
Private Const wdFindContinue = 1
Private Const wdFindStop = 0
Private Const wdReplaceOne = 1
Private Const wdReplaceAll = 2
Private Const wdSeekCurrentPageFooter = 10
Private Const wdSeekCurrentPageHeader = 9
Private Const wdSeekMainDocument = 0
Private Const wdPaneNone = 0
Private Const wdNormalView = 1
Private Const wdOutlineView = 2
Private Const wdPrintView = 3
Private Const wdToggle = &H98967E
Private Const wdStory = 6
Private Const wdGoToPage = 1
Private Const wdPageBreak = 7
Private Const wdLineBreak = 6
Private Const wdNumberOfPagesInDocument = 4
Private Const wdActiveEndAdjustedPageNumber = 1

Private WordAppl As Object 'Word.Application
Private Doc As Object
' Private Doc As Word.Document
Private mWordGuard As WordGuard.Guard

Private con As ADODB.Connection
Private queries As New Scripting.Dictionary
Private Parameters As New Scripting.Dictionary
Private DocumentReady As Boolean
'Private DocOpen As Boolean
Private mstErrMsg As String
Private mloErrNbr As Long

Private Sub Class_Initialize()
    
    On Error GoTo Fin

    Call StartWordAppl

Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: Class_Initialize" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
'    MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
'    If DocOpen Then
    If Not Doc Is Nothing Then
        Doc.Close wdDoNotSaveChanges
        DoEvents
        Set Doc = Nothing
    End If
    
    WordAppl.Quit wdDoNotSaveChanges
    DoEvents
    Set WordAppl = Nothing
    
    Set mWordGuard = Nothing
    DoEvents

Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: Class_Terminate" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub

Public Sub StartWordAppl()
    
    On Error GoTo Fin
    
    Set mWordGuard = Nothing
    DoEvents
    
    If Not WordAppl Is Nothing Then
        On Error Resume Next
        WordAppl.Quit wdDoNotSaveChanges
        DoEvents
        Set WordAppl = Nothing
        On Error GoTo Fin
    End If
    
    Set WordAppl = CreateObject("Word.Application") 'New Word.Application
    DoEvents
    
    Set mWordGuard = New WordGuard.Guard
    DoEvents
    
    WordAppl.Caption = "One1 - Automatic Word report tool :)"
    DoEvents
    
    Call mWordGuard.AssignAppl("OpusApp", WordAppl.Caption)
    DoEvents
    
    WordAppl.Caption = Empty
    DoEvents

'    Set Doc = Nothing
'    DocOpen = False
    mstErrMsg = ""
    mloErrNbr = 0
Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: StartWordAppl" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"

End Sub

Public Property Set Connection(aConnection As ADODB.Connection)
    Set con = aConnection
End Property

Public Property Get ErrMsg() As String
    ErrMsg = mstErrMsg
End Property

Public Property Get ErrNbr() As Long
    ErrNbr = mloErrNbr
End Property

Public Sub AddQuery(name As String, query As String)
    Call queries.Add(name, query)
End Sub

Public Sub AddParameter(name As String, param As String)
    Call Parameters.Add(name, param)
End Sub

Public Sub SetTemplate(Temp As String)
    
' Begin a new report
    
    On Error GoTo ErrHnd
    
'    queries.RemoveAll
'    Parameters.RemoveAll
    DocumentReady = False
    
    mWordGuard.StartProcessDoc
    
    If Not Doc Is Nothing Then
        On Error Resume Next ' if word is not active
        Doc.Close wdDoNotSaveChanges
        DoEvents
        Set Doc = Nothing
        On Error GoTo ErrHnd
    End If
    Set Doc = WordAppl.Documents.Add()
    Doc.ActiveWindow.View.Type = wdPrintView
'    Set Doc = CreateObject("Word.Document") 'New Word.Document
'    DocOpen = True
    DocumentReady = False
    
'    Dim docApp As Object 'Word.Application
'    Set docApp = Doc.Application
    
    Doc.Range(0, 0).Select
    WordAppl.Selection.InsertFile FileName:=Temp, Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
    Doc.Range(0, 0).Select
    Exit Sub
ErrHnd:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: SetTemplate" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & vbCrLf & _
                "Temp = " & Temp
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select

End Sub

Public Sub SaveReport(pstQueryFileName As String)
    
' Save the report to given file name or qury
    
    Dim j As Integer
    Dim lstFileName As String
    Dim lstQuery As String
    
    Dim lrsFileName As ADODB.Recordset

    On Error GoTo Fin
    If Not DocumentReady Then FillReport
    
    If nte(pstQueryFileName) = "" Then Exit Sub
    
    If Left(Trim(UCase(pstQueryFileName)), 6) = "SELECT" Then
        lstQuery = pstQueryFileName
    
        For j = 0 To Parameters.Count - 1
            lstQuery = Replace(lstQuery, "#" & Parameters.Keys(j) & "#", Parameters.Items(j))
        Next j
        
        Set lrsFileName = con.Execute(lstQuery)
        
        lstFileName = nte(lrsFileName(0))
        
        lrsFileName.Close
        Set lrsFileName = Nothing
    Else
        lstFileName = pstQueryFileName
    End If
    
    If IsNull(lstFileName) Then Exit Sub
    'this wil work on word 2003 and less.
  '  Doc.SaveAs FileName:=lstFileName, FileFormat:=wdFormatDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        
    'this wil work on word 2007 and above.
    Doc.SaveAs FileName:=lstFileName, FileFormat:=wdFormatPDF, _
        LockComments:=False, Password:="", AddToRecentFiles:=False, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False
        
    Exit Sub
Fin:
    If Err.Number = 5152 Then
        Err.Clear
        Exit Sub
    End If
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: SaveReport , QueryFileName= " & pstQueryFileName & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub

Public Sub PrintReport(PrinterDevice As String, pstCopiesQuery As String)

' Print the report

    Dim j As Integer
    Dim linCopies As Integer
    Dim lstQuery As String
    
    Dim lrsCopies As ADODB.Recordset

    On Error GoTo Fin
    If Not DocumentReady Then FillReport
    
    
    If Not IsNumeric(pstCopiesQuery) Then
        lstQuery = pstCopiesQuery
    
        For j = 0 To Parameters.Count - 1
            lstQuery = Replace(lstQuery, "#" & Parameters.Keys(j) & "#", Parameters.Items(j))
        Next j
    
        Set lrsCopies = con.Execute(lstQuery)
        
        If lrsCopies.EOF Then
            linCopies = 3
        Else
            If IsNull(lrsCopies(0)) Then
                linCopies = 0
            Else
                linCopies = CInt(lrsCopies(0))
            End If
            'linCopies = IIf(IsNull(lrsCopies(0)), 0, CInt(lrsCopies(0)))
        End If
        
        lrsCopies.Close
        Set lrsCopies = Nothing
    
    Else
        linCopies = CInt(pstCopiesQuery)
    End If
    
    
    If PrinterDevice <> "" Then
        Doc.Application.ActivePrinter = PrinterDevice
    End If
    
    '20.09.2006: do not print at all for zero number of copies:
    If linCopies = 0 Then Exit Sub
'    If linCopies = 0 Then linCopies = 1
    
    Doc.Application.PrintOut Copies:=linCopies, Background:=False
    Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: PrintReport , PrinterDevice= " & PrinterDevice & vbCrLf & "CopiesQuery= " & pstCopiesQuery & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub

Private Sub FillReport()
    
' Replace the queries fields on the report report
' For each query:
' 1. call QueryToDictionary function to get the data to a dictionary
' 2. use the InsertFieldValue/InsertRTF functions to insert the data to the Report
    
    Dim i As Integer
    Dim j As Integer
    Dim query As String
    Dim qdic As New Scripting.Dictionary
    Dim qdicMulti As New Scripting.Dictionary
'    Dim docApp As Object 'Word.Application
    
    On Error GoTo Fin
    For i = 0 To queries.Count - 1
        query = queries.Items(i)
        For j = 0 To Parameters.Count - 1
            query = Replace(query, "#" & Parameters.Keys(j) & "#", Parameters.Items(j))
        Next j
        qdic.RemoveAll
        qdicMulti.RemoveAll
        Call QueryToDictionary(query, qdic, qdicMulti)
        For j = 0 To qdicMulti.Count - 1
            If Mid(qdicMulti.Items(j), 1, 5) = "[RTF]" Then
                Call InsertRTF(queries.Keys(i) & "_" & qdicMulti.Keys(j), _
                    Mid(qdicMulti.Items(j), 6))
            Else
                Call InsertFieldValue(queries.Keys(i) & "_" & qdicMulti.Keys(j), _
                    CStr(nte(qdicMulti.Items(j))))
            End If
        Next j
        For j = 0 To qdic.Count - 1
            If Mid(qdic.Items(j), 1, 5) = "[RTF]" Then
                Call InsertRTF(queries.Keys(i) & "_" & qdic.Keys(j), _
                    Mid(qdic.Items(j), 6))
            Else
                Call InsertFieldValue(queries.Keys(i) & "_" & qdic.Keys(j), _
                    CStr(nte(qdic.Items(j))))
            End If
        Next j
    Next i
        
    Call FormatReport("</B>")
    Call FormatReport("</I>")
    Call FormatReport("</S08>")
    Call FormatReport("</S10>")
    Call FormatReport("</S12>")
    Call FormatReport("</S14>")
    Call FormatReport("</S16>")
    Call FormatReport("</S18>")
    Call FormatReport("LTRLTR")
    Call FormatReport("ילשילש")
    
    WordAppl.Selection.Find.ClearFormatting
    WordAppl.Selection.Find.Replacement.ClearFormatting
    With WordAppl.Selection.Find
        .Text = "<BR>"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    While WordAppl.Selection.Find.Execute
        WordAppl.Selection.TypeParagraph
    Wend
    
    PageBreaks
    
    DocumentReady = True
    qdic.RemoveAll
    qdicMulti.RemoveAll
    Set qdic = Nothing
    Set qdicMulti = Nothing
    Exit Sub
Fin:
    qdic.RemoveAll
    qdicMulti.RemoveAll
    Set qdic = Nothing
    Set qdicMulti = Nothing
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: FillReport " & vbCrLf & "Error Number:" & Err.Number & vbCrLf & _
                "Description" & Err.Description & vbCrLf & "query = " & query
    mloErrNbr = Err.Number
    Err.Raise Err.Number

'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub
Private Sub FormatReport(pFormatString As String)
    
    Dim lboFormatFound
    Dim lloStart As Long
    Dim lloEnd As Long
    
    lboFormatFound = True
    
    While lboFormatFound
        WordAppl.Selection.Find.ClearFormatting
        WordAppl.Selection.Find.Replacement.ClearFormatting
        With WordAppl.Selection.Find
            .Text = pFormatString
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        If InStr("LTRLTRילשילש", pFormatString) = 0 Then
            lboFormatFound = WordAppl.Selection.Find.Execute(Replace:=wdReplaceOne)
            lloStart = WordAppl.Selection.End
            lboFormatFound = WordAppl.Selection.Find.Execute(Replace:=wdReplaceOne)
            lloEnd = WordAppl.Selection.Start
            WordAppl.Selection.Start = lloStart
            WordAppl.Selection.End = lloEnd
        Else
            WordAppl.Selection.Find.Execute Replace:=wdReplaceAll
            lboFormatFound = False
        End If
        
        Select Case pFormatString
        Case "</B>"
            WordAppl.Selection.Font.Bold = wdToggle
            WordAppl.Selection.Font.BoldBi = wdToggle
        Case "</I>"
            WordAppl.Selection.Font.Italic = wdToggle
            WordAppl.Selection.Font.ItalicBi = wdToggle
        Case "LTRLTR", "ילשילש"
        Case Else
            If Mid(pFormatString, 3, 1) = "S" Then
                WordAppl.Selection.Font.Size = CInt(Mid(pFormatString, 4, 2))
            End If
        End Select
    Wend
End Sub

Private Sub PageBreaks()
    Dim curpage As Integer
    
    'add page breaks
    WordAppl.Selection.HomeKey Unit:=wdStory
    Do While WordAppl.Selection.Information(wdNumberOfPagesInDocument) <> 1
        WordAppl.Selection.GoToNext (wdGoToPage)
        WordAppl.Selection.Find.ClearFormatting
        With WordAppl.Selection.Find
            .Text = "<PAGE>"
            .Replacement.Text = ""
            .Forward = False
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        If WordAppl.Selection.Find.Execute Then
            curpage = WordAppl.Selection.Information(wdActiveEndAdjustedPageNumber)
            While curpage = WordAppl.Selection.Information(wdActiveEndAdjustedPageNumber)
                WordAppl.Selection.InsertBreak Type:=wdLineBreak
            Wend
        End If
        If WordAppl.Selection.Information(wdNumberOfPagesInDocument) = _
            WordAppl.Selection.Information(wdActiveEndAdjustedPageNumber) Then
            Exit Do
        End If
    Loop
    WordAppl.Selection.Find.ClearFormatting
    WordAppl.Selection.Find.Replacement.ClearFormatting
    With WordAppl.Selection.Find
        .Text = "<PAGE>"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    WordAppl.Selection.Find.Execute Replace:=wdReplaceAll
End Sub

Private Sub QueryToDictionary(query As String, dic As Scripting.Dictionary, dicMulti As Scripting.Dictionary)
    
' Executes the query parameter, and insert each field to dictionary
' field name is the key, data is the item data
    
    Dim rst As ADODB.Recordset
    Dim i As Integer
'    Dim Temp As String
    Dim strVal As String
    
    Const MULTI_STR As String = "_*"

    On Error GoTo Fin
    Set rst = con.Execute(query)
    
    ' If no result from query, return empty fields
    If rst.EOF Then
        For i = 0 To rst.Fields.Count - 1
            Call dic.Add(rst(i).name, "")
            Call dicMulti.Add(rst(i).name & MULTI_STR, "")
        Next i
        rst.Close
        Exit Sub
    End If
    
    For i = 0 To rst.Fields.Count - 1
        If rst(i).Type = adLongVarChar Or rst(i).Type = adLongVarWChar Then
            strVal = "[RTF]" & ReadBlob(rst(i))
        Else
            strVal = nte(rst(i).Value)
        End If
        Call dic.Add(rst(i).name, strVal)
        Call dicMulti.Add(rst(i).name & MULTI_STR, strVal)
    Next i
    rst.MoveNext
    While Not rst.EOF
        For i = 0 To rst.Fields.Count - 1
            If rst(i).Type = adLongVarChar Or rst(i).Type = adLongVarWChar Then
                strVal = ReadBlob(rst(i))
            Else
                strVal = nte(rst(i).Value)
            End If
            dicMulti(rst(i).name & MULTI_STR) = _
                                dicMulti(rst(i).name & MULTI_STR) & _
                                "<BR>" & strVal
'            Temp = dicMulti(rst(i).name & MULTI_STR)
'            Call dicMulti.Remove(rst(i).name & MULTI_STR)
'            Call dicMulti.Add(rst(i).name & MULTI_STR, Temp & "<BR>" & strVal)

        Next i
        rst.MoveNext
    Wend
    rst.Close
    Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: QueryToDictionary , query= " & query & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading
'    End Select
End Sub

Private Sub InsertRTF(TextToReplace As String, ReplacingRTF As String)
    Dim fs As New Scripting.FileSystemObject
    Dim RTFFile As Scripting.TextStream
    Dim tmpFile As String
'    Dim docApp As Object 'Word.Application
    
    On Error GoTo Fin
    tmpFile = fs.GetSpecialFolder(2) & "\" & fs.GetTempName
    tmpFile = Replace(tmpFile, ".tmp", ".rtf")
    Set RTFFile = fs.CreateTextFile(tmpFile)
    ReplacingRTF = Replace(ReplacingRTF, "\lang1033", "")
    ReplacingRTF = Replace(ReplacingRTF, "\lang1037", "")
    Call RTFFile.Write(ReplacingRTF)
    Call RTFFile.Close
    
'    Set docApp = Doc.Application
    WordAppl.Selection.Find.ClearFormatting
    With WordAppl.Selection.Find
        .Text = TextToReplace
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchKashida = False
        .MatchDiacritics = False
        .MatchAlefHamza = False
        .MatchControl = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    While WordAppl.Selection.Find.Execute
        WordAppl.Selection.InsertFile FileName:=tmpFile, Range:="", ConfirmConversions:= _
            False, Link:=False, Attachment:=False
    Wend
    
    Call fs.DeleteFile(tmpFile) '18/05/05
    Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: InsertRTF" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & vbCrLf & _
                "TextToReplace = " & TextToReplace & vbCrLf & "ReplacingRTF = " & ReplacingRTF
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub

Private Sub InsertFieldValue(TextToReplace As String, ReplacingText As String)
'    Dim docApp As Object 'Word.Application

    Dim lstReplacingText As String
    Dim lstReplacingChunk As String
    On Error GoTo Fin
    
    lstReplacingText = Trim(ReplacingText)
    
    While Len(lstReplacingText) > 200
        lstReplacingChunk = Mid(lstReplacingText, 1, 200) & TextToReplace
        InsertFieldChunk TextToReplace, lstReplacingChunk
        lstReplacingText = Mid(lstReplacingText, 201)
    Wend
    InsertFieldChunk TextToReplace, lstReplacingText
    
Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: InsertFieldValue" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & vbCrLf & _
                "TextToReplace = " & TextToReplace & vbCrLf & "ReplacingText = " & ReplacingText & vbCrLf & _
                "Replacing Chunk = " & lstReplacingChunk
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
'    End Select
End Sub

Private Sub InsertFieldChunk(TextToReplace As String, ReplacingText As String)

    Dim i As Integer

    On Error GoTo Fin
'    Set docApp = Doc.Application
    For i = 1 To 3
        Select Case i
        Case 1
            Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        Case 2
            Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        Case 3
            Doc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        End Select
        WordAppl.Selection.Find.ClearFormatting
        WordAppl.Selection.Find.Replacement.ClearFormatting
        With WordAppl.Selection.Find
            .Text = TextToReplace
            .Replacement.Text = ReplacingText
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchKashida = False
            .MatchDiacritics = False
            .MatchAlefHamza = False
            .MatchControl = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        WordAppl.Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    Exit Sub
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: InsertFieldChunk" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & vbCrLf & _
                "TextToReplace = " & TextToReplace & vbCrLf & "ReplacingText = " & ReplacingText
    mloErrNbr = Err.Number
    Err.Raise Err.Number

End Sub
Private Function nte(e As Variant) As Variant
    
    nte = IIf(IsNull(e), "", e)

End Function
    
Private Function ReadBlob(pFld As ADODB.Field) As String
    
    ' Function read a the clob data from the field
    '   using the stream object of the ADODB library
    
    Dim lStream As ADODB.Stream
    Dim lstData As String
   
    On Error GoTo Fin
    
    Set lStream = New ADODB.Stream
    lStream.Charset = csHeBrEw
    lStream.Type = adTypeText
    lStream.Open

    lStream.WriteText pFld.Value
    lStream.Position = 0
    lstData = lStream.ReadText
    
    lStream.Close
    Set lStream = Nothing
    
    ReadBlob = lstData
Exit Function
Fin:
    mstErrMsg = mstErrMsg & vbCrLf & "SUB: ReadBlob" & vbCrLf & "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description
    mloErrNbr = Err.Number
    Err.Raise Err.Number
'    Select Case Err.Number
'    Case 462   ' Word Closed
'    Case Else
'        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading
'    End Select
End Function

Public Sub EndReport()
    
    mWordGuard.EndProcessDoc
    
    If Not Doc Is Nothing Then
        Doc.Close wdDoNotSaveChanges
        DoEvents
        Set Doc = Nothing
    End If

End Sub

Public Sub RemoveParameters()
    
    Parameters.RemoveAll

End Sub

Public Sub RemoveQueries()
    
    queries.RemoveAll

End Sub

