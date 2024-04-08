VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirecWordReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements LSEXT.IEntityExtension

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function IEntityExtension_CanExecute(Parameters As LSEXT.IExtensionParameters) As LSEXT.ExecuteExtension
    IEntityExtension_CanExecute = exEnabled
End Function

Private Sub IEntityExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
    Dim ServiceProvider As LSSERVICEPROVIDERLib.NautilusServiceProvider
    Dim NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
    Dim con As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim rstWR As ADODB.Recordset
    Dim constr As String
    
    Set rst = Parameters("RECORDS")
    

    Set ServiceProvider = Parameters("SERVICE_PROVIDER")
    Set NtlsCon = ServiceProvider.QueryServiceProvider("DBConnection")

    constr = "Provider=OraOLEDB.Oracle" & _
        ";Data Source=" & NtlsCon.GetServerDetails & _
        ";User ID=" & NtlsCon.GetUsername & _
        ";Password=" & NtlsCon.GetPassword
        
    constr = NtlsCon.GetADOConnectionString
    con.Open constr
    con.CursorLocation = adUseClient
    con.Execute "SET ROLE LIMS_USER"
    Call ConnectSameSession(con, CDbl(Parameters("SESSION_ID")))

'    Dim queries As ADODB.Recordset
    Dim params As ADODB.Recordset
'    Dim Dest As ADODB.Recordset
    Dim QueriesParams As ADODB.Recordset
    Dim EnvelopeQuery As String
    Dim fParam As New frmParameter
    Dim i As Integer
'    Dim Template As String
    Dim lstComputerName As String
    Dim lloStrLen As Long
    Dim lstSQL As String
    Dim ldoDocId As Double
    
    lloStrLen = 50
    lstComputerName = Space(50)
    Call GetComputerName(lstComputerName, lloStrLen)
    lstComputerName = Left(lstComputerName, lloStrLen)
    
    Set rstWR = con.Execute("select * from lims_sys.u_wreport wr, lims_sys.u_wreport_user wru " & _
        "where wr.u_wreport_id = wru.u_wreport_id and " & _
        "wr.u_wreport_id = " & rst(0).Value)
    Set params = con.Execute("select phrase_name, phrase_description " & _
        "from lims_sys.phrase_header h, lims_sys.phrase_entry e " & _
        "where e.phrase_id = h.phrase_id and " & _
        "h.name = '" & rstWR("U_PARAM_PHRASE") & "' " & _
        "order by e.order_number")
    EnvelopeQuery = rstWR("U_ENVELOPE_QUERY")
    While Not params.EOF
        fParam.lblDescription.Caption = params("PHRASE_DESCRIPTION")
        fParam.txtValue.Text = ""
        fParam.Show vbModal
        EnvelopeQuery = Replace(EnvelopeQuery, "#" & params("PHRASE_NAME") & "#", fParam.txtValue.Text)
        params.MoveNext
    Wend
    params.Close
        
'--    EnvelopeQuery = Replace(EnvelopeQuery, "'", "''")
'
'--    lstSQL = "INSERT INTO lims.bg_print (report_type, report_id, envelope_query, created_on, workstation_name) " & _
'--             "VALUES ( 'DIRECT'," & rst(0) & "," & _
'--             "'" & EnvelopeQuery & "',sysdate, " & _
'--             "'" & lstComputerName & "') "
'--
'--    con.Execute lstSQL
'--
'--    rstWR.Close
'--    con.Close
    
    Set QueriesParams = con.Execute(EnvelopeQuery)
'    Set queries = con.Execute("select u_query, u_query_name, u_word_template, u_wreport_user.u_wreport_id " & _
'        "from lims_sys.u_wreport_query_user, lims_sys.u_wreport_user " & _
'        "where u_wreport_query_user.u_wreport_id = u_wreport_user.u_wreport_id and " & _
'        "u_wreport_user.u_wreport_id = " & rst(0).Value)
'    Set WordReport.Connection = con
'    Template = queries("U_WORD_TEMPLATE")
'    While Not queries.EOF
'        Call WordReport.AddQuery(queries("U_QUERY_NAME"), queries("U_QUERY"))
'        queries.MoveNext
'    Wend
'    queries.Close
'    If QueriesParams.EOF Then
'        MsgBox "No records selected."
'    End If
    
    EnvelopeQuery = Replace(EnvelopeQuery, "'", "''")
    
    While Not QueriesParams.EOF
'        WordReport.RemoveParameters
'        WordReport.SetTemplate (Template)

        ' Get sequence for the report
        lstSQL = "select lims.docs_seq.nextval doc_id from dual"
        ldoDocId = con.Execute(lstSQL)("doc_id").Value

        ' Insert report to bg_print
        lstSQL = "INSERT INTO lims.bg_print (report_type, report_id, doc_id, envelope_query, created_on, workstation_name) " & _
                 "VALUES ( 'DIRECT'," & rst(0) & "," & ldoDocId & "," & _
                 "'" & EnvelopeQuery & "',sysdate, " & _
                 "'" & lstComputerName & "') "

        con.Execute lstSQL

        ' Insert report_parameters to bg_print_params
        For i = 0 To QueriesParams.Fields.Count - 1
            lstSQL = "INSERT INTO lims.bg_print_params (doc_id, param_name, param_value) " & vbCrLf & _
                     "VALUES ( " & ldoDocId & "," & vbCrLf & _
                     "         '" & QueriesParams.Fields(i).name & "'," & vbCrLf & _
                     "         '" & Replace(QueriesParams.Fields(i).Value, "'", "''") & "') "
                    
            con.Execute lstSQL
            
        '            Call WordReport.AddParameter(QueriesParams.Fields(i).name, QueriesParams.Fields(i).Value)
        Next i
'        Set Dest = con.Execute("select * from lims_sys.u_wrdestination_user du where du.u_wreport_id = " & rst(0).Value)
'        While Not Dest.EOF
'            If Dest("U_TYPE") = "F" Then
'                Call WordReport.SaveReport(nvl(Dest("U_DEVICE_NAME"), ""))
'            Else
'                Call WordReport.PrintReport(nvl(Dest("U_DEVICE_NAME"), ""), nvl(Dest("U_COPIES"), 0))
'            End If
'            Dest.MoveNext
'        Wend
'        Dest.Close
        QueriesParams.MoveNext
    Wend
'    QueriesParams.Close
'    rstWR.Close
'    Set WordReport = Nothing
End Sub

Private Sub ConnectSameSession(ByVal aConnection As ADODB.Connection, ByVal aSessionID As Double)
    Dim aProc As New ADODB.Command
    Dim aSession As New ADODB.Parameter
    
    aProc.ActiveConnection = aConnection
    aProc.CommandText = "lims.lims_env.connect_same_session"
    aProc.CommandType = adCmdStoredProc
    
    aSession.Type = adDouble
    aSession.Direction = adParamInput
    aSession.Value = aSessionID
    aProc.Parameters.Append aSession
    
    aProc.Execute
    Set aSession = Nothing
    Set aProc = Nothing
End Sub

Private Function nvl(e As Variant, v As Variant) As Variant
    nvl = IIf(IsNull(e), v, e)
End Function


