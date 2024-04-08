VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WFWordReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements LSEXT.IWorkflowExtension

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub IWorkflowExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
    Dim ServiceProvider As LSSERVICEPROVIDERLib.NautilusServiceProvider
    Dim NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
    Dim con As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim lstComputerName As String
    Dim lloStrLen As Long
    Dim lstSQL As String
    Dim constr As String
    
    lloStrLen = 50
    lstComputerName = Space(50)
    Call GetComputerName(lstComputerName, lloStrLen)
    lstComputerName = Left(lstComputerName, lloStrLen)
    
    Set rst = Parameters("RECORDS")
    rst.MoveLast
'    con.Open "Provider=OraOLEDB.Oracle;Data Source=ntls;User ID=lims;Password=lims"
    
    Set ServiceProvider = Parameters("SERVICE_PROVIDER")
    Set NtlsCon = ServiceProvider.QueryServiceProvider("DBConnection")
    constr = "Provider=OraOLEDB.Oracle" & _
        ";Data Source=" & NtlsCon.GetServerDetails & _
        ";User ID=" & NtlsCon.GetUsername & _
        ";Password=" & NtlsCon.GetPassword
     
          
          If NtlsCon.GetServerIsProxy Then
    constr = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
          
    

        
        
        
    con.Open constr
    con.CursorLocation = adUseClient
    con.Execute "SET ROLE LIMS_USER"
    Call ConnectSameSession(con, CDbl(Parameters("SESSION_ID")))

    lstSQL = "INSERT INTO lims.bg_print (report_type, sdg_id, workflow_node_id, created_on, workstation_name) " & _
             "VALUES ( 'WF'," & rst(Parameters("TABLE_NAME") & "_ID") & "," & _
             Parameters("WORKFLOW_NODE_ID") & ",sysdate, " & _
             "'" & lstComputerName & "') "

    con.Execute lstSQL

    con.Close
    
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

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

