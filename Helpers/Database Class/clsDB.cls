VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'...Class requires references to the following.
'Microsoft ADO Ext. 6.0 for DDL and Security
'Microsoft ActiveX Data Objects 6.1 Library

Option Explicit

Private mError As String
Private cConnStatus As Boolean
Private cRSStatus As Boolean
Private sFullDBPath As String
Private sDataBaseName As String
Private sSqlString As String
Private iRecordCount As Integer

Dim Conn As ADODB.Connection      'Active Connection
Dim RS As ADODB.Recordset         'Active Recordset

Public Event Connected(Status As Boolean)


Public Function OpenConnection(sDataBaseName As String, sDBPath As String, Optional sDBPassword As String = "") As Boolean

  Dim strSQL As String
  
  Error = ""
  OpenConnection = False
  
  If Right(sDBPath, 1) <> "\" Then sDBPath = sDBPath & "\"
  DatabasePath = sDBPath
  DatabaseName = sDataBaseName
  
  Set Conn = New ADODB.Connection
  If sDBPassword = "" Then
    strSQL = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DatabasePath & DatabaseName & ";Jet OLEDB"
  Else
    strSQL = "Provider=Microsoft.jet.OLEDB.4.0;Data Source=" & DatabasePath & DatabaseName & ";Jet OLEDB:Database Password=" & sDBPassword
  End If
  On Error Resume Next
  Err.Clear
  Conn.Open strSQL
  Conn.CommandTimeout = 0
  
  If Err <> 0 Then
    Error = Err.Description
  End If
  
  ConnectionStatus = IIf(Error = "", True, False)
  OpenConnection = ConnectionStatus
  RaiseEvent Connected(ConnectionStatus)
  Connection = Conn.ConnectionString
  On Error GoTo 0

End Function

Private Sub Class_Initialize()

  ConnectionStatus = False
  RecordSetStatus = False

End Sub

Private Sub Class_Terminate()

  CloseConnection

End Sub

Public Function OpenRecordSet(sSQL As String, Optional CursorType As CursorTypeEnum = adOpenUnspecified, Optional LockType As LockTypeEnum = adLockUnspecified) As Boolean

  Error = ""
  OpenRecordSet = False
  If cConnStatus = False Then
    Error = "No Connection has been set"
    Exit Function
  End If
  
  Set RS = New ADODB.Recordset
  On Error Resume Next
  SqlString = sSQL
  Err.Clear
  RS.Open sSQL, Conn, CursorType, LockType
  
  RecordSetStatus = RS.State
  OpenRecordSet = cRSStatus
  
  
  If Err <> 0 Then
    Error = Err.Description
    On Error GoTo 0
    Exit Function
  End If
  
  OpenRecordSet = True
  Recordset = RS      '...set the RS to the class
  RecordCount = GetRecordCount
  On Error GoTo 0

End Function

Public Sub CloseRecordset()

  If cConnStatus = True Then
    On Error Resume Next
    RS.Close
    On Error GoTo 0
  End If
  Set RS = Nothing

End Sub

Public Sub CloseConnection()

  CloseRecordset
  If cConnStatus Then
    On Error Resume Next
    Conn.Close
    On Error GoTo 0
  End If
  RaiseEvent Connected(False)
  cConnStatus = False
  Set Conn = Nothing

End Sub

Private Function GetRecordCount() As Integer
  Dim tRS As New ADODB.Recordset
  
  If Not ConnectionStatus Then
    GetRecordCount = -2      'Not Connected
  Else
    tRS.Open SqlString, Connection, adOpenStatic, adLockReadOnly
    If tRS.EOF Then
      GetRecordCount = -1
    Else
      GetRecordCount = tRS.RecordCount
    End If
  End If
  Set tRS = Nothing

End Function

Public Property Get RecordCount() As Integer
  RecordCount = iRecordCount
End Property

Private Property Let RecordCount(ByVal vNewValue As Integer)
  iRecordCount = vNewValue
End Property

Private Property Get Connection() As ADODB.Connection
  Set Connection = Conn
End Property

Private Property Get ConnectionStatus() As Boolean
  ConnectionStatus = CBool(Conn.State)
  cConnStatus = ConnectionStatus
End Property

Private Property Let ConnectionStatus(ByVal vNewValue As Boolean)
  cConnStatus = vNewValue
End Property

Private Property Get RecordSetStatus() As Boolean
  RecordSetStatus = CBool(RS.State)
  cRSStatus = RecordSetStatus
End Property

Private Property Let RecordSetStatus(ByVal vNewValue As Boolean)
  cRSStatus = vNewValue
End Property

'...public so we can access this outside of class
Public Property Get Recordset() As ADODB.Recordset
  Set Recordset = RS
End Property

Private Property Let Recordset(ByVal vNewValue As ADODB.Recordset)
  Set RS = vNewValue
End Property

'...public so we can access this outside of class
Public Property Get Error() As String
  Error = Trim(mError)
End Property

Private Property Let Error(ByVal vNewValue As String)
  mError = vNewValue
End Property

'...public so we can access this outside of class
Public Property Get DatabaseName() As String
  DatabaseName = sDataBaseName
End Property

Private Property Let DatabaseName(ByVal vNewValue As String)
  sDataBaseName = vNewValue
End Property

Private Property Get SqlString() As String
  SqlString = sSqlString
End Property

Private Property Let SqlString(ByVal vNewValue As String)
  sSqlString = vNewValue
End Property

'...public so we can access this outside of class
Public Property Get DatabasePath() As String
  DatabasePath = sFullDBPath
End Property

Private Property Let DatabasePath(ByVal vNewValue As String)
  sFullDBPath = vNewValue
End Property


