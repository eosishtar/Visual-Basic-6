VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cn As New ADODB.Recordset
Dim rs As ADODB.Recordset

Dim Success As Boolean
Dim DBASE_STATUS As Boolean
Dim vError As String

Private Const DBPass = "]lpkz]q`ƒ*"
Private Const ENCRYPT_KEY = "hjvghchf"

Public Event Connected(Status As Boolean)
Public Event ConnString(CommandText As String)

'...used for modifing the table
Public Enum ModifyTableNum
  vAdd = 1
  vDelete = 2
End Enum

'... gets the full path of the database
Public Function GetDBPath() As String
  
  vError = ""
  GetDBPath = GetSetting(App.EXEName, "DataBase", "DataPath")
  If GetDBPath = "" Then
    vError = "Could not locate the database!"
  End If

End Function


'... saves the full path of the database
Public Function SaveDBPath(vDBPath As String) As Boolean

  SaveDBPath = False
  If Trim(vDBPath) = "" Then
    Exit Function
  End If
  
  SaveSetting App.EXEName, "DataBase", "DataPath", vDBPath
  SaveDBPath = True

End Function

'...returns the password for the database
Private Function DBasePassword() As String

  DBasePassword = DecryptPass(DBPass, ENCRYPT_KEY)
  
End Function

Public Function Connect(vFullDBPath As String, Optional vPassword As String, Optional UseSystemPass As Boolean) As Boolean
  Dim cmd As String

  vError = ""
  Connect = False
  
  If UseSystemPass Then
    vPassword = DBasePassword
  End If
  
  On Error Resume Next
  If vPassword = "" Then
    cmd = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & vFullDBPath & ";Jet OLEDB"
  Else
    cmd = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & vFullDBPath & ";Jet OLEDB:Database Password=" & vPassword
  End If
  
  If Err.Number = 0 Then
    Connect = True
  Else
    Err.Clear
    vError = "Could not establish connection to database!"
  End If
  
  '...successfully connected
  If Connect Then
    With cn
      .CursorLocation = adUseClient
      .ActiveConnection = cmd   '...connection string
        Cnn = cmd
      .Open                     '... open connection
      RaiseEvent ConnString(cmd)
      RaiseEvent Connected(True)
    End With
  End If
  On Error GoTo 0
  
End Function

Public Function OpenDB(vSql As String, vCursorType As CursorTypeEnum, vLockType As LockTypeEnum)
  Dim dBpath As String

  Connect GetDBPath, DBasePassword

  Set rs = New ADODB.Recordset

  '...open the record set
  rs.Open vSql, cn.ActiveConnection, vCursorType, vLockType
  DBASE_STATUS = True

End Function

Public Function CloseDB() As Boolean
  
  vError = ""
  CloseDB = False
  On Error Resume Next
  rs.Close
  cn.Close

  Set rs = Nothing
  Set cn = Nothing
  
  If Err.Number = 0 Then
    CloseDB = True
  Else
    vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
    Err.Clear
  End If

  RaiseEvent Connected(False)

End Function

Private Function DecryptPass(texti, sEncrypt) As String

  Dim DeCrypted As String
  Dim T As Integer, x1 As Integer, G As Integer, TT As Integer
  Dim sana As Long

  vError = ""
  On Error Resume Next
  For T = 1 To Len(sEncrypt)
         sana = Asc(Mid(sEncrypt, T, 1))
         x1 = x1 + sana
  Next
  
  x1 = Int((x1 * 0.1) / 6)
  sEncrypt = x1
  G = 0
  
  For TT = 1 To Len(texti)
      sana = Asc(Mid(texti, TT, 1))
      G = G + 1
      If G = 6 Then G = 0
      x1 = 0
      If G = 0 Then x1 = sana + (sEncrypt - 2)
      If G = 1 Then x1 = sana - (sEncrypt - 5)
      If G = 2 Then x1 = sana + (sEncrypt - 4)
      If G = 3 Then x1 = sana - (sEncrypt - 2)
      If G = 4 Then x1 = sana + (sEncrypt - 3)
      If G = 5 Then x1 = sana - (sEncrypt - 5)
      x1 = x1 - G
      DeCrypted = DeCrypted & Chr(x1)
  Next
  
  If Err.Number = 0 Then
    DecryptPass = DeCrypted
  Else
    Err.Clear
    vError = "An error occurred when trying to decrypt the password"
  End If
  
End Function

Public Function DBBackup(vDatabase As String, Optional vBackUpPath As String) As Boolean
  '...must add a reference to "Microsoft Scripting Runtime"
  
  Dim MyBackupPath As String
  Dim reSp As String
  Dim fso As Object
    
  Success = False
  vError = ""
  If vDatabase = "" Then Exit Function


    '...warn user, system needs sole access
    reSp = MsgBox("You are about to backup the database. Please make sure no one is connected!!  " & Chr(13) & vbTab & "Continue ?  ", vbCritical + vbYesNo, "Backup Database")
    If reSp = vbYes Then
      '...Set the label values
      MyBackupPath = App.Path
      If Right(MyBackupPath, 1) <> "\" Then MyBackupPath = MyBackupPath & "\"
      MyBackupPath = MyBackupPath & Format(Now, "yyyymmddHHnnss") & ".bak"
      
      Set fso = New FileSystemObject
      fso.CopyFile vDatabase, MyBackupPath
      
      If Err.Number <> 0 Then
        vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
      Else
        Success = True
      End If
      
      If Success = True Then
        DBBackup = True
        vBackUpPath = MyBackupPath
      End If
    End If
  
    Set fso = Nothing
    Screen.MousePointer = vbDefault
End Function

'... used to check if fields exist in current database
Public Function TableExists(vTable As String, ShowErrMsg As Boolean) As Boolean
  '...requires refrence to DAO Object 3.6 Library
  Dim dB As DAO.Database
  Dim tbd As DAO.TableDef
  Dim ws As Object
      
  Set ws = DBEngine.Workspaces(0)
  Set dB = ws.OpenDatabase(GetDBPath, False, False, "MS Access;PWD=" & DBasePassword)
  TableExists = True
  vError = ""
  
    On Error Resume Next
    Set tbd = dB.TableDefs(vTable)
    TableExists = (Err.Number = 0)
    Set tbd = Nothing
  
  If ShowErrMsg Then
    If Err.Number <> 0 Then
      vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
    End If
  End If
  
  '...clean up
  On Error GoTo 0
  Err.Clear
  Set ws = Nothing
  Set dB = Nothing
  Set tbd = Nothing
  
End Function

'... used to add or delete fields to existing tables.
Public Function ModifyTable(vAction As ModifyTableNum, vTable As String, vField As String, Optional vType As DAO.DataTypeEnum, Optional vFieldSize As Integer) As Boolean
    Dim dB As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim ws As Object
    
    'Initialize
    vError = ""
    ModifyTable = False
    Set ws = DBEngine.Workspaces(0)
    Set dB = ws.OpenDatabase(GetDBPath, False, False, "MS Access;PWD=" & DBasePassword)

  
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Set tdf = dB.TableDefs(vTable)
    
    
    If vAction = vAdd Then
      '...Add a field to the table.
      If vField <> "" Or vType <> "" Or vFieldSize <> "" Then
        tdf.Fields.Append tdf.CreateField(vField, vType, vFieldSize)
        If Err.Number <> 0 Then
          If Err.Number = 13 Then
            ModifyTable = True
          Else
            vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
          End If
        Else
          ModifyTable = True
        End If
      End If
    Else
      '...Delete a field from the table.
      tdf.Fields.Delete vField
      If Err.Number <> 0 Then
        vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
      Else
        ModifyTable = True
      End If
    End If
    
    On Error GoTo 0
    If ModifyTable Then
      If vAction = vAdd Then
        vError = "The field '" & vField & "' was successfully added to table '" & vTable & "'"
      Else
        vError = "The field '" & vField & "' was successfully deleted from table '" & vTable & "'"
      End If
    End If

    '...clean up
    Screen.MousePointer = vbNormal
    Set fld = Nothing
    Set tdf = Nothing
    Set dB = Nothing
    
End Function


'... used to create a new access database
Public Function CreateAccessDatabase(vDBName As String, Optional vUsePassword As Boolean) As Boolean
   
   Dim dB As DAO.Database
   Dim dBpath As String
   Dim vPassword As String
   
    vError = ""
    CreateAccessDatabase = False
    dBpath = App.Path
    If Right(dBpath, 1) <> "\" Then dBpath = dBpath & "\"

    If Not vUsePassword Then
      Set dB = DBEngine.CreateDatabase(dBpath & vDBName & ".mdb", dbLangGeneral)
    Else
      vPassword = DBasePassword
      Set dB = DBEngine.CreateDatabase(dBpath & vDBName & ".mdb", dbLangGeneral & ";pwd=" & vPassword)
    End If
   
    If Err.Number = 0 Then
      CreateAccessDatabase = True
    Else
      vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
    End If
  
End Function

Public Function DBCompact(ByVal DBName1 As String, Optional ByVal DBName2 As String)
'.... requires reference to Microsoft Jet and Replications Objects
  Dim JRO As New JRO.JetEngine
  Dim db_Old As String
  Dim db_New As String
  Dim myLen As Integer
  Dim reSp As String
  Dim s1 As Long, s2 As Long

  db_Old = DBName1
  myLen = Len(DBName1)
  myLen = myLen - 4
  db_New = Left(DBName1, myLen)
  db_New = db_New + "_Temp.mdb"

  'get file size before compact
  vError = ""
  Close #1
  Open db_Old For Binary As #1
  s1 = LOF(1)
  Close #1

  Set JRO = New JetEngine
  '...warn user, system needs sole access
  reSp = MsgBox("You are about to compact the database. Please make sure no one is connected!!  " & Chr(13) & vbTab & "Continue ?  ", vbCritical + vbYesNo, Screen.ActiveForm.Caption)
  If reSp = vbYes Then
  
    '...close the connections
    Call CloseDB
    
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    
    Set JRO = New JRO.JetEngine
      JRO.CompactDatabase "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & db_Old & ";Jet OLEDB:Database Password=[YOUR DB Password]", _
      "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & db_New & ";Jet OLEDB:Database Password=" & DBasePassword
    
    FileCopy db_Old, db_New
    Kill db_Old
    Name db_New As db_Old
    
    'get file size after compact
    Open db_Old For Binary As #2
    s2 = LOF(2)
    Close #2
    
    vError = "Compact complete " & vbCrLf & vbCrLf & "Size:    " & Round(s2 / 1024 / 1024, 2) & "Mb"
  End If
  
  Screen.MousePointer = vbNormal

On Error GoTo 0
  Set JRO = Nothing

End Function

Private Sub Class_Terminate()
  
  '... fail over to close db connection
  If DBASE_STATUS Then
    Call CloseDB
  End If

End Sub


Public Property Get OpenRecordSet() As ADODB.Recordset

  Set OpenRecordSet = rs

End Property

Public Property Let OpenRecordSet(ByVal vNewValue As ADODB.Recordset)

  Set rs = vNewValue

End Property

Public Function Error() As String
  Error = vError
End Function


