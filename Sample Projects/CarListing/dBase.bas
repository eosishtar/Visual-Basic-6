Attribute VB_Name = "dBase"
Option Explicit

Public rs As New ADODB.Recordset
Public cn As New ADODB.connection

Public cmd As String
Dim vError As String

Private Const DBPass = "]lpkz]q`ƒ*"
Public Const ENCRYPT_KEY = "hjvghchf"

Public Function GetDatabaseName() As String

  GetDatabaseName = GetSetting(App.EXEName, "Datapath", "Database")
  
  If GetDatabaseName = "" Then
    GetDatabaseName = GetAppPath & "DB.mdb"
      frmSettings.Show
  End If

End Function

Public Function GetDefPrinter() As String
  GetDefPrinter = GetSetting(App.EXEName, "Datapath", "DefaultPrinter")
End Function

Public Sub Go()

Call dbase
Call connection

    Set rs = New ADODB.Recordset
    
End Sub

Public Sub dbase()                         ' call this procedure to connect to dbase
  Dim vPassword As String
  
  vPassword = DBasePassword
  
  'cmd = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & GetDatabaseName & ";Jet OLEDB:Database Password=Starlight1"
  cmd = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & GetDatabaseName & ";Jet OLEDB:Database Password=" & vPassword
End Sub
Public Sub connection() ' run this procedure to open a connected dbase

On Error GoTo OmY

Set cn = New ADODB.connection
    With cn
        .CursorLocation = adUseClient
        .ConnectionString = cmd   'to connect to Access dbase
                                  'only to connect to SQL dbase
        .Open
    End With

Exit Sub

OmY:

MsgBox "Database not found...", vbCritical, "Database Manager..."
frmSettings.Show 1

End Sub

Private Function DBasePassword() As String

  DBasePassword = DecryptPass(DBPass, ENCRYPT_KEY)
  
End Function

Public Function SaveDBPath(vDBPath As String) As Boolean

  SaveDBPath = False
  If Trim(vDBPath) = "" Then
    Exit Function
  End If
  
  SaveSetting App.EXEName, "DataBase", "DataPath", vDBPath
  SaveDBPath = True

End Function

Private Function DecryptPass(texti, sEncrypt) As String

  Dim DeCrypted As String
  Dim t As Integer, x1 As Integer, G As Integer, TT As Integer
  Dim sana As Long

  vError = ""
  On Error Resume Next
  For t = 1 To Len(sEncrypt)
         sana = Asc(Mid(sEncrypt, t, 1))
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

