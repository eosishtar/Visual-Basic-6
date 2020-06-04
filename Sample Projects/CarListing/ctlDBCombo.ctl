VERSION 5.00
Begin VB.UserControl ctlDBCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DataSourceBehavior=   1  'vbDataSource
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
End
Attribute VB_Name = "ctlDBCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event ItemSelected(ItemID As Long, ItemName As String)
Dim vError As String
Dim TempRS As ADODB.Recordset

Public Function ListIndex(Index As Integer)
  Combo1.ListIndex = Index
End Function
Public Function ClearList()
  Combo1.Clear
End Function

Public Function AddItem(ItemToAdd As String)
  Combo1.AddItem ItemToAdd
End Function

Public Function PopulateList(Conn As ADODB.connection, tTable As String, tField As String, Blocked As Boolean, Optional Allocate As Boolean) As Boolean
'... use blocked = true when u want to display products that are not blocked.
Dim C1 As Integer
Dim C2 As Integer
Dim C3 As String
Dim strSql As String

PopulateList = False
vError = ""
On Error Resume Next
Err.Clear
Set TempRS = New ADODB.Recordset
If TempRS.State = 1 Then Set TempRS = Nothing

If Blocked = False Then
  strSql = "Select Distinct ID," & tField & " from " & tTable & " ORDER BY " & tField
Else
  If Allocate Then
    strSql = "Select Distinct ID, Blocked, Deleted," & tField & " from " & tTable & " WHERE [Blocked] = " & False & " AND [DELETED] = " & False & " ORDER BY " & tField
  Else
    strSql = "Select Distinct ID, Blocked," & tField & " from " & tTable & " WHERE [Blocked] = " & False & " ORDER BY " & tField
  End If
End If

TempRS.Open strSql, Conn, adOpenStatic, adLockOptimistic

Combo1.Clear

For C1 = 1 To TempRS.RecordCount
     C3 = TempRS.Fields(tField).Value
    Combo1.AddItem C3
    For C2 = 0 To Combo1.ListCount - 1
      If Combo1.List(C2) = C3 Then
        Combo1.ItemData(C2) = TempRS.Fields("ID").Value
      End If
    Next C2
    TempRS.MoveNext
  Next C1
  Combo1.Refresh
  TempRS.Close
  Set TempRS = Nothing
  If Err <> 0 Then
    vError = Err.Description
    On Error GoTo 0
  End If
  
  On Error GoTo 0
  PopulateList = True

End Function

Private Sub Combo1_Change()
  If Combo1.ListIndex = -1 Then Exit Sub
  RaiseEvent ItemSelected(Combo1.ItemData(Combo1.ListIndex), Combo1.List(Combo1.ListIndex))
End Sub

Private Sub Combo1_Click()
  If Combo1.ListIndex = -1 Then Exit Sub
  RaiseEvent ItemSelected(Combo1.ItemData(Combo1.ListIndex), Combo1.List(Combo1.ListIndex))
End Sub

Private Sub UserControl_Resize()
  Combo1.Move 0, 0, UserControl.Width
  If UserControl.Height <> Combo1.Height Then UserControl.Height = Combo1.Height
End Sub

Public Function Error() As String
  Error = vError
End Function

Public Property Get Text() As String
  Text = Combo1.List(Combo1.ListIndex)
End Property

Public Property Let Text(ByVal vNewValue As String)
  Combo1.Text = vNewValue
End Property

Public Property Get ID() As Long
  ID = Combo1.ItemData(Combo1.ListIndex)
End Property

Public Property Let ID(ByVal vNewValue As Long)

  Dim C4 As Long
  For C4 = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(C4) = vNewValue Then
      Combo1.ListIndex = C4
      Exit For
    End If
  Next C4
    
End Property

