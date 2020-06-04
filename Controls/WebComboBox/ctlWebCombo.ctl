VERSION 5.00
Begin VB.UserControl ctlWebCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ControlContainer=   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   10380
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ctlWebCombo.ctx":0000
      Left            =   0
      List            =   "ctlWebCombo.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   450
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "ctlWebCombo.ctx":0004
      Left            =   5280
      List            =   "ctlWebCombo.ctx":0006
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "ctlWebCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vSearchInCombo As Boolean
Private vEnabled As Boolean
Private vError As String
Private bSearchIsOpen As Boolean
Private TempRS As ADODB.Recordset
Private sSearchColor As OLE_COLOR
Private vbArrowPressed As Boolean
Private vSearchText As String
Private bAutoComplete As Boolean


Dim blnAuto As Boolean

Public Event ItemSelected(ItemID As Long, ItemName As String)
Public Event SearchOpen(SearchIsOpen As Boolean)

Public Enum ComboTypeEnum
  tDropDownCombo = 0
  tDropdownList = 2
End Enum
Private vComboType As ComboTypeEnum



Private Sub cmdOpen_Click()

  If bSearchIsOpen Then
    '...Close
    UserControl.Height = Text1.Height
    List1.Visible = False
    Text2.Visible = False
    bSearchIsOpen = False
    RaiseEvent SearchOpen(bSearchIsOpen)
  Else
    '...Open
    If vSearchInCombo Then
      UserControl.Height = Text1.Height + List1.Height + 50
      List1.Visible = True
      Text2.Visible = True
      Text2.Move Text1.Left + 25, Text1.Top + 25, Text1.Width - (cmdOpen.Width + 25), Text1.Height - 25
      List1.Move Text1.Left, Text1.Top + Text1.Height, Text1.Width
      If Text2.Visible And Text2.Enabled Then
        Text2.SetFocus
      End If
    Else
      UserControl.Height = Text1.Height + List1.Height
      List1.Visible = True
      List1.Move Text1.Left, Text1.Top + Text1.Height, Text1.Width
    End If
    Call SearchList("")     'Refresh List
    bSearchIsOpen = True
    RaiseEvent SearchOpen(bSearchIsOpen)
    
  End If

End Sub

Private Sub Combo1_Change()
  If Combo1.ListIndex = -1 Then Exit Sub
  RaiseEvent ItemSelected(Combo1.ItemData(Combo1.ListIndex), Combo1.List(Combo1.ListIndex))
End Sub

Private Sub Combo1_Click()
  If Combo1.ListIndex = -1 Then Exit Sub
  RaiseEvent ItemSelected(Combo1.ItemData(Combo1.ListIndex), Combo1.List(Combo1.ListIndex))
End Sub

Private Sub List1_Click()
  Dim iListIndex As Integer
  If List1.ListIndex = -1 Or vbArrowPressed Then Exit Sub
  Call cmdOpen_Click
  iListIndex = GetListIndex(List1.List(List1.ListIndex))      'Get List Index From ComboBox
  Text1.Text = Combo1.List(iListIndex)
  RaiseEvent ItemSelected(Combo1.ItemData(iListIndex), Combo1.List(iListIndex))
End Sub

Private Sub List1_DblClick()
  Dim iListIndex As Integer
  If List1.ListIndex = -1 Or vbArrowPressed Then Exit Sub
  Call cmdOpen_Click
  iListIndex = GetListIndex(List1.List(List1.ListIndex))      'Get List Index From ComboBox
  Text1.Text = Combo1.List(iListIndex)
  RaiseEvent ItemSelected(Combo1.ItemData(iListIndex), Combo1.List(iListIndex))
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
  vbArrowPressed = False

  If List1.ListCount > 0 Then
    If KeyAscii = 13 Then
      Call List1_Click
    End If
  End If

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  vbArrowPressed = False
End Sub

Private Sub Text1_GotFocus()
  cmdOpen.SetFocus
End Sub

Private Sub Text2_GotFocus()
  Text2.Text = vSearchText
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If Text2.Text = vSearchText Then
    Text2.Text = ""
  End If
  
  If KeyCode = vbKeyBack Or vbKeyDelete Then
    blnAuto = True
    Text2.SelText = ""
    blnAuto = False
  ElseIf KeyCode = vbKeyReturn Then
    Text2.SelStart = Len(Text2.Text)
  End If
  
  
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

  vbArrowPressed = False

  '...User pressed enter
  If KeyCode = 13 Then
      If List1.ListCount > 0 Then
        Call cmdOpen_Click
        Text1.Text = List1.List(0)
        cmdOpen.SetFocus
        RaiseEvent ItemSelected(List1.ItemData(0), List1.List(0))
      End If
    Exit Sub
  ElseIf KeyCode = vbKeyDown Then
    vbArrowPressed = True
    List1.SetFocus
    List1.ListIndex = 0
    
  End If
  Call SearchList(Text2.Text)
  
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.MousePointer = 0
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Text1.MousePointer = 12
End Sub


Private Sub UserControl_Initialize()
  Call ComboResize
  Call DoVisible(False)
  bSearchIsOpen = False
End Sub

Private Sub UserControl_Resize()
  Call ComboResize
End Sub

Public Function Error() As String
  Error = vError
End Function


' --- load property bags
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("vComboType", vComboType)
  Call PropBag.WriteProperty("vSearchInCombo", vSearchInCombo)
  Call PropBag.WriteProperty("vEnabled", vEnabled)
  Call PropBag.WriteProperty("sSearchColor", sSearchColor)
  Call PropBag.WriteProperty("bAutoComplete", bAutoComplete)
  Call PropBag.WriteProperty("vSearchText", vSearchText)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vComboType = PropBag.ReadProperty("vComboType", "0")
  vSearchInCombo = PropBag.ReadProperty("vSearchInCombo", False)
  vEnabled = PropBag.ReadProperty("vEnabled", True)
  sSearchColor = PropBag.ReadProperty("sSearchColor", &HC0C0FF)
  bAutoComplete = PropBag.ReadProperty("bAutoComplete", True)
  vSearchText = PropBag.ReadProperty("vSearchText", "Search...")
    
  '...set values from propery bag.
  Text2.BackColor = sSearchColor
  Text2.Text = vSearchText
  
  Call DoComboStyle(vComboType)
  Call DoEnabled(vEnabled)
  Call ComboResize
  
End Sub

Private Sub DoComboStyle(eComboStyle As ComboTypeEnum)

  Call DoVisible(False)
  
  Select Case eComboStyle
    Case 0    'DropDown Combo
      Combo1.Enabled = True
      Combo1.Visible = True
      Call ComboResize
    Case 2      'DropDown List
      Text1.Visible = True
      Text1.Enabled = True
      cmdOpen.Visible = True
      cmdOpen.Enabled = True
      Call ComboResize
  End Select
  
End Sub

Private Sub ComboResize()

  If vComboType = tDropDownCombo Then
    Combo1.Move 0, 0, UserControl.Width
    If UserControl.Height <> Text1.Height Then UserControl.Height = Text1.Height
    
  ElseIf vComboType = tDropdownList Then
    Text1.Move 0, 0, UserControl.Width, 315
    cmdOpen.Move Text1.Left + (Text1.Width - cmdOpen.Width), Text1.Top, cmdOpen.Width, Text1.Height
    
  End If
  
End Sub

Private Sub DoVisible(vValue As Boolean)

  List1.Visible = vValue
  
  Text1.Text = ""
  Text2.Text = ""
  
  Text1.Visible = vValue
  Text1.Enabled = vValue
  cmdOpen.Visible = vValue
  cmdOpen.Enabled = vValue

  Combo1.Enabled = vValue
  Combo1.Visible = vValue

End Sub

Private Sub DoEnabled(vValue As Boolean)
  
  If vValue Then
    List1.Visible = False
    Text2.Visible = False
    Text1.BackColor = vbWindowBackground
  Else
    Text1.BackColor = vbButtonFace
  End If
  
  Combo1.Enabled = vValue
  Text1.Enabled = vValue
  cmdOpen.Enabled = vValue
  
End Sub

Public Property Get Text() As String
  Text = Combo1.List(Combo1.ListIndex)
End Property

Public Property Let Text(ByVal vNewValue As String)
  
  On Error GoTo Error
  If vComboType = tDropDownCombo Then
    'normal combo
    Combo1.Text = vNewValue
  ElseIf vComboType = tDropdownList Then
    'custom dropdown
    Text1.Text = vNewValue
  End If
  On Error GoTo 0
  
Error:
  If Err.Number <> 0 Then
    MsgBox "Error : " & Err.Number & vbCrLf & Err.Description, vbCritical, App.Title
  End If
  Err.Clear
  On Error GoTo 0
End Property

Public Property Get ID() As Long
  
  On Error Resume Next
  ID = Combo1.ItemData(Combo1.ListIndex)
  
  If Err.Number <> 0 Then
    ID = 0
    Err.Clear
  End If
  On Error GoTo 0
  
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

Public Property Get STYLE_COMBO() As ComboTypeEnum
  STYLE_COMBO = vComboType
End Property

Public Property Let STYLE_COMBO(ByVal vNewValue As ComboTypeEnum)
  vComboType = vNewValue
  Call DoComboStyle(vComboType)
End Property

Public Property Get ENABLED_COMBO() As Boolean
  ENABLED_COMBO = vEnabled
End Property

Public Property Let ENABLED_COMBO(ByVal vNewValue As Boolean)
  vEnabled = vNewValue
  Call DoEnabled(vEnabled)
End Property

Public Property Get SEARCH() As Boolean
  SEARCH = vSearchInCombo
End Property

Public Property Let SEARCH(ByVal vNewValue As Boolean)
  vSearchInCombo = vNewValue
End Property

Public Property Get SEARCH_TEXT() As String
  SEARCH_TEXT = vSearchText
End Property

Public Property Let SEARCH_TEXT(ByVal vNewValue As String)
  vSearchText = vNewValue
  Text2.Text = vSearchText
End Property

Public Property Get SEARCH_AUTOCOMPLETE() As Boolean
  SEARCH_AUTOCOMPLETE = bAutoComplete
End Property

Public Property Let SEARCH_AUTOCOMPLETE(ByVal vNewValue As Boolean)
  bAutoComplete = vNewValue
End Property

Public Property Get SEARCH_COLOR() As OLE_COLOR
  SEARCH_COLOR = sSearchColor
End Property

Public Property Let SEARCH_COLOR(ByVal vNewValue As OLE_COLOR)
  sSearchColor = vNewValue
  Text2.BackColor = sSearchColor
End Property

Public Function ListIndex(Index As Integer)
  Combo1.ListIndex = Index
End Function

Public Function ClearList()
  Combo1.Clear
  List1.Clear
End Function

Public Function AddItem(ItemToAdd As String)
  Combo1.AddItem ItemToAdd
  List1.AddItem ItemToAdd
End Function

Public Function ListCount() As Integer
  ListCount = Combo1.ListCount
End Function

Public Function List(ListIndex As Integer) As String
  List = Combo1.List(ListIndex)
End Function

Public Function PopulateListSQL(Conn As ADODB.Connection, tTableName As String, tFieldName As String, BlockedItems As Boolean, Optional vCatType As Integer, Optional DeletedItems As Boolean) As Boolean
  '... use blocked = true when u want to display products that are not blocked.
  '... use allocate = true when u want to display products that are not deleted.
  Dim C1 As Integer
  Dim C2 As Integer
  Dim C3 As String
  Dim strSql As String

  PopulateListSQL = False
  vError = ""
  On Error Resume Next
  Err.Clear
  Set TempRS = New ADODB.Recordset
  If TempRS.State = 1 Then Set TempRS = Nothing

  If vCatType = 0 Then
    If BlockedItems = False Then
      strSql = "Select Distinct ID," & tFieldName & " from " & tTableName & " ORDER BY " & tFieldName
    Else
      If DeletedItems Then
        strSql = "Select Distinct ID, Blocked, Deleted," & tFieldName & " from " & tTableName & " WHERE [Blocked] = " & False & " AND [DELETED] = " & False & " ORDER BY " & tFieldName
      Else
        strSql = "Select Distinct ID, Blocked," & tFieldName & " from " & tTableName & " WHERE [Blocked] = " & False & " ORDER BY " & tFieldName
      End If
    End If
  Else
    'only for category
    strSql = "Select Distinct ID, CatType, Blocked, Deleted," & tFieldName & " from " & tTableName & " WHERE [CatType] = " & vCatType & " and [Blocked] = " & False & " AND [DELETED] = " & False & " ORDER BY " & tFieldName
  End If

  TempRS.Open strSql, Conn, adOpenStatic, adLockOptimistic

  Combo1.Clear
  List1.Clear
  
  For C1 = 1 To TempRS.RecordCount
      C3 = TempRS.Fields(tFieldName).Value
      Combo1.AddItem C3
      List1.AddItem C3
      For C2 = 0 To Combo1.ListCount - 1
        If Combo1.List(C2) = C3 Then
          Combo1.ItemData(C2) = TempRS.Fields("ID").Value
          List1.ItemData(C2) = TempRS.Fields("ID").Value
        End If
      Next C2
      TempRS.MoveNext
    Next C1
    Combo1.Refresh
    List1.Refresh
    TempRS.Close
  
  Set TempRS = Nothing
  If Err <> 0 Then
    vError = Err.Description
    On Error GoTo 0
  End If
  On Error GoTo 0
  PopulateListSQL = True

End Function

Private Sub SearchList(vText As String)

  Dim j As Integer
  Dim strPart As String
  Dim iLoop As Integer
  Dim iStart As Integer
  Dim strItem As String
  Dim HighlightFirstFound As Boolean
  Dim HighlightedItem As Integer
  Const RemoveChar As String = "'"

  List1.Clear
  HighlightFirstFound = True
  For j = 1 To Combo1.ListCount
    If InStr(1, Trim$(Replace(UCase(Combo1.List(j - 1)), RemoveChar, "")), Trim$(Replace(UCase(vText), RemoveChar, ""))) > 0 Or _
      InStr(1, Trim$(Replace(UCase(Combo1.List(j - 1)), RemoveChar, "")), Trim$(UCase(vText))) > 0 Then
      List1.AddItem Combo1.List(j - 1)
      
      If HighlightFirstFound Then
        HighlightedItem = j - 1
        HighlightFirstFound = False
      End If
      
    End If
  Next j
  
  '...Do the autocomplete option
  If bAutoComplete And Not HighlightFirstFound And Not blnAuto And Trim$(vText) <> "" Then
    iStart = Text2.SelStart                    '...get the part the user has typed (not selected)
    strPart = Left$(Text2.Text, iStart)
    
    For iLoop = 0 To List1.ListCount - 1
      strItem = UCase$(List1.List(iLoop))
        If UCase$(strItem) Like UCase$(strPart & "*") And UCase$(strItem) <> UCase$(Text2.Text) Then
          'blnAuto = True
          Text2.SelText = Mid$(List1.List(iLoop), iStart + 1)   '...add on the new ending
          Text2.SelStart = iStart                               '...reset the selection
          Text2.SelLength = Len(Text2.Text) - iStart
          'blnAuto = False
        Exit For
      End If
    Next iLoop
    
  End If

End Sub

Private Function GetListIndex(vText As String) As Integer

  Dim j As Integer

  For j = 1 To Combo1.ListCount
    If UCase(Combo1.List(j - 1)) = UCase(vText) Then
      GetListIndex = j - 1
      Exit For
    End If
  Next j

End Function
