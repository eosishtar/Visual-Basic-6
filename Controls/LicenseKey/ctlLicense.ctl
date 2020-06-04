VERSION 5.00
Begin VB.UserControl LicenseKeyControl 
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ScaleHeight     =   1020
   ScaleWidth      =   6870
   ToolboxBitmap   =   "ctlLicense.ctx":0000
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   0
      MaxLength       =   4
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   0
      MaxLength       =   4
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "LicenseKeyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const LicSpacing As Integer = 25
Const cTextWidth As Integer = 200

Dim i As Integer
Dim sMaxLicBoxes As Integer
Dim sLicenseKey As String
Dim sLicenseKeyLength As Integer


Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).SelStart = Len(Text1(Index))
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

  If Index = 1 And KeyCode = 17 And Shift = 2 Then       'User press Ctrl + V in 1st text box
    PasteLicenseKey (Clipboard.GetText)
  End If

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
  
  Dim Y As Integer
  
  KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Change all entry to Uppercase
  
  If KeyAscii = 8 Then
    If Len(Text1(Index)) = 0 Then     'no more data, skip to next text box
      If (Index) > 1 Then
        Text1(Index - 1).SetFocus
      End If
    Else
      Text1(Index).Text = Left(Text1(Index).Text, Len(Text1(Index).Text) - 1)
    End If
  End If

End Sub

Private Sub UserControl_Initialize()
  sMaxLicBoxes = 1
End Sub

Private Sub ClearLicenseBoxes()
  
  For i = 1 To sMaxLicBoxes
    Text1(i).Text = ""
    Text1(i).TabStop = True
    Text1(i).TabIndex = i
  Next i
  
End Sub

Private Sub PasteLicenseKey(sLicKey As String)
  Dim Cnt As Integer

  sLicenseKey = sLicKey
  If sLicenseKey = "" Then Exit Sub
  
  Cnt = 1
  For i = 1 To sMaxLicBoxes
    Text1(i).Text = Mid(sLicenseKey, Cnt, 4)
    Cnt = Cnt + sLicenseKeyLength
  Next i
    
End Sub

Public Property Get LicenseKey() As String
  LicenseKey = sLicenseKey
End Property

Public Property Let LicenseKey(ByVal vNewValue As String)
  sLicenseKey = vNewValue
  PasteLicenseKey (sLicenseKey)
End Property

Public Property Get LicKeyHolders() As Integer
  LicKeyHolders = sMaxLicBoxes
End Property

Public Property Let LicKeyHolders(ByVal vNewValue As Integer)
  Dim OldValue As Integer
  
  If vNewValue < 1 Then
    vNewValue = 1
  ElseIf vNewValue > 8 Then
    vNewValue = 8
  End If
  
  OldValue = sMaxLicBoxes
  sMaxLicBoxes = vNewValue
  
  Call LoadTextBoxes(OldValue, sMaxLicBoxes)
End Property

Public Property Get LicKeyLength() As Integer
  LicKeyLength = sLicenseKeyLength
End Property

Public Property Let LicKeyLength(ByVal vNewValue As Integer)
  If vNewValue < 1 Then
    vNewValue = 1
  ElseIf vNewValue > 8 Then
    vNewValue = 8
  End If
  sLicenseKeyLength = vNewValue
  SetMaxLength vNewValue
  Call ResizeControl
End Property

Private Sub SetMaxLength(vValue As Integer)
  For i = 1 To sMaxLicBoxes
    Text1(i).MaxLength = vValue
  Next i
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '...read the property bags details
  sMaxLicBoxes = PropBag.ReadProperty("sMaxLicBoxes", 1)
  sLicenseKey = PropBag.ReadProperty("sLicenseKey", "")
  sLicenseKeyLength = PropBag.ReadProperty("sLicenseKeyLength", 4)

  '..Set properties
  Call LoadTextBoxes(sMaxLicBoxes, sMaxLicBoxes)
  Call ClearLicenseBoxes
  Call SetMaxLength(sLicenseKeyLength)
  Call PasteLicenseKey(sLicenseKey)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  '...write the property bags details
  Call PropBag.WriteProperty("sMaxLicBoxes", sMaxLicBoxes)
  Call PropBag.WriteProperty("sLicenseKey", sLicenseKey)
  Call PropBag.WriteProperty("sLicenseKeyLength", sLicenseKeyLength)
  
End Sub

Private Sub UserControl_Resize()
  'ResizeControl
End Sub

Private Sub ResizeControl()
  Dim sTextBoxWidth As Single
  ' 1 Char 300 Width
  ' 2 Char 400 Width
  ' 3 Char 500 Width


  If UserControl.Height <> 315 Then UserControl.Height = 315
  sTextBoxWidth = cTextWidth + (sLicenseKeyLength * 100)     'Width of One Textbox
  
  UserControl.Width = (sTextBoxWidth * sMaxLicBoxes) + (LicSpacing * sMaxLicBoxes)
  Text1(1).Move 0, 0, sTextBoxWidth, UserControl.Height

  
  If sMaxLicBoxes > 1 Then
    For i = 2 To sMaxLicBoxes
      Text1(i).Move Text1(i - 1).Left + Text1(1).Width + LicSpacing, Text1(1).Top, Text1(1).Width, Text1(1).Height
    Next i
  End If
  

End Sub

Private Sub LoadTextBoxes(iUnloadAmt As Integer, iLoadAmt As Integer)
  
  '...unload controls
  For i = 1 To sMaxLicBoxes
    On Error Resume Next
    Unload Text1(i)
  Next i
  
  '...load new controls
  For i = 1 To sMaxLicBoxes
    Load Text1(i)
    Text1(i).Visible = True
  Next i
  
  On Error GoTo 0
  
  '...resize control
  ResizeControl
End Sub


