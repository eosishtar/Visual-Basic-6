VERSION 5.00
Begin VB.UserControl ctlTextbox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "ctlTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

Public Enum enumTextBoxChar
  tbAllCharacters = 0
  tbOnlyNumbers = 1
  tbOnlyCharacter = 2
End Enum
Private mTextBoxMode As enumTextBoxChar
Public Enum enumTextAlignment
  taLeft = 0
  taRight = 1
  taCentre = 2
End Enum
Private mTextAlignment As enumTextAlignment
Public Enum enumAppearance
  taFlat = 0
  ta3D = 1
End Enum
Private mAppearance As enumAppearance
Public Enum enumBorderStyle
  taNone = 0
  taFixedSingle = 1
End Enum
Private mBorderStyle As enumBorderStyle

Private mForceUcase As Boolean
Private mBackColor As OLE_COLOR
Private mForeColor As OLE_COLOR
Private mPasswordChar As String


Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Let Text(value As String)
  Text1.Text = value
  PropertyChanged "Text"
End Property

Public Property Get Text() As String
  Text = Text1.Text
End Property
   
Public Property Let TextBoxMode(value As enumTextBoxChar)
  mTextBoxMode = value
End Property

Public Property Get TextBoxMode() As enumTextBoxChar
  TextBoxMode = mTextBoxMode
End Property

Public Property Set Font(value As StdFont)
  With mFont
    .Name = value.Name
    .Size = value.Size
    .Bold = value.Bold
    .Italic = value.Italic
    .Strikethrough = value.Strikethrough
    .Underline = value.Underline
  End With
  Set UserControl.Font = mFont
  PropertyChanged "Font"
End Property

Public Property Get Font() As StdFont
  Set Font = mFont
End Property

Public Property Let MaxLength(value As Integer)
  Text1.MaxLength = value
  PropertyChanged "MaxLength"
End Property

Public Property Get MaxLength() As Integer
  MaxLength = Text1.MaxLength
End Property

Public Property Let ForceUCase(value As Boolean)
  mForceUcase = value
End Property

Public Property Get ForceUCase() As Boolean
  ForceUCase = mForceUcase
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = mBackColor
End Property

Public Property Let BackColor(vNewValue As OLE_COLOR)
  mBackColor = vNewValue
  Text1.BackColor = vNewValue
  PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = mForeColor
End Property

Public Property Let ForeColor(vNewValue As OLE_COLOR)
  mForeColor = vNewValue
  Text1.ForeColor = vNewValue
  PropertyChanged "ForeColor"
End Property

Public Property Let BorderStyle(vNewValue As enumBorderStyle)
  mBorderStyle = vNewValue
  Text1.BorderStyle = vNewValue
  PropertyChanged "BorderStyle"
End Property

Public Property Get BorderStyle() As enumBorderStyle
  BorderStyle = mBorderStyle
End Property

Public Property Let Alignment(value As enumTextAlignment)
  mTextAlignment = value
  Text1.Alignment = mTextAlignment
  PropertyChanged "Alignment"
End Property

Public Property Get Alignment() As enumTextAlignment
  Alignment = mTextAlignment
End Property

Public Property Let Appearance(vNewValue As enumAppearance)
  mAppearance = vNewValue
  Text1.Appearance = mAppearance
  PropertyChanged "Appearance"
End Property

Public Property Get Appearance() As enumAppearance
  Appearance = mAppearance
End Property

Public Property Get MultiLine() As Boolean
  MultiLine = Text1.MultiLine
End Property

Public Property Let PasswordChar(vNewValue As String)
  Text1.PasswordChar = vNewValue
  mPasswordChar = vNewValue
  PropertyChanged "PasswordChar"
End Property

Public Property Get PasswordChar() As String
  PasswordChar = mPasswordChar
End Property

Private Sub Text1_Change()
  RaiseEvent Change
End Sub

Private Sub Text1_Click()
  RaiseEvent Click
End Sub

Private Sub Text1_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'...65-90 for a upper caption
'...97-122 for a lower caption
'...8 for backspace
'...32 for space

  Select Case mTextBoxMode
    Case tbAllCharacters
      'Allow Anything
      
    Case tbOnlyNumbers
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
        'Only Numbers Allowed
      Else
        KeyAscii = 0
      End If
      
    Case tbOnlyCharacter
      Select Case KeyAscii
        Case 65 To 90, 97 To 122, 8, 32
          'Only Characters Allowed
        Case Else
          KeyAscii = 0
        End Select
      
  End Select
  
  '...Force to Uppercase
  If mForceUcase Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  End If
  RaiseEvent KeyPress(KeyAscii)
  
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
  Set mFont = New StdFont
  ResizeControl
End Sub

Private Sub UserControl_InitProperties()
  Set mFont = UserControl.Parent.Font
  Set Text1.Font = mFont
  BackColor = vbWhite
  ForeColor = Text1.ForeColor
  Enabled = True
End Sub

Private Sub mFont_FontChanged(ByVal PropertyName As String)
  Set UserControl.Font = mFont
  Set Text1.Font = mFont
  UserControl.Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Enabled Then
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
  Set mFont = PropBag.ReadProperty("Font", UserControl.Parent.Font)
  Text = PropBag.ReadProperty("Text", "Text1")
  Set Text1.Font = mFont
  Enabled = PropBag.ReadProperty("Enabled", True)
  MaxLength = PropBag.ReadProperty("MaxLength", 0)
  mTextBoxMode = PropBag.ReadProperty("mTextBoxMode", 0)
  mForceUcase = PropBag.ReadProperty("mForceUcase", False)
  mTextAlignment = PropBag.ReadProperty("TextAlignment", 0)
  mAppearance = PropBag.ReadProperty("Appearance", 1)
  mBorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  mBackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
  mForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
  mPasswordChar = PropBag.ReadProperty("PasswordChar", "")
  
  
  '...Set the selected properties
  Text1.BorderStyle = mBorderStyle
  Text1.Alignment = mTextAlignment
  Text1.Appearance = mAppearance
  Text1.BackColor = mBackColor
  Text1.ForeColor = mForeColor
  Text1.PasswordChar = mPasswordChar

End Sub

Private Sub UserControl_Resize()
  ResizeControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "Font", mFont
  PropBag.WriteProperty "Text", Text
  PropBag.WriteProperty "Enabled", Enabled
  PropBag.WriteProperty "MaxLength", MaxLength
  PropBag.WriteProperty "mTextBoxMode", mTextBoxMode
  PropBag.WriteProperty "mForceUcase", mForceUcase
  PropBag.WriteProperty "TextAlignment", mTextAlignment
  PropBag.WriteProperty "Appearance", mAppearance
  PropBag.WriteProperty "BorderStyle", mBorderStyle
  PropBag.WriteProperty "BackColor", mBackColor
  PropBag.WriteProperty "ForeColor", mForeColor
  PropBag.WriteProperty "PasswordChar", mPasswordChar

End Sub

Private Sub ResizeControl()
  Text1.Move 0, 0, UserControl.Width, UserControl.Height
End Sub
