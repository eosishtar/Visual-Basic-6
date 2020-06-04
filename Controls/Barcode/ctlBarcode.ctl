VERSION 5.00
Begin VB.UserControl ctlBarcode 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ScaleHeight     =   1560
   ScaleWidth      =   3255
   ToolboxBitmap   =   "ctlBarcode.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1305
      Left            =   120
      ScaleHeight     =   1305
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "ctlBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim bShowText As Boolean
Dim sBarText As String
Dim vCOLOR_BACK As OLE_COLOR
Public Enum BARCODE_SIZE
  sSmall = 0
  sMeduim = 1
  sLarge = 2
End Enum
Dim bBCodeSize As BARCODE_SIZE

Public Event BarcodeCreated()

Private Sub DrawBarcode(ByVal bc_string As String, obj As Object)

  Dim xpos!, Y1!, Y2!, dw%, Th!, tw, new_string$
  Dim C As String
  Dim bc_pattern As String
  Dim n As Integer
  Dim i As Integer
  Dim bc(90) As String

  If bc_string = "" Then obj.Cls: Exit Sub
  '...define barcode patterns
  bc(1) = "1 1221"            'pre-amble
  bc(2) = "1 1221"            'post-amble
  bc(48) = "11 221"           'digits
  bc(49) = "21 112"
  bc(50) = "12 112"
  bc(51) = "22 111"
  bc(52) = "11 212"
  bc(53) = "21 211"
  bc(54) = "12 211"
  bc(55) = "11 122"
  bc(56) = "21 121"
  bc(57) = "12 121"
                              'capital letters
  bc(65) = "211 12"           'A
  bc(66) = "121 12"           'B
  bc(67) = "221 11"           'C
  bc(68) = "112 12"           'D
  bc(69) = "212 11"           'E
  bc(70) = "122 11"           'F
  bc(71) = "111 22"           'G
  bc(72) = "211 21"           'H
  bc(73) = "121 21"           'I
  bc(74) = "112 21"           'J
  bc(75) = "2111 2"           'K
  bc(76) = "1211 2"           'L
  bc(77) = "2211 1"           'M
  bc(78) = "1121 2"           'N
  bc(79) = "2121 1"           'O
  bc(80) = "1221 1"           'P
  bc(81) = "1112 2"           'Q
  bc(82) = "2112 1"           'R
  bc(83) = "1212 1"           'S
  bc(84) = "1122 1"           'T
  bc(85) = "2 1112"           'U
  bc(86) = "1 2112"           'V
  bc(87) = "2 2111"           'W
  bc(88) = "1 1212"           'X
  bc(89) = "2 1211"           'Y
  bc(90) = "1 2211"           'Z
                              'Misc
  bc(32) = "1 2121"           'space
  bc(35) = ""                 '# cannot do!
  bc(36) = "1 1 1 11"         '$
  bc(37) = "11 1 1 1"         '%
  bc(43) = "1 11 1 1"         '+
  bc(45) = "1 1122"           '-
  bc(47) = "1 1 11 1"         '/
  bc(46) = "2 1121"           '.
  bc(64) = ""                 '@ cannot do!
  'A Fix made by changing 65 to 42.
  bc(42) = "1 1221"           '*
  
  bc_string = UCase(bc_string)
  
  'dimensions
  obj.ScaleMode = 3                               'pixels
  obj.Cls
  obj.Picture = Nothing
  dw = CInt(obj.ScaleHeight / 40)                 'space between bars
  If dw < 1 Then dw = 1
  'Debug.Print dw
  Th = obj.TextHeight(bc_string)                  'text height
  tw = obj.TextWidth(bc_string)                   'text width
  new_string = Chr$(1) & bc_string & Chr$(2)      'add pre-amble, post-amble
  
  Y1 = obj.ScaleTop
  Y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * Th
  obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
  
  
  'draw each character in barcode string
  xpos = obj.ScaleLeft
  For n = 1 To Len(new_string)
      C = Asc(Mid$(new_string, n, 1))
      If C > 90 Then C = 0
      bc_pattern$ = bc(C)
      
      'draw each bar
      For i = 1 To Len(bc_pattern$)
          Select Case Mid$(bc_pattern$, i, 1)
              Case " "
                  'space
                  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                  xpos = xpos + dw
                  
              Case "1"
                  'space
                  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                  xpos = xpos + dw
                  'line
                  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
                  xpos = xpos + dw
              
              Case "2"
                  'space
                  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                  xpos = xpos + dw
                  'wide line
                  obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
                  xpos = xpos + 2 * dw
          End Select
      Next
  Next
  
  '1 more space
  obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
  xpos = xpos + dw
  
  'final size and text
  obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
  obj.CurrentX = (obj.ScaleWidth - tw) / 2
  obj.CurrentY = Y2 + 0.25 * Th
  If bShowText Then
    obj.Print bc_string
  End If
  
  Select Case bBCodeSize
    Case 0
        Picture1.Height = Picture1.Height * (1.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 8
    Case 1
        Picture1.Height = Picture1.Height * (2.4 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 10
    Case 2
        Picture1.Height = Picture1.Height * (3 * 40 / Picture1.ScaleHeight)
        Picture1.FontSize = 14
  End Select
  
  UserControl.Width = obj.Width
  RaiseEvent BarcodeCreated     'Fired after successful creation of barcode

End Sub

Public Sub GenBarcode(ByVal sBarTaxt As String)
 Call DrawBarcode(sBarTaxt, Picture1)
End Sub

Public Sub ClearBarcode()
  Picture1.Cls
End Sub

Private Sub UserControl_Initialize()
  UserControl_Resize
End Sub

Private Sub UserControl_Resize()
  If bShowText Then
    Picture1.Move 0, 0, UserControl.Width, UserControl.Height - 100
  Else
    Picture1.Move 0, 0, UserControl.Width, UserControl.Height
  End If
End Sub

' --- load property bags
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("bShowText", bShowText)
  Call PropBag.WriteProperty("vCOLOR_BACK", vCOLOR_BACK)
  Call PropBag.WriteProperty("sBarText", sBarText)
   Call PropBag.WriteProperty("bBCodeSize", bBCodeSize)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  bShowText = PropBag.ReadProperty("bShowText", True)
  vCOLOR_BACK = PropBag.ReadProperty("vCOLOR_BACK", vbButtonFace)
  sBarText = PropBag.ReadProperty("sBarText", "")
  bBCodeSize = PropBag.ReadProperty("bBCodeSize", 1)
  
  Picture1.BackColor = vCOLOR_BACK
  BAR_TEXT = sBarText
End Sub

Public Property Get SHOW_TEXT() As Boolean
  SHOW_TEXT = bShowText
End Property

Public Property Let SHOW_TEXT(ByVal vNewValue As Boolean)
  bShowText = vNewValue
  Call UserControl_Resize
  Call GenBarcode(sBarText)
End Property

Public Property Get BAR_TEXT() As String
  BAR_TEXT = sBarText
End Property

Public Property Let BAR_TEXT(ByVal vNewValue As String)
  sBarText = vNewValue
  Call GenBarcode(sBarText)
End Property

Public Property Get BARCODE_BACKCOLOR() As OLE_COLOR
  BARCODE_BACKCOLOR = vCOLOR_BACK
End Property

Public Property Let BARCODE_BACKCOLOR(ByVal vNewValue As OLE_COLOR)
  vCOLOR_BACK = vNewValue
  Picture1.BackColor = vCOLOR_BACK
  Call GenBarcode(sBarText)
End Property

Public Property Get BARCODE_SIZE() As BARCODE_SIZE
  BARCODE_SIZE = bBCodeSize
End Property

Public Property Let BARCODE_SIZE(ByVal vNewValue As BARCODE_SIZE)
  bBCodeSize = vNewValue
  Call GenBarcode(sBarText)
End Property
