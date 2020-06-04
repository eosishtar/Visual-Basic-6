VERSION 5.00
Begin VB.UserControl ctlToggle 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ScaleHeight     =   450
   ScaleWidth      =   3315
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   325
      Left            =   1560
      TabIndex        =   2
      Top             =   10
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   0
      Top             =   75
      Width           =   315
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "ctlToggle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim vON As Boolean
Dim vON_OFF As Boolean
Dim vTEXT_ON As String
Dim vTEXT_OFF As String
Dim vFONT_SIZE As Single
Dim vBUTTON_COLOR As OLE_COLOR
Dim vCOLOR_ON As OLE_COLOR
Dim vCOLOR_OFF As OLE_COLOR
Dim vENABLED As Boolean


Public Event SWITCHCLICK(vValue As Boolean)

Private Sub Label2_Click()

  If Not vENABLED Then
    Exit Sub
  End If

  If vON Then
    vON = False
    Call ResizeControl
  Else
    vON = True
    Call ResizeControl
  End If

  '...raise event
  RaiseEvent SWITCHCLICK(vON)
  
End Sub

Private Sub UserControl_Initialize()

  Screen.MousePointer = vbNormal
  Call UserControl_Resize

End Sub


Private Sub UserControl_Resize()

If Label1.Height < 100 Then Label3.FontSize = 6.25
If Label1.Height < 200 Then Label3.FontSize = 8.25
If Label1.Height > 200 Then Label3.FontSize = 9
If Label1.Height > 300 Then Label3.FontSize = 10
If Label1.Height > 400 Then Label3.FontSize = 12
If Label1.Height > 500 Then Label3.FontSize = 14
If Label1.Height > 600 Then Label3.FontSize = 16
If Label1.Height > 700 Then Label3.FontSize = 18
If Label1.Height > 800 Then Label3.FontSize = 20

'...set min / max values for control
UserControl.FontSize = Label3.FontSize
If UserControl.Width < UserControl.TextWidth(Label3.Caption) Then UserControl.Width = UserControl.TextWidth(Label3.Caption)
If UserControl.Height < UserControl.TextHeight(Label3.Caption) Then UserControl.Height = UserControl.TextHeight(Label3.Caption)

'...move control into place
Label1.Move 0, 0, UserControl.Width, UserControl.Height
Call ResizeControl

End Sub


Public Sub ResizeControl()


  If Not vENABLED Then
    Label1.BackColor = vbButtonFace
    Exit Sub
  End If

  If vON Then
    Label3.Caption = vTEXT_ON
    'switch label
    
    
    Label2.Width = (Label1.Width / 2)
    Label2.Move Label1.Left, Label1.Top + 10, Label2.Width, Label1.Height - 50
    Label3.Move (UserControl.Width / 2) + ((UserControl.Width / 2) - Label3.Width) / 2, (UserControl.Height - Label3.Height) / 2
    
    'back colour label
    Label1.BackColor = vCOLOR_ON
  Else
    Label3.Caption = vTEXT_OFF
    'switch label
    
    Label2.Width = (Label1.Width / 2)
    Label2.Move Label1.Left + Label2.Width, Label1.Top + 10, Label2.Width, Label1.Height - 50
    Label3.Move ((UserControl.Width / 2) - Label3.Width) / 2, (UserControl.Height - Label3.Height) / 2
    
    'back colour label
    Label1.BackColor = vCOLOR_OFF
  End If


End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' read the property bags details
  vON_OFF = PropBag.ReadProperty("vON_OFF", True)
  vFONT_SIZE = PropBag.ReadProperty("vFONT_SIZE", "6.25")
  vTEXT_ON = PropBag.ReadProperty("vTEXT_ON", "ON")
  vTEXT_OFF = PropBag.ReadProperty("vTEXT_OFF", "OFF")
  vBUTTON_COLOR = PropBag.ReadProperty("vBUTTON_COLOR", vbGrayText)
  vCOLOR_ON = PropBag.ReadProperty("vCOLOR_ON", vbGreen)
  vCOLOR_OFF = PropBag.ReadProperty("vCOLOR_OFF", vbRed)
  vENABLED = PropBag.ReadProperty("vENABLED ", True)
  

  ' set default values
  If vFONT_SIZE = 0 Then vFONT_SIZE = 6.25
  If vBUTTON_COLOR = 0 Then vBUTTON_COLOR = &H8000000A
  If vCOLOR_ON = 0 Then vCOLOR_ON = vbGreen
  If vCOLOR_OFF = 0 Then vCOLOR_OFF = vbRed
  
  Label3.Visible = vON_OFF
  Label3.FontSize = vFONT_SIZE
  Label2.BackColor = vBUTTON_COLOR
  If vON Then Label1.BackColor = vCOLOR_ON Else Label1.BackColor = vCOLOR_OFF
  If vON Then Label3.Caption = vTEXT_ON Else Label3.Caption = vTEXT_OFF
  
  'after props have loaded, resize it
  Call UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' write the property bags details
  Call PropBag.WriteProperty("vON_OFF", vON_OFF)
  Call PropBag.WriteProperty("vFONT_SIZE", vFONT_SIZE)
  Call PropBag.WriteProperty("vBUTTON_COLOR", vBUTTON_COLOR)
  Call PropBag.WriteProperty("vCOLOR_ON", vCOLOR_ON)
  Call PropBag.WriteProperty("vCOLOR_OFF", vCOLOR_OFF)
  Call PropBag.WriteProperty("vTEXT_ON", vTEXT_ON)
  Call PropBag.WriteProperty("vTEXT_OFF", vTEXT_OFF)
  Call PropBag.WriteProperty("vENABLED", vENABLED)
  
End Sub

Public Property Get ON_OFF() As Boolean
  ON_OFF = vON_OFF
End Property

Public Property Let ON_OFF(ByVal vNewValue As Boolean)
  vON_OFF = vNewValue
  Label3.Visible = vON_OFF
End Property

Public Property Get FONT_SIZE() As Single
  FONT_SIZE = vFONT_SIZE
End Property

Public Property Let FONT_SIZE(ByVal vNewValue As Single)
  vFONT_SIZE = vNewValue
  Label3.FontSize = vFONT_SIZE
End Property

Public Property Get BUTTON_COLOR() As OLE_COLOR
  BUTTON_COLOR = vBUTTON_COLOR
End Property

Public Property Let BUTTON_COLOR(ByVal vNewValue As OLE_COLOR)
  vBUTTON_COLOR = vNewValue
  Label2.BackColor = vBUTTON_COLOR
End Property

Public Property Get COLOR_ON() As OLE_COLOR
  COLOR_ON = vCOLOR_ON
End Property

Public Property Let COLOR_ON(ByVal vNewValue As OLE_COLOR)
  vCOLOR_ON = vNewValue
  Label1.BackColor = vCOLOR_ON
End Property

Public Property Get COLOR_OFF() As OLE_COLOR
  COLOR_OFF = vCOLOR_OFF
End Property

Public Property Let COLOR_OFF(ByVal vNewValue As OLE_COLOR)
  vCOLOR_OFF = vNewValue
  Label1.BackColor = vCOLOR_OFF
End Property

Public Property Get TEXT_ON() As String
  TEXT_ON = vTEXT_ON
End Property

Public Property Let TEXT_ON(ByVal vNewValue As String)
  vTEXT_ON = vNewValue
  Label3.Caption = vTEXT_ON
End Property

Public Property Get TEXT_OFF() As String
  TEXT_OFF = vTEXT_OFF
End Property

Public Property Let TEXT_OFF(ByVal vNewValue As String)
  vTEXT_OFF = vNewValue
  Label3.Caption = vTEXT_OFF
End Property

Public Property Get xVALUE() As Boolean
  xVALUE = vON
End Property

Public Property Let xVALUE(ByVal vNewValue As Boolean)
  vON = vNewValue
  Call ResizeControl
End Property

Public Property Get xENABLED() As Boolean
  xENABLED = vENABLED
End Property

Public Property Let xENABLED(ByVal vNewValue As Boolean)
  vENABLED = vNewValue
  Call ResizeControl
End Property

