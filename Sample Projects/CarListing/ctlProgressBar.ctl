VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ScaleHeight     =   1875
   ScaleWidth      =   3735
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.Label lblPerc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 % Complete"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1335
      TabIndex        =   2
      Top             =   990
      Width           =   960
   End
   Begin VB.Label lblFront 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   248
      TabIndex        =   1
      Top             =   960
      Width           =   15
   End
   Begin VB.Label lblBack 
      BackColor       =   &H8000000C&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim vForeColor As OLE_COLOR
Dim vBackColor As OLE_COLOR
Dim vTextColor As OLE_COLOR
Dim vPercView As Boolean
Dim vPercCaption As String
Dim vUnloadProgBar As Integer
Public Event TimeOut()



Private Sub Timer1_Timer()
  RaiseEvent TimeOut
  Timer1.Enabled = False
End Sub

Private Sub UserControl_Initialize()
  Timer1.Enabled = False
End Sub

Private Sub UserControl_Terminate()
  Timer1.Enabled = False
End Sub

'
' --- load property bags
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("vPercView", vPercView)
  Call PropBag.WriteProperty("vForeColor", vForeColor)
  Call PropBag.WriteProperty("vBackColor", vBackColor)
  Call PropBag.WriteProperty("vTextColor", vTextColor)
  Call PropBag.WriteProperty("vPercCaption", vPercCaption)
  Call PropBag.WriteProperty("vUnloadProgBar", vUnloadProgBar)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vPercView = PropBag.ReadProperty("vPercView", True)
  vForeColor = PropBag.ReadProperty("vForeColor", vbBlue)
  vBackColor = PropBag.ReadProperty("vBackColor", &H8000000C)
  vTextColor = PropBag.ReadProperty("vTextColor", vbWhite)
  vPercCaption = PropBag.ReadProperty("vPercCaption", "% Complete")
  vUnloadProgBar = PropBag.ReadProperty("vUnloadProgBar", "1")
  
  '...set values from propery bag.
  lblFront.BackColor = vForeColor
  lblBack.BackColor = vBackColor
  lblPerc.Visible = vPercView
  lblPerc.ForeColor = vTextColor
  lblPerc.Caption = "0 " & vPercCaption
  
End Sub

Private Sub UserControl_Resize()

If lblBack.Height < 200 Then lblPerc.FontSize = 8.25
If lblBack.Height > 200 Then lblPerc.FontSize = 9
If lblBack.Height > 300 Then lblPerc.FontSize = 10
If lblBack.Height > 400 Then lblPerc.FontSize = 12
If lblBack.Height > 500 Then lblPerc.FontSize = 14
If lblBack.Height > 600 Then lblPerc.FontSize = 16
If lblBack.Height > 700 Then lblPerc.FontSize = 18
If lblBack.Height > 800 Then lblPerc.FontSize = 20

'...set min / max values for control
UserControl.Font = lblPerc.Font
If UserControl.Width < UserControl.TextWidth(lblPerc.Caption) Then UserControl.Width = UserControl.TextWidth(lblPerc.Caption)
If UserControl.Height < UserControl.TextHeight(lblPerc.Caption) Then UserControl.Height = UserControl.TextHeight(lblPerc.Caption)

'...move control into place
lblBack.Move 0, 0, UserControl.Width, UserControl.Height
lblFront.Move lblBack.Left, lblBack.Top, lblFront.Width, lblBack.Height
lblPerc.Move (lblBack.Width - lblPerc.Width) / 2, (lblBack.Height - lblPerc.Height) / 2

End Sub

'...this function set the width of the front label
Public Function SetPerc(CurrRec As Integer, TotalRec As Integer)
  
  If TotalRec = 0 Or (CurrRec = TotalRec) Then
    lblFront.Width = lblBack.Width
    SetPercLabel (100)
  Else
    If CurrRec / TotalRec * 100 <= 100 Then
      lblFront.Width = (CurrRec / TotalRec) * lblBack.Width
      SetPercLabel (Round(CurrRec / TotalRec * 100, 2))
    Else
      lblFront.Width = lblBack.Width
      SetPercLabel (100)
    End If
  End If
  
  If lblFront.Width = lblBack.Width Then
    lblPerc.Caption = 100 & " " & vPercCaption
    ' all record done. make invisble after set time
    Timer1.Interval = 2000
    Timer1.Enabled = True
  End If
  
End Function

'... this function sets the % of the label
Private Function SetPercLabel(CurrRec As Integer)

If lblPerc.ForeColor = vForeColor Then lblPerc.ForeColor = vBackColor
If lblPerc.ForeColor = vBackColor Then lblPerc.ForeColor = vForeColor

lblPerc.Caption = CurrRec & " " & vPercCaption

End Function

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = vForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
  vForeColor = vNewValue
  lblFront.BackColor = vForeColor
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = vBackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  vBackColor = vNewValue
  lblBack.BackColor = vBackColor
End Property

Public Property Get PercColor() As OLE_COLOR
  PercColor = vTextColor
End Property

Public Property Let PercColor(ByVal vNewValue As OLE_COLOR)
  vTextColor = vNewValue
  lblPerc.ForeColor = vTextColor
End Property


Public Property Get PercView() As Boolean
  PercView = vPercView
End Property

Public Property Let PercView(ByVal vNewValue As Boolean)
  vPercView = vNewValue
  lblPerc.Visible = vPercView
End Property

Public Property Get PercCaption() As String
  PercCaption = vPercCaption
End Property

Public Property Let PercCaption(ByVal vNewValue As String)
  vPercCaption = vNewValue
  lblPerc.Caption = vPercCaption
End Property


