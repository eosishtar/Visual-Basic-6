VERSION 5.00
Begin VB.UserControl ProgressBar 
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   ScaleHeight     =   4245
   ScaleWidth      =   7470
   ToolboxBitmap   =   "ctlProgressBar.ctx":0000
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
      TabIndex        =   0
      Top             =   990
      Width           =   960
   End
   Begin VB.Shape lblFront 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   240
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape lblBack 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   240
      Top             =   960
      Width           =   3135
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vForeColor As OLE_COLOR
Dim vBackColor As OLE_COLOR
Dim vTextColor As OLE_COLOR
Dim vPercView As Boolean
Dim vPercCaption As String
Dim vUnloadProgBar As Integer

Public Event TimeOut()
Public Event CurrentProgress(CurrentRecord As Integer, TotalRecords As Integer)

Public Enum PercShape
  sRectangle = 0
  sRoundedCorners = 4
End Enum
Dim vPercShape As PercShape

Public Enum PercType
  PercentageOfProgress = 0
  ActualItemOfItems = 1
End Enum
Dim vPercType As PercType

Private Sub Timer1_Timer()
  RaiseEvent TimeOut
  Timer1.Enabled = False
End Sub

Private Sub UserControl_Initialize()
  Timer1.Enabled = False
  ResetProgress
  UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
  Timer1.Enabled = False
End Sub

' --- load property bags
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("vPercView", vPercView)
  Call PropBag.WriteProperty("vForeColor", vForeColor)
  Call PropBag.WriteProperty("vBackColor", vBackColor)
  Call PropBag.WriteProperty("vTextColor", vTextColor)
  Call PropBag.WriteProperty("vPercCaption", vPercCaption)
  Call PropBag.WriteProperty("vUnloadProgBar", vUnloadProgBar)
  Call PropBag.WriteProperty("vPercType", vPercType)
  Call PropBag.WriteProperty("vPercShape", vPercShape)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vPercView = PropBag.ReadProperty("vPercView", True)
  vForeColor = PropBag.ReadProperty("vForeColor", vbBlue)
  vBackColor = PropBag.ReadProperty("vBackColor", &H8000000C)
  vTextColor = PropBag.ReadProperty("vTextColor", vbWhite)
  vPercCaption = PropBag.ReadProperty("vPercCaption", "% Complete")
  vUnloadProgBar = PropBag.ReadProperty("vUnloadProgBar", "1")
  vPercType = PropBag.ReadProperty("vPercType", "1")
  vPercShape = PropBag.ReadProperty("vPercShape", "0")
  
  '...set values from propery bag.
  lblFront.BackColor = vForeColor
  lblFront.Shape = vPercShape
  lblBack.BackColor = vBackColor
  lblBack.Shape = vPercShape
  lblPerc.Visible = vPercView
  lblPerc.ForeColor = vTextColor
  lblPerc.Caption = "0 " & vPercCaption
    
End Sub

Private Sub UserControl_Resize()

  '...set min / max values for control
  UserControl.Font = lblPerc.Font
  If UserControl.Width < UserControl.TextWidth(lblPerc.Caption) Then UserControl.Width = UserControl.TextWidth(lblPerc.Caption)
  If UserControl.Height < UserControl.TextHeight(lblPerc.Caption) Then UserControl.Height = UserControl.TextHeight(lblPerc.Caption)
  
  '...move control into place
  lblBack.Move 0, 0, UserControl.Width, UserControl.Height
  lblFront.Move lblBack.Left, lblBack.Top, lblFront.Width, lblBack.Height
  lblPerc.Move (lblBack.Width - lblPerc.Width) / 2, (lblBack.Height - lblPerc.Height) / 2

  DoLabelSize

End Sub

'...this function set the width of the front label
Public Function SetPerc(CurrRec As Integer, TotalRec As Integer)
  
  If TotalRec = 0 Or (CurrRec = TotalRec) Then
    lblFront.Width = lblBack.Width
    SetPercLabel 100, 100
  Else
    If CurrRec / TotalRec * 100 <= 100 Then
      lblFront.Width = (CurrRec / TotalRec) * lblBack.Width
      If vPercType = 0 Then
        SetPercLabel Round(CurrRec / TotalRec * 100, 2), TotalRec   'Displaying % of work done
      Else
        SetPercLabel Round(CurrRec), TotalRec                       'Displaying Actual Item of Items
      End If
      RaiseEvent CurrentProgress(CurrRec, TotalRec)
    Else
      lblFront.Width = lblBack.Width
      SetPercLabel 100, 100
    End If
  End If
  
  If lblFront.Width = lblBack.Width Then
    '...all record done - make invisble after set time
    If vPercType = 0 Then
      lblPerc.Caption = 100 & " " & Trim(vPercCaption)
    Else
      If vPercCaption = "" Then
        lblPerc.Caption = CurrRec & " of " & TotalRec
      Else
        lblPerc.Caption = CurrRec & " of " & TotalRec & " " & Trim(vPercCaption)
      End If
    End If
    Timer1.Interval = 2000
    Timer1.Enabled = True
  End If
  
End Function

Public Sub ResetProgress()
  SetPerc 0, 1
End Sub

'... this function sets the % of the label
Private Function SetPercLabel(CurrRec As Integer, TotalRec As Integer)

  If lblPerc.ForeColor = vForeColor Then lblPerc.ForeColor = vBackColor
  If lblPerc.ForeColor = vBackColor Then lblPerc.ForeColor = vForeColor
  
  If vPercType = 0 Then
    '...Customtext
    lblPerc.Caption = CurrRec & " " & Trim(vPercCaption)
  Else
    '...item 1 of 2
    If vPercCaption = "" Then
      lblPerc.Caption = CurrRec & " of " & TotalRec
    Else
      lblPerc.Caption = CurrRec & " of " & TotalRec & " " & Trim(vPercCaption)
    End If
  End If
  Call DoLabelSize

End Function

'...this function resize the label accordingly
Private Sub DoLabelSize()
  If lblBack.Height < 200 Then lblPerc.FontSize = 8.25
  If lblBack.Height > 200 And lblBack.Height < 300 Then lblPerc.FontSize = 9
  If lblBack.Height > 300 And lblBack.Height < 400 Then lblPerc.FontSize = 10
  If lblBack.Height > 400 And lblBack.Height < 500 Then lblPerc.FontSize = 12
  If lblBack.Height > 500 And lblBack.Height < 600 Then lblPerc.FontSize = 14
  If lblBack.Height > 600 And lblBack.Height < 700 Then lblPerc.FontSize = 16
  If lblBack.Height > 700 And lblBack.Height < 800 Then lblPerc.FontSize = 18
  If lblBack.Height > 800 Then lblPerc.FontSize = 20
  
  lblPerc.Move (lblBack.Width - lblPerc.Width) / 2, (lblBack.Height - lblPerc.Height) / 2
End Sub

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = vForeColor
End Property

Public Property Let ForeColor(ByVal vNewvalue As OLE_COLOR)
  vForeColor = vNewvalue
  lblFront.BackColor = vForeColor
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = vBackColor
End Property

Public Property Let BackColor(ByVal vNewvalue As OLE_COLOR)
  vBackColor = vNewvalue
  lblBack.BackColor = vBackColor
End Property

Public Property Get PERC_COLOR() As OLE_COLOR
  PERC_COLOR = vTextColor
End Property

Public Property Let PERC_COLOR(ByVal vNewvalue As OLE_COLOR)
  vTextColor = vNewvalue
  lblPerc.ForeColor = vTextColor
End Property

Public Property Get PERC_VIEW() As Boolean
  PERC_VIEW = vPercView
End Property

Public Property Let PERC_VIEW(ByVal vNewvalue As Boolean)
  vPercView = vNewvalue
  lblPerc.Visible = vPercView
End Property

Public Property Get PERC_CAPTION() As String
  PERC_CAPTION = vPercCaption
End Property

Public Property Let PERC_CAPTION(ByVal vNewvalue As String)
  vPercCaption = vNewvalue
  lblPerc.Caption = vPercCaption
End Property

Public Property Get PERC_TYPE() As PercType
  PERC_TYPE = vPercType
End Property

Public Property Let PERC_TYPE(ByVal vNewvalue As PercType)
  vPercType = vNewvalue
  SetPercLabel 0, 1
End Property

Public Property Get PERC_SHAPE() As PercShape
  PERC_SHAPE = vPercShape
End Property

Public Property Let PERC_SHAPE(ByVal vNewvalue As PercShape)
  vPercShape = vNewvalue
  lblBack.Shape = vNewvalue
End Property


