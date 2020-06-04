VERSION 5.00
Begin VB.UserControl ctlDone 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   900
   ScaleHeight     =   435
   ScaleWidth      =   900
   Begin VB.Label lblNo 
      Alignment       =   2  'Center
      Caption         =   "û"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblYes 
      Alignment       =   2  'Center
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "ctlDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Sub UserControl_Resize()

  lblYes.Move 0, 0, UserControl.Width, UserControl.Height
  lblNo.Move 0, 0, UserControl.Width, UserControl.Height
  lblYes.Visible = False

End Sub

Public Property Get Done() As Boolean

  Done = lblYes.Visible

End Property

Public Property Let Done(ByVal vNewValue As Boolean)

  lblYes.Visible = vNewValue
  lblNo.Visible = Not vNewValue
  'lblNo.Visible = False

End Property

Public Sub RedCross()

lblNo.Visible = True
lblYes.Visible = False

End Sub
