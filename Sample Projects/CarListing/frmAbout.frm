VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3630
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6360
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505.49
   ScaleMode       =   0  'User
   ScaleWidth      =   5972.369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Width           =   525
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblDescription 
         Caption         =   "App Description"
         ForeColor       =   &H00000000&
         Height          =   915
         Left            =   240
         TabIndex        =   2
         Top             =   1335
         Width           =   3165
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   4995
      Picture         =   "frmAbout.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2955
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   2
      X1              =   112.686
      X2              =   5972.369
      Y1              =   1863.588
      Y2              =   1863.588
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   112.686
      X2              =   5972.369
      Y1              =   1905.001
      Y2              =   1905.001
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub




Private Sub Form_Load()
  Dim ImgPath As String
  
  'test if image exists
  ImgPath = GetAppPath & "MyLogo.jpg"
  If Dir(ImgPath) = "" Then
    Image1.Picture = LoadPicture()
  Else
    Image1.Picture = LoadPicture(GetAppPath & "MyLogo.jpg")
  End If

  Me.Caption = "About " & App.Title
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle.Caption = App.Title
  lblDescription.Caption = "Developed By : Mark Lang"
  lblDescription.Visible = True
  
End Sub



