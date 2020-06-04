VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   360
      Picture         =   "frmLogin.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   525
      Left            =   3840
      Picture         =   "frmLogin.frx":18CF5
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2040
      PasswordChar    =   "#"
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   2512
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Tag             =   "Username"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   3495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   3735
      Left            =   120
      Top             =   120
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3550
      Left            =   200
      Top             =   210
      Width           =   5020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   2535
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdCancel_Click()
  
  EndProgram
  
End Sub

Private Sub cmdSave_Click()
Dim ErrString As String
Dim ErrMsg As String
Dim Success As Boolean

ErrMsg = "Required fields."
Success = False

For i = 1 To 2
  If Text1(i - 1).Text = "" Then
    ErrString = ErrString & "* " & Text1(i - 1).Tag & vbCrLf
  End If
Next i

If ErrString <> "" Then
  MsgBox ErrMsg & vbCrLf & ErrString, vbInformation + vbOKOnly, "System Error..."
  Exit Sub
End If

With rs
  sql = "Select * from Security WHERE Username = " & Chr(34) & CleanSQL(Text1(0).Text) & Chr(34)
  .Open sql, cn, adOpenKeyset, adLockOptimistic
  If .EOF Then
    .Close
    MsgBox "Incorrect login details entered.", vbInformation + vbOKOnly, "System Error..."
    Exit Sub
  End If
  
  '...check to see if user is blocked
 If rs!Blocked = True Then
  .Close
  MsgBox "User " & Trim$(Text1(0).Text) & " has been deleted.", vbInformation + vbOKOnly, ""
  Exit Sub
 End If
  
 If Decrypt(rs!Password, ENCRYPT_KEY) = Trim$(CleanSQL(Text1(1).Text)) Then
  'login success
    User.ID = rs!ID
    User.Username = rs!Username
    User.UserCode = rs!UserCode
    User.Password = rs!Password
    .Close
    Success = True
    MDIForm1.Caption = App.EXEName & " (Ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")" & "  /" & User.Username
    
    SaveSetting App.EXEName, "User", "LastUser", User.Username
    SetMenu True

  Else
  'login fail
    .Close
    MDIForm1.Caption = App.EXEName & " (Ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    MsgBox "Incorrect login details entered.", vbInformation + vbOKOnly, "System Error..."
    Text1(0).Text = ""
    Text1(0).TabStop = True
    Text1(0).SetFocus
    Text1(1).Text = ""
  End If
End With


If Success Then
  Unload Me
Else
  Set frmLogin = Nothing
End If
  

End Sub

Private Sub Form_Load()
  Dim Luser As String

  SetMenu False

  Luser = GetSetting(App.EXEName, "User", "LastUser")
  If Luser <> "" Then
    Text1(0).Text = Luser
    Text1(0).TabStop = False
  End If
  
  Call CenterForm(Me)

End Sub
