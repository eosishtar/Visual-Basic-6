VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Profiles"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   0
         Tag             =   "Username"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2400
         PasswordChar    =   "#"
         TabIndex        =   1
         Tag             =   "Password"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2400
         PasswordChar    =   "#"
         TabIndex        =   2
         Tag             =   "Confirm Password"
         Top             =   1800
         Width           =   2775
      End
      Begin CarListing.ctlDBCombo ctlDBCombo1 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
      End
      Begin VB.Label Label2 
         Caption         =   "User Code"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Username"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   1845
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   525
      Left            =   1560
      Picture         =   "frmUser.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5025
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   525
      Left            =   240
      Picture         =   "frmUser.frx":18CFE
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   6360
      Picture         =   "frmUser.frx":1B9D9
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   -1200
      Picture         =   "frmUser.frx":1E64C
      ScaleHeight     =   3015
      ScaleWidth      =   10455
      TabIndex        =   3
      Top             =   -1560
      Width           =   10455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Profiles "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Left            =   -1080
      TabIndex        =   12
      Top             =   1560
      Width           =   10395
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   4860
      Y2              =   4860
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   4785
      Y2              =   4800
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDelete_Click()
Dim Cancel As Integer
Dim tLoggedInUser As String

tLoggedInUser = Right(MDIForm1.Caption, Len(MDIForm1.Caption) - InStrRev(MDIForm1.Caption, "/"))
If Trim$(tLoggedInUser) = User.Username Then
  MsgBox User.Username & " is currently logged in. Logged in users can not be  deleted.", vbExclamation + vbOKOnly, Me.Caption
  Exit Sub
Else
  If Trim$(User.UserCode) = "0001" Then
    MsgBox "Super Administrator can not be deleted.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
  End If
End If

If MsgBox("Are you sure you want to delete user ''" & User.Username & "''?", vbQuestion + vbYesNo, "Delete ") = vbYes Then
  With rs
    sql = "Select * From Security WHERE [ID] = " & User.ID
      .Open sql, cn, adOpenForwardOnly, adLockOptimistic
        If .EOF Then rs.Close: Exit Sub
          rs!Blocked = True
        .Update
      .Close
      MsgBox User.Username & " has been successfully deleted!", vbInformation + vbOKOnly, "Deleted"
  End With
  
'  User.ID = 0
'  User.Username = ""
'  User.UserCode = ""
'  User.Password = ""
'  User.Deleted = False
  
  If ctlDBCombo1.PopulateList(cn, "Security", "UserCode", True) = False Then
    MsgBox ctlDBCombo1.Error
  End If
  ctlDBCombo1.AddItem "-- Add New --"
  Text1(0).Text = ""
  Text1(1).Text = ""
  Text1(2).Text = ""
  
  cmdDelete.Enabled = False
  
Else
  Cancel = 1
End If

End Sub

Private Sub cmdSave_Click()
  Dim ErrMsg As String
  Dim ErrString As String
  Dim i As Integer
  
  ErrMsg = "Required fields to be completed!" & vbCrLf
  For i = 1 To 3
    If Text1(i - 1).Text = "" Then
      ErrString = ErrString & " * " & Text1(i - 1).Tag & vbCrLf
    End If
  Next i
  
  If ErrString <> "" Then
    MsgBox ErrMsg & ErrString, vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
  End If
  
  If User.Username <> Text1(0).Text Then
    If ChkDuplicate(CleanSQL(Text1(0).Text), "Security") = True Then
      MsgBox "That username has already been used!.", vbExclamation + vbOKOnly, Me.Caption
      Exit Sub
    End If
  End If
  
  If Text1(1).Text <> Text1(2).Text Then
    MsgBox "The passwords you have entered do not exist.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
  End If

  With rs
    sql = "Select * from Security WHERE UserCode = " & Chr(34) & ctlDBCombo1.Text & Chr(34)
      .Open sql, cn, adOpenKeyset, adLockOptimistic
      
      If .EOF Then
        .AddNew
      End If
      
      If ctlDBCombo1.Text = "-- Add New --" Then
        rs!UserCode = GetNextUserCode
      End If
      rs!Username = Text1(0).Text
      rs!Password = Crypt(Text1(1).Text, ENCRYPT_KEY)
      rs!Blocked = False
      
      .Update
      .Close
      
  End With

  User.UserCode = ctlDBCombo1.Text
  User.Username = Text1(0).Text
  User.Password = Crypt(Text1(1).Text, ENCRYPT_KEY)
  User.Deleted = False
  
  MsgBox "User details have been successfully updated. ", vbInformation + vbOKOnly, Me.Caption
  
  Frame1.Caption = " Create New User "
  If ctlDBCombo1.PopulateList(cn, "Security", "UserCode", True) = False Then
    MsgBox ctlDBCombo1.Error
  End If
  ctlDBCombo1.AddItem "-- Add New --"
  Text1(0).Text = ""
  Text1(1).Text = ""
  Text1(2).Text = ""

End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub ctlDBCombo1_ItemSelected(ItemID As Long, ItemName As String)

  If Trim$(ItemName) = "-- Add New --" Then
    Text1(0).Text = ""
    Text1(1).Text = ""
    Text1(2).Text = ""
    cmdSave.Caption = "Create"
    'cmdSave.Picture = LoadPicture(GetAppPath & "Images\Add_16x16.jpg")
    cmdSave.Picture = MDIForm1.ImageList1.ListImages(1).Picture
    cmdDelete.Enabled = False
    Frame1.Caption = " Create New User "
  Else
    GetUserDetails ItemName
    Text1(0).Text = User.Username
    Text1(1).Text = User.Password
    Text1(2).Text = Text1(1).Text
    cmdSave.Caption = "Update"
    'cmdSave.Picture = LoadPicture(GetAppPath & "Images\Check_16x16.jpg")
    cmdSave.Picture = MDIForm1.ImageList1.ListImages(7).Picture
    cmdDelete.Enabled = True
    Frame1.Caption = " Update " & User.Username & " "
  End If

End Sub

Private Sub Form_Load()

  Call CenterForm(Me)

  Frame1.Caption = " Create New User "
  If ctlDBCombo1.PopulateList(cn, "Security", "UserCode", True) = False Then
    MsgBox ctlDBCombo1.Error
  End If
  ctlDBCombo1.AddItem "-- Add New --"

End Sub


