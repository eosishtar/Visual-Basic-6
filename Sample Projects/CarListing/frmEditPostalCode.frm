VERSION 5.00
Begin VB.Form frmEditPostalCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Postal Code"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   Icon            =   "frmEditPostalCode.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      Picture         =   "frmEditPostalCode.frx":000C
      ScaleHeight     =   3255
      ScaleWidth      =   10455
      TabIndex        =   13
      Top             =   -1920
      Width           =   10455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   6120
      Picture         =   "frmEditPostalCode.frx":1DE3A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   525
      Left            =   120
      Picture         =   "frmEditPostalCode.frx":20AAD
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4560
      Width           =   1245
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   525
      Left            =   1440
      Picture         =   "frmEditPostalCode.frx":23788
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   " Postal Code Details "
      Height          =   2055
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   5655
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2520
         TabIndex        =   10
         Tag             =   "Suburb Postal Code"
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2520
         TabIndex        =   9
         Tag             =   "Suburb"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         Tag             =   "Area Postal Code"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Tag             =   "Area"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Salesman Code"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2190
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Suburb Postal Code"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Tag             =   "Suburb Postal Code"
         Top             =   1605
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suburb"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Tag             =   "Suburb"
         Top             =   1245
         Width           =   510
      End
      Begin VB.Label Label4 
         Caption         =   "Area Postal Code"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Tag             =   "Area Postal Code"
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Area "
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Tag             =   "Area"
         Top             =   495
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Postal Code"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   -600
      TabIndex        =   12
      Top             =   1440
      Width           =   8715
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   -840
      X2              =   7560
      Y1              =   4305
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -840
      X2              =   7560
      Y1              =   4380
      Y2              =   4380
   End
End
Attribute VB_Name = "frmEditPostalCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cmdDelete_Click()
  Dim Cancel As Integer
  
  If MsgBox("Are you sure you want to delete this Postal Code ''" & PCode.PostalCode & "''?", vbQuestion + vbYesNo, "Delete ") = vbYes Then
    With rs
      sql = "Select * From PCodes WHERE [ID] = " & PCode.ID
        .Open sql, cn, adOpenForwardOnly, adLockOptimistic
          .Delete
          .Update
        .Close
        MsgBox PCode.PostalCode & " has been successfully deleted!", vbInformation + vbOKOnly, "Deleted"
    End With
    
    PCode.ID = 0
    PCode.PostalCode = ""
    Unload Me
    frmPCodes.LoadPCodes ""

  Else
    Cancel = 1
  End If

End Sub

Private Sub cmdSave_Click()
  Dim i As Integer
  Dim vUpdate As Boolean
    
  vUpdate = True
  
  If Text1(1).Text = "" And Text1(3).Text = "" Then
    MsgBox "Please enter at least one Area / Suburb.", vbInformation + vbOKOnly
    Exit Sub
  End If
  If Text1(2).Text = "" And Text1(4).Text = "" Then
    MsgBox "Please enter at least one Postal Code.", vbInformation + vbOKOnly
    Exit Sub
  End If
      
  With rs
     sql = "select ID,Area,AreaCode,Suburb,SuburbCode from PCodes where ID = " & PCode.ID
      .Open sql, cn, adOpenKeyset, adLockOptimistic
        If .EOF Then
          .AddNew
          vUpdate = False
        End If
        
        If Trim(Text1(1)) = "" Then Text1(1) = NonRequired                                 'Area
        If Trim(Text1(2)) = "" Then Text1(2) = NonRequired                                 'Area
        If Trim(Text1(3)) = "" Then Text1(3) = NonRequired                                 'Area
        If Trim(Text1(4)) = "" Then Text1(4) = NonRequired                                 'Area
        
        rs.Fields(1) = Text1(1).Text
        rs.Fields(2) = Text1(2).Text
        rs.Fields(3) = Text1(3).Text
        rs.Fields(4) = Text1(4).Text
      .Update
    .Close
  End With
  
  If Not vUpdate Then
    MsgBox "New Postal Code successfully added", vbInformation + vbOKOnly
  Else
    MsgBox "Postal Code successfully updated.", vbInformation + vbOKOnly
  End If
  
  Unload Me
  frmPCodes.LoadPCodes ""

End Sub

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  If PCode.ID = 0 Then
    Label1.Caption = "Add New Postal Code"
    Me.Caption = Label1.Caption
    cmdDelete.Enabled = False
  Else
    Label1.Caption = "Edit Postal Code"
    Me.Caption = Label1.Caption
    GetPostalDetails PCode.ID
  End If
  
End Sub

Public Sub GetPostalDetails(ID As Integer)
  Dim i As Integer
  
  With rs
    sql = "select ID,Area,AreaCode,SuburbCode,Suburb from PCodes where ID = " & ID
      .Open sql, cn, adOpenKeyset, adLockOptimistic
      If .EOF Then
        .Close
        'lblRecords.Caption = Counter(tRecCount, "Postal Codes")
        Exit Sub
      End If
      
      If Not IsNull(rs.Fields(1).Value) Then Text1(1) = rs.Fields(1).Value         'Area
      If Not IsNull(rs.Fields(2).Value) Then Text1(2) = rs.Fields(2).Value         'Area Code
      If Not IsNull(rs.Fields(4).Value) Then Text1(3) = rs.Fields(4).Value         'Suburb
      If Not IsNull(rs.Fields(3).Value) Then Text1(4) = rs.Fields(3).Value         'Sub Code
      
      PCode.ID = ID
      If Trim(Text1(2).Text) = "" Then
        PCode.PostalCode = Text1(4).Text
      Else
        If Trim(Text1(4).Text) = "" Then
          PCode.PostalCode = Text1(2).Text
        End If
      End If
      
    .Close
  End With
  
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
    Case 2, 4
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
      Else
        KeyAscii = 0
      End If
    Case 1, 3
      KeyAscii = Asc(UCase(Chr(KeyAscii)))    'Change all entry to Uppercase
      'Text1(Index).Text = Text1(Index).Text & UCase(KeyAscii)
    Case Else
    ' fail over
  End Select
  
End Sub
