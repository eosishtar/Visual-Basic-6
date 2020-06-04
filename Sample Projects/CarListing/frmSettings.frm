VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings Manager"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10455
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      Picture         =   "frmSettings.frx":1601A
      ScaleHeight     =   2505
      ScaleWidth      =   10425
      TabIndex        =   23
      Top             =   0
      Width           =   10455
   End
   Begin VB.CommandButton cmdShowMore 
      Caption         =   "More Options"
      Height          =   525
      Left            =   4560
      Picture         =   "frmSettings.frx":33E48
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9000
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "frmSettings.frx":36CEF
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel Record"
      Top             =   9000
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   525
      Left            =   9000
      Picture         =   "frmSettings.frx":399CA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save Record"
      Top             =   9000
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   " Database Path "
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   10095
      Begin VB.CommandButton cmdFind 
         Height          =   435
         Left            =   8640
         Picture         =   "frmSettings.frx":3C63D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Find Database"
         Top             =   360
         Width           =   1245
      End
      Begin VB.TextBox txtDataPath 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         TabIndex        =   1
         Top             =   390
         Width           =   8295
      End
   End
   Begin VB.CommandButton cmdShowLess 
      Caption         =   "Back"
      Height          =   525
      Left            =   4560
      Picture         =   "frmSettings.frx":3F2DC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9000
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame4 
      Caption         =   " Settings "
      Height          =   4455
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   10095
      Begin VB.CommandButton cmdCompact 
         Caption         =   "DB Compact"
         Height          =   525
         Left            =   1440
         Picture         =   "frmSettings.frx":42085
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Compact the Database"
         Top             =   3840
         Width           =   1245
      End
      Begin VB.CommandButton cmdDatabaseBU 
         Caption         =   "DB Backup"
         Height          =   525
         Left            =   120
         Picture         =   "frmSettings.frx":4520E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Backup the Database"
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Frame Frame2 
         Caption         =   " Default Printer "
         Height          =   945
         Left            =   150
         TabIndex        =   18
         Top             =   360
         Width           =   4515
         Begin VB.ComboBox cboPrinter 
            Height          =   315
            Left            =   330
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   360
            Width           =   3705
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Company Details "
         Height          =   3825
         Left            =   4830
         TabIndex        =   7
         Top             =   390
         Width           =   5175
         Begin VB.TextBox txtCompName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1650
            TabIndex        =   12
            Top             =   360
            Width           =   3315
         End
         Begin VB.TextBox txtContactPerson 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1650
            TabIndex        =   11
            Top             =   855
            Width           =   3315
         End
         Begin VB.TextBox txtTelephone 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1650
            TabIndex        =   10
            Top             =   1350
            Width           =   3315
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            Height          =   1335
            Left            =   1650
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   2340
            Width           =   3315
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1650
            TabIndex        =   8
            Top             =   1845
            Width           =   3315
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Company Name"
            Height          =   225
            Left            =   120
            TabIndex        =   17
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Telephone"
            Height          =   225
            Left            =   120
            TabIndex        =   16
            Top             =   1410
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Person"
            Height          =   225
            Left            =   120
            TabIndex        =   15
            Top             =   930
            Width           =   1245
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Address"
            Height          =   225
            Left            =   120
            TabIndex        =   14
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Email"
            Height          =   225
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdLogo 
         Caption         =   "Select Logo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3510
         TabIndex        =   6
         ToolTipText     =   "Select Logo"
         Top             =   3150
         Width           =   1050
      End
      Begin VB.Image imgLogo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2235
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1380
         Width           =   4515
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "More Settings"
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   10095
      Begin VB.Frame Frame6 
         Caption         =   "Banking Details"
         Height          =   1695
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   3615
         Begin VB.TextBox txtBank 
            Appearance      =   0  'Flat
            Height          =   1095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Terms and Conditions"
         Height          =   2055
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   9855
         Begin VB.TextBox txtTerms 
            Appearance      =   0  'Flat
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   480
            Width           =   9615
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings Manager"
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
      Height          =   525
      Left            =   0
      TabIndex        =   21
      Top             =   2640
      Width           =   10395
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   10350
      Y1              =   8805
      Y2              =   8805
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   10350
      Y1              =   8760
      Y2              =   8760
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Response As Integer
Dim Success As Boolean
Dim Products() As String
Dim vError As String

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCompact_Click()
  DBCompact Main.Database
  If cn = "" Then dbase.Go
End Sub

Private Sub cmdDatabaseBU_Click()
    
  Dim MyBackupPath As String
  Dim Resp As String
  Dim vPath As String
    
  Success = False
  If Main.Database = "" Then Exit Sub
  
  '...warn user, system needs sole access
  Resp = MsgBox("You are about to backup the database. Please make sure no one is connected!!  " & Chr(13) & vbTab & "Continue ?  ", vbCritical + vbYesNo, "Backup Database")
  If Resp = vbYes Then

  '...Check path exist first
  
   vPath = GetAppPath & "BackUp\"
    
    MyBackupPath = GetAppPath & "BackUp\" & Format(Now, "yyyymmddHHnnss") & "_" & App.EXEName & ".bak"
    
    On Error Resume Next
    Set fso = New FileSystemObject
    fso.CopyFile Main.Database, MyBackupPath
    SaveSetting App.EXEName, "DataPath", "LastBackUp", Format(Now, "yyyymmddHHnnss")
    Success = True
    On Error GoTo 0

    If Err = 0 And Success = True Then
      MsgBox "Database succesfully backed up to " & MyBackupPath, vbInformation + vbOKOnly, Me.Caption
    Else
      MsgBox "An error has occurred. " & vbCrLf & Err.Number & " - " & Err.Description, vbInformation + vbOKOnly, Me.Caption
      Err.Clear
    End If
  End If
  
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub cmdFind_Click()

  With CD1
    .Filter = "Database Files (*.mdb)|*.mdb"
    .ShowOpen
    Main.Database = .FileName
  End With
  
  If Main.Database = "" Then Exit Sub
    
  SaveSetting App.EXEName, "DataPath", "DataBase", Main.Database
  txtDataPath.Text = Main.Database
  
End Sub

Private Sub cmdLogo_Click()
  imgLogo.Visible = True
   
  With CD1
    .DialogTitle = "Select a Logo"
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist  ' Set flags
    .Filter = "Jpeg Logo's (*.jpg)|*.jpg|"
    .ShowOpen
    Main.Logo = .FileName
'      .CancelError = True
  End With
 
  If Main.Logo = "" Then Exit Sub
  
  Set fso = New FileSystemObject
  fso.CopyFile Main.Logo, GetAppPath & "MyLogo.jpg", True
  imgLogo.Picture = LoadPicture(GetAppPath & "MyLogo.jpg")
  SaveSetting App.EXEName, "DataPath", "LogoPath", Main.Logo

End Sub

Private Sub cmdSave_Click()
  Dim i As Integer
  Dim ErrString As String
  Dim ErrItems As String
  Dim sString As String
  
  ErrString = "The following errors were encountered."
  ErrItems = ""

'  If txtDataPath.Text = "" Or cboPrinter.Text = "" Then
  If txtDataPath.Text = "" Then
      MsgBox "Please complete all fields first.", vbInformation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    If txtDataPath.Text = "" Then cmdFind.TabIndex = 0: Exit Sub
    If cboPrinter = "" Then cboPrinter.TabIndex = 0: Exit Sub
  End If
  
  If txtCompName.Text = "" Then
    MsgBox "Please enter the company name first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    txtCompName.SetFocus
    Exit Sub
  End If
  
  If txtContactPerson.Text = "" Then
    MsgBox "Please enter the company contact person first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    txtContactPerson.SetFocus
    Exit Sub
  End If
  
  If txtTelephone.Text = "" Then
    MsgBox "Please enter the company telephone first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    txtTelephone.SetFocus
    Exit Sub
  End If
  
  If txtAddress.Text = "" Then
    MsgBox "Please enter the company address first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    txtAddress.SetFocus
    Exit Sub
  End If
  
  If txtEmail.Text = "" Then
    MsgBox "Please enter the company email address first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions False
    txtEmail.SetFocus
    Exit Sub
  Else
    '...validate email
    If Not ValidEmail(txtEmail.Text) Then
      Response = MsgBox("It seems that your email address is incorrect. Are you sure you want to continue?", vbInformation + vbYesNo, Me.Caption)
      If Response <> 6 Then
        ShowMoreOptions False
        txtEmail.SetFocus
        Exit Sub
      End If
    End If
  End If
  
  If txtBank.Text = "" Then
    MsgBox "Please enter the banking details first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions True
    txtBank.SetFocus
    Exit Sub
  End If
  
  If txtTerms.Text = "" Then
    MsgBox "Please enter the terms and conditions first.", vbExclamation + vbOKOnly, Me.Caption
    ShowMoreOptions True
    txtTerms.SetFocus
    Exit Sub
  End If
 
  '...save company settings
    sql = "Select * from CompanyDetails where ID = 1"
    With rs
      .Open sql, cn, adOpenKeyset, adLockOptimistic

      If .EOF Then
        .AddNew
      End If
      
      rs!CompanyName = Trim(txtCompName.Text)
      rs!Telephone = Trim(txtTelephone.Text)
      rs!Address = Trim(txtAddress.Text)
      rs!Person = Trim(txtContactPerson.Text)
      rs!Email = Trim(txtEmail.Text)
      rs!BankingDetails = Trim(txtBank.Text)
      rs!TermsConditions = Trim(txtTerms.Text)
        
      rs.Update
    rs.Close
    End With
  
  '... save everything into the class
  Company.Name = Trim(txtCompName.Text)
  Company.Person = Trim(txtContactPerson.Text)
  Company.Address = Trim(txtAddress.Text)
  Company.Telephone = Trim(txtTelephone.Text)
  Company.Email = Trim(txtEmail.Text)
  Company.BankingDetails = txtBank.Text
  Company.TermsConditions = txtTerms.Text
 
  Main.Database = txtDataPath.Text
  Main.Printer = cboPrinter.Text
  SaveSetting App.EXEName, "DataPath", "DataBase", Main.Database
  SaveSetting App.EXEName, "DataPath", "DefaultPrinter", Main.Printer
  
  MsgBox "Company Settings have been saved successfully.", vbInformation + vbOKOnly, Me.Caption
  Unload Me
  
End Sub

Private Sub cmdShowLess_Click()

  ShowMoreOptions False

End Sub

Private Sub cmdShowMore_Click()

  ShowMoreOptions True
  
End Sub

Private Function CloseDB() As Boolean
  
  vError = ""
  CloseDB = False
  On Error Resume Next
  
  If rs.State = 1 Then rs.Close
  If cn.State = 1 Then cn.Close

  Set rs = Nothing
  Set cn = Nothing
  
  If Err.Number = 0 Then
    CloseDB = True
  Else
    vError = "An error has occurred. " & vbCrLf & vbCrLf & "Error Nr:  " & Err.Number & vbCrLf & "Error Desc:  " & Err.Description
    Err.Clear
  End If

End Function

Public Function DBCompact(ByVal DBName1 As String, Optional ByVal DBName2 As String)
'.... requires reference to Microsoft Jet and Replications Objects
  Dim JRO As New JRO.JetEngine
  Dim db_Old As String
  Dim db_New As String
  Dim myLen As Integer
  Dim Resp As String
  Dim s1 As Long, s2 As Long

  db_Old = DBName1
  myLen = Len(DBName1)
  myLen = myLen - 4
  db_New = Left(DBName1, myLen)
  db_New = db_New + "_Temp.mdb"

  'get file size before compact
  vError = ""
  Close #1
  Open db_Old For Binary As #1
  s1 = LOF(1)
  Close #1

  Set JRO = New JetEngine
  '...warn user, system needs sole access
  Resp = MsgBox("You are about to compact the database. Please make sure no one is connected!!  " & Chr(13) & vbTab & "Continue ?  ", vbCritical + vbYesNo, Screen.ActiveForm.Caption)
  If Resp = vbYes Then
    Screen.MousePointer = vbHourglass
  
    '...close the connections
    If Not CloseDB Then
      Screen.MousePointer = vbNormal
      MsgBox vError
      Exit Function
    End If
    
    On Error Resume Next
    Set JRO = New JRO.JetEngine
      JRO.CompactDatabase "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & db_Old & ";Jet OLEDB:Database Password=Starlight1", _
      "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & db_New & ";Jet OLEDB:Database Password=Starlight1"
    FileCopy db_Old, db_New
    Kill db_Old
    Name db_New As db_Old
    
    'get file size after compact
    Open db_Old For Binary As #2
    s2 = LOF(2)
    Close #2
    
    vError = "Compact complete " & vbCrLf & vbCrLf & "Size:    " & Round(s2 / 1024 / 1024, 2) & "Mb"
  End If
  
  Screen.MousePointer = vbNormal
  MsgBox vError

On Error GoTo 0
  Set JRO = Nothing

End Function

Private Sub Form_Load()
  Dim prt As Printer
  
  txtDataPath.Locked = True
  Call CenterForm(Me)
  For Each prt In Printers
    cboPrinter.AddItem prt.DeviceName
  Next
  
  If Main.Database = "" Then Exit Sub
  
  GetCompanyDetails
  txtCompName.Text = Company.Name
  txtContactPerson.Text = Company.Person
  txtAddress.Text = Company.Address
  txtTelephone.Text = Company.Telephone
  txtEmail.Text = Company.Email
  txtBank.Text = Company.BankingDetails
  txtTerms.Text = Company.TermsConditions
  txtDataPath.Text = Main.Database
  
  cboPrinter.ListIndex = SetComboText(cboPrinter, Main.Printer)
  
  '...display selected logo
  On Error Resume Next
  imgLogo.Picture = LoadPicture(GetAppPath & "MyLogo.jpg")
  On Error GoTo 0
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set fso = Nothing
End Sub

Private Sub txtAddress_GotFocus()

  cmdSave.Default = False

End Sub

Private Sub txtAddress_LostFocus()

  cmdSave.Default = True

End Sub

Private Sub ShowMoreOptions(ShowVar As Boolean)

  cmdShowMore.Visible = Not ShowVar
  cmdShowLess.Visible = ShowVar
  Frame5.Visible = ShowVar
  Frame4.Visible = Not ShowVar

End Sub

Private Sub txtBank_GotFocus()

  cmdSave.Default = False

End Sub

Private Sub txtBank_LostFocus()

  cmdSave.Default = True

End Sub

Private Sub txtTerms_GotFocus()

  cmdSave.Default = False

End Sub

Private Sub txtTerms_LostFocus()

  cmdSave.Default = True

End Sub
