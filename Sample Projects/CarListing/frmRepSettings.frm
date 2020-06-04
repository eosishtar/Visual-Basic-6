VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Settings"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   Icon            =   "frmRepSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   10470
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      Picture         =   "frmRepSettings.frx":1601A
      ScaleHeight     =   2505
      ScaleWidth      =   10425
      TabIndex        =   13
      Top             =   -120
      Width           =   10455
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   " Menu Options "
      Height          =   3375
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   7335
      Begin CarListing.ctlDBCombo ctlDBCombo1 
         Height          =   315
         Left            =   2790
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
      End
      Begin VB.CheckBox chkSubHead 
         Alignment       =   1  'Right Justify
         Caption         =   "Blocked"
         Height          =   195
         Left            =   1320
         TabIndex        =   5
         Top             =   2130
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkSubHeadTots 
         Alignment       =   1  'Right Justify
         Caption         =   "Include Sub Heading Totals"
         Height          =   195
         Left            =   1320
         TabIndex        =   4
         Top             =   2460
         Width           =   1695
      End
      Begin VB.CheckBox chkHourPat 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Hours Patrolled"
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   2790
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   2790
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         _Version        =   393216
         Format          =   114229251
         CurrentDate     =   42239
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   2790
         TabIndex        =   7
         Top             =   1095
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   609
         _Version        =   393216
         Format          =   3801091
         CurrentDate     =   42239
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Report from "
         Height          =   195
         Left            =   690
         TabIndex        =   10
         Top             =   645
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Report to"
         Height          =   195
         Left            =   690
         TabIndex        =   9
         Top             =   1140
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Category Type"
         Height          =   195
         Left            =   690
         TabIndex        =   8
         Top             =   1605
         Visible         =   0   'False
         Width           =   1845
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "E&xport"
      Default         =   -1  'True
      Height          =   525
      Left            =   9060
      Picture         =   "frmRepSettings.frx":33E48
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "frmRepSettings.frx":36D92
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7080
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report Manager"
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
      TabIndex        =   14
      Top             =   2520
      Width           =   10395
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
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
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   2115
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   6885
      Y2              =   6885
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   6840
      Y2              =   6840
   End
End
Attribute VB_Name = "frmRepSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSubHead_Click()

  Select Case vReportMenu
    Case 5      'voucher report menu
      If chkSubHead.Value <> 0 Then
        chkSubHeadTots.Enabled = False
        chkSubHeadTots.Value = 0
      Else
        chkSubHeadTots.Enabled = True
        chkSubHeadTots.Value = 0
      End If
  End Select
  
End Sub

Private Sub chkSubHeadTots_Click()
  
  Select Case vReportMenu
    Case 5      'voucher report menu
      If chkSubHeadTots.Value <> 0 Then
        chkSubHead.Enabled = False
        chkSubHead.Value = 0
      Else
        chkSubHead.Enabled = True
        chkSubHead.Value = 0
      End If
  End Select
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub Form_Load()

  Call CenterForm(Me)
  DTPicker1.Value = Format(Date, "dd MMMM yyyy")
  DTPicker2.Value = Format(Date, "dd MMMM yyyy")

  '...set the menu options
  List1.AddItem "Client Listing"
  List1.AddItem "Deal Listing"
  
  MenuOptionsVisible False

End Sub

'... menu options
Private Sub List1_Click()
  Dim i As Integer

  For i = 1 To List1.ListCount
    If List1.Selected(i - 1) = True Then
      SetupReportMenu i
      Exit For
    End If
  Next i

End Sub

Public Sub MenuOptionsVisible(vValue As Boolean)
  '...all controls set to false first here
  
  Label1.Visible = vValue
  Label3.Visible = vValue
  Label4.Visible = vValue
  chkSubHead.Visible = vValue
  chkSubHeadTots.Visible = vValue
  chkHourPat.Visible = vValue
  DTPicker1.Visible = vValue
  DTPicker2.Visible = vValue
  ctlDBCombo1.Visible = vValue
'  ctlDBCombo2.Visible = vValue
  
End Sub

Public Sub SetupReportMenu(vMenu As Integer)
  
  MenuOptionsVisible False
  Frame1.Caption = List1.List(vMenu - 1) & " Options... "
  vReportMenu = vMenu     'global for reporting
  ReportHead = List1.List(vMenu - 1)
  
  
  Select Case vMenu
    Case 1      'client listing
      Label1.Visible = True
      Label1.Caption = "Date From"
      DTPicker1.Visible = True
      DTPicker1.Top = Label1.Top - ((DTPicker1.Height - Label1.Height) / 2)
      DTPicker1.Value = Format(Date, "dd MMMM yyyy")
      
      Label3.Visible = True
      Label3.Caption = "Date To"
      DTPicker2.Visible = True
      DTPicker2.Top = Label3.Top - ((DTPicker2.Height - Label3.Height) / 2)
      DTPicker2.Value = Format(Date, "dd MMMM yyyy")
      
    Case 2   'Deal listing
    
      Label1.Visible = True
      Label1.Caption = "Date From"
      DTPicker1.Visible = True
      DTPicker1.Top = Label1.Top - ((DTPicker1.Height - Label1.Height) / 2)
      DTPicker1.Value = Format(Date, "dd MMMM yyyy")
      
      Label3.Visible = True
      Label3.Caption = "Date To"
      DTPicker2.Visible = True
      DTPicker2.Top = Label3.Top - ((DTPicker2.Height - Label3.Height) / 2)
      DTPicker2.Value = Format(Date, "dd MMMM yyyy")
      
      Label4.Visible = True
      Label4.Caption = "Deal"
      ctlDBCombo1.Visible = True
      ctlDBCombo1.Top = Label4.Top - ((ctlDBCombo1.Height - Label4.Height) / 2)
      If ctlDBCombo1.PopulateList(cn, "Deals", "ID", False) = False Then
        MsgBox ctlDBCombo1.Error
      End If
      ctlDBCombo1.AddItem "-- All Deals --"
      ctlDBCombo1.ListIndex 0
    
    Case 3      'product listing
      Label1.Visible = True
      Label1.Caption = "Products"
'      ctlDBCombo1.ClearList
'      ctlDBCombo1.Visible = True
'      ctlDBCombo1.Top = Label1.Top - ((ctlDBCombo1.Height - Label1.Height) / 2)
'      ctlDBCombo1.AddItem "-- All Products --"
'      ctlDBCombo1.ListIndex 0
      
      chkSubHead.Visible = True
      chkSubHead.Caption = "Exclude Blocked"
      chkSubHead.Top = ctlDBCombo1.Top + ctlDBCombo1.Height + 100
          
    Case 4      'show rating
      Label1.Visible = True
      Label1.Caption = "Date From"
      DTPicker1.Visible = True
      DTPicker1.Top = Label1.Top - ((DTPicker1.Height - Label1.Height) / 2)
      DTPicker1.Value = Format(Date, "dd MMMM yyyy")
      
      Label3.Visible = True
      Label3.Caption = "Date To"
      DTPicker2.Visible = True
      DTPicker2.Top = Label3.Top - ((DTPicker2.Height - Label3.Height) / 2)
      DTPicker2.Value = Format(Date, "dd MMMM yyyy")
      
      Label4.Visible = True
      Label4.Caption = "Show"
'      ctlDBCombo1.Visible = True
'      ctlDBCombo1.Top = Label4.Top - ((ctlDBCombo1.Height - Label4.Height) / 2)
'      If ctlDBCombo1.PopulateList(cn, "Category", "Description", False) = False Then
'        MsgBox ctlDBCombo1.Error
'      End If
'      ctlDBCombo1.AddItem "-- All Shows --"
'      ctlDBCombo1.ListIndex 0
      
      chkSubHead.Visible = True
      chkSubHead.Caption = "Show Avg Rating"
      chkSubHead.Top = ctlDBCombo1.Top + ctlDBCombo1.Height + 100
      
    Case 5      'voucher
      Label1.Visible = True
      Label1.Caption = "Date From"
      DTPicker1.Visible = True
      DTPicker1.Top = Label1.Top - ((DTPicker1.Height - Label1.Height) / 2)
      DTPicker1.Value = Format(Date, "dd MMMM yyyy")
      
      Label3.Visible = True
      Label3.Caption = "Date To"
      DTPicker2.Visible = True
      DTPicker2.Top = Label3.Top - ((DTPicker2.Height - Label3.Height) / 2)
      DTPicker2.Value = Format(Date, "dd MMMM yyyy")
      
      Label4.Visible = True
      Label4.Caption = "Show"
'      ctlDBCombo1.Visible = True
'      ctlDBCombo1.Top = Label4.Top - ((ctlDBCombo1.Height - Label4.Height) / 2)
'      If ctlDBCombo1.PopulateList(cn, "Category", "Description", False) = False Then
'        MsgBox ctlDBCombo1.Error
'      End If
'      ctlDBCombo1.AddItem "-- All Shows --"
'      ctlDBCombo1.ListIndex 0
      
      chkSubHead.Visible = True
      chkSubHead.Caption = "Include Winners"
      chkSubHead.Top = ctlDBCombo1.Top + ctlDBCombo1.Height + 100
      
      chkSubHeadTots.Visible = True
      chkSubHeadTots.Caption = "All Winners Only"
      chkSubHeadTots.Top = chkSubHead.Top + chkSubHead.Height + 100
      
  End Select


End Sub

Private Sub cmdExport_Click()
  Dim theSql As String
  Dim RptInt As Integer
  Dim DateFrom As Single
  Dim DateTo As Single
  
  ' check n make sure something is selected
  If List1.ListIndex = -1 Then
    MsgBox "Please select a report first.", vbInformation + vbOKOnly, Me.Caption
    Exit Sub
  End If
  
  DateFrom = GetDateVal(DTPicker1)
  DateTo = GetDateVal(DTPicker2)

  Select Case vReportMenu
    Case 1    'client listing
      theSql = "SELECT Deals.SellerFirstName, Deals.SellerLastName, Deals.SellerIDNumber, Deals.SellerCompanyName, Deals.SellerCompanyRegNr, Deals.SellerContactNr, Deals.SellerAltContactNr, Deals.SellerEmailAddress, Deals.SellerNotes, Deals.DateCreated FROM Deals WHERE (Deals.DateCreated >= " & Chr(39) & DateFrom & Chr(39) & " And Deals.DateCreated <= " & Chr(39) & DateTo & Chr(39) & ") ORDER BY [DateCreated]"
      RptInt = 3
         
    Case 2    'Deal listing
      If ctlDBCombo1.Text = "-- All Deals --" Then
        theSql = "SELECT Deals.ID, Deals.VehicleID, Deals.VehicleRegNr, Deals.VehicleDateBought, Deals.VehicleDateSold, Deals.VehicleCost, Deals.VehicleService, Deals.VehicleSold,Deals.VehicleNotes as Profit FROM Deals Where (Deals.VehicleDateBought >= " & Chr(39) & DateFrom & Chr(39) & " And Deals.VehicleDateBought <= " & Chr(39) & DateTo & Chr(39) & ")ORDER BY [Deals.VehicleDateBought]"
      Else
        theSql = "SELECT Deals.ID, Deals.VehicleID, Deals.VehicleRegNr, Deals.VehicleDateBought, Deals.VehicleDateSold, Deals.VehicleCost, Deals.VehicleService, Deals.VehicleSold,Deals.VehicleNotes as Profit FROM Deals Where Deals.ID = " & Val(ctlDBCombo1.Text) & " ORDER BY [Deals.VehicleDateBought]"
      End If
      RptInt = 4
      
'    Case 3    'product listing
'      If chkSubHead.Value <> 0 Then
'        theSql = "SELECT Incidents.Code, Incidents.Description, Incidents.BrochureLink, Incidents.Blocked From Incidents WHERE Incidents.Blocked = False ORDER BY Incidents.Description"
'      Else
'        theSql = "SELECT Incidents.Code, Incidents.Description, Incidents.BrochureLink, Incidents.Blocked From Incidents ORDER BY Incidents.Description"
'      End If
'    Case 4   'show rating
'      If ctlDBCombo1.Text = "-- All Shows --" Then
'        theSql = "SELECT  Comments.vDate as ShowDate, Category.Code, Category.Description, Comments.cName, Comments.RatingScale, Comments.SalesManCode FROM Comments INNER JOIN Category ON Comments.ShowDayCode = Category.Code Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ")ORDER BY [vDate]"
'      Else
'        GetCategoryDetails (ctlDBCombo1.ID)
'        theSql = "SELECT Comments.vDate as ShowDate, Category.Code, Category.Description, Comments.cName, Comments.RatingScale, Comments.SalesManCode FROM Comments INNER JOIN Category ON Comments.ShowDayCode = Category.Code Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Comments.ShowDayCode = " & Chr(39) & ShowDay.Code & Chr(39) & " ORDER BY [vDate]"
'      End If
'    Case 5    'voucher report
'      If chkSubHeadTots.Value <> 0 Then
'        If ctlDBCombo1.Text = "-- All Shows --" Then    'All Winners Only / all shows
'          theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Comments.VoucherNr, Comments.RaffleWinner FROM Comments Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Comments.RaffleWinner = " & True & " ORDER BY [vDate]"
'        Else
'          ' All Winners Only / Selected Show
'          theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Category.Description, Comments.VoucherNr, Comments.RaffleWinner FROM Comments INNER JOIN Category ON Comments.ShowDayCode = Category.Code Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Comments.RaffleWinner = " & True & " ORDER BY [vDate]"
'        End If
'      Else
'        If chkSubHead.Value <> 0 Then 'include winners
'          If ctlDBCombo1.Text = "-- All Shows --" Then
'            'include Winners / all shows
'            theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Comments.VoucherNr, Comments.RaffleWinner FROM Comments Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") ORDER BY [vDate]"
'          Else
'            'include Winners / Selected Show
'            theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Category.Description, Comments.VoucherNr, Comments.RaffleWinner FROM Comments INNER JOIN Category ON Comments.ShowDayCode = Category.Code Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Category.Description = " & Chr(39) & ctlDBCombo1.Text & Chr(39) & " ORDER BY [vDate]"
'          End If
'        Else
'          If ctlDBCombo1.Text = "-- All Shows --" Then
'            'non Winners / all shows
'            theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Comments.VoucherNr, Comments.RaffleWinner FROM Comments Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Comments.RaffleWinner = " & False & " ORDER BY [vDate]"
'          Else
'            ' non winners, specific show
'            theSql = "SELECT Comments.vDate as ShowDate, Comments.cDay as ShowDay, Comments.cName as Name, Comments.cContact as Contact, Comments.cEmail as Email, Category.Description, Comments.VoucherNr, Comments.RaffleWinner FROM Comments INNER JOIN Category ON Comments.ShowDayCode = Category.Code Where (Comments.vDate >= " & Chr(39) & DateFrom & Chr(39) & " And Comments.vDate <= " & Chr(39) & DateTo & Chr(39) & ") AND Category.Description = " & Chr(39) & ctlDBCombo1.Text & Chr(39) & " AND Comments.RaffleWinner = " & False & " ORDER BY [vDate]"
'          End If
'        End If
'      End If
  End Select

  If theSql <> "" Then
    Call ExcelDump(theSql, RptInt)
    Screen.MousePointer = vbDefault
  Else
    MsgBox "Failed to call the Excel Report Writer.", vbInformation + vbOKOnly, "System Error..."
  End If
'
End Sub




