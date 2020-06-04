VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deal Listing"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14325
   Icon            =   "frmClients.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14325
   Begin VB.Timer Timer1 
      Left            =   10320
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   " Vehicle Listing "
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   13815
      Begin CarListing.ctlProgressBar ctlProgressBar1 
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   4200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   450
         vPercView       =   -1  'True
         vForeColor      =   65535
         vBackColor      =   12632256
         vTextColor      =   0
         vPercCaption    =   "% complete"
         vUnloadProgBar  =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Buyer Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Buyer Contact"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Seller Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Seller Contact"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vehicle"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   5400
         Width           =   3735
      End
   End
   Begin CarListing.ctlMenu ctlMenu1 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1931
      vAdd_Enable     =   -1  'True
      vEdit_Enable    =   0   'False
      vFind_Enable    =   -1  'True
      vView_Enable    =   0   'False
      vEmail_Enable   =   0   'False
      vExport_Enable  =   -1  'True
      vPrint_Enable   =   0   'False
      vDelete_Enable  =   0   'False
      vClose_Enable   =   -1  'True
      vOK_Enable      =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClients.frx":1601A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClients.frx":18BF1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   14160
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   14110
      Y1              =   7750
      Y2              =   7750
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ctlMenu1_AddNew()

'  Static FormCount As Long
'  Dim frmD As AddClient
'  FormCount = FormCount + 1
'  Set frmD = New AddClient
'  Deal.ID = 0
'  frmD.Show
  Deal.ID = 0
  AddClient.Show

End Sub

Private Sub ctlMenu1_CloseWindow()
  Unload Me
End Sub

Private Sub ctlMenu1_Export()
  sql = "SELECT Deals.ID, Deals.BuyerFirstname, Deals.BuyerLastName, Deals.BuyerContactNr, Deals.SellerFirstName, Deals.SellerLastName, Deals.SellerContactNr, Deals.VehicleID, Deals.VehicleRegNr, Deals.VehicleCost, Deals.VehicleService, Deals.VehicleSold, Deals.DealClosed FROM Deals"
  Call ExcelDump(sql, 2)
End Sub

Private Sub ctlMenu1_Reset()
  LoadClients 0, ""
End Sub

Private Sub ctlMenu1_Search(vSearchText As String)
  LoadClients 0, vSearchText
End Sub

Private Sub ctlProgressBar1_TimeOut()
  ctlProgressBar1.Visible = False
  lblRecords.Caption = Counter(ListView1.ListItems.Count, "Deals")
  lblRecords.Visible = True
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  Call LoadClients(0, "")

End Sub

Public Function LoadClients(ID As Integer, Optional SearchText As String)

  Dim i As Integer
  Dim tRecCount As Integer
  
  ListView1.ListItems.Clear
  lblRecords.Visible = False
  i = 0
  ctlProgressBar1.Visible = True
  
  With rs         '0      1               2                 3               4               5               6               7
    sql = "SELECT ID, BuyerFirstname, BuyerLastName, BuyerContactNr, SellerFirstName, SellerLastName, SellerContactNr, VehicleID, DealClosed FROM Deals"
      .Open sql, cn, adOpenKeyset, adLockOptimistic
      If .EOF Then
        .Close
        lblRecords.Caption = Counter(ListView1.ListItems.Count, "Deals")
        Exit Function
      End If
      
      tRecCount = rs.RecordCount
      ctlProgressBar1.SetPerc i, tRecCount
      
      Do While Not .EOF
        If InStr(1, UCase(rs!ID), UCase(SearchText)) > 0 Or InStr(1, UCase(rs!BuyerFirstName), UCase(SearchText)) > 0 Or InStr(1, UCase(rs!BuyerLastName), UCase(SearchText)) > 0 Or InStr(1, UCase(rs!SellerFirstName), UCase(SearchText)) > 0 Or InStr(1, UCase(rs!SellerLastName), UCase(SearchText)) > 0 Then
            If rs!DealClosed = True Then
              Set itmx = ListView1.ListItems.Add(, , rs!ID, , 1)      'active
            Else
              Set itmx = ListView1.ListItems.Add(, , rs!ID, , 2)      'not active
            End If
          
            '...Buyer FirstName & Buyer LastName
            If Not IsNull(rs.Fields("BuyerLastName").Value) Then itmx.SubItems(1) = rs.Fields("BuyerFirstname").Value & " " & rs.Fields("BuyerLastName").Value Else itmx.SubItems(1) = rs.Fields("BuyerFirstname").Value
            '...Buyer Contact
            If Not IsNull(rs.Fields("BuyerContactNr").Value) Then itmx.SubItems(2) = rs.Fields("BuyerContactNr").Value
            '...Seller First Name & Seller Last Name
            If Not IsNull(rs.Fields("SellerLastName").Value) Then itmx.SubItems(3) = rs.Fields("SellerFirstName").Value & " " & rs.Fields("SellerLastName").Value Else itmx.SubItems(3) = rs.Fields("SellerFirstName").Value
            '...Seller Contact
            If Not IsNull(rs.Fields("SellerContactNr").Value) Then itmx.SubItems(4) = rs.Fields("SellerContactNr").Value
            '...Vehicle Detail
            GetMasterVehicleDetails Val(rs!VehicleID)
            itmx.SubItems(5) = "(" & MasterVehicle.ModelYear & ")" & "   " & MasterVehicle.VehicleMake & " " & MasterVehicle.VehicleModel
        End If
          .MoveNext
          i = i + 1
          ctlProgressBar1.SetPerc i, tRecCount
        Loop
      .Close
            
      AltLVBackground ListView1, frmClients
      lblRecords.Caption = Counter(ListView1.ListItems.Count, "Deals")
      Call AutosizeColumns(ListView1)     ' resize all the columns
      
  End With
    
End Function



Private Sub ListView1_DblClick()

  Deal.ID = ListView1.SelectedItem
  AddClient.Show
  
End Sub

