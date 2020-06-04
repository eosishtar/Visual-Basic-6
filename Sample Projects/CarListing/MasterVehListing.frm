VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form MasterVehListing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Master Vehicle Listing "
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14325
   Icon            =   "MasterVehListing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   14325
   Begin CarListing.ctlMenu ctlMenu1 
      Height          =   1095
      Left            =   240
      TabIndex        =   4
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
   Begin VB.Frame Frame1 
      Caption         =   " Vehicle Listing "
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   13815
      Begin CarListing.ctlProgressBar ctlProgressBar1 
         Height          =   255
         Left            =   1800
         TabIndex        =   3
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
         TabIndex        =   1
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
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   5400
         Width           =   3735
      End
   End
   Begin VB.Timer Timer1 
      Left            =   10320
      Top             =   720
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   14160
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   14160
      Y1              =   7875
      Y2              =   7875
   End
End
Attribute VB_Name = "MasterVehListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const SearchWait = 1000



Private Sub Command1_Click()

End Sub

Private Sub ctlMenu1_AddNew()
  GetMasterVehicleDetails 0
  AddVehicle.Show 1
End Sub

'...dont forget to set dirty
'...ctlMenu1.DIRTY True


Private Sub ctlMenu1_CloseWindow()
  Unload Me
End Sub


Private Sub ctlMenu1_Export()
  sql = "SELECT MasterListing.ID, MasterListing.ModelYear, MasterListing.VehicleMake, MasterListing.VehicleModel, MasterListing.VehicleDisplacement AS CC, MasterListing.VehicleType AS VehType, MasterListing.RatedHorsepower AS HP, MasterListing.NrofCylinders AS Cyc, MasterListing.EngineCode, MasterListing.TransmissionType, MasterListing.NrofGears as Gears, MasterListing.DriveSystemDescription, MasterListing.FuelType, MasterListing.TurboType " & _
        "FROM MasterListing"

  Call ExcelDump(sql, 1)
End Sub

Private Sub ctlMenu1_Reset()
  LoadCars ""
End Sub


Private Sub ctlMenu1_Resized()
  If Frame1.Top <> ctlMenu1.Top + ctlMenu1.Height + 100 Then Frame1.Top = ctlMenu1.Top + ctlMenu1.Height + 100
End Sub


Private Sub ctlMenu1_Search(vSearchText As String)

 Timer1.Interval = 0
 Timer1.Interval = SearchWait
 Timer1.Tag = vSearchText
 
End Sub

Private Sub ctlProgressBar1_TimeOut()
  ctlProgressBar1.Visible = False
End Sub

Private Sub Form_Load()
  
  Call CenterForm(Me)
  DoEvents
  Timer1.Interval = 1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Timer1.Interval = 0
  
End Sub


Private Sub ListView1_DblClick()
  
  GetMasterVehicleDetails ListView1.SelectedItem
  
'  Dim temp As String
'  temp = "MasterVehicle.ID = " & MasterVehicle.ID & vbCrLf & _
'        "MasterVehicle.ModelYear = " & MasterVehicle.ModelYear & vbCrLf & _
'        "MasterVehicle.VehicleMake = " & MasterVehicle.VehicleMake & vbCrLf & _
'        "MasterVehicle.VehicleModel = " & MasterVehicle.VehicleModel & vbCrLf & _
'        "MasterVehicle.VehicleDisplacement = " & MasterVehicle.VehicleDisplacement & vbCrLf & _
'        "MasterVehicle.VehicleType = " & MasterVehicle.VehicleType & vbCrLf & _
'        "MasterVehicle.RatedHorsePower = " & MasterVehicle.RatedHorsePower & vbCrLf & _
'        "MasterVehicle.NrOrCylinders = " & MasterVehicle.NrOrCylinders & vbCrLf & _
'        "MasterVehicle.EngineCode = " & MasterVehicle.EngineCode & vbCrLf & _
'        "MasterVehicle.NrOfGears = " & MasterVehicle.NrOfGears & vbCrLf & _
'        "MasterVehicle.DriveSystemCode = " & MasterVehicle.DriveSystemCode & vbCrLf & _
'        "MasterVehicle.DriveSystemDescription = " & MasterVehicle.DriveSystemCode
'
    AddVehicle.Show 1
  
End Sub

Private Sub Timer1_Timer()
  
  Timer1.Interval = 0
  LoadCars Timer1.Tag

End Sub

Private Sub cmdLoad_Click()

  LoadCars
   
End Sub


Public Function LoadCars(Optional SearchText As String)

  Dim sString As String
  Dim itmx As Object
  Dim i, j, k, X   As Integer
  Dim vTotRec As Integer

  lblRecords.Caption = ""
  DoEvents
  sql = "SELECT [ID],[ModelYear], [VehicleMake],[VehicleModel], [VehicleType],[VehicleDisplacement],[RatedHorsepower],[DriveSystemCode],[TransmissionTypeCode] From MasterListing"
  With rs
    .Open sql, cn, adOpenKeyset, adLockOptimistic

    If rs.EOF Then
        rs.Close
      Exit Function
    End If
    
    ctlProgressBar1.Visible = True
    rs.MoveLast
    vTotRec = rs.RecordCount
    X = 1
    rs.MoveFirst
   
    '...set listview parameters
    ListView1.ColumnHeaders.Clear
    ListView1.ListItems.Clear
    ListView1.View = lvwReport
    ListView1.BorderStyle = ccFixedSingle
    ListView1.FullRowSelect = True
    ListView1.GridLines = True

    '   count the columns and add them to the listview
    For i = 0 To rs.Fields.Count - 1
      ListView1.ColumnHeaders.Add , , rs.Fields(i).Name
    Next i
    
    '   count the rows and add the items and subitems
      rs.MoveFirst
      For j = 1 To vTotRec
        If InStr(1, UCase(rs.Fields("VehicleMake").Value), UCase(SearchText)) > 0 Or InStr(1, UCase(rs.Fields("VehicleModel").Value), UCase(SearchText)) > 0 Or InStr(1, Str(rs.Fields("ModelYear").Value), UCase(SearchText)) > 0 Then
          Set itmx = ListView1.ListItems.Add(, , rs.Fields(0).Value)
            For k = 1 To ListView1.ColumnHeaders.Count - 1
                  On Error Resume Next
                itmx.SubItems(k) = rs.Fields(k).Value
            Next k
        End If
        rs.MoveNext
        
        '...do progress bar
        X = X + 1
        ctlProgressBar1.SetPerc X, vTotRec
        
      Next j
      rs.Close
    End With
    
    AltLVBackground ListView1, MasterVehListing
    lblRecords.Caption = ListView1.ListItems.Count & " records loaded..."
    Call AutosizeColumns(ListView1)     ' resize all the columns
    ListView1.ColumnHeaders(1).Width = 0
    
End Function


