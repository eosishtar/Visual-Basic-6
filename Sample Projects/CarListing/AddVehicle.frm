VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AddVehicle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Vehicle"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   Icon            =   "AddVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " *** Car ID Code ***"
      Height          =   5415
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   9255
      Begin VB.Frame Frame4 
         Caption         =   " Notes "
         Height          =   3375
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   9015
         Begin VB.Frame Frame5 
            Caption         =   " Book Value Details "
            Height          =   855
            Left            =   120
            TabIndex        =   34
            Top             =   2400
            Width           =   8775
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5280
               TabIndex        =   36
               Top             =   360
               Width           =   2895
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   285
               Left            =   240
               TabIndex        =   35
               Top             =   360
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   503
               _Version        =   393216
               Format          =   104529921
               CurrentDate     =   42482
            End
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   1845
            Index           =   9
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   33
            Tag             =   "Vehicle Model"
            Top             =   480
            Width           =   8775
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Year Model"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   " Vehicle Details "
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   9015
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   5
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            Left            =   6240
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "Fuel Type"
            Top             =   345
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   2
            Tag             =   "Vehicle Make"
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Tag             =   "Vehicle Model"
            Top             =   735
            Width           =   3255
         End
         Begin VB.Label Label1 
            Caption         =   "Turbo"
            Height          =   195
            Index           =   12
            Left            =   5160
            TabIndex        =   31
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Fuel"
            Height          =   195
            Index           =   11
            Left            =   5160
            TabIndex        =   29
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Make"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Model"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   765
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Vehicle Specifications "
         Height          =   3375
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   9015
         Begin VB.CommandButton cmdClearPic 
            Height          =   315
            Left            =   8160
            Picture         =   "AddVehicle.frx":1601A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Clear Image"
            Top             =   2880
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2880
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2160
            Width           =   2535
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   2160
            TabIndex        =   12
            Top             =   2520
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   2160
            TabIndex        =   10
            Top             =   1800
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   2160
            TabIndex        =   9
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   2160
            TabIndex        =   8
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   2160
            TabIndex        =   6
            Top             =   360
            Width           =   2535
         End
         Begin CarListing.ctlThumbnail Thumb 
            Height          =   2655
            Left            =   4800
            TabIndex        =   30
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4683
         End
         Begin VB.Label Label1 
            Caption         =   "Drive System"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   26
            Top             =   2925
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nr of Gears"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   25
            Top             =   2565
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Transmission"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   24
            Top             =   2205
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Engine Code"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   1845
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nr of Cyclinders"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   22
            Top             =   1485
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Rated Horsepower"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   21
            Top             =   1125
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Type"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   20
            Top             =   765
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Displacement"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   19
            Top             =   405
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
         Height          =   195
         Index           =   10
         Left            =   6360
         TabIndex        =   28
         Top             =   300
         Width           =   735
      End
   End
   Begin CarListing.ctlMenu ctlMenu1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1508
      vAdd_Enable     =   0   'False
      vEdit_Enable    =   -1  'True
      vFind_Enable    =   0   'False
      vView_Enable    =   0   'False
      vEmail_Enable   =   0   'False
      vExport_Enable  =   -1  'True
      vPrint_Enable   =   -1  'True
      vDelete_Enable  =   -1  'True
      vClose_Enable   =   -1  'True
      vOK_Enable      =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   -180
      X2              =   13155
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -360
      X2              =   13320
      Y1              =   7080
      Y2              =   7080
   End
End
Attribute VB_Name = "AddVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImageToSave As String
Dim Cnter As Integer
Dim DIRTY As Boolean
Dim vTday As Single

Private Sub cmdClearPic_Click()

  ImageToSave = "#CLEAR#"
  Thumb.ClearPicture
  DIRTY = True

End Sub

Private Sub Combo1_Change(Index As Integer)

  DIRTY = True

End Sub

Private Sub Combo1_Click(Index As Integer)

  DIRTY = True

End Sub

Private Sub ctlMenu1_CloseWindow()
  Unload Me
End Sub

Private Sub ctlMenu1_Delete()
  Dim sString As String

  sql = "SELECT * From MasterListing where ID = " & MasterVehicle.ID
  With rs
    .Open sql, cn, adOpenKeyset, adLockOptimistic

    If .EOF Then
        .Close
        MsgBox "An error has occurred. No changes were made.", vbCritical + vbOKOnly, Me.Caption
      Exit Sub
    End If
    
    .Delete
    
  .Update
  .Close
  End With
  
  MsgBox "Vehicle successfully deleted.", vbInformation + vbOKOnly, Me.Caption
  Unload Me
  MasterVehListing.LoadCars ""
    
End Sub

Private Sub ctlMenu1_Edit()
  If Frame4.Visible Then
    Frame4.Visible = False
  Else
    Frame4.Visible = True
  End If
End Sub

Private Sub ctlMenu1_Export()
    
  If MasterVehicle.ID <> 0 Then
    'sql = "Select * from MasterListing where ID = " & MasterVehicle.ID
    sql = "SELECT MasterListing.ID, MasterListing.ModelYear, MasterListing.VehicleMake, MasterListing.VehicleModel, MasterListing.VehicleDisplacement AS CC, MasterListing.VehicleType AS VehType, MasterListing.RatedHorsepower AS HP, MasterListing.NrofCylinders AS Cyc, MasterListing.EngineCode, MasterListing.TransmissionType, MasterListing.NrofGears as Gears, MasterListing.DriveSystemDescription, MasterListing.FuelType, MasterListing.TurboType " & _
          "FROM MasterListing WHERE ID = " & MasterVehicle.ID

    Call ExcelDump(sql, 1)
  End If

End Sub

Private Sub ctlMenu1_Save()
  Dim vErrString As String
  Dim v As Integer
  Dim ContinueOK As Boolean
  Dim tString As String
  Dim TT As String
  Dim Recs As Integer
  
  '... validation
  ContinueOK = True
  For v = 0 To 1
    Select Case v
      Case 0, 1     'make, model
        If Text1(v).Text = "" Then
          vErrString = vErrString & " * " & Text1(v).Tag & vbCrLf
          ContinueOK = False
        End If
    End Select
  Next v
  
  If Combo1(0).ListIndex = -1 Then
    vErrString = vErrString & " * " & Combo1(0).Tag & vbCrLf
    ContinueOK = False
  End If
  
  If Combo1(4).ListIndex = -1 Or Trim$(Combo1(4).Text) = "" Then
    vErrString = vErrString & " * " & Combo1(4).Tag & vbCrLf
    ContinueOK = False
  End If
  
  If Not ContinueOK Then
    MsgBox "Please complete the following items first before you continue?" & vbCrLf & vErrString, vbInformation + vbOKOnly, Me.Caption
    Exit Sub
  End If
    
  If ImageToSave <> "" Then
    If ImageToSave = "#CLEAR#" Then
      Recs = OpenDocStore("V" & Trim(Str(MasterVehicle.ID)))
      If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
      If Recs = 0 Then MsgBox "Image not found": CloseDocStore: Exit Sub
      If Recs <> 1 Then MsgBox "More than one Image found": CloseDocStore: Exit Sub
      DocStoreRS.Delete
      MasterVehicle.VehicleImage = ""
      CloseDocStore
    Else
      Recs = OpenDocStore("V" & Trim(Str(MasterVehicle.ID)))
      If Recs = -1 Then
        MsgBox DocError
      Else
        If Recs = 0 Then
          On Error GoTo 0
          DocStoreRS.AddNew
          DocStoreRS.Fields("DocID") = "V" & Trim(Str(MasterVehicle.ID))
          DocStoreRS.Fields("DocName") = "vehiclepic"
          DocStoreRS.Fields("DocExtension") = "syspic"
        End If
        FileToField ImageToSave, "DocContents", "DocExtension"
        DocStoreRS.Update
      End If
      MasterVehicle.VehicleImage = "V" & Trim(Str(MasterVehicle.ID))
      CloseDocStore
    End If
  End If

  sql = "Select * from MasterListing where ID = " & MasterVehicle.ID
  With rs
    .Open sql, cn, adOpenStatic, adLockPessimistic
  
    If .EOF Then
      .AddNew
    End If
    
    rs!ModelYear = Combo1(0).Text
    rs!VehicleMake = UCase(Text1(0).Text)
    rs!VehicleModel = UCase(Text1(1).Text)
    
    If Trim$(Text1(2).Text) = "" Then Text1(2).Text = 0
    rs!VehicleDisplacement = Text1(2).Text
    'If Trim$(Combo1(1).Text) = "" Then Combo1(1).Text = NonRequired
    rs!VehicleType = Combo1(1).Text
    If Trim$(Text1(4).Text) = "" Then Text1(4).Text = 0
    rs!RatedHorsePower = Text1(4).Text
    If Trim$(Text1(5).Text) = "" Then Text1(5).Text = 0
    rs!NrOfCylinders = Text1(5).Text
    If Trim$(Text1(6).Text) = "" Then Text1(6).Text = NonRequired
    rs!EngineCode = Text1(6).Text
    
    '...get the trans type
    If Trim$(Combo1(2).Text) <> "" Then
      tString = Combo1(2).Text
      TT = Trim$(Left(tString, InStrRev(tString, "-") - 1))
      rs!TransmissionTypeCode = TT
      rs!TransmissionType = GetTransmission(TT)
    End If
    If Trim$(Text1(8).Text) = "" Then Text1(8).Text = 0
    rs!NrOfGears = Text1(8).Text
    
    If Trim$(Combo1(3).Text) <> "" Then
      tString = Combo1(3).Text
      TT = Trim$(Left(tString, InStrRev(tString, "-") - 1))
      rs!DriveSystemCode = TT
      rs!DriveSystemDescription = GetDriveSys(TT)
    End If
    
    '...save ther picture
    If Thumb.HasPicture Then
      rs!VehicleImage = MasterVehicle.VehicleImage
    Else
      rs!VehicleImage = NonRequired
    End If
    '...save the notes
    If Trim$(Text1(9).Text) = "" Then Text1(9).Text = NonRequired
    rs!Notes = Text1(9).Text
    
    '...save the turbotype
    rs!FuelType = Combo1(4).Text
    If Combo1(5).ListIndex = -1 Then
      rs!TurboType = NonRequired
    Else
      rs!TurboType = Combo1(5).Text
    End If
    
    'save the book value details
    
    rs!BookValue = Val(Format(Text2.Text, "### ### ### ### ##0.00"))
    rs!BookDate = GetDateVal(DTPicker1)
    
    rs.Update
    rs.Close
  End With
  
  DIRTY = False
  MsgBox Trim$(UCase(Text1(0).Text)) & " " & Trim$(UCase(Text1(1).Text)) & " saved successfully.", vbInformation + vbOKOnly, Me.Caption
  Unload Me
  MasterVehListing.LoadCars ""

End Sub

Private Sub Form_Load()

  Dim vYear As String
  Dim tmpPath As String
  Dim Recs As Integer

  
  vTday = DateValue(Now)

  vYear = Format(Now, "yyyy")
  vYear = vYear + 1
  Combo1(0).Clear
  For Cnter = 1 To Max_Vehicle_Years
    vYear = vYear - 1
    Combo1(0).AddItem vYear
  Next Cnter
  Thumb.BorderStyle = bFixed
  
  '...add vehicle type
  Combo1(1).AddItem ""
  Combo1(1).AddItem "Car"
  Combo1(1).AddItem "Both"
  Combo1(1).AddItem "Truck"
  Combo1(1).AddItem "SUV"
  Combo1(1).AddItem "Scooter"
  Combo1(1).AddItem "Bike"
  
  '...add veh trans
  Combo1(2).AddItem ""
  Combo1(2).AddItem "A - Automatic"
  Combo1(2).AddItem "M - Manual"
  Combo1(2).AddItem "AM - Automated Manual"
  Combo1(2).AddItem "SCV - Selectable Continuously Variable"
  Combo1(2).AddItem "SA - Semi Automatic"
  Combo1(2).AddItem "CVT - Continuously Variable"
  Combo1(2).AddItem "OT - Other"
  Combo1(2).AddItem "AMS - Automated Manual Selectable"
  
  '...add Drive System
  Combo1(3).AddItem ""
  Combo1(3).AddItem "R - 2 Wheel Drive, Rear"
  Combo1(3).AddItem "F - 2 Wheel Drive, Front"
  Combo1(3).AddItem "A - All Wheel Drive"
  Combo1(3).AddItem "4 - 4 Wheel Drive"
  Combo1(3).AddItem "P - Part time 4 Wheel Drive"
  
  '...fuel type
  Combo1(4).AddItem ""
  Combo1(4).AddItem "Diesel"
  Combo1(4).AddItem "Petrol"
  Combo1(4).AddItem "Other"
  
  '...turbo type
  Combo1(5).AddItem ""
  Combo1(5).AddItem "N/A"
  Combo1(5).AddItem "SuperCharged"
  Combo1(5).AddItem "Turbo"
  Combo1(5).AddItem "Twin Turbo"
  Combo1(5).AddItem "Both"
  Combo1(5).AddItem "Other"

  'bookvalue details
  DTPicker1.Value = vTday
  Text2.Text = 0
  
  
  If MasterVehicle.ID <> 0 Then
    Frame1.Caption = " Car ID : " & MasterVehicle.ID & " "
    Me.Caption = Frame1.Caption
    Combo1(0).Text = Trim$(MasterVehicle.ModelYear)
    Text1(0).Text = UCase(Trim$(MasterVehicle.VehicleMake))
    Text1(1).Text = UCase(Trim$(MasterVehicle.VehicleModel))
    Text1(2).Text = Trim$(MasterVehicle.VehicleDisplacement)
    Combo1(1).ListIndex = SetComboText(Combo1(1), MasterVehicle.VehicleType)
    Text1(4).Text = Trim$(MasterVehicle.RatedHorsePower)
    Text1(5).Text = Trim$(MasterVehicle.NrOrCylinders)
    Text1(6).Text = Trim$(MasterVehicle.EngineCode)
    Combo1(2).ListIndex = SetComboText(Combo1(2), Left(MasterVehicle.TransmissionCode, 2))
    Text1(8).Text = Trim$(MasterVehicle.NrOfGears)
    Combo1(3).ListIndex = SetComboText(Combo1(3), Left(MasterVehicle.DriveSystemCode, 1))
    Text1(9).Text = Trim$(MasterVehicle.Notes)
    
    Combo1(4).ListIndex = SetComboText(Combo1(4), MasterVehicle.FuelType)
    Combo1(5).ListIndex = SetComboText(Combo1(5), MasterVehicle.TurboType)

    Text2.Text = FormatNumber(MasterVehicle.BookValue, 2)
    DTPicker1.Value = Format(MasterVehicle.BookDate, "dd MMMM yyyy")

  End If

  ImageToSave = ""
  If Trim$(MasterVehicle.VehicleImage) = "" Then
    cmdClearPic.Visible = False
  Else
    cmdClearPic.Visible = True
    tmpPath = App.Path
    If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
    tmpPath = tmpPath & "tmp.img"
    Recs = OpenDocStore("V" & Trim(Str(MasterVehicle.ID)))
    If Recs = -1 Then
      MsgBox DocError
    Else
      If Recs = 0 Then
        Thumb.ClearPicture
      Else
        FileFromField tmpPath, "DocContents", "DocExtension"
        Thumb.SetPicture LoadPicture(tmpPath)
        On Error Resume Next
        Kill tmpPath
        On Error GoTo 0
      End If
    End If
    CloseDocStore
  End If
  DIRTY = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim Resp As VbMsgBoxResult

  If DIRTY Then
    Resp = MsgBox("Exit without saving?", vbQuestion + vbYesNo, Me.Caption)
    If Resp <> vbYes Then
      Cancel = 1
      Exit Sub
    End If
  End If
  RemoveTempFiles

End Sub

Private Sub Text1_Change(Index As Integer)

  DIRTY = True

End Sub

Private Sub Text1_GotFocus(Index As Integer)

  Text1(Index).SelStart = 0
  Text1(Index).SelLength = Trim$(Len(Text1(Index)))

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
    Case 0, 1
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 2, 4, 5, 8
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
      'allow
      Else
        KeyAscii = 0
      End If
  End Select
  
End Sub



Private Sub Text2_Change()
  
  DTPicker1.Value = vTday

End Sub

Private Sub Text2_GotFocus()
  Text2.SelStart = 0
  Text2.SelLength = Len(Trim(Text2.Text))
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

  If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
  Else
    KeyAscii = 0
  End If
  
End Sub

Private Sub Text2_LostFocus()

  Text2.Text = FormatNumber(Text2.Text, 2)
  
End Sub

Private Sub Thumb_DblClick()

  Dim Recs As Integer, tmpPath As String

  tmpPath = App.Path
  If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
  tmpPath = tmpPath & "tmp.img"
  Recs = OpenDocStore("V" & Trim(Str(MasterVehicle.ID)))
  If Recs > 0 Then
    FileFromField tmpPath, "DocContents", "DocExtension"
    OpenFile2 tmpPath
    DoEvents
  End If
  CloseDocStore

End Sub

Private Sub Thumb_NewDropImage(NewPath As String)

  ImageToSave = NewPath
  Thumb.SetPicture LoadPicture(NewPath)
  DIRTY = True
  cmdClearPic.Visible = True

End Sub
