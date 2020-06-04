VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Link a MasterVehicle"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9450
   Icon            =   "PCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9450
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   9840
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   2160
   End
   Begin VB.TextBox txtSearchPcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton cmdReset 
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      Picture         =   "PCodes.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   195
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "PCodes.frx":191A3
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1245
   End
   Begin VB.CommandButton cmdLoadPCodesAdmin 
      Appearance      =   0  'Flat
      Caption         =   "Load PCodes Admin"
      Height          =   585
      Left            =   11880
      Picture         =   "PCodes.frx":1BE7E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   2520
      Top             =   2160
   End
   Begin VB.CommandButton cmdFilterPCodes 
      Appearance      =   0  'Flat
      Caption         =   "Load PCodes Customer"
      Height          =   585
      Left            =   11880
      Picture         =   "PCodes.frx":1EB1D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtCompPC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   " Link a MasterVehicle ID "
      Height          =   4695
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   9135
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Area"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Area Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Suburb"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Suburb Code"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Sales Code"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin CarListing.ctlProgressBar ctlProgressBar1 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   5520
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      vPercView       =   -1  'True
      vForeColor      =   65535
      vBackColor      =   12632256
      vTextColor      =   0
      vPercCaption    =   "% complete"
      vUnloadProgBar  =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   5925
      Y2              =   5925
   End
   Begin VB.Label lblRecords 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   720
   End
End
Attribute VB_Name = "frmPCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const SearchWait = 1000
Private Const BLANK_ = "     "
Dim tmpSbItm As String
Dim DoContinue As Boolean
Dim Cnter As Integer
Dim vTemp As String
Dim vTemp2 As String
Dim vTemp3 As String

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdReset_Click()
  txtSearchPcode.Text = ""
  lblRecords.Visible = False
  LoadMVehicles ""
End Sub

Private Sub ctlProgressBar1_TimeOut()
  ctlProgressBar1.Visible = False
  lblRecords.Visible = True
End Sub

Private Sub Form_Load()
  lblRecords.Caption = ""
  
  '...timer2 will load pcodes, otherwise it takes a while before anything happens
  Screen.MousePointer = vbHourglass
  Timer2.Enabled = True
  Me.Top = (Screen.Height - (Me.Height - 1000)) / 4
  Me.Left = (Screen.Width - Me.Width) / 2
  
  ListView1.ColumnHeaders.Clear
  
End Sub




Private Sub ListView1_DblClick()
    
    
  If ctlProgressBar1.Visible = True Then ctlProgressBar1.Visible = False
    
    On Error Resume Next
    tmpSbItm = ListView1.SelectedItem
    If Err = 0 Then DoContinue = True Else DoContinue = False
    On Error GoTo 0

    If DoContinue Then
      AddClient.Text1(12).Text = ListView1.SelectedItem
      'set details for vehicle
      AddClient.lblDetails.Caption = "(" & ListView1.SelectedItem.SubItems(1) & ") " & ListView1.SelectedItem.SubItems(2) & " " & ListView1.SelectedItem.SubItems(3)
      
      'grab the book value details
      vTemp = List1.List(ListView1.SelectedItem - 1)
      vTemp2 = Right(vTemp, Len(vTemp) - InStrRev(vTemp, "-"))      'got the date
      vTemp3 = Left(vTemp, Len(vTemp) - InStr(vTemp, "-"))          'got the value
      
      If Trim(vTemp3) = "" Or Trim(vTemp3) = 0 Then
        AddClient.lblBookValue.Caption = "Book Value : ( Not Set )"
      Else
        AddClient.lblBookValue.Caption = "Book Value : " & FormatNumber(vTemp3, 2) & " (" & Format(vTemp2, "dd MMMM yyyy") & ")"
      End If
      
      Unload Me
    End If

End Sub

'loads the postal codes
Private Sub Timer2_Timer()
  Timer2.Interval = 1000
  LoadMVehicles ""
  Timer2.Enabled = False
  
  Screen.MousePointer = vbNormal

End Sub

Private Sub Timer1_Timer()

  Timer1.Interval = 0
  LoadMVehicles Timer1.Tag

End Sub

Private Sub txtSearchPcode_KeyUp(KeyCode As Integer, Shift As Integer)
 
 Timer1.Interval = 0
 Timer1.Interval = SearchWait
 Timer1.Tag = txtSearchPcode
 
End Sub


Public Function LoadMVehicles(Optional SearchText As String)
  Dim sString As String
  Dim itmx As Object
  Dim i, j, k, x   As Integer
  Dim vTotRec As Integer
  Dim BookVal As String
  

  lblRecords.Caption = ""
  sql = "SELECT [ID],[ModelYear], [VehicleMake],[VehicleModel], [VehicleType],[VehicleDisplacement],[BookValue],[BookDate] From MasterListing"
  With rs
    .Open sql, cn, adOpenKeyset, adLockOptimistic

    If rs.EOF Then
        rs.Close
      Exit Function
    End If
    
    ctlProgressBar1.Visible = True
    rs.MoveLast
    vTotRec = rs.RecordCount
    x = 1
    rs.MoveFirst
   
    '...set listview parameters
    ListView1.ColumnHeaders.Clear
    ListView1.ListItems.Clear
    ListView1.View = lvwReport
    ListView1.BorderStyle = ccFixedSingle
    ListView1.FullRowSelect = True
    ListView1.GridLines = True

    '   count the columns and add them to the listview
    For i = 0 To rs.Fields.Count - 3      'remove book value and book date in column header
      ListView1.ColumnHeaders.Add , , rs.Fields(i).Name
    Next i
    
    '   count the rows and add the items and subitems
      rs.MoveFirst
      For j = 1 To vTotRec
       BookVal = ""
        If InStr(1, UCase(rs.Fields("VehicleMake").Value), UCase(SearchText)) > 0 Or InStr(1, UCase(rs.Fields("VehicleModel").Value), UCase(SearchText)) > 0 Or InStr(1, Str(rs.Fields("ModelYear").Value), UCase(SearchText)) > 0 Then
          Set itmx = ListView1.ListItems.Add(, , rs.Fields(0).Value)
            For k = 1 To ListView1.ColumnHeaders.Count + 1 'add book value and book date in column header
              'On Error Resume Next
              If k = 6 Then
                If Not IsNull(rs.Fields(k).Value) Then BookVal = rs.Fields(k).Value
              ElseIf k = 7 Then
                If Not IsNull(rs.Fields(k).Value) Then List1.AddItem BookVal & "-" & rs.Fields(k).Value
              Else
                If Not IsNull(rs.Fields(k).Value) Then itmx.SubItems(k) = rs.Fields(k).Value
              End If
            Next k
        End If
        rs.MoveNext
        
        '...do progress bar
        x = x + 1
        ctlProgressBar1.SetPerc x, vTotRec
        
      Next j
      rs.Close
    End With
    
    AltLVBackground ListView1, frmPCodes
    lblRecords.Caption = ListView1.ListItems.Count & " records loaded..."
    Call AutosizeColumns(ListView1)     ' resize all the columns
    ListView1.ColumnHeaders(1).Width = 0
    

End Function









