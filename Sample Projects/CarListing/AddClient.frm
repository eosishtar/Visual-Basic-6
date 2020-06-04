VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form AddClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12495
   Icon            =   "AddClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12495
   Begin CarListing.ctlThumbnail Thumb 
      Height          =   2655
      Left            =   8280
      TabIndex        =   96
      Top             =   3240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4683
   End
   Begin VB.CommandButton cmdNext 
      Height          =   315
      Left            =   10320
      Picture         =   "AddClient.frx":1601A
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Next Image"
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   315
      Left            =   9600
      Picture         =   "AddClient.frx":18EC1
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   "Previous Image"
      Top             =   6000
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.CommandButton cmdPrintDeal 
      Caption         =   "Print Deal"
      Height          =   525
      Left            =   8400
      Picture         =   "AddClient.frx":1BC6A
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Print Deal"
      Top             =   8400
      Width           =   1245
   End
   Begin VB.CommandButton cmdCloseDeal 
      Caption         =   "Close Deal"
      Height          =   525
      Left            =   9720
      Picture         =   "AddClient.frx":1EDB7
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Close Deal"
      Top             =   8400
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   120
      TabIndex        =   21
      Top             =   7320
      Width           =   12135
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   295
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Vehicle Sale Price"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   295
         Left            =   5460
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Total Vehicle Maintenance"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   295
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Vehicle Purchase Price"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Sold Price"
         Height          =   255
         Left            =   8880
         TabIndex        =   26
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Maintenance"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Vehicle Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   280
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Client"
      Height          =   525
      Left            =   1560
      Picture         =   "AddClient.frx":21974
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Delete"
      Top             =   8400
      Width           =   1125
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   525
      Left            =   11040
      Picture         =   "AddClient.frx":24658
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Save Records"
      Top             =   8400
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   525
      Left            =   120
      Picture         =   "AddClient.frx":272CB
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Close"
      Top             =   8400
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdClearPic 
      Height          =   315
      Left            =   11640
      Picture         =   "AddClient.frx":29FA6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clear Image"
      Top             =   6000
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   -240
      Picture         =   "AddClient.frx":2CC81
      ScaleHeight     =   3015
      ScaleWidth      =   12735
      TabIndex        =   0
      Top             =   -1200
      Width           =   12735
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblSold 
         BackStyle       =   0  'Transparent
         Caption         =   "SOLD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   2535
         Left            =   1440
         TabIndex        =   97
         Top             =   1800
         Visible         =   0   'False
         Width           =   4335
      End
   End
   Begin VB.Frame frame1 
      Caption         =   " View Expenses "
      Height          =   4575
      Index           =   5
      Left            =   2160
      TabIndex        =   36
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdRemoveAtt 
         Height          =   370
         Left            =   5520
         Picture         =   "AddClient.frx":4AAAF
         Style           =   1  'Graphical
         TabIndex        =   105
         TabStop         =   0   'False
         ToolTipText     =   "Remove Attachment"
         Top             =   240
         Width           =   370
      End
      Begin VB.CommandButton cmdAddAtt 
         Height          =   370
         Left            =   4800
         Picture         =   "AddClient.frx":4D793
         Style           =   1  'Graphical
         TabIndex        =   103
         TabStop         =   0   'False
         ToolTipText     =   "Add Attachment"
         Top             =   240
         Width           =   370
      End
      Begin VB.CommandButton cmdViewAtt 
         Height          =   370
         Left            =   5160
         Picture         =   "AddClient.frx":50461
         Style           =   1  'Graphical
         TabIndex        =   101
         TabStop         =   0   'False
         ToolTipText     =   "View Attachment"
         Top             =   240
         Visible         =   0   'False
         Width           =   370
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Columns         =   4
         Height          =   2955
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   5775
      End
      Begin VB.CommandButton cmdAddDoc 
         Height          =   370
         Left            =   5160
         Picture         =   "AddClient.frx":5337C
         Style           =   1  'Graphical
         TabIndex        =   102
         TabStop         =   0   'False
         ToolTipText     =   "Add Document"
         Top             =   240
         Width           =   370
      End
      Begin VB.CommandButton cmdDeleteDoc 
         Height          =   370
         Left            =   5520
         Picture         =   "AddClient.frx":5604A
         Style           =   1  'Graphical
         TabIndex        =   100
         TabStop         =   0   'False
         ToolTipText     =   "Remove Document"
         Top             =   240
         Width           =   370
      End
      Begin VB.CommandButton cmdChangeAtt 
         Height          =   370
         Left            =   4800
         Picture         =   "AddClient.frx":58D2E
         Style           =   1  'Graphical
         TabIndex        =   104
         TabStop         =   0   'False
         ToolTipText     =   "Change Attachment"
         Top             =   240
         Visible         =   0   'False
         Width           =   370
      End
      Begin VB.Label lblNotes 
         Height          =   675
         Left            =   240
         TabIndex        =   92
         Top             =   3720
         Width           =   5520
      End
   End
   Begin VB.Frame frame1 
      Height          =   4575
      Index           =   2
      Left            =   2160
      TabIndex        =   72
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdLinkCarID 
         Height          =   315
         Left            =   3840
         Picture         =   "AddClient.frx":5BEB7
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Link to Master Vehicle"
         Top             =   470
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1020
         Index           =   17
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   16
         Left            =   2160
         TabIndex        =   81
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   15
         Left            =   2160
         TabIndex        =   79
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   14
         Left            =   2160
         TabIndex        =   77
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   13
         Left            =   2160
         TabIndex        =   75
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   295
         Index           =   12
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   295
         Index           =   0
         Left            =   2160
         TabIndex        =   87
         Top             =   3480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   529
         _Version        =   393216
         Format          =   104071169
         CurrentDate     =   42471
      End
      Begin VB.Label Label1 
         Caption         =   "--/--/----"
         Height          =   270
         Index           =   27
         Left            =   2160
         TabIndex        =   89
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Date Sold"
         Height          =   270
         Index           =   20
         Left            =   360
         TabIndex        =   88
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Date Bought"
         Height          =   270
         Index           =   18
         Left            =   360
         TabIndex        =   86
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Notes"
         Height          =   270
         Index           =   17
         Left            =   360
         TabIndex        =   84
         Top             =   2295
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Odometer Reading"
         Height          =   270
         Index           =   16
         Left            =   360
         TabIndex        =   82
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Engine Nr"
         Height          =   270
         Index           =   15
         Left            =   360
         TabIndex        =   80
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Vin Nr"
         Height          =   270
         Index           =   14
         Left            =   360
         TabIndex        =   78
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Registration Nr"
         Height          =   270
         Index           =   13
         Left            =   360
         TabIndex        =   76
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Linked Vehicle ID"
         Height          =   270
         Index           =   12
         Left            =   360
         TabIndex        =   74
         Top             =   495
         Width           =   1335
      End
   End
   Begin VB.Frame frame1 
      Caption         =   " Buyers Details "
      Height          =   4575
      Index           =   1
      Left            =   2160
      TabIndex        =   47
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   25
         Left            =   2160
         TabIndex        =   56
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   24
         Left            =   2160
         TabIndex        =   55
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   23
         Left            =   2160
         TabIndex        =   54
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   22
         Left            =   2160
         TabIndex        =   53
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   52
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   19
         Left            =   2160
         TabIndex        =   51
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   11
         Left            =   2160
         TabIndex        =   50
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   10
         Left            =   2160
         TabIndex        =   49
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1020
         Index           =   9
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   3360
         Width           =   3495
      End
      Begin CarListing.ctlDone ctlDone1 
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   70
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   265
         Index           =   26
         Left            =   360
         TabIndex        =   65
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   270
         Index           =   25
         Left            =   360
         TabIndex        =   64
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "ID Number"
         Height          =   270
         Index           =   24
         Left            =   360
         TabIndex        =   63
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Company Name"
         Height          =   270
         Index           =   23
         Left            =   360
         TabIndex        =   62
         Top             =   1572
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Company Reg Nr"
         Height          =   270
         Index           =   22
         Left            =   360
         TabIndex        =   61
         Top             =   1932
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Contact Number"
         Height          =   270
         Index           =   19
         Left            =   360
         TabIndex        =   60
         Top             =   2292
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alt. Contact Number"
         Height          =   270
         Index           =   11
         Left            =   360
         TabIndex        =   59
         Top             =   2655
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Email Address"
         Height          =   270
         Index           =   10
         Left            =   360
         TabIndex        =   58
         Top             =   3015
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Notes"
         Height          =   270
         Index           =   9
         Left            =   360
         TabIndex        =   57
         Top             =   3375
         Width           =   1695
      End
   End
   Begin VB.Frame frame1 
      Caption         =   " Sellers Details "
      Height          =   4575
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1020
         Index           =   8
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   7
         Left            =   2160
         TabIndex        =   43
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   6
         Left            =   2160
         TabIndex        =   41
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   5
         Left            =   2160
         TabIndex        =   9
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   4
         Left            =   2160
         TabIndex        =   8
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   3
         Left            =   2160
         TabIndex        =   7
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   1
         Left            =   2160
         TabIndex        =   5
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   0
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin CarListing.ctlDone ctlDone1 
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   71
         Top             =   1125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         Caption         =   "Notes"
         Height          =   270
         Index           =   8
         Left            =   360
         TabIndex        =   46
         Top             =   3375
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Email Address"
         Height          =   270
         Index           =   7
         Left            =   360
         TabIndex        =   44
         Top             =   3015
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Alt. Contact Number"
         Height          =   270
         Index           =   6
         Left            =   360
         TabIndex        =   42
         Top             =   2655
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Contact Number"
         Height          =   270
         Index           =   3
         Left            =   360
         TabIndex        =   40
         Top             =   2292
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Company Reg Nr"
         Height          =   270
         Index           =   5
         Left            =   360
         TabIndex        =   14
         Top             =   1932
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Company Name"
         Height          =   270
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   1572
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "ID Number"
         Height          =   270
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Last Name"
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   265
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   495
         Width           =   1335
      End
   End
   Begin VB.Frame frame1 
      Caption         =   " Service Details "
      Height          =   4575
      Index           =   3
      Left            =   2160
      TabIndex        =   68
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   123
         Top             =   3840
         Width           =   5775
      End
      Begin VB.CheckBox chkService 
         Caption         =   "Warranty Plan"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   122
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkService 
         Caption         =   "Maintenance Plan"
         Height          =   255
         Index           =   1
         Left            =   1980
         TabIndex        =   121
         Top             =   360
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   " Warranty Plan "
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   116
         Top             =   2760
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   29
            Left            =   240
            TabIndex        =   118
            Top             =   480
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Index           =   2
            Left            =   3480
            TabIndex        =   117
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   42460
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Valid Until"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   120
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Limited KM "
            Height          =   270
            Index           =   31
            Left            =   480
            TabIndex        =   119
            Top             =   255
            Width           =   1935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Maintenance Plan "
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   111
         Top             =   1680
         Width           =   5775
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   113
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   42460
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   28
            Left            =   240
            TabIndex        =   112
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Limited KM "
            Height          =   270
            Index           =   30
            Left            =   480
            TabIndex        =   115
            Top             =   255
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Valid Until"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   114
            Top             =   270
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Service Plan "
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Width           =   5775
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   295
            Index           =   26
            Left            =   240
            TabIndex        =   107
            Top             =   480
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   300
            Index           =   0
            Left            =   3480
            TabIndex        =   108
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   104660993
            CurrentDate     =   42460
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Valid Until"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   110
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Limited KM "
            Height          =   270
            Index           =   29
            Left            =   480
            TabIndex        =   109
            Top             =   255
            Width           =   1935
         End
      End
      Begin VB.CheckBox chkService 
         Caption         =   "Service Plan"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frame1 
      Caption         =   " Add New Expense "
      Height          =   4575
      Index           =   6
      Left            =   2160
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdClearAttachment 
         Height          =   370
         Left            =   2760
         Picture         =   "AddClient.frx":5EB56
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Clear Attachment"
         Top             =   3480
         Visible         =   0   'False
         Width           =   370
      End
      Begin VB.CommandButton cmdAttachment 
         Height          =   370
         Left            =   2280
         Picture         =   "AddClient.frx":6183A
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Add Attachment"
         Top             =   3480
         Width           =   370
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   780
         Index           =   18
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   27
         Left            =   2280
         TabIndex        =   67
         Top             =   1920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   295
         Index           =   20
         Left            =   2280
         TabIndex        =   66
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdSaveExpense 
         Caption         =   "&Save"
         Height          =   525
         Left            =   4680
         Picture         =   "AddClient.frx":64508
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save Record"
         Top             =   3960
         Width           =   1245
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1560
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   2280
         TabIndex        =   31
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         Format          =   210305025
         CurrentDate     =   42460
      End
      Begin VB.Label Label1 
         Caption         =   "Notes"
         Height          =   270
         Index           =   28
         Left            =   480
         TabIndex        =   91
         Top             =   2655
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Expense (Inc. VAT)"
         Height          =   270
         Index           =   21
         Left            =   480
         TabIndex        =   34
         Top             =   2295
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Type of Expense"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Date of Expense"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   1230
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Add New Expense"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Label lblModified 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   12000
      TabIndex        =   125
      Top             =   1920
      Width           =   405
   End
   Begin VB.Label lblBookValue 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12000
      TabIndex        =   124
      Top             =   2730
      Width           =   405
   End
   Begin VB.Label lblCount 
      Caption         =   "0 of 0"
      Height          =   255
      Left            =   8280
      TabIndex        =   95
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   -960
      X2              =   13080
      Y1              =   8205
      Y2              =   8205
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   -960
      X2              =   12840
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Label lblDetails 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "(Year) Make, Model, Series"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9765
      TabIndex        =   17
      Top             =   2280
      Width           =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Menu Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "New Client"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   5115
   End
End
Attribute VB_Name = "AddClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DateSet As Boolean
Dim CurrPic As Integer
Dim PictureCount As Integer
Dim ExpensesMemo(0 To 256) As String
Dim DateChange(1 To 3) As Boolean
Dim i As Integer
Dim DIRTY As Boolean
Dim Resp As Integer
Dim vTday As Date

Private Sub cmdAddExpense_Click()
  
    '...clear all frames
  On Error Resume Next
  For i = 0 To Frame1.Count - 1
    Frame1(i).Visible = False
  Next i
  On Error GoTo 0
  
  Combo2.Clear
  Combo2.AddItem "Vehicle Purchase"     '...must only show for 1st purchase
  Combo2.AddItem "Body Part"
  Combo2.AddItem "License / Registration"
  Combo2.AddItem "Auto Spares"
  Combo2.AddItem "Dent Repair"
  Combo2.AddItem "Paint Repair"
  Combo2.AddItem "Other"
  
End Sub

Private Sub chkService_Click(Index As Integer)

  If chkService(Index).Value = 0 Then
    SetService Index, False
  Else
    SetService Index, True
  End If


End Sub

Public Sub SetService(Index As Integer, OnOROff As Boolean, Optional TODAY As Boolean)

  Dim X As Integer
  
  If TODAY Then
    For X = 0 To 2
      'set dates today for service
      DTPicker3(X).Value = vTday
    
      'if service set already, dont set dirty flag
      If Deal.ServicePlanEnabled(X + 1) Then
        DateChange(X + 1) = True
      Else
        DateChange(X + 1) = False
      End If
    Next X
  End If
  
  If OnOROff = False Then
    Frame3(Index).Enabled = False
    If Index = 0 Then
      Text1(26).BackColor = vbButtonFace
      Text1(26).Text = 0
      Text1(26).TabStop = True
      DTPicker3(0).TabStop = True
    ElseIf Index = 1 Then
      Text1(28).BackColor = vbButtonFace
      Text1(28).Text = 0
      Text1(28).TabStop = True
      DTPicker3(1).TabStop = True
    ElseIf Index = 2 Then
      Text1(29).BackColor = vbButtonFace
      Text1(29).Text = 0
      Text1(29).TabStop = True
      DTPicker3(2).TabStop = True
    End If
  Else
    Frame3(Index).Enabled = True
    If Index = 0 Then
      Text1(26).BackColor = vbWindowBackground
      Text1(26).Text = ""
      Text1(26).TabStop = False
      DTPicker3(0).TabStop = True
    ElseIf Index = 1 Then
      Text1(28).BackColor = vbWindowBackground
      Text1(28).Text = ""
      Text1(28).TabStop = False
      DTPicker3(1).TabStop = True
    ElseIf Index = 2 Then
      Text1(29).BackColor = vbWindowBackground
      Text1(29).Text = ""
      DTPicker3(2).TabStop = True
      Text1(29).TabStop = False
    End If
  End If

  If Text1(26).BackColor = vbButtonFace And Text1(28).BackColor = vbButtonFace And Text1(29).BackColor = vbButtonFace Then
    Text5.BackColor = vbButtonFace
    Text5.Text = ""
  Else
    Text5.BackColor = vbWindowBackground
  End If
    
End Sub


Private Sub cmdAttachment_Click()

  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Please select the document to attach"
  CommonDialog1.Filter = "*.*"
  CommonDialog1.InitDir = GetDocumentsFolder(Me)
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err <> 0 Then On Error GoTo 0: Exit Sub
  On Error GoTo 0
  cmdAttachment.Tag = CommonDialog1.FileName
  cmdAttachment.ToolTipText = "Document Attached"
  cmdClearAttachment.Visible = True

End Sub

Private Sub cmdClearAttachment_Click()

  If MsgBox("Are you sure you want to clear the attachment?", vbYesNo + vbQuestion, Me.Caption) <> vbYes Then Exit Sub
  cmdClearAttachment.Visible = False
  cmdAttachment.Tag = ""
  cmdAttachment.ToolTipText = "Add Attachment"

End Sub

Private Sub cmdClose_Click()
  
  Unload Me
  
End Sub

Private Sub cmdCloseDeal_Click()

  If Deal.VehicleCost = 0 Then
    MsgBox "You cannot close the deal until a vehicle has a purchase price.", vbInformation + vbOKOnly
  Else
    CloseDeal.cldDone = False
    CloseDeal.cldBuyerFirstName = Text1(25).Text
    CloseDeal.cldBuyerLastName = Text1(24).Text
    CloseDeal.cldBuyerContact = Text1(19).Text
    CloseDeal.cldBuyerID = Text1(23).Text
    CloseDeal.cldBuyAmount = Deal.VehicleSold
    CloseDeal.cldReg = Deal.VehicleRegNr
    frmCloseDeal.Show vbModal, MDIForm1
    If Not CloseDeal.cldDone Then Exit Sub
  '  Deal.BuyerFirstName = CloseDeal.cldBuyerFirstName
  '  Deal.BuyerLastName = CloseDeal.cldBuyerLastName
  '  Deal.BuyerIDNumber = CloseDeal.cldBuyerID
  '  Deal.BuyerContactNr = CloseDeal.cldBuyerContact
  '  Deal.VehicleSold = CloseDeal.cldBuyAmount
  '  Deal.VehicleDateSold = CloseDeal.cldBuyDate
  '  Deal.DealClosed = True
    cmdSave_Click
    frmClients.LoadClients 0
  End If

End Sub

Private Sub cmdDelete_Click()
  Dim Resp As Integer
  Dim vErr As String

  Resp = MsgBox("Are you sure you want to delete Deal " & Deal.ID & " and all related images and documents?", vbInformation + vbYesNo, "Delete Deal " & Deal.ID)
  If Resp = 6 Then
    Call DeleteDeal(Deal.ID, vErr)
  
    If vErr = "" Then
      MsgBox "Deal " & Deal.ID & " was successfully removed.", vbInformation + vbOKOnly, "Delete Deal"
      Deal.ID = 0
      frmClients.LoadClients 0, ""
      Unload Me
    Else
      MsgBox "An error has occurred." & vbCrLf & vErr, vbInformation + vbOKOnly, "Delete Deal"
    End If
  End If
  
End Sub

Private Sub cmdLinkCarID_Click()

  ' show the form to link mastercar ID
  frmPCodes.Show

End Sub

Private Sub cmdNext_Click()

  CurrPic = CurrPic + 1
  ShowPicture
  DIRTY = True

End Sub

Private Sub cmdPrev_Click()

  CurrPic = CurrPic - 1
  ShowPicture
  DIRTY = True

End Sub

Private Sub cmdPrintDeal_Click()

  Call PrintInvoice(Deal.ID)

End Sub

Private Sub cmdSave_Click()

                ' items, error  'multi array, 1 place holder, for 2 subs
  Dim ValidItems(1 To 4, 1 To 2)
  Dim vErrString As String
  Dim i As Integer
  Dim ContinueOK As Boolean
  
  '
  '     - start validation
  '
  vErrString = ""
  ContinueOK = True
  ValidItems(1, 1) = Trim(Text1(0).Text)
  ValidItems(1, 2) = "Seller First Name"
  
  ValidItems(2, 1) = Trim(Text1(5).Text)
  ValidItems(2, 2) = "Seller Contact Number"
  
  ValidItems(3, 1) = Trim(Text1(12).Text)
  ValidItems(3, 2) = "Link Car ID"
  
  ValidItems(4, 1) = DTPicker1.Value
  ValidItems(4, 2) = "Date Bought Vehicle"
  
  
  For i = 1 To UBound(ValidItems)
    If i = 4 Then
      If Not DateSet Then
        vErrString = vErrString & "*" & ValidItems(i, 2) & vbCrLf
        ContinueOK = False
      End If
    Else
      If Trim$(ValidItems(i, 1)) = "" Then
        vErrString = vErrString & "*" & ValidItems(i, 2) & vbCrLf
        ContinueOK = False
      End If
    End If
  Next i
  
  
  'check if service is done
  CheckService vErrString
  
  If vErrString <> "" Then
    MsgBox "Please complete the required fields first. " & vbCrLf & vErrString, vbInformation + vbOKOnly, Me.Caption
    Exit Sub
  End If
  
  '
  '     - end validation
  '
  
  Call TextToType
  
  DateSet = True
  
  With rs
    sql = "Select * from Deals WHERE ID = " & Deal.ID
    .Open sql, cn, adOpenKeyset, adLockOptimistic
      If .EOF Then
        rs.AddNew
      End If
      
      rs!BuyerFirstName = Deal.BuyerFirstName
      rs!BuyerLastName = Deal.BuyerLastName
      rs!BuyerIDNumber = Deal.BuyerIDNumber
      rs!BuyerCompanyName = Deal.BuyerCompanyName
      rs!BuyerCompanyRegNr = Deal.BuyerCompanyRegNr
      rs!BuyerContactNr = Deal.BuyerContactNr
      rs!BuyerAltContactNr = Deal.BuyerAltContactNr
      rs!BuyerEmailAddress = Deal.BuyerEmailAddress
      rs!BuyerNotes = Deal.BuyerNotes
    
      rs!SellerFirstName = Deal.SellerFirstName
      rs!SellerLastName = Deal.SellerLastName
      rs!SellerIDNumber = Deal.SellerIDNumber
      rs!SellerCompanyName = Deal.SellerCompanyName
      rs!SellerCompanyRegNr = Deal.SellerCompanyRegNr
      rs!SellerContactNr = Deal.SellerContactNr
      rs!SellerAltContactNr = Deal.SellerAltContactNr
      rs!SellerEmailAddress = Deal.SellerEmailAddress
      rs!SellerNotes = Deal.SellerNotes
      
      rs!VehicleID = Deal.VehicleID
      rs!VehicleRegNr = Deal.VehicleRegNr
      rs!VehicleVINNr = Deal.VehicleVINNr
      rs!VehicleEngineNr = Deal.VehicleEngineNr
      rs!VehicleKM = Deal.VehicleKM
      rs!VehicleNotes = Deal.VehicleNotes
      rs!VehicleImage = Deal.VehicleImage
      rs!VehicleCost = Deal.VehicleCost
      rs!VehicleService = Deal.VehicleService
      rs!VehicleDateBought = Deal.VehicleDateBought
      rs!VehicleDateSold = Deal.VehicleDateSold
      rs!VehicleSold = Deal.VehicleSold
      
      'save the service details
      rs!ServicePlan1 = Deal.ServicePlanEnabled(1)
      rs!ServicePlan2 = Deal.ServicePlanEnabled(2)
      rs!ServicePlan3 = Deal.ServicePlanEnabled(3)
      rs!ServiceKM1 = Deal.ServiceKMs(1)
      rs!ServiceKM2 = Deal.ServiceKMs(2)
      rs!ServiceKM3 = Deal.ServiceKMs(3)
      rs!ServiceDate1 = Deal.ServiceDate(1)
      rs!ServiceDate2 = Deal.ServiceDate(2)
      rs!ServiceDate3 = Deal.ServiceDate(3)
      rs!ServiceNotes = Deal.ServiceNotes
      
      rs!DateModified = Deal.DateModified
      rs!UserModified = Deal.UserModified
      rs!DateCreated = Deal.DateCreated
      rs!UserCreated = Deal.UserCreated
      rs!DealClosed = Deal.DealClosed
      
      .Update
    .Close
    
    If Deal.ID = 0 Then
      MsgBox "New Deal was successfully created.", vbInformation
    Else
      MsgBox "Deal " & Deal.ID & " was successfully saved.", vbInformation
    End If
    
  End With

  Thumb.Visible = True
  lblCount.Visible = True
  CheckPicButtons
  DIRTY = False
  Unload Me
  frmClients.LoadClients 0, ""

End Sub

Private Sub cmdSaveExpense_Click()
  
  Dim vErrStr As String, Recs As Integer, ExpID As Integer, nfExt As String

  vErrStr = ""
  If Combo2.ListIndex = -1 Then
    vErrStr = vErrStr & "* Type of Expense" & vbCrLf
  ElseIf Combo2.Text = "Other" And Trim(Text1(27).Text) = "" Then
    vErrStr = vErrStr & "* Type of Expense" & vbCrLf
  End If

  If Trim(Text1(20).Text) = "" Then
    vErrStr = vErrStr & "* Value of Expense" & vbCrLf
  End If
  
  'cant save until deal has been saved
  If Deal.ID = 0 Then
    MsgBox "Please save the Deal first before linking expenses to this deal.", vbInformation + vbOKCancel, Me.Caption
    Exit Sub
  End If

  If vErrStr = "" Then
    
    With rs
      .Open "Select * from VehExpenses", cn, adOpenKeyset, adLockOptimistic
        .AddNew
          rs!DealID = Deal.ID     '...link to the deal
          rs!ExpenseDate = GetDateVal(DTPicker1)
          If Combo2.ListIndex = 6 Then
            rs!ExpenseType = Trim$(Text1(27).Text)
          Else
            rs!ExpenseType = Combo2.Text
          End If
          rs!ExpenseValue = Val(Text1(20).Text)
          If Trim$(Text1(18).Text) = "" Then Text1(18).Text = NonRequired
          rs!Notes = Text1(18).Text
        .Update
        ExpID = rs!ID
      .Close
    End With
    
    If cmdAttachment.Tag <> "" Then
      nfExt = Right(cmdAttachment.Tag, Len(cmdAttachment.Tag) - InStrRev(cmdAttachment.Tag, "."))
      Recs = OpenDocStore("EA" & Trim(Str(ExpID)))
      If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
      If Recs > 0 Then MsgBox "Document name already exists for this Deal": CloseDocStore: Exit Sub
      On Error GoTo 0
      DocStoreRS.AddNew
      DocStoreRS.Fields("DocID") = "EA" & Trim(Str(ExpID))
      DocStoreRS.Fields("DocName") = "expenseattachment"
      DocStoreRS.Fields("DocExtension") = nfExt
      FileToField cmdAttachment.Tag, "DocContents", "DocExtension"
      DocStoreRS.Update
      CloseDocStore
    End If
    
    'if its a vehicle expense, then save it
    If Trim$(Combo2.Text) = "Vehicle Purchase" Then
      Deal.VehicleCost = Val(Text1(20).Text)
      Text2.Text = FormatNumber(Text1(20).Text, 2)
    Else
      Deal.VehicleService = Deal.VehicleService + Val(Text1(20).Text)
      Text3.Text = FormatNumber(Deal.VehicleService, 2)
      'update service cost
      With rs
        .Open "select * from Deals where ID = " & Deal.ID
          If .EOF Then
            .Close
          Else
            rs!VehicleService = Deal.VehicleService
            .Update
          .Close
          End If
      End With
    End If
    
  Else
    MsgBox "Please complete the following fields first." & vbCrLf & vErrStr, vbInformation
    Exit Sub
  End If
  
  ' load combo for expenses
  LoadExpenses
  MsgBox "Expense added to Deal ID " & Deal.ID, vbInformation
  Combo2.ListIndex = -1
  Text1(20).Text = ""
  Text1(18).Text = ""
  cmdAttachment.ToolTipText = "Add Attachment"
  cmdAttachment.Tag = ""
  cmdClearAttachment.Visible = False
  DIRTY = False

End Sub

Private Sub Combo2_Click()

  If Trim(Combo2.Text) = "Other" Then
    Text1(27).Visible = True
    Text1(27).Text = ""
    Text1(27).TabStop = True
    Text1(27).TabIndex = 3
    
  Else
    Text1(27).Visible = False
    Text1(27).Text = NonRequired
    Text1(27).TabStop = False
  End If

End Sub

Private Sub DTPicker2_Change(Index As Integer)

  If Index = 0 Then
    DateSet = True
  End If

End Sub

Private Sub DTPicker3_CloseUp(Index As Integer)
  DateChange(Index + 1) = True
End Sub

Private Sub Form_Load()

  Dim vYear As String

  '...set defaults
  DIRTY = False
  Call CenterForm(Me)
  ctlDone1(0).Done = False
  ctlDone1(1).Done = False
  Thumb.BorderStyle = bFixed
  
  SetService 0, False, True
  SetService 1, False, True
  SetService 2, False, True
  
  'print option only available after deal close
  cmdPrintDeal.Enabled = False
  cmdCloseDeal.Enabled = False
  
  vTday = Now()
  DTPicker1.Value = vTday
  DTPicker2(0).Value = vTday
  
  lblSold.Visible = False
  If Deal.ID = 0 Then
    lblHeader.Caption = "New Deal"
    lblDetails.Caption = ""
    lblBookValue.Caption = ""
    lblModified.Caption = ""
    Me.Caption = "Add New Deal"
    DateSet = False
  Else
    DateSet = True
    DTPicker2(0).Enabled = False            'cant change bought date after first save
    Call GetDeal(Deal.ID)                   '... GetDetails
    Call GetMasterVehicleDetails(Val(Deal.VehicleID))        '...Get Vehicle details
    lblHeader.Caption = "Deal ID " & Deal.ID
    lblDetails.Caption = "(" & MasterVehicle.ModelYear & ") " & MasterVehicle.VehicleMake & " " & MasterVehicle.VehicleModel
    If MasterVehicle.BookValue = 0 Then
      lblBookValue.Caption = "Book Value : ( Not Set )"
    Else
      lblBookValue.Caption = "Book Value : " & FormatNumber(MasterVehicle.BookValue, 2) & " (" & Format(MasterVehicle.BookDate, "dd MMMM yyyy") & ")"
    End If
    lblModified.Caption = "Last Modified : " & Format(Deal.DateModified, "dd MMMM yyyy")
    Me.Caption = "Deal ID " & Deal.ID
    Call TypeToText
    If ValidID(Deal.SellerIDNumber) Then
      ctlDone1(1).Done = True
    End If
    If ValidID(Deal.BuyerIDNumber) Then
      ctlDone1(0).Done = True
    End If
    
    Text2.Text = FormatNumber(Deal.VehicleCost, 2)
    Text3.Text = FormatNumber(Deal.VehicleService, 2)
    Text4.Text = FormatNumber(Deal.VehicleSold, 2)
        
    cmdCloseDeal.Enabled = True
    
        
    'if deal is closed, prevent saving
    If Deal.DealClosed Then
      lblSold.Visible = True
      Label1(27).Caption = Format(Deal.VehicleDateSold, "dd MMMM yyyy")
      Call CloseDealOFF
    End If
  End If

  '...set menu options
  List1.Clear
  List1.AddItem "Sellers Details", 0
  List1.AddItem "Buyers Details", 1
  List1.AddItem "Vehicle Details", 2
  List1.AddItem "Service Details", 3
  List1.AddItem "   ---------   ", 4
  List1.AddItem "Add Expenses", 5
  
  '..only load options after save
  If Deal.ID <> 0 Then
    List1.AddItem "View Expenses", 6
    List1.AddItem "View Documents", 7
    cmdLinkCarID.Enabled = False      'cannot link new car after save
  End If
  
  '...clear all frames
  On Error Resume Next
  For i = 0 To Frame1.Count - 1
    Frame1(i).Visible = False
  Next i
  On Error GoTo 0
  '...set first frame visible
  Frame1(0).Visible = True
  Call DoTabOrder(0)
  
  ' load combo for expenses
  LoadExpenses
      
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim mRsp As VbMsgBoxResult

  'if deal is already closed, then dont ask, just exit
  If Deal.DealClosed Then
    DIRTY = False
  End If

  If DIRTY Then
    mRsp = MsgBox("Exit without saving?", vbQuestion + vbYesNo, Me.Caption)
    If mRsp = vbYes Then
      RemoveTempFiles
    Else
      Cancel = 1
      Exit Sub
    End If
  End If


End Sub

Private Sub List1_Click()

  Dim i As Integer
  
  If List1.ListIndex = 4 Then Exit Sub
  
  On Error Resume Next
  Frame1(6).Visible = False
  For i = 0 To Frame1.Count - 1
    Frame1(i).Visible = False
  Next i
  On Error GoTo 0
 
'  List1.AddItem "Sellers Details", 0
'  List1.AddItem "Buyers Details", 1
'  List1.AddItem "Vehicle Details", 2
'  List1.AddItem "Service Details", 3
'  List1.AddItem "   ---------   ", 4
'  List1.AddItem "Add Expenses", 5
'  List1.AddItem "View Expenses", 6
'  List1.AddItem "Add Document", 7
'  List1.AddItem "View Document", 8
 
  Select Case List1.ListIndex
    Case 0, 1, 2, 3
      Frame1(List1.ListIndex).Visible = True
      Frame1(List1.ListIndex).Caption = List1.List(List1.ListIndex)
    Case 5
      Frame1(6).Visible = True
    Case 6, 7, 8
      Frame1(5).Visible = True
      Frame1(5).Caption = List1.List(List1.ListIndex)
      cmdDeleteDoc.Visible = False
      cmdChangeAtt.Visible = False
      cmdViewAtt.Visible = False
      cmdRemoveAtt.Visible = False
      cmdAddAtt.Visible = False
      If List1.ListIndex = 6 Then
        Call GetExpenses(Deal.ID)
        cmdAddDoc.Visible = False
      Else
        Call GetDocs(Deal.ID)
        cmdAddDoc.Visible = True
      End If
    Case Else
  End Select
  
  Call DoTabOrder(List1.ListIndex)
    
End Sub

Private Sub List2_Click()
  
  If List2.ListIndex = -1 Then Exit Sub
  If List1.List(List1.ListIndex) = "View Expenses" Then
    'show the tooltip for the selected note
    lblNotes.Caption = ""
    lblNotes.Caption = ExpensesMemo(List2.ListIndex)
    lblNotes.Refresh
    CheckExpenseAtt
  Else
    cmdDeleteDoc.Visible = True
  End If

End Sub

Private Sub List2_DblClick()

  Dim Recs As Integer, tmpPath As String

  If List1.List(List1.ListIndex) = "View Expenses" Then
    If cmdViewAtt.Visible = True Then cmdViewAtt_Click
  Else
    If List2.ListIndex = -1 Then Exit Sub
    tmpPath = App.Path
    If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
    tmpPath = tmpPath & "tmp.img"
    Recs = OpenDocStore("DD" & Trim(Str(Deal.ID)), List2.List(List2.ListIndex))
    If Recs = -1 Then MsgBox DocError: Exit Sub
    If Recs <> 1 Then MsgBox "Document not found": Exit Sub
    FileFromField tmpPath, "DocContents", "DocExtension"
    OpenFile2 tmpPath
    CloseDocStore
  End If

End Sub

Private Sub Text1_Change(Index As Integer)

  ' change was made
  DIRTY = True
  
  If Index = 2 Or Index = 23 Then
    CheckID Text1(Index).Text, Index
  End If



End Sub


Private Sub CheckID(vID As String, vControl As Integer)


  Select Case vControl
    Case 2
      If ValidID(vID) Then
        ctlDone1(1).Done = True
      Else
        ctlDone1(1).Done = False
      End If
    Case 23
      If ValidID(vID) Then
        ctlDone1(0).Done = True
      Else
        ctlDone1(0).Done = False
      End If
  End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
  Text1(Index).SelStart = 0
  Text1(Index).SelLength = Len(Trim(Text1(Index).Text))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

  ' only allow numbers
  Select Case Index       '26,28,29 = KM for service
    Case 5, 6, 11, 19, 16, 26, 28, 29
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Then
      Else
        KeyAscii = 0
      End If
    '..only allow numbers for costings, money etc
    Case 13, 14, 15
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case 20
      If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 46 Then
        '...only allow 1 comma
        If InStr(1, Text1(Index).Text, ".") > 0 And KeyAscii = 46 Then
          KeyAscii = 0
        End If
      Else
        KeyAscii = 0
      End If
    Case Else
    ' fail over
  End Select

End Sub

Private Sub Text1_LostFocus(Index As Integer)

  Dim v As ValidatorClass
  Set v = New ValidatorClass
  Dim ErrString As String
  Dim vControl As Integer
  Dim ShowMsg As Boolean
  
  ShowMsg = True

  'Dont show msg if field blank
  If Trim$(Text1(Index).Text) = "" Then
    ShowMsg = False
  End If

  Select Case Index
    Case 2, 23     'ID Number
      v.Validate oID_Number, Text1(Index).Text, ErrString
      Select Case Index
        Case 2
          If ErrString = "" Then
            ctlDone1(1).Done = True
          Else
            ctlDone1(1).Done = False
          End If
        Case 23
          If ErrString = "" Then
            ctlDone1(0).Done = True
          Else
            ctlDone1(0).Done = False
          End If
        End Select
     Case 7, 10    'email addresses
        If Trim$(Text1(Index).Text) <> "" Then
          If ValidEmail(Text1(Index).Text) = False Then
            MsgBox "The Email Address you have entered is invalid.", vbInformation
          End If
        End If
      Case 20     'validate cost,maintainance etc
        'If Trim$(Text1(Index).Text) <> "" Then
        '  Text1(Index).Text = FormatNumber(Text1(Index).Text, 2)
        'End If
  End Select

  'display any errors
  If ShowMsg And ErrString <> "" Then
    MsgBox ErrString, vbInformation + vbOKOnly
  End If
  Set v = Nothing
  
End Sub

Private Sub Text5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Text5.ToolTipText = Text5.Text
End Sub

Private Sub Thumb_DblClick()

  Dim Recs As Integer, tmpPath As String

  tmpPath = App.Path
  If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
  tmpPath = tmpPath & "tmp.img"
  Recs = OpenDocStore("D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(CurrPic)))
  If Recs > 0 Then
    FileFromField tmpPath, "DocContents", "DocExtension"
    OpenFile2 tmpPath
    DoEvents
  End If
  CloseDocStore

End Sub

Private Sub Thumb_NewDropImage(NewPath As String)

  Dim Recs As Integer

  If CurrPic <= PictureCount And PictureCount <> 0 Then
    If MsgBox("Would you like to overwrite the existing image?", vbYesNo, "Overwrite Image?") <> vbYes Then Exit Sub
  End If
  Thumb.SetPicture LoadPicture(NewPath)
  If CurrPic = 0 Then CurrPic = 1
  Recs = OpenDocStore("D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(CurrPic)))
  If Recs = -1 Then
    MsgBox DocError
  Else
    If Recs = 0 Then
      On Error GoTo 0
      DocStoreRS.AddNew
      DocStoreRS.Fields("DocID") = "D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(CurrPic))
      DocStoreRS.Fields("DocName") = "dealpic"
      DocStoreRS.Fields("DocExtension") = "syspic"
      PictureCount = PictureCount + 1
    End If
    FileToField NewPath, "DocContents", "DocExtension"
    DocStoreRS.Update
  End If
  Deal.VehicleImage = Trim(Str(CurrPic))
  DIRTY = True
  CloseDocStore
  CheckPicButtons

End Sub

Private Sub Timer1_Timer()

  Timer1.Interval = 0
  
  If Deal.ID = 0 Then
    Thumb.Visible = False
    lblCount.Visible = False
    cmdNext.Visible = False
    cmdPrev.Visible = False
    cmdClearPic.Visible = False
  End If
  PictureCount = OpenDocStore("D" & Trim(Str(Deal.ID))): CloseDocStore
  If PictureCount = 0 Then
    cmdNext.Visible = False
    cmdPrev.Visible = False
    cmdClearPic.Visible = False
    lblCount.Caption = "No Pictures"
    If Not Deal.DealClosed Then
      Thumb.ClearPicture
    End If
  Else
    If Deal.VehicleImage = "" Then
      CurrPic = 1
    Else
      CurrPic = Val(Deal.VehicleImage)
      If CurrPic > PictureCount Then CurrPic = PictureCount
    End If
    ShowPicture
  End If
  
End Sub

'...From the text to type
Public Sub TextToType()
  Dim X As Integer
  
  If Trim$(Text1(0).Text) <> "" Then Deal.SellerFirstName = Text1(0).Text Else Deal.SellerFirstName = NonRequired
  If Trim$(Text1(1).Text) <> "" Then Deal.SellerLastName = Text1(1).Text Else Deal.SellerLastName = NonRequired
  If Trim$(Text1(2).Text) <> "" Then Deal.SellerIDNumber = Text1(2).Text Else Deal.SellerIDNumber = NonRequired
  If Trim$(Text1(3).Text) <> "" Then Deal.SellerCompanyName = Text1(3).Text Else Deal.SellerCompanyName = NonRequired
  If Trim$(Text1(4).Text) <> "" Then Deal.SellerCompanyRegNr = Text1(4).Text Else Deal.SellerCompanyRegNr = NonRequired
  If Trim$(Text1(5).Text) <> "" Then Deal.SellerContactNr = Text1(5).Text Else Deal.SellerContactNr = NonRequired
  If Trim$(Text1(6).Text) <> "" Then Deal.SellerAltContactNr = Text1(6).Text Else Deal.SellerAltContactNr = NonRequired
  If Trim$(Text1(7).Text) <> "" Then Deal.SellerEmailAddress = Text1(7).Text Else Deal.SellerEmailAddress = NonRequired
  If Trim$(Text1(8).Text) <> "" Then Deal.SellerNotes = Text1(8).Text Else Deal.SellerNotes = NonRequired

  If Trim$(Text1(25).Text) <> "" Then Deal.BuyerFirstName = Text1(25).Text Else Deal.BuyerFirstName = NonRequired
  If Trim$(Text1(24).Text) <> "" Then Deal.BuyerLastName = Text1(24).Text Else Deal.BuyerLastName = NonRequired
  If Trim$(Text1(23).Text) <> "" Then Deal.BuyerIDNumber = Text1(23).Text Else Deal.BuyerIDNumber = NonRequired
  If Trim$(Text1(22).Text) <> "" Then Deal.BuyerCompanyName = Text1(22).Text Else Deal.BuyerCompanyName = NonRequired
  If Trim$(Text1(21).Text) <> "" Then Deal.BuyerCompanyRegNr = Text1(21).Text Else Deal.BuyerCompanyRegNr = NonRequired
  If Trim$(Text1(19).Text) <> "" Then Deal.BuyerContactNr = Text1(19).Text Else Deal.BuyerContactNr = NonRequired
  If Trim$(Text1(11).Text) <> "" Then Deal.BuyerAltContactNr = Text1(11).Text Else Deal.BuyerAltContactNr = NonRequired
  If Trim$(Text1(10).Text) <> "" Then Deal.BuyerEmailAddress = Text1(10).Text Else Deal.BuyerEmailAddress = NonRequired
  If Trim$(Text1(9).Text) <> "" Then Deal.BuyerNotes = Text1(9).Text Else Deal.BuyerNotes = NonRequired
  
  Deal.VehicleID = Text1(12).Text
  If Trim$(Text1(13).Text) <> "" Then Deal.VehicleRegNr = Text1(13).Text Else Deal.VehicleRegNr = NonRequired
  If Trim$(Text1(14).Text) <> "" Then Deal.VehicleVINNr = Text1(14).Text Else Deal.VehicleVINNr = NonRequired
  If Trim$(Text1(15).Text) <> "" Then Deal.VehicleEngineNr = Text1(15).Text Else Deal.VehicleEngineNr = NonRequired
  If Trim$(Text1(16).Text) <> "" Then Deal.VehicleKM = Text1(16).Text Else Deal.VehicleKM = NonRequired
  If Trim$(Text1(17).Text) <> "" Then Deal.VehicleNotes = Text1(17).Text Else Deal.VehicleNotes = NonRequired
  
  If DTPicker2(0) <> "" Then Deal.VehicleDateBought = GetDateVal(DTPicker2(0)) Else DTPicker2(0) = NonRequired

  'set service info
  For X = 1 To 3
    Deal.ServicePlanEnabled(X) = CBool(chkService(X - 1).Value)
    If X = 1 Then
      Deal.ServiceKMs(X) = Text1(26).Text
    ElseIf X = 2 Then
      Deal.ServiceKMs(X) = Text1(28).Text
    ElseIf X = 3 Then
      Deal.ServiceKMs(X) = Text1(29).Text
    End If
    Deal.ServiceDate(X) = GetDateVal(DTPicker3(X - 1))
  Next X
  
  If Trim$(Text5.Text) <> "" Then Deal.ServiceNotes = Text5.Text Else Deal.ServiceNotes = NonRequired
  Deal.DateModified = TodaysDate      'change after each save
  Deal.UserModified = User.ID         'active user that made change
  
  '...only for new deals
  If Deal.ID = 0 Then
    Deal.DateCreated = TodaysDate
    Deal.UserCreated = User.ID
    Deal.DealClosed = False
  End If
  
  'over deal if this deal has been closed
  If CloseDeal.cldDone Then
    Deal.BuyerFirstName = CloseDeal.cldBuyerFirstName
    Deal.BuyerLastName = CloseDeal.cldBuyerLastName
    Deal.BuyerIDNumber = CloseDeal.cldBuyerID
    Deal.BuyerContactNr = CloseDeal.cldBuyerContact
    Deal.VehicleSold = CloseDeal.cldBuyAmount
    Deal.VehicleDateSold = CloseDeal.cldBuyDate
    Deal.DealClosed = True
  End If
 
End Sub

'...From type to textboxes
Public Sub TypeToText()
  Dim X As Integer


  Text1(0).Text = Deal.SellerFirstName
  Text1(1).Text = Deal.SellerLastName
  Text1(2).Text = Deal.SellerIDNumber
  Text1(3).Text = Deal.SellerCompanyName
  Text1(4).Text = Deal.SellerCompanyRegNr
  Text1(5).Text = Deal.SellerContactNr
  Text1(6).Text = Deal.SellerAltContactNr
  Text1(7).Text = Deal.SellerEmailAddress
  Text1(8).Text = Deal.SellerNotes

  Text1(25).Text = Deal.BuyerFirstName
  Text1(24).Text = Deal.BuyerLastName
  Text1(23).Text = Deal.BuyerIDNumber
  Text1(22).Text = Deal.BuyerCompanyName
  Text1(21).Text = Deal.BuyerCompanyRegNr
  Text1(19).Text = Deal.BuyerContactNr
  Text1(11).Text = Deal.BuyerAltContactNr
  Text1(10).Text = Deal.BuyerEmailAddress
  Text1(9).Text = Deal.BuyerNotes
  
  Text1(12).Text = Deal.VehicleID
  Text1(13).Text = Deal.VehicleRegNr
  Text1(14).Text = Deal.VehicleVINNr
  Text1(15).Text = Deal.VehicleEngineNr
  Text1(16).Text = Deal.VehicleKM
  Text1(17).Text = Deal.VehicleNotes
  
  If Deal.VehicleDateBought > 0 Then
    DTPicker2(0).Value = Format(Deal.VehicleDateBought, "dd MMMM yyyy")
  Else
    DTPicker2(0).Value = 0
  End If
  
  'set service info
  For X = 1 To 3
    If Deal.ServicePlanEnabled(X) Then
      chkService(X - 1).Value = 1
      DTPicker3(X - 1).Value = Format(Deal.ServiceDate(X), "dd MMMM yyyy")
    Else
      chkService(X - 1).Value = 0
      DTPicker3(X - 1).Value = vTday
    End If
    
    If X = 1 Then
      Text1(26).Text = Deal.ServiceKMs(X)
    ElseIf X = 2 Then
      Text1(28).Text = Deal.ServiceKMs(X)
    ElseIf X = 3 Then
      Text1(29).Text = Deal.ServiceKMs(X)
    End If
  Next X
  Text5.Text = Deal.ServiceNotes

  DIRTY = False   ' have to set it to false bcoz of text chnage

End Sub

Private Sub ShowPicture()

  Dim Recs As Integer, tmpPath As String

  tmpPath = App.Path
  If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
  tmpPath = tmpPath & "tmp.img"
  Recs = OpenDocStore("D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(CurrPic)))
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
  Deal.VehicleImage = Trim(Str(CurrPic))
  CloseDocStore
  CheckPicButtons

End Sub

Public Sub CheckPicButtons()

  If Deal.ID = 0 Then Exit Sub
  If CurrPic = 1 Then cmdPrev.Visible = False Else cmdPrev.Visible = True
  If CurrPic <= PictureCount Then cmdNext.Visible = True Else cmdNext.Visible = False
  If CurrPic > PictureCount Then
    lblCount.Caption = CurrPic & " of " & CurrPic & " (New)"
    cmdClearPic.Visible = False
  Else
    If PictureCount = 0 Then
      cmdNext.Visible = False
      cmdPrev.Visible = False
      cmdClearPic.Visible = False
      lblCount.Caption = "No Pictures"
    Else
      lblCount.Caption = CurrPic & " of " & PictureCount
      cmdClearPic.Visible = True
    End If
  End If

End Sub

Private Sub cmdClearPic_Click()

  Dim Recs As Integer, vCounter As Integer
  
  If MsgBox("Are you sure you want to remove this image?", vbInformation + vbYesNo, "Remove Image?") <> vbYes Then Exit Sub
  Recs = OpenDocStore("D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(CurrPic)))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs = 0 Then MsgBox "Image not found": CloseDocStore: Exit Sub
  If Recs <> 1 Then MsgBox "More than one Image found": CloseDocStore: Exit Sub
  If PictureCount = 1 Then
    DocStoreRS.Delete
    PictureCount = 0
    CurrPic = 0
    Thumb.ClearPicture
    Deal.VehicleImage = ""
    DIRTY = True
    CloseDocStore
    CheckPicButtons
  ElseIf CurrPic = PictureCount Then
    DocStoreRS.Delete
    PictureCount = PictureCount - 1
    CurrPic = CurrPic - 1
    CloseDocStore
    ShowPicture
    CheckPicButtons
  Else
    DocStoreRS.Delete
    CloseDocStore
    For vCounter = CurrPic To PictureCount - 1
      Recs = OpenDocStore("D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(vCounter + 1)))
      If Recs = -1 Then
        MsgBox DocError
      ElseIf Recs = 0 Then
        MsgBox "Image not found"
      ElseIf Recs <> 1 Then
        MsgBox "More than one Image found"
      Else
        DocStoreRS.Fields("DocID") = "D" & Trim(Str(Deal.ID)) & "-" & Trim(Str(vCounter))
        DocStoreRS.Update
      End If
      CloseDocStore
    Next vCounter
    PictureCount = PictureCount - 1
    CurrPic = CurrPic - 1
    If CurrPic = 0 Then CurrPic = 1
    ShowPicture
    CheckPicButtons
  End If

End Sub

Private Function GetExpenses(ID As Integer)
    
  Dim AfterTypeTab As String
  Dim AmountString As String
  Dim ExpenseItemCount As Integer

  List2.Clear
  List2.OLEDropMode = 0
  List2.Columns = 4
  ExpenseItemCount = 0
  With rs
    .Open "Select * from VehExpenses WHERE DealID = " & ID, cn, adOpenKeyset, adLockOptimistic
      Do While Not .EOF
        ExpenseItemCount = ExpenseItemCount + 1
        AfterTypeTab = vbTab & vbTab
        If Me.TextWidth(rs!ExpenseType) > 1400 Then
          AfterTypeTab = vbTab
        ElseIf Me.TextWidth(rs!ExpenseType) < 700 Then
          AfterTypeTab = AfterTypeTab & vbTab
        End If
        AmountString = FormatNumber(rs!ExpenseValue, 2)
        Do Until Me.TextWidth(AmountString) >= 1040
          AmountString = " " & AmountString
        Loop
        List2.AddItem "  " & ExpenseItemCount & vbTab & Format(rs!ExpenseDate, "dd MMMM yyyy") & vbTab & rs!ExpenseType & AfterTypeTab & AmountString
        List2.ItemData(List2.NewIndex) = rs!ID
        ExpensesMemo(List2.NewIndex) = Trim$(rs!Notes)
        .MoveNext
      Loop
    .Close
  End With
    
End Function

Private Function GetDocs(ID As Integer)
    
  Dim AfterTypeTab As String, Recs As Integer, RecZ As Integer

  List2.Clear
  List2.Columns = 1
  List2.OLEDropMode = 1
  Recs = OpenDocStore("DD" & Trim(Str(Deal.ID)))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Function
  If Recs = 0 Then CloseDocStore: Exit Function
  If Not DocStoreRS.BOF Then DocStoreRS.MoveFirst
  For RecZ = 1 To Recs
    List2.AddItem DocStoreRS.Fields("DocName")
    DocStoreRS.MoveNext
  Next RecZ
  CloseDocStore
    
End Function

Private Sub cmdDeleteDoc_Click()

  Dim Recs As Integer

  If MsgBox("Are you sure you want to remove this document?", vbYesNo + vbQuestion, "Remove Document?") <> vbYes Then Exit Sub
  Recs = OpenDocStore("DD" & Trim(Str(Deal.ID)), List2.List(List2.ListIndex))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs = 0 Then MsgBox "Document not found": CloseDocStore: Exit Sub
  If Recs > 1 Then MsgBox "More than one document found": CloseDocStore: Exit Sub
  DocStoreRS.Delete
  CloseDocStore
  List2.RemoveItem List2.ListIndex
  cmdDeleteDoc.Visible = False

End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

  If List1.List(List1.ListIndex) = "View Expenses" Then Exit Sub
  DoAddDoc Data.Files(1)

End Sub

Private Sub cmdAddDoc_Click()

  If List1.List(List1.ListIndex) = "View Expenses" Then Exit Sub
  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Please select the document to add"
  CommonDialog1.Filter = "*.*"
  CommonDialog1.InitDir = GetDocumentsFolder(Me)
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err <> 0 Then On Error GoTo 0: Exit Sub
  On Error GoTo 0
  DoAddDoc CommonDialog1.FileName

End Sub

Private Sub DoAddDoc(DocToAddPath As String)

  Dim NewFileName As String, nfExt As String, Recs As Integer

  nfExt = Right(DocToAddPath, Len(DocToAddPath) - InStrRev(DocToAddPath, "."))
  NewFileName = DocToAddPath
  NewFileName = Left(NewFileName, InStrRev(NewFileName, ".") - 1)
  NewFileName = Right(NewFileName, Len(NewFileName) - InStrRev(NewFileName, "\"))
  NewFileName = InputBox("Please give a name for this document:", "Add Document", NewFileName)
  If NewFileName = "" Then Exit Sub
  Recs = OpenDocStore("DD" & Trim(Str(Deal.ID)), NewFileName)
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs > 0 Then MsgBox "Document name already exists for this Deal": CloseDocStore: Exit Sub
  List2.AddItem NewFileName
  List2.ListIndex = List2.NewIndex
  On Error GoTo 0
  DocStoreRS.AddNew
  DocStoreRS.Fields("DocID") = "DD" & Trim(Str(Deal.ID))
  DocStoreRS.Fields("DocName") = NewFileName
  DocStoreRS.Fields("DocExtension") = nfExt
  FileToField DocToAddPath, "DocContents", "DocExtension"
  DocStoreRS.Update
  CloseDocStore

End Sub

Public Sub LoadExpenses()

  '..adding expense
  Combo2.Clear
  If Deal.VehicleCost = 0 Then
    Combo2.AddItem "Vehicle Purchase"     '...must only show for 1st purchase
  End If
  If Deal.ID <> 0 Then
    Combo2.AddItem "Body Part", 0
    Combo2.AddItem "License / Registration", 1
    Combo2.AddItem "Auto Spares", 2
    Combo2.AddItem "Dent Repair", 3
    Combo2.AddItem "Paint Repair", 4
    Combo2.AddItem "Car Wash / Valet", 5
    Combo2.AddItem "Other", 6
  End If

End Sub

Private Sub CheckExpenseAtt()

  Dim Recs As Integer

  Recs = OpenDocStore("EA" & Trim(Str(List2.ItemData(List2.ListIndex))))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs = 0 Then
    cmdAddAtt.Visible = True
    cmdChangeAtt.Visible = False
    cmdViewAtt.Visible = False
    cmdRemoveAtt.Visible = False
  Else
    cmdAddAtt.Visible = False
    cmdChangeAtt.Visible = True
    cmdViewAtt.Visible = True
    cmdRemoveAtt.Visible = True
  End If
  CloseDocStore

End Sub

Private Sub cmdViewAtt_Click()

  Dim Recs As Integer, tmpPath As String

  tmpPath = App.Path
  If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
  tmpPath = tmpPath & "tmp.img"
  Recs = OpenDocStore("EA" & Trim(Str(List2.ItemData(List2.ListIndex))))
  If Recs > 0 Then
    FileFromField tmpPath, "DocContents", "DocExtension"
    OpenFile2 tmpPath
    DoEvents
  End If
  CloseDocStore

End Sub

Private Sub cmdRemoveAtt_Click()

  Dim Recs As Integer

  If MsgBox("Are you sure you want to remove this attachment?", vbYesNo + vbQuestion, Me.Caption) <> vbYes Then Exit Sub
  Recs = OpenDocStore("EA" & Trim(Str(List2.ItemData(List2.ListIndex))))
  If Recs > 0 Then DocStoreRS.Delete
  CloseDocStore
  cmdAddAtt.Visible = True
  cmdChangeAtt.Visible = False
  cmdViewAtt.Visible = False
  cmdRemoveAtt.Visible = False

End Sub

Private Sub cmdAddAtt_Click()

  Dim nfExt As String, Recs As Integer

  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Please select the attachment to add"
  CommonDialog1.Filter = "*.*"
  CommonDialog1.InitDir = GetDocumentsFolder(Me)
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err <> 0 Then On Error GoTo 0: Exit Sub
  On Error GoTo 0
  nfExt = Right(CommonDialog1.FileName, Len(CommonDialog1.FileName) - InStrRev(CommonDialog1.FileName, "."))
  Recs = OpenDocStore("EA" & Trim(Str(List2.ItemData(List2.ListIndex))))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs > 0 Then MsgBox "Attachment name already exists for this Expense": CloseDocStore: Exit Sub
  DocStoreRS.AddNew
  DocStoreRS.Fields("DocID") = "EA" & Trim(Str(List2.ItemData(List2.ListIndex)))
  DocStoreRS.Fields("DocName") = "expenseattachment"
  DocStoreRS.Fields("DocExtension") = nfExt
  FileToField CommonDialog1.FileName, "DocContents", "DocExtension"
  DocStoreRS.Update
  CloseDocStore
  cmdAddAtt.Visible = False
  cmdChangeAtt.Visible = True
  cmdViewAtt.Visible = True
  cmdRemoveAtt.Visible = True

End Sub

Private Sub cmdChangeAtt_Click()

  Dim nfExt As String, Recs As Integer

  CommonDialog1.CancelError = True
  CommonDialog1.DialogTitle = "Please select the new attachment"
  CommonDialog1.Filter = "*.*"
  CommonDialog1.InitDir = GetDocumentsFolder(Me)
  On Error Resume Next
  CommonDialog1.ShowOpen
  If Err <> 0 Then On Error GoTo 0: Exit Sub
  On Error GoTo 0
  nfExt = Right(CommonDialog1.FileName, Len(CommonDialog1.FileName) - InStrRev(CommonDialog1.FileName, "."))
  Recs = OpenDocStore("EA" & Trim(Str(List2.ItemData(List2.ListIndex))))
  If Recs = -1 Then MsgBox DocError: CloseDocStore: Exit Sub
  If Recs = 0 Then MsgBox "Attachment not found for this Expense": CloseDocStore: Exit Sub
  DocStoreRS.Fields("DocExtension") = nfExt
  FileToField CommonDialog1.FileName, "DocContents", "DocExtension"
  DocStoreRS.Update
  CloseDocStore

End Sub

Public Sub CloseDealOFF()
  Dim i As Integer
  Dim ctl As Control

  
  'blank all textboxes
  For i = 0 To Text1.Count - 1
    On Error Resume Next
      Text1(i).BackColor = vbButtonFace
      Text1(i).Enabled = False
    On Error GoTo 0
  Next i

  'disable all butttons
  For Each ctl In AddClient.Controls
    If TypeOf ctl Is CommandButton Then
      If ctl.Enabled = True Then
        ctl.Enabled = False
      End If
    End If
  Next

  'lock all service info
  chkService(0).Enabled = False
  chkService(1).Enabled = False
  chkService(2).Enabled = False
  DTPicker3(0).Enabled = False
  DTPicker3(1).Enabled = False
  DTPicker3(2).Enabled = False
  Text5.Enabled = False
  Text5.BackColor = vbButtonFace

  Thumb.DragMode = vbManual

  cmdClose.Enabled = True
  cmdPrintDeal.Enabled = True
  cmdPrev.Enabled = True
  cmdNext.Enabled = True
  
  'enable cmd for view expenses
  cmdViewAtt.Enabled = True
  cmdAddAtt.Enabled = True
  
  
End Sub


Public Function CheckService(vErrString As String)
Dim r As Integer

For r = 1 To 3
  If chkService(r - 1).Value <> 0 Then
    
    Deal.ServicePlanEnabled(r) = True
    If r = 1 Then
      If Trim(Text1(26).Text) = 0 Or Trim(Text1(26).Text) = "" Then
        vErrString = vErrString & "* Limited KMs " & vbCrLf
      Else
        Deal.ServiceKMs(r) = Text1(26).Text
      End If
    ElseIf r = 2 Then
      If Trim(Text1(28).Text) = 0 Or Trim(Text1(28).Text) = "" Then
        vErrString = vErrString & "* Limited KMs " & vbCrLf
      Else
        Deal.ServiceKMs(r) = Text1(28).Text
      End If
    ElseIf r = 3 Then
      If Trim(Text1(29).Text) = 0 Or Trim(Text1(29).Text) = "" Then
        vErrString = vErrString & "* Limited KMs " & vbCrLf
      Else
        Deal.ServiceKMs(r) = Text1(29).Text
      End If
    End If
    
    If DateChange(r) = True Then
      Deal.ServiceDate(r) = GetDateVal(DTPicker3(r - 1))
    Else
      vErrString = vErrString & "* Select Date" & vbCrLf
    End If
    
  End If
Next r

End Function

Public Function DeleteDeal(vID As Integer, vErrString As String)

  Dim sql2 As New ADODB.Recordset
  Dim sql3 As New ADODB.Recordset

  Screen.MousePointer = vbHourglass
  vErrString = ""
  'first delete actual deal
  sql = "SELECT * From Deals where ID = " & vID
  With rs
    .Open sql, cn, adOpenKeyset, adLockOptimistic
    If .EOF Then
      vErrString = vErrString & " * Deal ID not found."
    Else
      .Delete
      .Update
      '...delete docstore deal pics
      sql3.Open "delete from DocStore where left(DocID," & Len(Trim(Str(vID))) + 1 & ") = 'D" & Trim(Str(vID)) & "'", cn, adOpenKeyset, adLockOptimistic
      With sql2
        .Open "SELECT * From VehExpenses where DealID = " & vID, cn, adOpenKeyset, adLockOptimistic
        If .RecordCount <> 0 Then
          Do Until sql2.RecordCount = 0
            If sql2.BOF = False Then sql2.MoveFirst
            '...delete expense attachement from docstore
            sql3.Open "delete from DocStore where left(DocID," & Len(Trim(Str(sql2!ID))) + 2 & ") = 'EA" & Trim(Str(sql2!ID)) & "'", cn, adOpenKeyset, adLockOptimistic
            .Delete
            .Update
          Loop
        End If
        .Close
      End With
    End If
    .Close
  End With
  
  Set sql2 = Nothing
  Set sql3 = Nothing
  
  Screen.MousePointer = vbNormal
  
End Function


Public Sub DoTabOrder(vIndex As Integer)
Dim ctl As Control
Dim Cnt As Integer
Dim Ind As Integer

'  List1.AddItem "Sellers Details", 0
'  List1.AddItem "Buyers Details", 1
'  List1.AddItem "Vehicle Details", 2
'  List1.AddItem "Service Details", 3
'  List1.AddItem "   ---------   ", 4
'  List1.AddItem "Add Expenses", 5
'  List1.AddItem "View Expenses", 6
'  List1.AddItem "View Documents", 7
 
 Ind = 1
 
 '...switch off all tabstops
  For Each ctl In AddClient.Controls
    On Error Resume Next
    If ctl.TabStop Then
      ctl.TabStop = False
    End If
    On Error GoTo 0
  Next

Select Case vIndex
  Case 0
    For Cnt = 1 To 9
      Text1(Cnt - 1).TabStop = True
      Text1(Cnt - 1).TabIndex = Ind
      Ind = Ind + 1
    Next Cnt
  Case 1
    For Cnt = 25 To 21 Step -1
      Text1(Cnt).TabStop = True
      Text1(Cnt).TabIndex = Ind
      Ind = Ind + 1
    Next Cnt
    Text1(19).TabStop = True
    Text1(19).TabIndex = Ind
    Ind = Ind + 1
    For Cnt = 11 To 9 Step -1
      Text1(Cnt).TabStop = True
      Text1(Cnt).TabIndex = Ind
      Ind = Ind + 1
    Next Cnt
  Case 2
    cmdLinkCarID.TabStop = True
    cmdLinkCarID.TabIndex = Ind
    Ind = Ind + 1
    For Cnt = 13 To 17
      Text1(Cnt).TabStop = True
      Text1(Cnt).TabIndex = Ind
      Ind = Ind + 1
    Next Cnt
  Case 3

    Text1(26).TabStop = True
    Text1(26).TabIndex = Ind
    Ind = Ind + 1
    DTPicker3(0).TabStop = True
    DTPicker3(0).TabIndex = Ind
    Ind = Ind + 1
    Text1(28).TabStop = True
    Text1(28).TabIndex = Ind
    Ind = Ind + 1
    DTPicker3(1).TabStop = True
    DTPicker3(1).TabIndex = Ind
    Ind = Ind + 1
    Text1(29).TabStop = True
    Text1(29).TabIndex = Ind
    Ind = Ind + 1
    DTPicker3(2).TabStop = True
    DTPicker3(2).TabIndex = Ind
  
  Case 4
  Case 5
    DTPicker1.TabStop = True
    DTPicker1.TabIndex = Ind
    Ind = Ind + 1
    Combo2.TabStop = True
    Combo2.TabIndex = Ind
    Ind = Ind + 1
    Text1(20).TabStop = True
    Text1(20).TabIndex = Ind
    Ind = Ind + 1
    Text1(18).TabStop = True
    Text1(18).TabIndex = Ind
    Ind = Ind + 1
    cmdSaveExpense.TabStop = True
    cmdSaveExpense.TabIndex = Ind
  Case 6
  Case 7
  Case Else

  
End Select

    cmdClose.TabStop = True
    cmdClose.TabIndex = Ind + 1
    cmdSave.TabStop = True
    cmdSave.TabIndex = Ind + 2

End Sub
