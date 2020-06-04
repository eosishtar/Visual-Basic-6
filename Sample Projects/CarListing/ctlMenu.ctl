VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlMenu 
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   9300
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":2CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":59C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":864C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":B340
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":E0FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":10D31
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":13908
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":16862
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":1978D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":1C43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":1F2F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":220AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":25245
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":27E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":2ABBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":2DA8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":30B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":33AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":36A67
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlMenu.ctx":39927
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1450
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6700
      Begin VB.CommandButton cmdSearch 
         Height          =   400
         Left            =   3720
         Picture         =   "ctlMenu.ctx":3CA84
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Search"
         Top             =   922
         Width           =   400
      End
      Begin VB.CommandButton cmdEmail 
         Height          =   400
         Left            =   2805
         Picture         =   "ctlMenu.ctx":3F92B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Email "
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox Text1 
         Height          =   325
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   400
         Left            =   3795
         Picture         =   "ctlMenu.ctx":42A2E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdFind 
         Height          =   400
         Left            =   1470
         Picture         =   "ctlMenu.ctx":45B7B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Find"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdView 
         Height          =   400
         Left            =   2325
         Picture         =   "ctlMenu.ctx":4881A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "View"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdEdit 
         Height          =   400
         Left            =   615
         Picture         =   "ctlMenu.ctx":4B735
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Notes"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton Command6 
         Height          =   400
         Left            =   4200
         Picture         =   "ctlMenu.ctx":4E35B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Reset"
         Top             =   922
         Width           =   400
      End
      Begin VB.CommandButton cmdExport 
         Height          =   400
         Left            =   3300
         Picture         =   "ctlMenu.ctx":514E4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Export to Excel"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdDelete 
         Height          =   400
         Left            =   5130
         Picture         =   "ctlMenu.ctx":5442E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdOK 
         Height          =   400
         Left            =   6120
         Picture         =   "ctlMenu.ctx":57112
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "OK"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Height          =   400
         Left            =   5625
         Picture         =   "ctlMenu.ctx":59D85
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Close"
         Top             =   240
         Width           =   400
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   400
         Left            =   120
         Picture         =   "ctlMenu.ctx":5CA60
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add New"
         Top             =   240
         Width           =   400
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000006&
         X1              =   2090
         X2              =   2090
         Y1              =   305
         Y2              =   560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000006&
         X1              =   1235
         X2              =   1235
         Y1              =   305
         Y2              =   560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000006&
         X1              =   4680
         X2              =   4680
         Y1              =   310
         Y2              =   565
      End
   End
End
Attribute VB_Name = "ctlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim vAdd_Enable As Boolean
Dim vEdit_Enable As Boolean
Dim vFind_Enable As Boolean
Dim vView_Enable As Boolean
Dim vEmail_Enable As Boolean
Dim vExport_Enable As Boolean
Dim vPrint_Enable As Boolean
Dim vDelete_Enable As Boolean
Dim vClose_Enable As Boolean
Dim vOK_Enable As Boolean
Dim DIRTY_MODE As Boolean
Dim res As Integer

Public Event AddNew()
Public Event Edit()
Public Event Search(vSearchText As String)
Public Event View()
Public Event Email()
Public Event Export()
Public Event PrintPreview()
Public Event Delete()
Public Event CloseWindow()
Public Event Save()
Public Event Resized()
Public Event Reset()

Public Sub DIRTY(vValue As Boolean)
  DIRTY_MODE = vValue
End Sub

Private Sub cmdAdd_Click()
  Call NEW_MODE(True)
  RaiseEvent AddNew
End Sub

Private Sub cmdClose_Click()
  
  If DIRTY_MODE Then
    res = MsgBox("Change will not be saved. Would you like to continue?", vbInformation + vbYesNo, Screen.ActiveForm.Caption)
    If res <> 6 Then
      ' changes not to be saved
      Exit Sub
    End If
  End If

  EDIT_MODE False
  NEW_MODE False
  DIRTY_MODE = False
  RaiseEvent CloseWindow

End Sub

Private Sub cmdDelete_Click()
  res = MsgBox("Would you like to delete this?", vbInformation + vbYesNo, Screen.ActiveForm.Caption)
  If res <> 6 Then
    ' changes not to be saved
    Exit Sub
  End If

  EDIT_MODE False
  NEW_MODE False
  DIRTY_MODE = False
  
  RaiseEvent Delete
End Sub

Private Sub cmdEdit_Click()
  'Call EDIT_MODE(True)
  RaiseEvent Edit
End Sub

Private Sub cmdEmail_Click()
  RaiseEvent Email
End Sub

Private Sub cmdExport_Click()
  RaiseEvent Export
End Sub

Private Sub cmdOK_Click()

  'EDIT_MODE False
  'NEW_MODE False
  DIRTY_MODE = False
  
  RaiseEvent Save

End Sub

Private Sub cmdPrint_Click()
  RaiseEvent PrintPreview
End Sub

Private Sub cmdSearch_Click()
  If Trim(Text1.Text) <> "" Then
    RaiseEvent Search(Text1.Text)
  End If
End Sub

Private Sub cmdView_Click()
  RaiseEvent View
End Sub

Private Sub Command6_Click()
  '...clear search bar
  Text1.Text = ""
  RaiseEvent Reset
End Sub

Private Sub cmdFind_Click()
  '...opem search bar
  If Frame1.Height = 800 Then
    Frame1.Height = 1450
    Text1.Text = ""
    Text1.SetFocus
  Else
    Frame1.Height = 800
  End If
  If UserControl.Height <> Frame1.Height Then
    UserControl.Height = Frame1.Height
    RaiseEvent Resized
  End If

End Sub


Private Sub Text1_GotFocus()
  cmdSearch.Default = True
End Sub

Private Sub Text1_LostFocus()
  cmdSearch.Default = False
End Sub

Private Sub UserControl_Initialize()
  Frame1.Height = 800
  LoadCmdImages
End Sub

Private Sub UserControl_Terminate()
  DIRTY_MODE = False
End Sub

' --- load property bags
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("vAdd_Enable", vAdd_Enable)
  Call PropBag.WriteProperty("vEdit_Enable", vEdit_Enable)
  Call PropBag.WriteProperty("vFind_Enable", vFind_Enable)
  Call PropBag.WriteProperty("vView_Enable", vView_Enable)
  Call PropBag.WriteProperty("vEmail_Enable", vEmail_Enable)
  Call PropBag.WriteProperty("vExport_Enable", vExport_Enable)
  Call PropBag.WriteProperty("vPrint_Enable", vPrint_Enable)
  Call PropBag.WriteProperty("vDelete_Enable", vDelete_Enable)
  Call PropBag.WriteProperty("vClose_Enable", vClose_Enable)
  Call PropBag.WriteProperty("vOK_Enable", vOK_Enable)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vAdd_Enable = PropBag.ReadProperty("vAdd_Enable", True)
  vEdit_Enable = PropBag.ReadProperty("vEdit_Enable", True)
  vFind_Enable = PropBag.ReadProperty("vFind_Enable", True)
  vView_Enable = PropBag.ReadProperty("vView_Enable", True)
  vEmail_Enable = PropBag.ReadProperty("vEmail_Enable", True)
  vExport_Enable = PropBag.ReadProperty("vExport_Enable", True)
  vPrint_Enable = PropBag.ReadProperty("vPrint_Enable", True)
  vDelete_Enable = PropBag.ReadProperty("vDelete_Enable", True)
  vClose_Enable = PropBag.ReadProperty("vClose_Enable", True)
  vOK_Enable = PropBag.ReadProperty("vOK_Enable", True)
  
  '...set values from propery bag.
  cmdAdd.Enabled = vAdd_Enable
  cmdEdit.Enabled = vEdit_Enable
  cmdFind.Enabled = vFind_Enable
  cmdView.Enabled = vView_Enable
  cmdEmail.Enabled = vEmail_Enable
  cmdExport.Enabled = vExport_Enable
  cmdPrint.Enabled = vPrint_Enable
  cmdDelete.Enabled = vDelete_Enable
  cmdClose.Enabled = vClose_Enable
  cmdOK.Enabled = vOK_Enable
  
End Sub

Public Property Get Cmd_Add() As Boolean
  Cmd_Add = vAdd_Enable
End Property

Public Property Let Cmd_Add(ByVal vNewValue As Boolean)
  vAdd_Enable = vNewValue
  cmdAdd.Enabled = vAdd_Enable
End Property

Public Property Get Cmd_Edit() As Boolean
  Cmd_Edit = vEdit_Enable
End Property

Public Property Let Cmd_Edit(ByVal vNewValue As Boolean)
  vEdit_Enable = vNewValue
  cmdEdit.Enabled = vEdit_Enable
End Property

Public Property Get Cmd_Find() As Boolean
  Cmd_Find = vFind_Enable
End Property

Public Property Let Cmd_Find(ByVal vNewValue As Boolean)
  vFind_Enable = vNewValue
  cmdFind.Enabled = vFind_Enable
End Property

Public Property Get Cmd_View() As Boolean
  Cmd_View = vView_Enable
End Property

Public Property Let Cmd_View(ByVal vNewValue As Boolean)
  vView_Enable = vNewValue
  cmdView.Enabled = vView_Enable
End Property

Public Property Get Cmd_Email() As Boolean
  Cmd_Email = vEmail_Enable
End Property

Public Property Let Cmd_Email(ByVal vNewValue As Boolean)
  vEmail_Enable = vNewValue
  cmdEmail.Enabled = vEmail_Enable
End Property

Public Property Get Cmd_Export() As Boolean
  Cmd_Export = vExport_Enable
End Property

Public Property Let Cmd_Export(ByVal vNewValue As Boolean)
  vExport_Enable = vNewValue
  cmdExport.Enabled = vExport_Enable
End Property

Public Property Get Cmd_Print() As Boolean
  Cmd_Print = vPrint_Enable
End Property

Public Property Let Cmd_Print(ByVal vNewValue As Boolean)
  vPrint_Enable = vNewValue
  cmdPrint.Enabled = vPrint_Enable
End Property

Public Property Get Cmd_Close() As Boolean
  Cmd_Close = vClose_Enable
End Property

Public Property Let Cmd_Close(ByVal vNewValue As Boolean)
  vClose_Enable = vNewValue
  cmdClose.Enabled = vClose_Enable
End Property

Public Property Get Cmd_Delete() As Boolean
  Cmd_Delete = vDelete_Enable
End Property

Public Property Let Cmd_Delete(ByVal vNewValue As Boolean)
  vDelete_Enable = vNewValue
  cmdDelete.Enabled = vDelete_Enable
End Property

Public Property Get Cmd_OK() As Boolean
  Cmd_OK = vOK_Enable
End Property

Public Property Let Cmd_OK(ByVal vNewValue As Boolean)
  vOK_Enable = vNewValue
  cmdOK.Enabled = vOK_Enable
End Property

Public Sub NEW_MODE(vValue As Boolean)
    
  cmdEdit.Enabled = Not vValue
  'cmdFind.Enabled = Not vValue
  cmdView.Enabled = Not vValue
  cmdEmail.Enabled = Not vValue
  cmdExport.Enabled = Not vValue
  cmdPrint.Enabled = Not vValue
  cmdDelete.Enabled = Not vValue

End Sub

Public Sub EDIT_MODE(vValue As Boolean)

  cmdAdd.Enabled = Not vValue
  'cmdFind.Enabled = Not vValue
  cmdView.Enabled = Not vValue
  cmdEmail.Enabled = Not vValue
  cmdExport.Enabled = Not vValue
  cmdPrint.Enabled = Not vValue
  cmdDelete.Enabled = Not vValue
  
End Sub

Private Sub LoadCmdImages()

  cmdAdd.Picture = ImageList1.ListImages(1).Picture
  cmdEdit.Picture = ImageList1.ListImages(6).Picture
  cmdFind.Picture = ImageList1.ListImages(10).Picture
  cmdView.Picture = ImageList1.ListImages(9).Picture
  cmdEmail.Picture = ImageList1.ListImages(17).Picture
  cmdExport.Picture = ImageList1.ListImages(8).Picture
  cmdPrint.Picture = ImageList1.ListImages(21).Picture
  cmdDelete.Picture = ImageList1.ListImages(4).Picture
  cmdClose.Picture = ImageList1.ListImages(2).Picture
  cmdOK.Picture = ImageList1.ListImages(3).Picture
  cmdSearch.Picture = ImageList1.ListImages(11).Picture
  Command6.Picture = ImageList1.ListImages(13).Picture

End Sub









