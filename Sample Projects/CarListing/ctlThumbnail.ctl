VERSION 5.00
Begin VB.UserControl ctlThumbnail 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label lblBlank 
      AutoSize        =   -1  'True
      Caption         =   "[ Drag Image Here ]"
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1395
   End
   Begin VB.Image Img 
      Height          =   1335
      Left            =   1080
      OLEDropMode     =   1  'Manual
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "ctlThumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum MyBorders
  bNone = 0
  bFixed = 1
End Enum
Dim vPicturePath As String
Dim vBorderStyle As MyBorders
Dim PicW As Long
Dim PicH As Long
Dim PicSet As Boolean
Dim vError As String
Dim picWidth As Long
Dim picHeight As Long
Public Enum FitTypeEnum
  fitAll = 0
  fitNoSpace = 1
End Enum
Dim vFitType As FitTypeEnum
Public Event DblClick()
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event NewDropImage(NewPath As String)

Public Property Get FitType() As FitTypeEnum

  FitType = vFitType

End Property

Public Property Let FitType(ByVal vNewValue As FitTypeEnum)

  vFitType = vNewValue
  SizePic

End Property

Private Sub Img_DblClick()

  RaiseEvent DblClick

End Sub

Private Sub Img_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

  RaiseEvent MouseUp(Button, Shift, x, Y)
  
End Sub

Private Sub Img_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)

  RaiseEvent NewDropImage(Data.Files(1))

End Sub

Private Sub UserControl_DblClick()

  RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

  RaiseEvent MouseUp(Button, Shift, x, Y)

End Sub

Private Sub UserControl_Resize()

  SizePic
  lblBlank.Move (UserControl.Width - lblBlank.Width) / 2, (UserControl.Height - lblBlank.Height) / 2
  
End Sub

Public Function Error() As String

  Error = vError

End Function

Public Function SetPicture(PicContent As IPicture) As Boolean

  UserControl.Refresh
  SetPicture = False
  vError = ""
  Img.Picture = LoadPicture()
  lblBlank.Visible = True
  Img.Move 0, 0, UserControl.Width, UserControl.Height
  vPicturePath = ""
  PicSet = False
  Img.Stretch = False
  On Error Resume Next
  Set Img.Picture = PicContent
  If Err <> 0 Then
    vError = Err.Description
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  picWidth = Img.Width
  picHeight = Img.Height
  Img.Stretch = True
  PicSet = True
  SizePic
  SetPicture = True
  lblBlank.Visible = False

End Function

Private Sub SizePic()

  If PicSet = False Then Exit Sub
  
  If vFitType = fitAll Then
    PicW = UserControl.ScaleWidth
    PicH = (UserControl.ScaleWidth / picWidth) * picHeight
    If PicH > UserControl.ScaleHeight Then
      PicH = UserControl.ScaleHeight
      PicW = (UserControl.ScaleHeight / picHeight) * picWidth
    End If
  Else
    PicW = UserControl.ScaleWidth
    PicH = (UserControl.ScaleWidth / picWidth) * picHeight
    If PicH < UserControl.ScaleHeight Then
      PicH = UserControl.ScaleHeight
      PicW = (UserControl.ScaleHeight / picHeight) * picWidth
    End If
  End If
  Img.Move (UserControl.ScaleWidth - PicW) / 2, (UserControl.ScaleHeight - PicH) / 2, PicW, PicH

End Sub

Public Property Get BorderStyle() As MyBorders

  BorderStyle = vBorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewValue As MyBorders)

  vBorderStyle = vNewValue
  UserControl.BorderStyle = vNewValue
  SizePic

End Property

Public Sub ClearPicture()

  Img.Picture = LoadPicture()
  lblBlank.Visible = True
  Img.Move 0, 0, UserControl.Width, UserControl.Height
  PicSet = False
  Img.Stretch = False
  vPicturePath = ""

End Sub

Public Function HasPicture() As Boolean

  HasPicture = PicSet

End Function

Public Property Get PicturePath() As String

  PicturePath = vPicturePath

End Property

Public Property Let PicturePath(ByVal vNewValue As String)

  Dim tmpPic As IPicture
  
  vPicturePath = vNewValue
  On Error Resume Next
  Set tmpPic = LoadPicture(vNewValue)
  If Err <> 0 Then MsgBox "Could not set PicturePath: " & Err.Description: On Error GoTo 0: Exit Property
  On Error GoTo 0
  If SetPicture(tmpPic) = False Then
    MsgBox "Could not set PicturePath: " & vError
  Else
    vPicturePath = vNewValue
  End If

End Property
