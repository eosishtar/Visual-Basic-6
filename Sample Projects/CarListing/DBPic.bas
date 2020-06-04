Attribute VB_Name = "modDBPic"
Option Explicit
Global DocStoreRS As ADODB.Recordset
Global DocError As String
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Const MAX_PATH As Integer = 260
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Public Function FileFromField(SaveAsPath As String, picField As String, picExtField As String) As Boolean

  Dim fLen As Long, abBytes() As Byte, iFileNum As Integer, lFileLength As Long, fExt As String

  FileFromField = False
  DocError = ""
  On Error Resume Next
  fLen = 0
  fLen = LenB(DocStoreRS.Fields(picField))
  fExt = DocStoreRS.Fields(picExtField)
  If UCase(Right(SaveAsPath, Len(SaveAsPath) - InStrRev(SaveAsPath, "."))) <> UCase(fExt) Then SaveAsPath = Left(SaveAsPath, InStrRev(SaveAsPath, ".")) & fExt
  On Error GoTo 0
  If fLen > 0 Then
    abBytes = DocStoreRS.Fields(picField).GetChunk(fLen)
    iFileNum = FreeFile
    Open SaveAsPath For Binary As #iFileNum
    lFileLength = 0
    lFileLength = UBound(abBytes)
    Put #iFileNum, , abBytes()
    Close #iFileNum
  End If
  If Err = 0 Then FileFromField = True Else DocError = Err.Description
  On Error GoTo 0

End Function

Public Function FileToField(LoadPath As String, picField As String, picExtField As String) As Boolean

  Dim iFileNum As Integer, lFileLength As Long, abBytes() As Byte
  
  FileToField = False
  DocError = ""
  On Error Resume Next
  iFileNum = FreeFile
  Open LoadPath For Binary Access Read As #iFileNum
  lFileLength = LOF(iFileNum)
  ReDim abBytes(lFileLength)
  Get #iFileNum, , abBytes()
  Close #iFileNum
  DocStoreRS.Fields(picField).AppendChunk abBytes()
  DocStoreRS.Fields(picExtField).Value = Right(LoadPath, Len(LoadPath) - InStrRev(LoadPath, "."))
  If Err = 0 Then FileToField = True Else DocError = Err.Description
  On Error GoTo 0

End Function

Public Function OpenDocStore(DocSubsetID As String, Optional DocFriendlyName As String) As Integer

  OpenDocStore = -1
  Set DocStoreRS = New ADODB.Recordset
  If DocFriendlyName = "" Then
    DocStoreRS.Open "select * from DocStore where Left(DocID," & Len(DocSubsetID) & ") = '" & DocSubsetID & "'", cn, adOpenKeyset, adLockOptimistic
  Else
    DocStoreRS.Open "select * from DocStore where Left(DocID," & Len(DocSubsetID) & ") = '" & DocSubsetID & "' and DocName = '" & DocFriendlyName & "'", cn, adOpenKeyset, adLockOptimistic
  End If
  OpenDocStore = DocStoreRS.RecordCount

End Function

Public Sub CloseDocStore()

  DocStoreRS.Close
  Set DocStoreRS = Nothing

End Sub

Public Function GetDocumentsFolder(ReturnForm As Form) As String
    
  Dim sPath As String
  Dim IDL As ITEMIDLIST
  
  GetDocumentsFolder = ""
  If SHGetSpecialFolderLocation(ReturnForm.hwnd, 5, IDL) = 0 Then
    sPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
        GetDocumentsFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
    End If
  End If
  If Right(GetDocumentsFolder, 1) <> "\" Then GetDocumentsFolder = GetDocumentsFolder & "\"

End Function
