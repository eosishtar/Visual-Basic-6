Attribute VB_Name = "modCommonFuncs"
Option Explicit
Public ImageText As String
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim TempRS As ADODB.Recordset
Declare Function ShellExecute& Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd&, ByVal lpszOp$, ByVal lpszFile$, ByVal lpszParams$, ByVal lpszdir$, ByVal fsShowCmd&)
Dim mlHwnd   As Long
Dim XyX As Integer
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Const MAX_PATH As Integer = 260
Private Declare Function ShellExecute2 Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub EndProgram()

  Set TempRS = Nothing
  Set rs = Nothing
  Set cn = Nothing
  
  End
  
End Sub

Public Function OpenFile2(Target As String) As Boolean

  Dim lHWnd As Long
  Dim lAns As Long
  
  lAns = ShellExecute2(lHWnd, "open", Target, vbNullString, vbNullString, 1)
  OpenFile2 = (lAns > 32)

End Function

Public Function SetComboText(vCombo As ComboBox, vItem As String) As Integer
  Dim vCnter As Integer
  
  SetComboText = -1
  For vCnter = 0 To vCombo.ListCount - 1
    If UCase(Left(vCombo.List(vCnter), Len(vItem))) = UCase(vItem) Then SetComboText = vCnter: Exit For
  Next vCnter

End Function

Public Function GetUserDetails(uUserCode As String)
  
  Set TempRS = New ADODB.Recordset
    sql = "select * from Security WHERE UserCode = " & Chr(34) & uUserCode & Chr(34)
    TempRS.Open sql, cn, adOpenForwardOnly, adLockOptimistic
    
    
    
    If TempRS.EOF Then
      TempRS.Close
    Else
    
      User.ID = TempRS!ID
      User.UserCode = TempRS!UserCode
      User.Username = TempRS!Username
      User.Password = TempRS!Password
        
      TempRS.Close
    End If
            
  Set TempRS = Nothing
  
End Function

Public Function ValidEmail(ByVal strCheck As String) As Boolean
'...function to validate email adress
  
  Dim bCK As Boolean
  Dim strDomainType As String
  Dim strDomainName As String
  Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
  Dim i As Integer
  
  If Trim$(strCheck) = "0" Then
    ValidEmail = True
    Exit Function
  End If
  
  bCK = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
  If Not bCK Then GoTo ExitFunction
  
  bCK = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
  If Not bCK Then GoTo ExitFunction
  
  ' Check for invalid characters.
  If Len(strCheck) > Len(sInvalidChars) Then
      For i = 1 To Len(sInvalidChars)
          If InStr(strCheck, Mid(sInvalidChars, i, 1)) > 0 Then
              bCK = False
              GoTo ExitFunction
          End If
      Next
  Else
      For i = 1 To Len(strCheck)
          If InStr(sInvalidChars, Mid(strCheck, i, 1)) > 0 Then
              bCK = False
              GoTo ExitFunction
          End If
      Next
  End If
  
  If InStr(1, strCheck, "@") > 1 Then 'Check for an @ symbol
      bCK = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
  Else
      bCK = False
  End If
  If Not bCK Then GoTo ExitFunction
  
  strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
  bCK = Not InStr(1, strCheck, "@") > 0 'Check to see if there are too many @'s
  If Not bCK Then GoTo ExitFunction
  
  strDomainType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
  bCK = Len(strDomainType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
  If Not bCK Then GoTo ExitFunction
  
  On Error Resume Next
  strCheck = Left(strCheck, Len(strCheck) - Len(strDomainType) - 1)
  If Err.Number <> 0 Then
    bCK = False
    Err.Clear
  End If
  On Error GoTo 0
  Do Until InStr(1, strCheck, ".") <= 1
      If Len(strCheck) >= InStr(1, strCheck, ".") Then
          strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
      Else
          bCK = False
          GoTo ExitFunction
      End If
  Loop
  If strCheck = "." Or Len(strCheck) = 0 Then bCK = False
  
ExitFunction:
  ValidEmail = bCK
End Function


Public Sub GetCompanyDetails()

  Company.ID = 0
  Company.Address = ""
  Company.Email = ""
  Company.Name = ""
  Company.Person = ""
  Company.Telephone = ""
  Company.BankingDetails = ""
  Company.TermsConditions = ""

  If cn = "" Then dbase.Go      'if no database installed
  Set TempRS = New ADODB.Recordset
    sql = "select * from CompanyDetails where [ID] = 1"
      TempRS.Open sql, cn, adOpenDynamic, adLockOptimistic
        If TempRS.EOF Then
          TempRS.Close
          MsgBox "Company details not found...", vbCritical + vbOKOnly, Screen.ActiveForm.Caption
          Exit Sub
        End If

        Company.ID = 1
        If Not IsNull(TempRS!CompanyName) Then Company.Name = TempRS!CompanyName
        If Not IsNull(TempRS!Telephone) Then Company.Telephone = TempRS!Telephone
        If Not IsNull(TempRS!Person) Then Company.Person = TempRS!Person
        If Not IsNull(TempRS!Address) Then Company.Address = TempRS!Address
        If Not IsNull(TempRS!Email) Then Company.Email = TempRS!Email
        If Not IsNull(TempRS!BankingDetails) Then Company.BankingDetails = TempRS!BankingDetails
        If Not IsNull(TempRS!TermsConditions) Then Company.TermsConditions = TempRS!TermsConditions

      TempRS.Close
  Set TempRS = Nothing
    
  '...get the logo path
  Main.Logo = GetSetting(App.EXEName, "Datapath", "LogoPath")

  
End Sub

Public Function GetAppPath()
  GetAppPath = App.Path
  If Right(GetAppPath, 1) <> "\" Then GetAppPath = GetAppPath & "\"
End Function

Public Function CenterForm(vForm As Form)
  'vForm.Top = (Screen.Height - vForm.Height) / 4
  'vForm.Left = (Screen.Width - vForm.Width) / 2
  
  vForm.Top = (MDIForm1.Height - vForm.Height) / 4
  vForm.Left = (MDIForm1.Width - vForm.Width) / 2
  
End Function

Public Sub KillAllActiveWindows()
  Dim yy As Integer

  yy = 1
  Do While yy < Forms.Count
    If Forms(yy).MDIChild Then
      Unload Forms(yy)
    Else
      yy = yy + 1
    End If
  Loop

End Sub

Public Sub SetDragPicture(Pic As PictureBox)
  
  ImageText = "[ Drag image in here ]"

  '...set this to manual
  Pic.OLEDropMode = vbOLEDropManual
  Pic.Visible = True
     
  '...Clear the image in "drag" pic box
  Pic.Picture = LoadPicture()

  '...Position the label [ Drag Inside Here ]
  Pic.CurrentX = (Pic.Width - Pic.TextWidth(ImageText)) / 2
  Pic.CurrentY = (Pic.Height - Pic.TextHeight(ImageText)) / 2
  Pic.Print ImageText
      
End Sub

Public Function GetMasterVehicleDetails(ID As Integer) As Boolean

  Set TempRS = New ADODB.Recordset
  
    MasterVehicle.ID = 0
    MasterVehicle.ModelYear = 0
    MasterVehicle.VehicleMake = ""
    MasterVehicle.VehicleModel = ""
    MasterVehicle.VehicleDisplacement = 0
    MasterVehicle.VehicleType = ""
    MasterVehicle.RatedHorsePower = 0
    MasterVehicle.NrOrCylinders = 0
    MasterVehicle.EngineCode = ""
    MasterVehicle.TransmissionCode = ""
    MasterVehicle.TransmissionDesc = ""
    MasterVehicle.NrOfGears = 0
    MasterVehicle.DriveSystemCode = ""
    MasterVehicle.DriveSystemDescription = ""
    MasterVehicle.VehicleImage = ""
    MasterVehicle.Notes = ""
    MasterVehicle.TurboType = ""
    MasterVehicle.FuelType = ""
    MasterVehicle.BookDate = 0
    MasterVehicle.BookValue = 0
    

    sql = "Select * from MasterListing where ID = " & ID
    TempRS.Open sql, cn, adOpenKeyset, adLockOptimistic
    
    If TempRS.EOF Then
      TempRS.Close
      Exit Function
    End If
    
    MasterVehicle.ID = ID
    If Not IsNull(TempRS!ModelYear) Then MasterVehicle.ModelYear = TempRS!ModelYear
    If Not IsNull(TempRS!VehicleMake) Then MasterVehicle.VehicleMake = TempRS!VehicleMake
    If Not IsNull(TempRS!VehicleModel) Then MasterVehicle.VehicleModel = TempRS!VehicleModel
    If Not IsNull(TempRS!VehicleDisplacement) Then MasterVehicle.VehicleDisplacement = TempRS!VehicleDisplacement
    If Not IsNull(TempRS!VehicleType) Then MasterVehicle.VehicleType = TempRS!VehicleType
    If Not IsNull(TempRS!RatedHorsePower) Then MasterVehicle.RatedHorsePower = TempRS!RatedHorsePower
    If Not IsNull(TempRS!NrOfCylinders) Then MasterVehicle.NrOrCylinders = TempRS!NrOfCylinders
    If Not IsNull(TempRS!EngineCode) Then MasterVehicle.EngineCode = TempRS!EngineCode
    If Not IsNull(TempRS!TransmissionTypeCode) Then MasterVehicle.TransmissionCode = TempRS!TransmissionTypeCode
    If Not IsNull(TempRS!TransmissionType) Then MasterVehicle.TransmissionDesc = TempRS!TransmissionType
    If Not IsNull(TempRS!NrOfGears) Then MasterVehicle.NrOfGears = TempRS!NrOfGears
    If Not IsNull(TempRS!DriveSystemCode) Then MasterVehicle.DriveSystemCode = TempRS!DriveSystemCode
    If Not IsNull(TempRS!DriveSystemDescription) Then MasterVehicle.DriveSystemDescription = TempRS!DriveSystemDescription
    If Not IsNull(TempRS!VehicleImage) Then MasterVehicle.VehicleImage = TempRS!VehicleImage
    If Not IsNull(TempRS!Notes) Then MasterVehicle.Notes = TempRS!Notes
    If Not IsNull(TempRS!FuelType) Then MasterVehicle.FuelType = TempRS!FuelType
    If Not IsNull(TempRS!TurboType) Then MasterVehicle.TurboType = TempRS!TurboType
    If Not IsNull(TempRS!BookValue) Then MasterVehicle.BookValue = TempRS!BookValue
    If Not IsNull(TempRS!BookDate) Then MasterVehicle.BookDate = TempRS!BookDate
  
  TempRS.Close
  
  Set TempRS = Nothing
  
End Function

Public Function GetTransmission(vValue As String) As String
  Select Case Trim$(vValue)
    Case "A": GetTransmission = "Automatic"
    Case "M": GetTransmission = "Manual"
    Case "AM": GetTransmission = "Automated Manual"
    Case "SCV": GetTransmission = "Selectable Continuously Variable"
    Case "SA": GetTransmission = "Semi Automatic"
    Case "CVT": GetTransmission = "Continuously Variable"
    Case "OT": GetTransmission = "Other"
    Case "AMS": GetTransmission = "Automated Manual": GetTransmission = "Selectable"
    Case Else: GetTransmission = ""
  End Select
End Function

Public Function GetDriveSys(vValue As String) As String
  Select Case Trim$(vValue)
    Case "R": GetDriveSys = "2 Wheel Drive, Rear"
    Case "F": GetDriveSys = "2 Wheel Drive, Front"
    Case "A": GetDriveSys = "All Wheel Drive"
    Case "4": GetDriveSys = "4 Wheel Drive"
    Case "P": GetDriveSys = "Part time 4 Wheel Drive"
    Case Else: GetDriveSys = ""
  End Select
End Function

Public Function GenImageCode() As String

  Dim Col As Collection                '...The collection we will use to store the numbers
  Dim Ar(1 To 10) As Integer           '...The array to store the values from the collection
  Dim i As Integer
  Dim X As Integer
  Dim MaxColl As Integer
  Randomize                            '...Just once to ensure that we get random values
  
  Set Col = New Collection
  MaxColl = 10
  
    For i = 1 To MaxColl               '...The possible numbers that we can have as a result is all the numbers from 1 to 100 so
      Col.Add i
    Next i
    
    For i = 1 To MaxColl
      X = RandomInteger(1, Col.Count)   '...Get a random item from the collection (that exists for sure)
      Ar(i) = Col.Item(X)               '...Add it to the array
      Col.Remove X                      '...Remove it so we don't add it again
    Next i
   
    For i = 1 To UBound(Ar)
      GenImageCode = GenImageCode & CStr(Ar(i))
    Next i
    
    Set Col = Nothing

End Function

'...The random number generator code
Private Function RandomInteger(Lowerbound As Integer, Upperbound As Integer) As Integer
  RandomInteger = Int((Upperbound - Lowerbound + 1) * Rnd + Lowerbound)
End Function

Function CleanSQL(OldString As String)
   Dim RegEx As New RegExp
   RegEx.Global = True
   RegEx.IgnoreCase = True
   RegEx.Pattern = "[(?*"",\\<>&#~%{}+_.@:\/!;]+"
   
   CleanSQL = RegEx.Replace(OldString, "")
End Function

Public Function ChkDuplicate(uItem As String, uTable As String) As Boolean
  Dim chkItem As String
  Dim setField As String
  ChkDuplicate = False
  
  If uTable = "Security" Then setField = "Username"
  
  chkItem = Trim(CleanSQL(uItem))
  
  Set TempRS = New ADODB.Recordset
    sql = "select * from " & uTable & " where " & setField & " = " & Chr(34) & chkItem & Chr(34)
    
    On Error Resume Next
    Err.Clear
    TempRS.Open sql, cn, adOpenForwardOnly, adLockOptimistic
    If Err = -2147217913 Then
      sql = "select * from " & uTable & " where " & setField & " = " & Val(chkItem)
      TempRS.Open sql, cn, adOpenForwardOnly, adLockOptimistic
    End If
    
    On Error GoTo 0
      If TempRS.EOF Then
        TempRS.Close
        ChkDuplicate = False
      Else
        ChkDuplicate = True
        TempRS.Close
      End If
      
  Set TempRS = Nothing

End Function

Public Function GetNextUserCode() As String
  Dim TempRs1 As ADODB.Recordset
  Set TempRs1 = New ADODB.Recordset
  
    GetNextUserCode = ""
    sql = "SELECT [UserCode] from Security ORDER BY [UserCode]"

    TempRs1.Open sql, cn, adOpenDynamic, adLockOptimistic
    If TempRs1.RecordCount = 0 Then
      GetNextUserCode = MakeFour(1)
    Else
      TempRs1.MoveLast
      GetNextUserCode = MakeFour(Val(Right(TempRs1.Fields("UserCode").Value, 4)) + 1)
    End If

    TempRs1.Close
  Set TempRs1 = Nothing
  
End Function

Public Function MakeFour(daNum As Long) As String
  
  MakeFour = Trim(Str(daNum))
  Do Until Len(MakeFour) = 4
    MakeFour = "0" & MakeFour
  Loop

End Function

Public Function PadString(daStr As String, daLen As Integer) As String
  
  PadString = Trim(daStr)
  Do Until Len(PadString) = daLen
    PadString = PadString & " "
  Loop

End Function

Public Function Counter(cCount As Integer, fFormName As String) As String
  Dim strResult As String

  If cCount = 0 Then
    strResult = "There are no Items in the " & fFormName & " list"
  ElseIf cCount = 1 Then
    strResult = "There is " & cCount & " Item in the " & fFormName & " list"
  ElseIf cCount > 1 Then
    strResult = "There are " & cCount & " Items in the " & fFormName & " list"
  End If
  
  Counter = strResult
  
End Function

Public Function Crypt(texti, sEncrypt) As String
Dim Crypted As String
Dim t As Integer, x1 As Integer, G As Integer, TT As Integer
Dim sana As Long
  
  'On Error Resume Next
  For t = 1 To Len(sEncrypt)
         sana = Asc(Mid(sEncrypt, t, 1))
         x1 = x1 + sana
  Next
  
  x1 = Int((x1 * 0.1) / 6)
  sEncrypt = x1
  G = 0
  For TT = 1 To Len(texti)
      sana = Asc(Mid(texti, TT, 1))
      G = G + 1
      If G = 6 Then G = 0
      x1 = 0
      If G = 0 Then x1 = sana - (sEncrypt - 2)
      If G = 1 Then x1 = sana + (sEncrypt - 5)
      If G = 2 Then x1 = sana - (sEncrypt - 4)
      If G = 3 Then x1 = sana + (sEncrypt - 2)
      If G = 4 Then x1 = sana - (sEncrypt - 3)
      If G = 5 Then x1 = sana + (sEncrypt - 5)
      
      x1 = x1 + G
      Crypted = Crypted & Chr(x1)
  Next
  
  Crypt = Crypted

End Function

Public Function Decrypt(texti, sEncrypt) As String
  Dim DeCrypted As String
  Dim t As Integer, x1 As Integer, G As Integer, TT As Integer
  Dim sana As Long

  'On Error Resume Next
  For t = 1 To Len(sEncrypt)
         sana = Asc(Mid(sEncrypt, t, 1))
         x1 = x1 + sana
  Next
  
  x1 = Int((x1 * 0.1) / 6)
  sEncrypt = x1
  G = 0
  
  For TT = 1 To Len(texti)
      sana = Asc(Mid(texti, TT, 1))
      G = G + 1
      If G = 6 Then G = 0
      x1 = 0
      If G = 0 Then x1 = sana + (sEncrypt - 2)
      If G = 1 Then x1 = sana - (sEncrypt - 5)
      If G = 2 Then x1 = sana + (sEncrypt - 4)
      If G = 3 Then x1 = sana - (sEncrypt - 2)
      If G = 4 Then x1 = sana + (sEncrypt - 3)
      If G = 5 Then x1 = sana - (sEncrypt - 5)
      x1 = x1 - G
      DeCrypted = DeCrypted & Chr(x1)
  Next
  
  Decrypt = DeCrypted
  
End Function


Public Function StripOut(From As String, What As String) As String
 
   Dim i As Integer

   StripOut = From
   For i = 1 To Len(What)
       StripOut = Replace(StripOut, Mid$(What, i, 1), "")
   Next i
 
End Function

Public Sub AltLVBackground(lv As ListView, frm As Form)

'---------------------------------------------------------------------------------
' Purpose   : Alternates row colors in a ListView control
' Method    : Creates a picture box and draws the desired color scheme in it, then
'             loads the drawn image as the listviews picture.
'---------------------------------------------------------------------------------
Dim lH      As Long
Dim lSM     As Byte
Dim picAlt  As PictureBox
Dim BackColorOne As OLE_COLOR
Dim BackColorTwo As OLE_COLOR

BackColorOne = &HE0E0E0     'grey
'BackColorTwo = &H80FFFF     'yellow
'BackColorTwo = 16777088     'light blue
BackColorTwo = 12632319       'light red

    With lv
        If .View = lvwReport And .ListItems.Count Then
            Set picAlt = frm.Controls.Add("VB.PictureBox", "picAlt")
            lSM = .Parent.ScaleMode
            .Parent.ScaleMode = vbTwips
            .PictureAlignment = lvwTile
            lH = .ListItems(1).Height
            With picAlt
                .BackColor = BackColorOne
                .AutoRedraw = True
                .Height = lH * 2
                .BorderStyle = 0
                .Width = 10 * Screen.TwipsPerPixelX
                picAlt.Line (0, lH)-(.ScaleWidth, lH * 2), BackColorTwo, BF
                Set lv.Picture = .Image
            End With
            Set picAlt = Nothing
            frm.Controls.Remove "picAlt"
            lv.Parent.ScaleMode = lSM
        End If
    End With
End Sub

Public Sub AutosizeColumns(ByVal TargetListView As ListView)

  Const SET_COLUMN_WIDTH As Long = 4126
  Const AUTOSIZE_USEHEADER As Long = -2
  Dim lngColumn As Long

  For lngColumn = 0 To (TargetListView.ColumnHeaders.Count - 1)
    Call SendMessage(TargetListView.hwnd, SET_COLUMN_WIDTH, lngColumn, ByVal AUTOSIZE_USEHEADER)
  Next lngColumn
  
End Sub

Public Function TodaysDate() As String
  Dim TODAY As Single
  
  TODAY = DateValue(Now())
  TodaysDate = TODAY
  
End Function


Public Sub GetDeal(ID As Integer)
  Dim X As Integer
  
  'reset the class
  Deal.BuyerFirstName = ""
  Deal.BuyerLastName = ""
  Deal.BuyerIDNumber = ""
  Deal.BuyerCompanyName = ""
  Deal.BuyerCompanyRegNr = ""
  Deal.BuyerContactNr = ""
  Deal.BuyerAltContactNr = ""
  Deal.BuyerEmailAddress = ""
  Deal.BuyerNotes = ""
  Deal.SellerFirstName = ""
  Deal.SellerLastName = ""
  Deal.SellerIDNumber = ""
  Deal.SellerCompanyName = ""
  Deal.SellerCompanyRegNr = ""
  Deal.SellerContactNr = ""
  Deal.SellerAltContactNr = ""
  Deal.SellerEmailAddress = ""
  Deal.SellerNotes = ""
  
  Deal.VehicleID = 0
  Deal.VehicleRegNr = ""
  Deal.VehicleVINNr = ""
  Deal.VehicleEngineNr = ""
  Deal.VehicleKM = ""
  Deal.VehicleNotes = ""
  Deal.VehicleImage = ""
  Deal.VehicleCost = 0
  Deal.VehicleService = 0
  Deal.VehicleDateBought = ""
  Deal.VehicleDateSold = ""
  Deal.VehicleSold = 0
  
  For X = 1 To 3
    Deal.ServicePlanEnabled(X) = False
    Deal.ServiceKMs(X) = 0
    Deal.ServiceDate(X) = 0
  Next X
  
  Deal.ServiceNotes = ""
  Deal.DateModified = ""
  Deal.UserModified = ""
  Deal.DateCreated = ""
  Deal.UserCreated = ""
  Deal.DealClosed = False

  Set TempRS = New ADODB.Recordset
    sql = "select * from Deals where [ID] = " & ID
      TempRS.Open sql, cn, adOpenDynamic, adLockOptimistic
        If TempRS.EOF Then
          TempRS.Close
          MsgBox "Deal ID not found!", vbCritical + vbOKOnly, Screen.ActiveForm.Caption
          Exit Sub
        End If

        Deal.ID = ID
        If Not IsNull(TempRS!BuyerFirstName) Then Deal.BuyerFirstName = TempRS!BuyerFirstName
        If Not IsNull(TempRS!BuyerLastName) Then Deal.BuyerLastName = TempRS!BuyerLastName
        If Not IsNull(TempRS!BuyerIDNumber) Then Deal.BuyerIDNumber = TempRS!BuyerIDNumber
        If Not IsNull(TempRS!BuyerCompanyName) Then Deal.BuyerCompanyName = TempRS!BuyerCompanyName
        If Not IsNull(TempRS!BuyerCompanyRegNr) Then Deal.BuyerCompanyRegNr = TempRS!BuyerCompanyRegNr
        If Not IsNull(TempRS!BuyerContactNr) Then Deal.BuyerContactNr = TempRS!BuyerContactNr
        If Not IsNull(TempRS!BuyerAltContactNr) Then Deal.BuyerAltContactNr = TempRS!BuyerAltContactNr
        If Not IsNull(TempRS!BuyerEmailAddress) Then Deal.BuyerEmailAddress = TempRS!BuyerEmailAddress
        If Not IsNull(TempRS!BuyerNotes) Then Deal.BuyerNotes = TempRS!BuyerNotes
        
        If Not IsNull(TempRS!SellerFirstName) Then Deal.SellerFirstName = TempRS!SellerFirstName
        If Not IsNull(TempRS!SellerLastName) Then Deal.SellerLastName = TempRS!SellerLastName
        If Not IsNull(TempRS!SellerIDNumber) Then Deal.SellerIDNumber = TempRS!SellerIDNumber
        If Not IsNull(TempRS!SellerCompanyName) Then Deal.SellerCompanyName = TempRS!SellerCompanyName
        If Not IsNull(TempRS!SellerCompanyRegNr) Then Deal.SellerCompanyRegNr = TempRS!SellerCompanyRegNr
        If Not IsNull(TempRS!SellerContactNr) Then Deal.SellerContactNr = TempRS!SellerContactNr
        If Not IsNull(TempRS!SellerAltContactNr) Then Deal.SellerAltContactNr = TempRS!SellerAltContactNr
        If Not IsNull(TempRS!SellerEmailAddress) Then Deal.SellerEmailAddress = TempRS!SellerEmailAddress
        If Not IsNull(TempRS!SellerNotes) Then Deal.SellerNotes = TempRS!SellerNotes
        
        If Not IsNull(TempRS!VehicleID) Then Deal.VehicleID = TempRS!VehicleID
        If Not IsNull(TempRS!VehicleRegNr) Then Deal.VehicleRegNr = TempRS!VehicleRegNr
        If Not IsNull(TempRS!VehicleVINNr) Then Deal.VehicleVINNr = TempRS!VehicleVINNr
        If Not IsNull(TempRS!VehicleEngineNr) Then Deal.VehicleEngineNr = TempRS!VehicleEngineNr
        If Not IsNull(TempRS!VehicleKM) Then Deal.VehicleKM = TempRS!VehicleKM
        If Not IsNull(TempRS!VehicleNotes) Then Deal.VehicleNotes = TempRS!VehicleNotes
        If Not IsNull(TempRS!VehicleImage) Then Deal.VehicleImage = TempRS!VehicleImage
        If Not IsNull(TempRS!VehicleCost) Then Deal.VehicleCost = TempRS!VehicleCost
        If Not IsNull(TempRS!VehicleService) Then Deal.VehicleService = TempRS!VehicleService
        If Not IsNull(TempRS!VehicleDateBought) Then Deal.VehicleDateBought = TempRS!VehicleDateBought
        If Not IsNull(TempRS!VehicleSold) Then Deal.VehicleSold = TempRS!VehicleSold
        If Not IsNull(TempRS!VehicleDateSold) Then Deal.VehicleDateSold = TempRS!VehicleDateSold
        
        If Not IsNull(TempRS!ServicePlan1) Then Deal.ServicePlanEnabled(1) = TempRS!ServicePlan1
        If Not IsNull(TempRS!ServicePlan2) Then Deal.ServicePlanEnabled(2) = TempRS!ServicePlan2
        If Not IsNull(TempRS!ServicePlan3) Then Deal.ServicePlanEnabled(3) = TempRS!ServicePlan3
        If Not IsNull(TempRS!ServiceKM1) Then Deal.ServiceKMs(1) = TempRS!ServiceKM1
        If Not IsNull(TempRS!ServiceKM2) Then Deal.ServiceKMs(2) = TempRS!ServiceKM2
        If Not IsNull(TempRS!ServiceKM3) Then Deal.ServiceKMs(3) = TempRS!ServiceKM3
        If Not IsNull(TempRS!ServiceDate1) Then Deal.ServiceDate(1) = TempRS!ServiceDate1
        If Not IsNull(TempRS!ServiceDate2) Then Deal.ServiceDate(2) = TempRS!ServiceDate2
        If Not IsNull(TempRS!ServiceDate3) Then Deal.ServiceDate(3) = TempRS!ServiceDate3
        If Not IsNull(TempRS!ServiceNotes) Then Deal.ServiceNotes = TempRS!ServiceNotes
        
        
        
        If Not IsNull(TempRS!DateModified) Then Deal.DateModified = TempRS!DateModified
        If Not IsNull(TempRS!UserModified) Then Deal.UserModified = TempRS!UserModified
        If Not IsNull(TempRS!DateCreated) Then Deal.DateCreated = TempRS!DateCreated
        If Not IsNull(TempRS!UserCreated) Then Deal.UserCreated = TempRS!UserCreated
        If Not IsNull(TempRS!DealClosed) Then Deal.DealClosed = TempRS!DealClosed

      TempRS.Close
  Set TempRS = Nothing
    
End Sub

Public Function OpenFile(vFileToOpen As String)
  Dim Cancel As Integer

  'attempt opening the file
  If vFileToOpen <> "" Then
    XyX = ShellExecute(mlHwnd, "Open", vFileToOpen, 0, 0, 1)
    If XyX <= 32 Then
      MsgBox "An error occured and the document could not be loaded.", vbCritical + vbOKOnly, "Error opening the file..."
    End If
  Else
    If Err = 32755 Then
      Cancel = 1
      Err.Clear
    End If
  End If
  
End Function

Public Sub RemoveTempFiles()

  Dim tmpPath As String

  tmpPath = App.Path
  If Right(tmpPath, 1) <> "\" Then tmpPath = tmpPath & "\"
  tmpPath = tmpPath & "tmp."
  On Error Resume Next
  Kill tmpPath & "jpg"
  Kill tmpPath & "bmp"
  Kill tmpPath & "gif"
  Kill tmpPath & "ico"
  Kill tmpPath & "jpeg"
  Kill tmpPath & "png"
  Kill tmpPath & "pdf"
  Kill tmpPath & "tif"
  Kill tmpPath & "tiff"
  Kill tmpPath & "txt"
  Kill tmpPath & "doc"
  Kill tmpPath & "docx"
  Kill tmpPath & "xls"
  Kill tmpPath & "xlsx"
  Kill tmpPath & "csv"
  On Error GoTo 0

End Sub

Public Function GetDateVal(vDPicker As DTPicker) As String

  Dim vTheDay As Single
  
  vTheDay = DateValue(vDPicker)
  GetDateVal = vTheDay
  
End Function

Public Function ValidID(vID As String) As Boolean

  Dim v As ValidatorClass
  Set v = New ValidatorClass
  Dim ErrString As String
  Dim vControl As Integer
  
  ValidID = False
  v.Validate oID_Number, vID, ErrString
  
  If ErrString = "" Then
    ValidID = True
  End If
  
  Set v = Nothing

End Function

Public Sub SetMenu(vValue As Boolean)
  
  MDIForm1.Toolbar1.Enabled = vValue
  MDIForm1.mnuFile.Enabled = vValue
  MDIForm1.mnuEdit.Enabled = vValue
  MDIForm1.mnuView.Enabled = vValue

End Sub
