VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidatorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ValidateType
  oID_Number = 1
  oUIF_NR = 2
End Enum


Public Function Validate(vOption As ValidateType, vValue As String, vErrString As String) As String
  ReDim TempNr(1) As String
  ReDim TestNr(1) As String
  Dim TotalValue As Integer
  Dim TotalValue2 As Integer
  Dim TotalValue3 As String
  Dim i As Integer
  Dim ValOption As Integer
  Dim ValText As String
  Dim CheckSum As Integer


  '...establish the validation type first
  Select Case vOption
    Case 1      '...ID Number
      ValOption = 13
      ValText = "ID Number"
    Case 2      '...UIF Number
      ValOption = 8
      ValText = "UIF Number"
  End Select
    
    ReDim TempNr(1 To ValOption)
    ReDim TestNr(1 To ValOption)
  
    '...test the length
    vErrString = ""
    vValue = Trim$(vValue)
    If Len(Trim(vValue)) = 0 Or Not IsNumeric(vValue) Then
      vErrString = "Please enter a " & ValText
      Exit Function
    Else
      If Len(Trim(vValue)) <> ValOption Then
        vErrString = ValText & " length is incorrect"
        Exit Function
      End If
    End If
  
    '...spilt the nr into the array
    For i = 1 To ValOption
      TempNr(i) = Right(Left(vValue, i), 1)
      '...store the checksum here
      If i = ValOption Then CheckSum = TempNr(i)
    Next i
  
    TotalValue = 0
    TotalValue2 = 0
    TotalValue3 = ""

  '... validate the type
  Select Case vOption

    Case 1      '...ID Number
      '...Add all the digits of the ID number in the odd positions (except CheckSUM)
      For i = 1 To ValOption Step 2
        If i <> ValOption Then
          TotalValue2 = TotalValue2 + TempNr(i)
        End If
      Next i
      
      '...Take all the even digits as one number and multiply that by 2:
      For i = 2 To ValOption Step 2
        If i <> ValOption Then
          TotalValue3 = TotalValue3 + TempNr(i)
        End If
      Next i
      TotalValue3 = Trim$(Val(TotalValue3) * 2)
      
      '...Spilt the digits and then add the digits of this number together (in b)
      ReDim TempNr(1 To Len(TotalValue3))
      For i = 1 To UBound(TempNr)
        TempNr(i) = Right(Left(TotalValue3, i), 1)
      Next i
      For i = 1 To Len(TotalValue3)
        TotalValue = TotalValue + TempNr(i)
      Next i
      
      '...Add TotalValue2 and TotalValue together
      TotalValue = TotalValue + TotalValue2
      '...Subtract the second character from TotalValue from 10, this number should now equal the control character
      TotalValue = 10 - Right(TotalValue, 1)
     
      
    Case 2      '...UIF Number
      '... Multiple the Value and then MOD it
      TestNr(1) = (TempNr(1)) * 1 Mod 11
      TestNr(2) = (TempNr(2)) * 2 Mod 11
      TestNr(3) = (TempNr(3) * 4) Mod 11
      TestNr(4) = (TempNr(4) * 5) Mod 11
      TestNr(5) = (TempNr(5) * 7) Mod 11
      TestNr(6) = (TempNr(6) * 8) Mod 11
      TestNr(7) = (TempNr(7) * 2) Mod 11

      '...return the total value except the checksum
      For i = 1 To ValOption - 1
        TotalValue = TotalValue + TestNr(i)
      Next i
      '...now MOD the total value
      TotalValue = TotalValue Mod 10
   
  End Select

  '...check that the total value = checksum
  If TotalValue <> CheckSum Then
    vErrString = "The " & ValText & " you have entered is invalid."
  End If

  '...return error if any
  If vErrString <> "" Then
    Validate = vErrString
  End If

End Function

'         ID VALIDATION
'      ----------------------------------
'        {YYMMDD}{G}{SSS}{C}{A}{Z}
'        YYMMDD: Date of birth
'        G : Gender. 0-4 Female; 5-9 Male.
'        SSS : Sequence No. for DOB/G combination.
'        C : Citizenship. 0 SA; 1 Other.
'        A : Usually 8, or 9 (can be other values)
'        Z : CheckSum digit.



Public Function IDValid(IDNumber As String) As Integer

  '... Total A - add digits in odd positions of Id number together (I.e. 1st + 3rd + 5th + 7th + 9th + 11th)
  '... Total B - add digits in even positions as an integer and multiply it by 2.  Then add all the individual digits together
  '... Total C - add total A and B together
  '... Control digit (CD) = 10 - last digit (units) of C.  If total C is a multiple of 10, the CD will be 0

  Dim X As Integer
  Dim TrimID As String
  Dim StringB As String
  Dim TotC As Double
  Dim TotB As Double
  Dim TotA As Integer
  Dim ControlString As String
  Dim Control As Integer

  IDValid = False

  TrimID = Trim$(IDNumber)
  If Len(TrimID) = 13 Then
    For X = 1 To 11 Step 2
      TotA = TotA + Val(Mid$(TrimID, X, 1))
    Next

    For X = 2 To 12 Step 2
      StringB = StringB + Mid$(TrimID, X, 1)
    Next
        
    TotB = Val(StringB) * 2
    StringB = Trim$(Str$(TotB))
        
    TotB = 0
    For X = 1 To Len(StringB)
      TotB = TotB + Val(Mid$(StringB, X, 1))
    Next
        
    TotC = TotA + TotB
    ControlString = Str$(TotC)

    If Len(Trim$(ControlString)) > 1 Then
      Control = 10 - Val(Right$(ControlString, 1))
      If Control > 9 Then
        Control = Val(Right$(Str$(Control), 1))
      End If
    End If

    If Control = Val(Right$(TrimID, 1)) Then
      IDValid = True
    End If
  End If

End Function







