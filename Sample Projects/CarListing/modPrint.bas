Attribute VB_Name = "modPrint"
Option Explicit
Dim RepRS As New ADODB.Recordset
'Global RepCN As New ADODB.connection
Global cmd As String
Global mYdb As String
Global lcTotal As Integer


Public Sub PrintInvoice(ID As Integer)

'  Set RepCN = New ADODB.connection
  Set RepRS = New ADODB.Recordset
  Dim intCtrl As Integer
  Dim z As Integer
  Dim ImgPath_ As String
  Dim vDate As Single
  Dim SerString As String
  Dim SerCnt As Integer
  
'    '----temp stuff
'  RepCN.CursorLocation = adUseClient
'  mYdb = App.Path & "\DB.mdb"
'  RepCN.ConnectionString = "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & mYdb & ";Jet OLEDB:Database Password=Starlight1"
'  RepCN.Open
'  If RepRS.State = 1 Then Set RepRS = Nothing
'  Call GetCompanyDetails
'  '----temp stuff
  
  vDate = Now()
  ImgPath_ = App.Path
  If Right(ImgPath_, 1) <> "\" Then ImgPath_ = ImgPath_ & "\"
  Call GetDeal(ID)
  Call GetMasterVehicleDetails(Deal.VehicleID)

    With RepRS
      sql = "Select ID, VehicleMake, ModelYear,VehicleModel from MasterListing WHERE ID = " & Deal.VehicleID
        .Open sql, cn, adOpenKeyset, adLockOptimistic
          If .EOF Then
            .Close
            MsgBox "No Invoice Items found", vbExclamation, "System error ..."
            Exit Sub
          End If
    End With

  With rptInvoice
     .Hide
     .Title = "Invoice " & Deal.ID & " ..."
     
    Set .DataSource = RepRS
     .DataMember = ""
     
          With .Sections("Section1").Controls
            For intCtrl = 1 To .Count
                If TypeOf .Item(intCtrl) Is RptTextBox Then
                    .Item(intCtrl).DataMember = ""
                    .Item(intCtrl).DataField = RepRS(z + 1).Name
                    z = z + 1
                End If
            Next intCtrl
            
          If TypeOf .Item("lblModelYear") Is RptLabel Then
            .Item("lblModelYear").Caption = MasterVehicle.ModelYear
          End If
          
          If TypeOf .Item("lblModel") Is RptLabel Then
            .Item("lblModel").Caption = MasterVehicle.VehicleMake & " " & MasterVehicle.VehicleModel
          End If
          
          If TypeOf .Item("lblPrice") Is RptLabel Then
            .Item("lblPrice").Caption = FormatCurrency(Deal.VehicleSold, 2, vbTrue, vbTrue, vbTrue)
          End If
          
          SerString = ""
          For SerCnt = 1 To 3
            If Deal.ServicePlanEnabled(SerCnt) Then
              If SerCnt = 1 Then
                'ServicePlan
                SerString = SerString & PadString("Service Plan", 25) & "   Up to " & Deal.ServiceKMs(SerCnt) & " KMs or " & Format(Deal.ServiceDate(SerCnt), "dd MMMM yyyy") & vbCrLf
              ElseIf SerCnt = 2 Then
                'Maintenance
                SerString = SerString & PadString("Maintenance Plan", 25) & "Up to " & Deal.ServiceKMs(SerCnt) & " KMs or " & Format(Deal.ServiceDate(SerCnt), "dd MMMM yyyy") & vbCrLf
              ElseIf SerCnt = 3 Then
                'Warranty
                SerString = SerString & PadString("Warranty Plan", 25) & "  Up to " & Deal.ServiceKMs(SerCnt) & " KMs or " & Format(Deal.ServiceDate(SerCnt), "dd MMMM yyyy") & vbCrLf
              End If
            End If
          Next SerCnt
          
          'print all service details
          If TypeOf .Item("lblService") Is RptLabel Then
            .Item("lblService").Caption = SerString
          End If
          
          
          
          
      End With

      If Dir(ImgPath_ & "MyLogo.jpg") <> "" Then
        Set rptInvoice.Sections("Section4").Controls("imglogo").Picture = LoadPicture(ImgPath_ & "MyLogo.jpg")
      End If
          
' ------------------------------------------- SECTION 4 ---------------------------------------
          
     With .Sections("Section4").Controls
        If TypeOf .Item("lblcompany") Is RptLabel Then
          .Item("lblcompany").Caption = Company.Name
          If Len(Company.Name) > 20 Then .Item("lblcompany").Font.Size = 16
        End If
     End With
      
     '...Top Part of Report (Section 4)
      With .Sections("Section4").Controls
        If TypeOf .Item("lblContactPerson") Is RptLabel Then
          .Item("lblContactPerson").Caption = "Contact Person : " & Company.Person
        End If
      End With
      
      With .Sections("Section4").Controls
        If TypeOf .Item("lblCompfax") Is RptLabel Then
          .Item("lblCompfax").Caption = "Telephone : " & Company.Telephone
        End If
      End With
      
      With .Sections("Section4").Controls
        If TypeOf .Item("lblcompemail") Is RptLabel Then
          .Item("lblcompemail").Caption = "Email : " & Company.Email
        End If
      End With
     
     
     
     With .Sections("Section4").Controls
      If TypeOf .Item("lblinvoice") Is RptLabel Then
        .Item("lblinvoice").Caption = Deal.ID
      End If
     End With
  
     With .Sections("Section4").Controls
      If TypeOf .Item("lbldate") Is RptLabel Then
        .Item("lbldate").Caption = Format(vDate, "dd MMMM yyyy")
      End If
     End With
         

     With .Sections("Section4").Controls
          If TypeOf .Item("lblclient") Is RptLabel Then
          .Item("lblclient").Caption = "Invoice To : "
          End If
     End With
     With .Sections("Section4").Controls
          If TypeOf .Item("lblClientPerson") Is RptLabel Then
          .Item("lblClientPerson").Caption = "Contact : " & Deal.BuyerFirstName & " " & Deal.BuyerLastName
          End If
     End With
     With .Sections("Section4").Controls
          If TypeOf .Item("lblClientCell") Is RptLabel Then
          .Item("lblClientCell").Caption = "Tel : " & Deal.SellerContactNr
          End If
     End With
     With .Sections("Section4").Controls
          If TypeOf .Item("lblclientemail") Is RptLabel Then
          .Item("lblclientemail").Caption = "Email : " & Deal.SellerEmailAddress
          End If
     End With


      With .Sections("Section4").Controls
           If TypeOf .Item("lblClientVat") Is RptLabel Then
              If Trim(Deal.BuyerAltContactNr) = "" Then
                .Item("lblClientVat").Caption = ""
              Else
                .Item("lblClientVat").Caption = "Alt. Tel : " & Deal.BuyerContactNr
              End If
           End If
      End With

  ' ------------------------------------------- SECTION 4 ---------------------------------------

     
     '=============================================================== vehicle details
     
    With .Sections("Section4").Controls
          If TypeOf .Item("lblreg") Is RptLabel Then
          .Item("lblreg").Caption = "Reg: " & Deal.VehicleRegNr
          End If
     End With
     
     With .Sections("Section4").Controls
          If TypeOf .Item("lblvin") Is RptLabel Then
          .Item("lblvin").Caption = "Vin No: " & Deal.VehicleVINNr
          End If
     End With

    With .Sections("Section4").Controls
      If TypeOf .Item("lblKM") Is RptLabel Then
       .Item("lblKM").Caption = "Odometer: " & Deal.VehicleKM
      End If
    End With
  
     '===============================================================
  
  '   Bottom Part of Report (Section 5)
     With .Sections("Section5").Controls
          If TypeOf .Item("lblsubtotal") Is RptLabel Then
          .Item("lblsubtotal").Caption = FormatCurrency(Deal.VehicleSold, 2, vbTrue, vbTrue, vbTrue)
          End If
     End With
     
 
     With .Sections("Section5").Controls
          If TypeOf .Item("lblvat") Is RptLabel Then
          .Item("lblvat").Caption = FormatCurrency(0, 2, vbTrue, vbTrue, vbTrue)
          End If
     End With
  
     With .Sections("Section5").Controls
          If TypeOf .Item("lbltotal") Is RptLabel Then
          .Item("lbltotal").Caption = FormatCurrency(Deal.VehicleSold, 2, vbTrue, vbTrue, vbTrue)
          End If
     End With
     
    With .Sections("Section5").Controls
          If TypeOf .Item("lblbank") Is RptLabel Then
          .Item("lblbank").Caption = Company.BankingDetails
          End If
     End With
     
     With .Sections("Section5").Controls
          If TypeOf .Item("lblfooter") Is RptLabel Then
          .Item("lblfooter").Caption = Company.TermsConditions
          End If
     End With
  
  .Refresh
     
  ' rptInvoice.Terminate will close recordset
  On Error Resume Next
  'RepRS.Close
  .Show
  On Error GoTo 0
  End With
  'Set RepRS = Nothing
  

End Sub


