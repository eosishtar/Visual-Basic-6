Attribute VB_Name = "modTypes"
Option Explicit


Public Type PostalCodeClass
  ID As Integer
  PostalCode As String
  SalesMan As String
End Type

Public Type VehicleType
  ID As Integer
  ModelYear As Integer
  VehicleMake As String
  VehicleModel As String
  VehicleDisplacement As String
  VehicleType As String
  RatedHorsePower As Single
  NrOrCylinders As Integer
  EngineCode As String
  TransmissionCode As String
  TransmissionDesc As String
  NrOfGears As Integer
  DriveSystemCode As String
  DriveSystemDescription As String
  VehicleImage As String
  Notes As String
  FuelType As String
  TurboType As String
  BookDate As String
  BookValue As Double
End Type

Public Type MainCore
  Database As String
  Printer As String
  Logo As String
End Type

Public Type CompanyDetails
  ID As Integer
  Name As String
  Address As String
  Telephone As String
  Person As String
  Email As String
  BankingDetails As String
  TermsConditions As String
End Type

Public Type UserClass
  ID As Integer
  UserCode As String
  Username As String
  Password As String
  Deleted As Boolean
End Type

Public Type ClientClass
  ID As Integer
  BuyerFirstName As String
  BuyerLastName As String
  BuyerIDNumber As String
  BuyerCompanyName As String
  BuyerCompanyRegNr As String
  BuyerContactNr As String
  BuyerAltContactNr As String
  BuyerEmailAddress As String
  BuyerNotes As String
  SellerFirstName As String
  SellerLastName As String
  SellerIDNumber As String
  SellerCompanyName As String
  SellerCompanyRegNr As String
  SellerContactNr As String
  SellerAltContactNr As String
  SellerEmailAddress As String
  SellerNotes As String
  
  VehicleID As Integer
  VehicleRegNr As String
  VehicleVINNr As String
  VehicleEngineNr As String
  VehicleKM As String
  VehicleNotes As String
  VehicleImage As String
  VehicleCost As Double
  VehicleService As Double
  VehicleDateBought As String
  VehicleDateSold As String
  VehicleSold As Double
  
  ServicePlanEnabled(1 To 3) As Boolean
  ServiceKMs(1 To 3) As Single
  ServiceDate(1 To 3) As String
  ServiceNotes As String
  
  DateModified As String
  UserModified As String
  DateCreated As String
  UserCreated As String
  DealClosed As Boolean
End Type

Public Type ExpenseClass
  DealID As Integer     'linked to the deal/client ID
  DateOfExpense As String
  NatureOfExpense As String
  ValueOfExpense As Double
End Type

Public Type CloseDealInfo
  cldBuyerFirstName As String
  cldBuyerLastName As String
  cldBuyerID As String
  cldBuyerContact As String
  cldBuyDate As String
  cldBuyAmount As Double
  cldDone As Boolean
  cldReg As String
End Type
