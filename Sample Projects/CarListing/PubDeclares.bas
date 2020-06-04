Attribute VB_Name = "PubDeclares"
Option Explicit



Global Const NonRequired = "     "
Global Const ModelsDir_$ = "\Models\"
Global Const DataDir_$ = "\Data\"
Public Pgm_Name As String

Global Const Max_Vehicle_Years = 35
Global TempRS As Recordset

Global vReportMenu As Integer
Global ReportHead As String

Global sql As String
Global fso As FileSystemObject
Global itmx As Object

' public types
Global MasterVehicle As VehicleType
Global Main As MainCore
Global Company As CompanyDetails
Global User As UserClass
Global Deal As ClientClass
Global Expense As ExpenseClass
Global PCode As PostalCodeClass
Global CloseDeal As CloseDealInfo

