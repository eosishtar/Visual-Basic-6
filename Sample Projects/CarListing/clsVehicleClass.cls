VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVehicleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mV As clsVehicleClass


Public Property Get ID() As Integer
  ID = mV.ID
End Property

Public Property Let ID(ByVal vNewValue As Integer)
  mV.ID = vNewValue
End Property

Public Property Get ModelYear() As Integer
  ModelYear = mV.ModelYear
End Property

Public Property Let ModelYear(ByVal vNewValue As Integer)
  mV.ModelYear = vNewValue
End Property

Public Property Get VehicleMake() As String
  VehicleMake = mV.VehicleMake
End Property

Public Property Let VehicleMake(ByVal vNewValue As String)
  mV.VehicleMake = vNewValue
End Property







  ID As Integer
  ModelYear As Integer
  VehicleMake As String
  VehicleModel As String
  VehicleDisplacement As String
  VehicleType As String
  RatedHorsePower As Single
  NrOrCylinders As Integer
  EngineCode As String
  NrOfGears As Integer
  DriveSystemCode As String
  DriveSystemDescription As String
