VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlDBCombo ctlDBCombo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim DatabasePath As String


Private Sub Form_Load()

  Set cn = New ADODB.Connection
  
  DatabasePath = App.Path & "\DB.mdb"
  With cn
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .Open "Provider=microsoft.jet.OLEDB.4.0; ;" & "Data Source = " & DatabasePath & ";Jet OLEDB"
  End With


  ctlDBCombo1.PopulateList cn, "Category", "Description", False
  

End Sub


Private Sub Form_Unload(Cancel As Integer)

  Set cn = Nothing

End Sub
