Attribute VB_Name = "modListView"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub AutosizeColumns(ByVal TargetListView As ListView)
Const SET_COLUMN_WIDTH As Long = 4126
Const AUTOSIZE_USEHEADER As Long = -2
Dim lngColumn As Long

For lngColumn = 0 To (TargetListView.ColumnHeaders.Count - 1)
Call SendMessage(TargetListView.hwnd, SET_COLUMN_WIDTH, lngColumn, ByVal AUTOSIZE_USEHEADER)

Next lngColumn
End Sub


Public Sub PopListView(ByVal myListView As ListView, ByVal mySql As String)
  Dim db As clsDBase
  Dim itMx As Object
    
  Set db = New clsDBase
    
  On Error GoTo ErrHandler
  
  db.OpenDB mySql, adOpenKeyset, adLockOptimistic

    If db.OpenRecordSet.EOF Then
        db.OpenRecordSet.Close
      Exit Sub
    End If
   
'   set listview parameters
      myListView.ColumnHeaders.Clear
      myListView.ListItems.Clear
      myListView.View = lvwReport
      myListView.BorderStyle = ccFixedSingle
      myListView.FullRowSelect = True
      myListView.GridLines = True

'   count the columns and add them to the listview
      For i = 0 To db.OpenRecordSet.Fields.Count - 1
        myListView.ColumnHeaders.Add , , db.OpenRecordSet.Fields(i).Name
      Next

'   count the rows and add the items and subitems
      db.OpenRecordSet.MoveFirst
      For j = 1 To rs.RecordCount
          Set itMx = myListView.ListItems.Add(, , db.OpenRecordSet.Fields(0).Value)
            For k = 1 To myListView.ColumnHeaders.Count - 1
                  On Error Resume Next
                itMx.SubItems(k) = db.OpenRecordSet.Fields(k).Value
            Next k
        db.OpenRecordSet.MoveNext
      Next j
      
    db.OpenRecordSet.Close
    Call AutosizeColumns(myListView)     ' resize all the columns
    



ErrHandler:
If Err.Number = 2147217900 Then
    smsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
  MsgBox (smsg)   ' only admins should see this message. will inform user if sql query is correct
End If

End Sub




