Attribute VB_Name = "done"
'Jake Ayoub 12/18/2020
'Updated 1/6/2021
'the code below will delete the Pasted_Sheet sheet and clear the contents of Credentialing_Work_History
Sub done()
    Dim new_sheet As String
    new_sheet = "Pasted_Data"
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = new_sheet Then
            Worksheets(i).Delete
        End If
            Next
    Sheets("Credentialing_Work_History").Cells.Interior.Color = xlNone
    Sheets("Credentialing_Work_History").UsedRange.ClearContents
    Range("A1").Select
    MsgBox ("Good Job !!! Have a nice day")
    
End Sub


























