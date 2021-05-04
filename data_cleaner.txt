Attribute VB_Name = "data_cleaner"
'this macro will clean the Fastaff_facilities database

Sub data_cleaner()

'create variables that will be used in this macro
Dim current_col As String
Dim city_col_name As String
Dim lastRowGap As String
Dim lastRowGap_true As String
Dim delete_row As String
Dim range_used As String
Dim counter As Integer

'code will run n times because it fixes a sneaky bug
counter = 0
While counter < 3
    counter = counter + 1
    'open fastaff facilities database
    With Fastaff_Facilities
        'get the range of company_names
        lastRowGap = Range("A2").End(xlDown).Row
        lastRowGap_true = "A1:A" & lastRowGap
        'loop through each company and check if it is a special value
        For Each cell In Range(lastRowGap_true)
            'this function is located in module 4 of this workbook
            If IsInArray(cell.Value, Array("Time Off", "Personal Time", "Gap In Employment", "TIME OFF", "Time Away from Work", "Time Off ", "Time off", "Personal Time off", "Personal Time Off", "Personal Time off", "time off", "Gap in Employment", "Employment Gap", "Not working", "not working")) = True Then
                'get the row to delete, and then delete the entire row
                delete_row = cell.Row
                Rows(delete_row).EntireRow.Delete
            End If
        Next
    End With
Wend

Dim city_range As String
Dim state_range As String


city_range = getColRange("Company_City")
state_range = getColRange("Company_State")

For Each n In Range(city_range)
    If n.Value = "" Then
        n.Value = "Not in Database"
    End If
Next

For Each j In Range(state_range)
    If j.Value = "" Then
        j.Value = "Not in Database"
    End If
Next

'-- Updated 12/23/2020 This function will Compare name, city and state to remove duplicates
Range("A:D").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
Worksheets("Fastaff_Facilities").Columns("A:D").EntireColumn.AutoFit

End Sub
