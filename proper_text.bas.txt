Attribute VB_Name = "Proper_Text"
'- Created by Jake Ayoub 1/6/2021 - Updated 1/26/2021
'-First Part: This function will check for every name inside the database and turn it into a proper case.
'Second part :  This function removes extra spaces from the end of the string.
Sub text_fix()

Dim name_rng As Range
Dim name_name_cell As Range
Dim name_name_selection As String
Dim city_rng As Range
Dim city_name_cell As Range
Dim city_name_selection As String

Dim zip_rng As String
Dim trim_row As Integer
Dim trim_rng As Range
Dim trimcell As Range

With Credentialing_Work_History
    ' First Part
    name_selection = getColRange("Company_Name")
    Set name_rng = Range(name_selection)
    For Each name_cell In name_rng
        name_cell.Value = WorksheetFunction.Proper(name_cell.Value)
    Next
    
    city_selection = getColRange("Company_City")
    Set city_rng = Range(city_selection)
    For Each city_cell In city_rng
        city_cell.Value = WorksheetFunction.Proper(city_cell.Value)
    Next
    'Second Part
    zip_rng = getColRange("Company_Name")
    ' To 'Find the last used cell in Col A
    trim_row = Range(zip_rng).End(xlDown).Row
    
    'Declare the range used by having the coordinates of rows and column till the last cell used.
    Set trim_rng = Range(Cells(2, 9), Cells(trim_row, 9))
    ' Loop through the range and remove any trailing space
    For Each trimcell In trim_rng
        trimcell = RTrim(trimcell)
    'Go to the next Cell
    Next trimcell
    
End With

End Sub



