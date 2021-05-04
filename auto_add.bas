Attribute VB_Name = "auto_add"
'this is a macro that connects the file from credentials gap to fastaff facilities database

Sub auto_add()
    
    Application.ScreenUpdating = False
    
    Dim rngZipStr As String
    Dim rngName As String
    Dim rngCity As String
    Dim rngState As String
    
    Dim database_sheetname As String
    Dim gaps_sheetname As String
    Dim last_row_database As Integer
    
    Dim name_range_gap As String
    Dim zip_range_gap As String
    Dim city_range_gap As String
    Dim state_range_gap As String
    
    Dim name_col_gap As String
    Dim zip_col_gap As String
    Dim city_col_gap As String
    Dim state_col_gap As String
    
    Dim zip_first As String
    Dim city_first As String
    Dim state_first As String
    
    Dim not_in_db_row As Integer
    Dim not_in_db_zip As String
    Dim not_in_db_city As String
    Dim not_in_db_state As String
    Dim lastRowFacilities As Integer
    
    Dim db_name_col As String
    Dim db_zip_col As String
    Dim db_city_col As String
    Dim db_state_col As String
    
    Dim new_name As String
    Dim new_zip As String
    Dim new_city As String
    Dim new_state As String
    
    Dim counter As Integer
    
    database_sheetname = "Fastaff_Facilities"
    gaps_sheetname = "Credentialing_Work_History"
    lastRowGap = Range("A2").End(xlDown).Row
    last_row_database = ThisWorkbook.Sheets(database_sheetname).Range("A2").End(xlDown).Row
    
    With Credentialing_Work_History
        ThisWorkbook.Sheets("Credentialing_Work_History").Activate
        
        'checks if in correct sheet
        If gaps_sheetname <> ActiveSheet.Name Then
            MsgBox ("Not in correct active sheet")
            Exit Sub
        End If
        
        'assign variables with column range given column name
        name_range_gap = getColRange("Company_Name")
        zip_range_gap = getColRange("Company_Postal_Code")
        city_range_gap = getColRange("Company_City")
        state_range_gap = getColRange("Company_State")
        
        'assign variables with column component of cell coordinate
        name_col_gap = getColStr("Company_Name")
        zip_col_gap = getColStr("Company_Postal_Code")
        city_col_gap = getColStr("Company_City")
        state_col_gap = getColStr("Company_State")
        
        'assign complete coordinates of first value given column name
        zip_first = "" & zip_col_gap & "2"
        city_first = "" & city_col_gap & "2"
        state_first = "" & state_col_gap & "2"
        
        'tells us if zip code exists or not. if not, will return "not in database", otherwise, the true value will JakeMagically appear
        Worksheets("Credentialing_Work_History").Range(zip_first).FormulaR1C1 = "=IFERROR(VLOOKUP([@[Company_Name]],Fastaff_Facilities,4,0),""Not in Database"")"
        Range(zip_first).Select
        selection.AutoFill Destination:=Range(zip_range_gap), Type:=xlFillDefault
        
        'tells us the last row of the fastaff facilities database
        Set rngZip = Range(zip_range_gap)
        last_row_database = ThisWorkbook.Sheets(database_sheetname).Range("A2").End(xlDown).Row
        
        'looping through the zip range in credential work history
        For Each cell In rngZip
            With Fastaff_Facilities
                ThisWorkbook.Sheets(database_sheetname).Activate
                
                'checks if in correct sheet
                If database_sheetname <> ActiveSheet.Name Then
                    MsgBox ("Not in correct active sheet")
                    Exit Sub
                End If
                
                'check if zip code is not in database
                If cell = "Not in Database" Then
                    
                    'catch row
                    not_in_db_row = cell.Row
                    
                    'catch name, city, state and zip and create coordinates (concatenation in this project is only used for coordinate creation)
                    not_in_db_name = "" & name_col_gap & not_in_db_row
                    not_in_db_zip = "" & zip_col_gap & not_in_db_row
                    not_in_db_city = "" & city_col_gap & not_in_db_row
                    not_in_db_state = "" & state_col_gap & not_in_db_row
                    
                    'get col coordinate given column name
                    db_name_col = getColStr("Company_Name")
                    db_zip_col = getColStr("Company_Postal_Code")
                    db_state_col = getColStr("Company_State")
                    db_city_col = getColStr("Company_City")
                    
                    'again create the coordinates
                    new_name = "" & db_name_col & last_row_database
                    new_zip = "" & db_zip_col & last_row_database
                    new_city = "" & db_city_col & last_row_database
                    new_state = "" & db_state_col & last_row_database
                
                    'check is in correct sheet, should be database
                    ThisWorkbook.Sheets(gaps_sheetname).Activate
                    If gaps_sheetname <> ActiveSheet.Name Then
                        MsgBox ("Not in correct active sheet")
                        Exit Sub
                    End If
                    
                    'add the not in database zip to the end of the fastaff facilites database
                    Range(not_in_db_name).Copy Destination:=Sheets(database_sheetname).Range(new_name).Offset(1, 0)
                    Range(not_in_db_zip).Copy Destination:=Sheets(database_sheetname).Range(new_zip).Offset(1, 0)
                    Range(not_in_db_city).Copy Destination:=Sheets(database_sheetname).Range(new_city).Offset(1, 0)
                    Range(not_in_db_state).Copy Destination:=Sheets(database_sheetname).Range(new_state).Offset(1, 0)
                    
                    'we need the the last row + 1 because we don't wanna override the last row
                    last_row_database = last_row_database + 1
                End If
            End With
        Next
    End With
    
    'vlookup sneaky problem
'    Worksheets("Credentialing_Work_History").Range(city_first).FormulaR1C1 = "=IFERROR(VLOOKUP([@[Company_Name]],Fastaff_Facilities,2,0),""Not in Database"")"
'    Range(city_first).Select
'    Selection.AutoFill Destination:=Range(city_range_gap), Type:=xlFillDefault
'
'    Worksheets("Credentialing_Work_History").Range(state_first).FormulaR1C1 = "=IFERROR(VLOOKUP([@[Company_Name]],Fastaff_Facilities,3,0),""Not in Database"")"
'    Range(state_first).Select
'    Selection.AutoFill Destination:=Range(state_range_gap), Type:=xlFillDefault
        
End Sub

