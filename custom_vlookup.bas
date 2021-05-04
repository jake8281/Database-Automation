Attribute VB_Name = "custom_vlookup"
'completed 1/4/2021 1:01am
'created by Jake Ayoub
'this function is for linking the data in credentialing_work_history with the right value from database based on its matching criteria

Sub custom_vlookup()
    Application.ScreenUpdating = False
    
    Dim gap_aggregate As String
    Dim database_aggregate As String
    Dim gap_col_range As String
    Dim database_col_range As String
    
    Dim gap_first_zip As String
    Dim gap_range_zip As String
    
    Dim gap_row As String
    Dim database_row As String
    
    Dim gap_city As String
    Dim gap_state As String
    Dim database_city As String
    Dim database_state As String
    
    Dim gap_col_zip As String
    Dim cell_to_formulate As String
    
    Dim database_zip As String
    
    
    'get information from credentialing work history (gap) worksheet
    'some of these are not used because I have ideas for improvement
    With Worksheets("Credentialing_Work_History").Activate
        gap_col_range = getColRange("Company_Name")
        gap_first_zip = "" & getColStr("Company_Postal_Code") & 2
        gap_range_zip = getColRange("Company_Postal_Code")
    End With
    
    'get information from database
    With Worksheets("Fastaff_Facilities").Activate
        database_col_range = getColRange("Company_Name")
    End With

    'this is the big nest
    'step 1: loop through the company name on the gap sheet
    'step 2: loop through the company name on the database
    'step 3: check with if statement if both company names are the same
    'step 4: create the variables I need to make the aggregates
    'step 5: check if aggregates are equal
    'step 6: if they are equal, then the zip code should be copied to the gap sheet
    
    'step 1
    For Each gap_cell In Worksheets("Credentialing_Work_History").Range(gap_col_range)
        'step 2
        For Each database_cell In Worksheets("Fastaff_Facilities").Range(database_col_range)
            'step 3
            If gap_cell = database_cell Then
                gap_row = gap_cell.Row
                
                'step 4
                With Worksheets("Credentialing_Work_History").Activate
                    gap_city = "" & getColStr("Company_City") & gap_row
                    gap_state = "" & getColStr("Company_State") & gap_row
                    gap_aggregate = "" & gap_cell.Value & Range(gap_city).Value & Range(gap_state).Value
                End With
                
                With Worksheets("Fastaff_Facilities").Activate
                    database_row = database_cell.Row
                    database_city = "" & getColStr("Company_City") & database_row
                    database_state = "" & getColStr("Company_State") & database_row
                    database_aggregate = "" & database_cell.Value & Range(database_city).Value & Range(database_state).Value
                End With
                
                'step 5
                If gap_aggregate = database_aggregate Then
                    With Worksheets("Credentialing_Work_History").Activate
                        gap_col_zip = getColStr("Company_Postal_Code")
                        cell_to_formulate = "" & gap_col_zip & gap_row
                        
                        With Worksheets("Fastaff_Facilities").Activate
                            database_zip = "" & getColStr("Company_Postal_Code") & database_row
                        End With
                        
                        'step 6
                        Worksheets("Credentialing_Work_History").Range(cell_to_formulate) = Worksheets("Fastaff_Facilities").Range(database_zip).Value
                    End With
                End If
            End If
        Next database_cell
    Next gap_cell
    
    With Worksheets("Credentialing_Work_History").Activate
        Range("A1").Select
    End With
   MsgBox ("Make sure all Zip codes have Green icon beside them in Credentialing_Work_History")
End Sub
