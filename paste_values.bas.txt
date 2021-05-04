Attribute VB_Name = "paste_values"
'Created Jake Ayoub 12/21/2020
Sub paste_values()
    Dim name_range As String
    Dim city_range As String
    Dim state_range As String
    Dim zip_range As String
    
    Dim name_col As String
    Dim city_col As String
    Dim state_col As String
    Dim zip_col As String
    Dim new_sheet As String
    
    With Credentialing_Work_History
        ThisWorkbook.Sheets("Credentialing_Work_History").Activate
    
        name_col = getColStr("Company_Name")
        city_col = getColStr("Company_City")
        state_col = getColStr("Company_State")
        zip_col = getColStr("Company_Postal_Code")

        Dim name_rng As String
        Dim city_rng As String
        Dim state_rng As String
        Dim zip_rng As String
        
        name_rng = "" & name_col & ":" & name_col
        city_rng = "" & city_col & ":" & city_col
        state_rng = "" & state_col & ":" & state_col
        zip_rng = "" & zip_col & ":" & zip_col
        
        new_sheet = "Pasted_Data"

        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = new_sheet Then
                Worksheets(i).Delete
            End If
        Next

        With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = new_sheet
        End With
        
        ThisWorkbook.Sheets("Credentialing_Work_History").Activate
        
        With Credentialing_Work_History
            Range(name_rng).Copy
            Worksheets(new_sheet).Range(name_rng).PasteSpecial xlPasteValues
            
            Range(city_rng).Copy
            Worksheets(new_sheet).Range(city_rng).PasteSpecial xlPasteValues
            
            Range(state_rng).Copy
            Worksheets(new_sheet).Range(state_rng).PasteSpecial xlPasteValues
            
            Range(zip_rng).Copy
            Worksheets(new_sheet).Range(zip_rng).PasteSpecial xlPasteValues
        End With
    End With
        
    Dim counter As Integer
    Dim column As String
    Dim empty_string As String

End Sub
