Attribute VB_Name = "is_in_array_function"
'-- Created By Ahmed Ayoub  12/18/2020

'this function checks if a string is in an array and returns true or false
Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    'make variable that will be used in loop
    Dim i
    'loop through provided list and check if stringToBeFound is inside
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

