Attribute VB_Name = "Mathematics"
Function deg2Rad(deg_in As Currency) As Currency
    deg2Rad = deg_in / 57.2957795130823
End Function
 
Function rad2Deg(rad_in As Currency) As Currency
    rad2Deg = 57.2957795130823 * rad_in
End Function

Function quickSortAndTrackIndices(array_in() As Variant) As Variant
    Dim indices() As Long
    ReDim indices(LBound(array_in) To UBound(array_in))
    Dim sorted_array() As Variant
    ReDim sorted_array(LBound(array_in) To UBound(array_in), 1 To 2)
    
    ' Initialize indices array
    Dim i As Long
    For i = LBound(array_in) To UBound(array_in)
        indices(i) = i
    Next i
    
    ' Perform the quicksort
    Call quickSort(array_in, indices, LBound(array_in), UBound(array_in))
    
    ' Store the sorted values and their original indices
    For i = LBound(array_in) To UBound(array_in)
        sorted_array(i, 1) = array_in(i)
        sorted_array(i, 2) = indices(i)
    Next i
    
    quickSortAndTrackIndices = sorted_array
End Function

Private Sub quickSort(array_in() As Variant, indices() As Long, i_left As Long, i_right As Long)
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim temp As Variant
    Dim temp_index As Long
    
    i = i_left
    j = i_right
    pivot = array_in((i_left + i_right) \ 2)
    
    Do While i <= j
        Do While array_in(i) < pivot
            i = i + 1
        Loop
        Do While array_in(j) > pivot
            j = j - 1
        Loop
        
        If i <= j Then
            ' Swap values
            temp = array_in(i)
            array_in(i) = array_in(j)
            array_in(j) = temp
            
            ' Swap indices
            temp_index = indices(i)
            indices(i) = indices(j)
            indices(j) = temp_index
            
            i = i + 1
            j = j - 1
        End If
    Loop
    
    If i_left < j Then quickSort array_in, indices, i_left, j
    If i < i_right Then quickSort array_in, indices, i, i_right
End Sub

