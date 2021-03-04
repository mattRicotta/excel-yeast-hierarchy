Private Sub Worksheet_Change(ByVal Target As Range)
    'Exit Sub if target cell is not "Selected Yeast"
    If Not Target.Address = Range("SelectedYeast").Address Then
        Exit Sub
    End If
    
    'Declare Variables
    Dim yst_names As Range
    Dim yst_parents As Range
    Dim yst_ids As Range
    Dim yst_parent As Long
    Dim lineage As New Collection
    
    'Set range variables
    Set yst_names = Range("Table1[Full_Name]")
    Set yst_parents = Range("Table1[Parent_ID]")
    Set yst_ids = Range("Table1[ID]")
    
    'Get number of table rows
    Dim row_count As Long
    row_count = yst_names.Rows().Count
    
    'Find parent and add to collection
    yst_parent = Range("SelectedYeast").Value
    Do Until yst_parent = 0
        For i = 0 To row_count
            If yst_ids.Item(i) = yst_parent Then
                lineage.Add yst_names(i)
                yst_parent = yst_parents.Item(i)
                Exit For
            End If
        Next
        Debug.Print "Parent Yeast #" & lineage.Count & ": " & lineage.Item(lineage.Count)
    Loop
    lineage.Remove 1
    
    'Print yeast parent lineage
    Dim write_range As Range
    Set write_range = Range("YeastLineage")
    For i = 0 To row_count
        If Len(write_range.Offset(i, 0)) > 0 Then
            write_range.Offset(i, 0).Clear
            write_range.Offset(i, 1).Clear
        End If
    Next
    For i = 1 To lineage.Count
        write_range.Offset(i - 1, 0) = "Gen " & lineage.Count - i
        write_range.Offset(i - 1, 1) = lineage.Item(i)
    Next
End Sub
