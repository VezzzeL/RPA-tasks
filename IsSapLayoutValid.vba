Function IsSapLayoutValid(session As SAPFEWSELib.GuiSession, transanctionId As String, layout As Variant) As Boolean
    On Error Resume Next
    Dim columnFound As Boolean
    Dim missingColumns As String
    Dim i As Integer
    Dim j As Integer
    Dim columnsQuantity As Integer
    
    missingColumns = ""
    columnsQuantity = session.FindById(transanctionId).ColumnCount
    
    For i = LBound(layout) To UBound(layout)
        columnFound = False
        For j = 0 To columnsQuantity - 1
            If session.FindById(transanctionId).columnOrder.item(j) = layout(i) Then
                columnFound = True
                Exit For
            End If
        Next j
        If Not columnFound Then
            missingColumns = missingColumns & layout(i) & ", "
        End If
    Next i

    If Len(missingColumns) > 0 Then
        MsgBox "Layout is not valid, missing " & missingColumns, vbCritical
        IsLayoutValid = False
    Else
        IsLayoutValid = True
    End If

    On Error GoTo 0
End Function