Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim originalEntry As Range
    Dim rowCount As Long, currentRow As Long, currentCol As Long
    Dim key1FormulaTemplate As String, key2FormulaTemplate As String
    
    key1FormulaTemplate = "C[Row]&H[Row]"
    key2FormulaTemplate = "D[Row]&I[Row]"
    
    rowCount = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    currentRow = Target.Row
    currentCol = Target.Column
    
    If currentCol = 3 Then
      Call TryToFindEmployee(Target, rowCount)
    End If
    
    If currentCol = 8 Or currentCol = 9 Then
        Set originalEntry = CheckForDublicates("A", key1FormulaTemplate, rowCount, currentRow)
        If originalEntry Is Nothing Then
         Set originalEntry = CheckForDublicates("B", key2FormulaTemplate, rowCount, currentRow)
        End If
        
        If Not originalEntry Is Nothing Then
           Dim response As Integer
           response = MsgBox("Such an entry is already in the list. would you like to see it?", vbYesNo, "Duplicate found")
           If response = vbYes Then
            originalEntry.Select
           End If
        End If
    End If
End Sub
Sub TryToFindEmployee(changed_cell As Range, rowCount As Long)
    Dim currentKey As String
    
    currentKey = changed_cell.Value2
    
    Dim keyRange As Range
    Set keyRange = ActiveSheet.Range("C2:C" & rowCount)
    Dim foundRange As Range
    Set foundRange = keyRange.Find(currentKey, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    
    Dim response As Integer
    If Not (foundRange Is Nothing) Then
      If foundRange.Row <> changed_cell.Row Then
        Cells(changed_cell.Row, 4).Value = Cells(foundRange.Row, 4).Value
        Cells(changed_cell.Row, 5).Value = Cells(foundRange.Row, 5).Value
        Cells(changed_cell.Row, 6).Value = Cells(foundRange.Row, 6).Value
      End If
    End If
End Sub
Function CheckForDublicates(keyColumn As String, keyFormulaTemplate As String, rowCount As Long, currentRow As Long) As Range
    Dim keys() As String
    keys = Split(keyFormulaTemplate, "&")
    Dim keyFormula As String, item As Variant
    Dim currentKey As String
    
    For Each item In keys
     keyFormula = Replace(item, "[Row]", currentRow)
     currentKey = currentKey & Range(keyFormula).Value2
    Next item
    
    Dim keyRange As Range
    Set keyRange = ActiveSheet.Range(keyColumn & "2:" & keyColumn & rowCount)
    Dim foundRange As Range
    Set foundRange = keyRange.Find(currentKey, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    
    Dim response As Integer
    If Not (foundRange Is Nothing) Then
      If foundRange.Row <> currentRow Then
       Set CheckForDublicates = foundRange
      Else
       Set CheckForDublicates = Nothing
      End If
    End If
End Function
