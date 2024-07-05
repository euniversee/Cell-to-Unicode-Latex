Attribute VB_Name = "Module2"
Function FormulaToUnicode(cell As Range) As String
    Dim formula As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim result As String
    Dim numerator As String
    Dim denominator As String
    
    formula = cell.formula

    If Left(formula, 1) = "=" Then
        formula = Mid(formula, 2)
    End If
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "[A-Z]+\d+"

    Set matches = regex.Execute(formula)

    For Each match In matches
        Dim cellRef As String
        Dim cellValue As Variant
        
        cellRef = match.Value
        cellValue = Range(cellRef).Value
        
        formula = Replace(formula, cellRef, cellValue)
    Next match

    formula = Replace(formula, "*", "×")
    formula = Replace(formula, "/", "÷")
    formula = Replace(formula, "^", "^")

    Dim leftCell As Range
    Set leftCell = cell.Offset(0, -1)
    Dim leftValue As Variant
    leftValue = leftCell.Value
    
    If InStr(formula, "÷") > 0 Then
        Dim parts() As String
        parts = Split(formula, "÷")
        numerator = parts(0)
        denominator = parts(1)
        
        numerator = Replace(numerator, "^", "^")
        
        result = leftValue & "=(" & numerator & " / " & denominator & ")"
    Else
        result = leftValue & "=(" & formula & ")"
    End If

    FormulaToUnicode = result
End Function

