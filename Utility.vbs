Function CheckSolverAddin() As Boolean

    'If the product of roles and users are greater than the Solver limit (200)
    'Check for open solver instead
    Dim ExpectedDecisionVariables As Integer
    ExpectedDecisionVariables = CountRoles() * CountUsers()
    
    If (ExpectedDecisionVariables > 199) Then
        CheckSolverAddin = CheckAddin("OpenSolver")
        If Not CheckSolverAddin Then
            DisplayOpenSolverInstructions
        End If
    Else
        CheckSolverAddin = CheckAddin("Solver add-in")
        If Not CheckSolverAddin Then
            DisplaySolverInstructions
        End If
    End If
End Function

Function CheckAddin(s As String) As Boolean
    Dim x As Variant
    On Error Resume Next
    x = AddIns(s).Installed
    On Error GoTo 0
    If IsEmpty(x) Then
        CheckAddin = False
    Else
        CheckAddin = True
    End If
End Function

Sub DisplayOpenSolverInstructions()
    Worksheets(OPENSOLVER_INSTRUCTIONS_NAME).Visible = True
    Worksheets(OPENSOLVER_INSTRUCTIONS_NAME).Activate
End Sub

Sub DisplaySolverInstructions()
    Worksheets(SOLVER_INSTRUCTIONS_NAME).Visible = True
    Worksheets(SOLVER_INSTRUCTIONS_NAME).Activate
End Sub

'https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'Checks if range is empty
Function NotEmpty(nullableRange As range) As Boolean
    NotEmpty = Not nullableRange Is Nothing
End Function

'Checks if one cell is equal to the other
'https://stackoverflow.com/questions/28071459/compare-2-cells-in-excel-by-using-vbas
Function CellEquality(cell1 As range, cell2 As range) As Boolean
    CellEquality = IIf([cell1] = [cell2], True, False)
End Function

'Maps date frequencies to vba acceptable format
Function MapFrequencyToVBAFormat(frequency As String)
    Dim standardFrequency As String
    standardFrequency = Trim(UCase(frequency))
    
    Select Case standardFrequency
        Case "DAY"
            MapFrequencyToVBAFormat = "d"
        Case "WEEK"
            MapFrequencyToVBAFormat = "ww"
        Case "MONTH"
            MapFrequencyToVBAFormat = "m"
        Case "QUARTER"
            MapFrequencyToVBAFormat = "q"
        Case "YEAR"
            MapFrequencyToVBAFormat = "yyyy"
    End Select
End Function

Sub InsertAndCopyColumn(worksheetName, range)
    With Worksheets(worksheetName)
        .Columns(range).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove 'insert new column
        .Columns(range - 1).Copy Destination:=.Columns(range)
    End With
End Sub

