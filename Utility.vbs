'https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'https://peltiertech.com/Excel/SolverVBA.html
Function CheckSolver() As Boolean
  '' Adjusted for Application.Run() to avoid Reference problems with Solver
  '' Peltier Technical Services, Inc., Copyright © 2007. All rights reserved.
  '' Returns True if Solver can be used, False if not.

  Dim bSolverInstalled As Boolean

  '' Assume true unless otherwise
  CheckSolver = True

  On Error Resume Next
  ' check whether Solver is installed
  bSolverInstalled = Application.AddIns("Solver Add-In").Installed
  Err.Clear

  If bSolverInstalled Then
    ' uninstall temporarily
    Application.AddIns("Solver Add-In").Installed = False
    ' check whether Solver is installed (should be false)
    bSolverInstalled = Application.AddIns("Solver Add-In").Installed
  End If

  If Not bSolverInstalled Then
    ' (re)install Solver
    Application.AddIns("Solver Add-In").Installed = True
    ' check whether Solver is installed (should be true)
    bSolverInstalled = Application.AddIns("Solver Add-In").Installed
  End If

  If Not bSolverInstalled Then
    MsgBox "Solver not found. This workbook will not work.", vbCritical
    CheckSolver = False
  End If

  If CheckSolver Then
    ' initialize Solver
    Application.Run "Solver.xlam!Solver.Solver2.Auto_open"
  End If

  On Error GoTo 0

End Function

'https://stackoverflow.com/questions/9879825/how-to-add-a-reference-programmatically
Sub AddReferences(wbk As Workbook)
    ' Run DebugPrintExistingRefs in the immediate pane, to show guids of existing references
    AddRef wbk, "{00025E01-0000-0000-C000-000000000046}", "DAO"
    AddRef wbk, "{00020905-0000-0000-C000-000000000046}", "Word"
    AddRef wbk, "{91493440-5A91-11CF-8700-00AA0060263B}", "PowerPoint"
End Sub

Sub AddRef(wbk As Workbook, sGuid As String, sRefName As String)
    Dim i As Integer
    On Error GoTo EH
    With wbk.VBProject.References
        For i = 1 To .Count
            If .Item(i).name = sRefName Then
               Exit For
            End If
        Next i
        If i > .Count Then
           .AddFromGuid sGuid, 0, 0 ' 0,0 should pick the latest version installed on the computer
        End If
    End With
EX: Exit Sub
EH: MsgBox "Error in 'AddRef'" & vbCrLf & vbCrLf & Err.Description
    Resume EX
    Resume ' debug code
End Sub

Function InputWithExit(prompt As String, boxName As String)
    Dim result As String
    result = Trim(WorksheetFunction.Proper(InputBox(prompt, boxName)))
    If result = "" Then End
    InputWithExit = result
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

