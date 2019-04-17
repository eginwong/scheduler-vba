Const USER_ROW = 5
Const ROLE_ROW = 4
Const SCHEDULE_DATE_CELL = "README!$J$8"
Const OBJECTIVE_CELL = "B2"

Sub AddUserToEngineSht(readmeNameCellRefString)
    Dim roleCount As Long
    roleCount = CountRoles
    
    Dim userRow As Long

    With Worksheets(ENGINE_NAME)
        .Rows(USER_ROW).Insert Shift:=xlShiftUp, CopyOrigin:=1 'insert new row
        .range("A" & USER_ROW).Formula = readmeNameCellRefString
        
        Dim userCapabilityCol As String
        Dim capabilitiesCapabilityCol As String
        
        'Insert user into priority function table (left)
        For i = 1 To roleCount
            userCapabilityCol = Col_Letter(i + 1)
            .Cells(USER_ROW, i + 1).Formula = "=IF(AND(README!$G9=TRUE,CAPABILITIES!" & Col_Letter(i + 2) & (USER_ROW + 2) & "= 1), " _
                & SCHEDULE_DATE_CELL & "-" & USERS_NAME & "!" & userCapabilityCol & "4, -100000000)"
        Next i
        .range(Col_Letter(3 + roleCount) & USER_ROW).Formula = readmeNameCellRefString
        
        'insert zeros into calculation table (right)
        For i = 1 To roleCount
            userCapabilityCol = Col_Letter(i + 1)
            .Cells(USER_ROW, i + 3 + roleCount) = 0
        Next i
        
        'update aggregation formulas in calculation table (right)
        Dim leftCol As String
        Dim rightCol As String
        leftCol = Col_Letter(4 + roleCount)
        rightCol = Col_Letter(3 + 2 * roleCount)
        .Cells(USER_ROW, (2 * roleCount + 4)).Formula = "=SUM(" & leftCol & USER_ROW & ":" & rightCol & USER_ROW & ")"
        
        Dim lastRowWithUser As Long
        lastRowWithUser = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        Call UpdateObjectiveFunction
        
        Dim actualCellRow As Long
        actualCellRow = .range("ACTUAL").Row
            
        'update Actuals
        For i = 1 To roleCount
            userCapabilityCol = Col_Letter(i + 3 + roleCount)
            .Cells(actualCellRow, i + 3 + roleCount).Formula = "=SUM(" & userCapabilityCol & USER_ROW & ":" & userCapabilityCol & lastRowWithUser & ")"
        Next i
    End With
End Sub

Sub RemoveUserFromEngineSht(userName)
    Dim engineCellWithUser As range
    Set engineCellWithUser = Worksheets(ENGINE_NAME).Columns("A").Find(userName)
    If NotEmpty(engineCellWithUser) Then
        engineCellWithUser.EntireRow.Delete
    End If
End Sub

'No ref needed as it is a copy from another column
Sub AddRoleToEngineSht()
    Dim roleCount As Long
    roleCount = CountRoles
    
    Dim firstNewRoleCol As Long
    firstNewRoleCol = roleCount + 1
    
    Dim secondNewRoleCol As Long
    secondNewRoleCol = 2 * roleCount + 3

    InsertAndCopyColumn ENGINE_NAME, firstNewRoleCol
    InsertAndCopyColumn ENGINE_NAME, secondNewRoleCol
    
    With Worksheets(ENGINE_NAME)
        'Update role totals
        Dim roleTotalRange As range
        Set roleTotalRange = .Cells(ROLE_ROW, 4 + 2 * roleCount)
        
        Dim secondRoleCol As String
        secondRoleCol = Col_Letter(roleCount + 4)
        
        Dim i As Long
        For i = 1 To CountUsers
            .Cells(roleTotalRange.Row + i, roleTotalRange.Column).Formula = "=SUM(" & secondRoleCol & (roleTotalRange.Row + i) & _
                ":" & Col_Letter(secondNewRoleCol) & (roleTotalRange.Row + i) & ")"
        Next i
    End With
    
    Call UpdateObjectiveFunction
End Sub

Sub RemoveRoleFromEngineSht(roleName)
    Dim engineCellWithRole As range
    
    Set engineCellWithRole = Worksheets(ENGINE_NAME).Rows(ROLE_ROW).Find(What:=roleName, LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If NotEmpty(engineCellWithRole) Then
        engineCellWithRole.EntireColumn.Delete
    End If
End Sub

Function SolveSchedule(SolutionMethod As String)
    
    'Ideally, we refactor this into separate methods but I want to avoid creating extra Subprocedures to retrieve the ranges for the
    'different constraints
    With Worksheets(ENGINE_NAME)
        'CONSTRAINT: align role counts
        Dim actualRoleConstraint As String
        actualRoleConstraint = Col_Letter(.range("ACTUAL").Column + 1) & .range("ACTUAL").Row _
            & ":" & Col_Letter(.range("ACTUAL").Column + CountRoles()) & .range("ACTUAL").Row
            
        Dim expectedRoleConstraint As String
        expectedRoleConstraint = Col_Letter(.range("EXPECTED").Column + 1) & .range("EXPECTED").Row _
            & ":" & Col_Letter(.range("EXPECTED").Column + CountRoles()) & .range("EXPECTED").Row
        
        'CONSTRAINT: working area as binary
        Dim firstUserRow As Long
        firstUserRow = USER_ROW
        Dim firstRoleCol As Long
        firstRoleCol = 4 + CountRoles()
        
        Dim workingAreaConstraint As String
        workingAreaConstraint = Col_Letter(firstRoleCol) & firstUserRow _
            & ":" & Col_Letter(firstRoleCol + CountRoles() - 1) & (firstUserRow + CountUsers() - 1)
        
        'CONSTRAINT: maximum role per person
        Dim maxRolePerPersonConstraint As String
        maxRolePerPersonConstraint = Col_Letter(firstRoleCol + CountRoles()) & firstUserRow _
            & ":" & Col_Letter(firstRoleCol + CountRoles()) & (firstUserRow + CountUsers() - 1)
              
        Dim solverResult As Variant

        Select Case SolutionMethod
            Case "OpenSolver"
                OpenSolver.AddConstraint .range(expectedRoleConstraint), RelationEQ, .range(actualRoleConstraint)
                OpenSolver.AddConstraint .range(workingAreaConstraint), RelationBIN
                OpenSolver.AddConstraint .range(maxRolePerPersonConstraint), RelationLE, , 1
                
                'OBJECTIVE
                OpenSolver.SetObjectiveFunctionCell .range(OBJECTIVE_CELL)
                OpenSolver.SetDecisionVariables .range(workingAreaConstraint)
                OpenSolver.SetObjectiveSense MaximiseObjective
                OpenSolver.SetChosenSolver "CBC"
                
                solverResult = OpenSolver.RunOpenSolver(False, True)
                SolveSchedule = ParseOpenSolverReturnCodes(solverResult) And CheckPositiveObjective
        
            Case "Solver"
                Application.Run "SolverReset"
                Application.Run "SolverAdd", expectedRoleConstraint, 2, actualRoleConstraint
                Application.Run "SolverAdd", workingAreaConstraint, 5
                Application.Run "SolverAdd", maxRolePerPersonConstraint, 1, 1
                Application.Run "SolverOk", OBJECTIVE_CELL, 1, "0", workingAreaConstraint, 2
                
                solverResult = Application.Run("SolverSolve", True)
                SolveSchedule = ParseSolverReturnCodes(solverResult) And CheckPositiveObjective
        End Select
    End With
End Function

Sub UpdateObjectiveFunction()
    Dim roleCount As Long
    roleCount = CountRoles
    
    With Worksheets(ENGINE_NAME)
        Dim lastRowWithUser As Long
        lastRowWithUser = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        .range("B2").Formula = "=SUMPRODUCT($B$" & USER_ROW & ":$" & Col_Letter(1 + roleCount) & "$" _
            & lastRowWithUser & ",$" & Col_Letter(4 + roleCount) & "$" & USER_ROW & ":$" & Col_Letter(3 + 2 * roleCount) & "$" & lastRowWithUser & ")"
    End With
End Sub

Function ParseSolverReturnCodes(code)
    ParseSolverReturnCodes = False
    If (IsError(code)) Then
        MsgBox "It is possible that you have hit the limit on the model. To continue using the scheduler, either remove users/roles or follow instructions in the OPENSOLVER_INSTRUCTION spreadsheet to install OpenSolver.", vbCritical
        'Trigger new spreadsheet with instructions!
        Worksheets(OPENSOLVER_INSTRUCTIONS_NAME).Visible = True
        Worksheets(OPENSOLVER_INSTRUCTIONS_NAME).Activate
        'Immediately exit
        Exit Function
    End If
    Select Case code
    Case 0 To 2
        ParseSolverReturnCodes = True
    'Case 3 To 13 are False
    End Select
End Function

Function ParseOpenSolverReturnCodes(code)
    ParseOpenSolverReturnCodes = False
    If code = 0 Then
        ParseOpenSolverReturnCodes = True
    End If
End Function

'Non-negative objective cell causes Solver to fail
'Need to check value through VBA instead
Function CheckPositiveObjective()
    CheckPositiveObjective = Worksheets(ENGINE_NAME).range(OBJECTIVE_CELL).Value > 0
End Function
