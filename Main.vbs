Sub CREATE_USER()
    Application.ScreenUpdating = False
    
    Dim userName As String
    'Retrieve userName from input
    userName = InputWithExit("What is the name of the new user?", "Add User")
    
    'Check if user already exists
    If UniqueUser(userName) Then
        Dim readmeCellRefString As String
        readmeCellRefString = AddNewUserToReadMeSht(userName)
        Call AddUserToCapabilitiesSht(readmeCellRefString)
        Call AddUserToUsersSht(readmeCellRefString)
        Call AddUserToEngineSht(readmeCellRefString)
        Worksheets(README_NAME).Activate
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub REMOVE_USER()
    Application.ScreenUpdating = False
    
    Dim userName As String
    'Retrieve userName from input
    userName = InputWithExit("Which user would you like to remove?", "Remove User")
    
    'Validation, does the user even exist?
    If Not UniqueUser(userName) Then
        Call RemoveUserFromCapabilitiesSht(userName)
        Call RemoveUserFromUsersSht(userName)
        Call RemoveUserFromEngineSht(userName)
        Call RemoveUserFromReadMeSht(userName)
    Else
        MsgBox "User does not exist", vbCritical
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub GENERATE_SCHEDULE()
    If Not (CheckSolverAddin()) Then
        MsgBox "Requisite dependencies, Solver or OpenSolver, are not installed. Please install those first.", vbCritical
        DisplaySolverInstructions
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Call ToggleUnavailable
    
    'Need to be in ENGINE worksheet to run Solver
    Worksheets(ENGINE_NAME).Activate
    Dim solverResults As Boolean
    solverResults = SolveSchedule(CheckSolverProgram)
    
    Call ToggleAvailable
    
    If solverResults Then
        Call PrintSchedule
        Call UpdateNextDateScheduled
        Call sourceSheet.Activate
    Else
        MsgBox "Schedule is impossible given the number of roles and available people!", vbCritical
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub ADD_ROLE()
    Application.ScreenUpdating = False
    
    Dim roleName As String
    'Retrieve role from input
    roleName = InputWithExit("What is the name of the new role?", "Add Role")
    
    'Check if role already exists
    If UniqueRole(roleName) Then
        Call AddRoleToCapabilitiesSht(roleName)
        Call AddRoleToUsersSht
        Call AddRoleToEngineSht
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub REMOVE_ROLE()
    Application.ScreenUpdating = False
    
    Dim roleName As String
    'Retrieve role from input
    roleName = InputWithExit("Which role would you like to remove?", "Remove Role")
    
    If Not UniqueRole(roleName) Then
        'Run twice because there are two columns with the role
        Call RemoveRoleFromEngineSht(roleName)
        Call RemoveRoleFromEngineSht(roleName)
        Call RemoveRoleFromUsersSht(roleName)
        'must be called in this order, otherwise role name is gone
        Call RemoveRoleFromCapabilitiesSht(roleName)
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Function CheckSolverProgram() As String
    'Set default for solver program
    CheckSolverProgram = "Solver"
    'VB ghetto way of doing try-catch block
    On Error Resume Next
    Err.Clear
    'Will throw an error if OpenSolver is not loaded as a plugin
    If AddIns("OpenSolver").Installed Then
        CheckSolverProgram = "OpenSolver"
    End If
End Function
