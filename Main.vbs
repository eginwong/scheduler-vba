Public Const CAPABILITIES_NAME As String = "CAPABILITIES"
Public Const ENGINE_NAME As String = "ENGINE"
Public Const README_NAME As String = "README"
Public Const SCHEDULE_NAME As String = "SCHEDULE"
Public Const USERS_NAME As String = "USERS"
Public Const UNAVAILABILITY_NAME As String = "UNAVAILABILITY"

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
        MsgBox ("User does not exist")
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub GENERATE_SCHEDULE()
    Application.ScreenUpdating = False
    Dim sourceSheet As Worksheet
    Set sourceSheet = ActiveSheet
    
    Call ToggleUnavailable
    
    'Need to be in ENGINE worksheet to run Solver
    Worksheets(ENGINE_NAME).Activate
    Dim solverResults As Boolean
    solverResults = SolveSchedule
    
    Call ToggleAvailable
    
    If solverResults Then
        Call PrintSchedule
        Call UpdateNextDateScheduled
        Call sourceSheet.Activate
    Else
        MsgBox ("Schedule is impossible given the number of roles and available people")
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

