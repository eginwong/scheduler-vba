Const ROLE_ROW As Integer = 4
Const STARTING_USER_ROW As Integer = 7
Const USER_COLUMN As String = "B"

Function CountRoles() As Long
    CountRoles = WorksheetFunction.CountA(Worksheets(CAPABILITIES_NAME).Rows(ROLE_ROW)) - 1 'subtract title of row
End Function

Function UniqueRole(roleName) As Boolean
    With Worksheets(CAPABILITIES_NAME)
        UniqueRole = Not NotEmpty(.Rows(ROLE_ROW).Find(roleName))
    End With
End Function

Sub AddUserToCapabilitiesSht(readmeNameCellRefString)
    With Worksheets(CAPABILITIES_NAME)
        .Rows(STARTING_USER_ROW).Insert Shift:=xlShiftUp, CopyOrigin:=1
        .range(USER_COLUMN & STARTING_USER_ROW).Formula = readmeNameCellRefString
    
        Dim i As Integer
        For i = 1 To CountRoles
            .Cells(STARTING_USER_ROW, i + 2) = 0 'offset for info columns
        Next i
    End With
End Sub

Sub RemoveUserFromCapabilitiesSht(userName)
    Dim cellWithuser As range
    Set cellWithuser = Worksheets(CAPABILITIES_NAME).Columns(USER_COLUMN).Find(userName, LookIn:=xlValues)
    If NotEmpty(cellWithuser) Then
        cellWithuser.EntireRow.Delete
    End If
End Sub

Function AddRoleToCapabilitiesSht(roleName)
    'update CAPABILITIES
    With Worksheets(CAPABILITIES_NAME)
        Dim newRoleCol As Long
        newRoleCol = 3 + CountRoles() 'offset for information columns
        .Columns(newRoleCol).Insert CopyOrigin:=xlFormatFromLeftOrAbove
        .Cells(ROLE_ROW, newRoleCol) = roleName
        .Cells(ROLE_ROW + 1, newRoleCol) = 1 'set default expected number to 1
        
        Dim i As Integer
        For i = 1 To CountUsers
            .Cells(STARTING_USER_ROW + i - 1, newRoleCol) = 0 'reset to 0 for new roles
        Next i
        
        AddRoleToCapabilitiesSht = "=" & CAPABILITIES_NAME & "!" & Col_Letter(newRoleCol) & ROLE_ROW
    End With
End Function

Sub RemoveRoleFromCapabilitiesSht(roleName)
    'update CAPABILITIES
    With Worksheets(CAPABILITIES_NAME)
        Dim roleRng As range
        Dim protectedRange As range
        Set protectedRange = .range("A" & ROLE_ROW)
        Set roleRng = .Rows(ROLE_ROW).Find(What:=roleName, _
                                                                    After:=protectedRange, _
                                                                   SearchOrder:=xlByColumns)
                                                                   
        If NotEmpty(roleRng) And Not CellEquality(roleRng, protectedRange) Then
            .Columns(roleRng.Column).EntireColumn.Delete
        End If
    End With
End Sub

