Const TITLE_COL = 1

Sub AddUserToUsersSht(readmeNameCellRefString)
    Dim startingCell As range
    Dim newUserRow As Long

    With Worksheets(USERS_NAME)
        Set startingCell = .Columns(TITLE_COL).Find("Roles")
        
        If NotEmpty(startingCell) Then
            newUserRow = startingCell.Row + 1
            .Rows(newUserRow).Insert Shift:=xlShiftUp, CopyOrigin:=1
            .Cells(newUserRow, 1).Formula = readmeNameCellRefString 'hardcode
        End If
    End With
End Sub

Sub RemoveUserFromUsersSht(userName)
    Dim offendingCell As range
    
    With Worksheets(USERS_NAME)
        Set offendingCell = .Columns(TITLE_COL).Find(userName)
        
        If NotEmpty(offendingCell) Then
            offendingCell.EntireRow.Delete
        End If
    End With
End Sub

Sub AddRoleToUsersSht()
    Dim startingCell As range
    Dim newCol As Long

    With Worksheets(USERS_NAME)
        Set startingCell = .Columns(TITLE_COL).Find("Roles")
        newCol = CountRoles() + startingCell.Column
        InsertAndCopyColumn USERS_NAME, newCol
        'Clear inserted column of any previous data from the last column, may want to pull this out later
        .range(Col_Letter(newCol) & (TITLE_COL + 1) & ":" & Col_Letter(newCol) & (TITLE_COL + CountUsers())).ClearContents
    End With
End Sub

Sub RemoveRoleFromUsersSht(roleName)
    Dim startingCell As range
    Dim offendingRole As range

    With Worksheets(USERS_NAME)
        Set startingCell = .Columns(TITLE_COL).Find("Roles")
        Set offendingRole = .Rows(startingCell.Row).Find(roleName)
        
        If NotEmpty(offendingRole) Then
            offendingRole.EntireColumn.Delete
        End If
    End With
End Sub
