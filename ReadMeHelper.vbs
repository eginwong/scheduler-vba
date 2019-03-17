Const STARTING_USER_COL = 6
Const USER_ROW = 9

Function CountUsers() As Long
    CountUsers = WorksheetFunction.CountA(Worksheets(README_NAME).Columns(STARTING_USER_COL))
End Function

Function UniqueUser(userName) As Boolean
    With Worksheets(README_NAME)
        UniqueUser = Not NotEmpty(.Columns(STARTING_USER_COL).Find(userName))
    End With
End Function

'Adds input username to README sheet and returns the cell reference as a String
Function AddNewUserToReadMeSht(userName)
    'insert into README to show completion
    With Worksheets(README_NAME)
        .Rows(USER_ROW).Insert Shift:=xlShiftUp, CopyOrigin:=1
        .range(Col_Letter(STARTING_USER_COL) & USER_ROW) = userName
        .range(Col_Letter(STARTING_USER_COL + 1) & USER_ROW) = True
        AddNewUserToReadMeSht = "=" & README_NAME & "!" & Col_Letter(STARTING_USER_COL) & USER_ROW
    End With
End Function

Sub RemoveUserFromReadMeSht(userName)
    Dim readMeCellWithUser As range
    With Worksheets(README_NAME)
        Set readMeCellWithUser = .Columns(STARTING_USER_COL).Find(userName)
        If NotEmpty(readMeCellWithUser) Then
            readMeCellWithUser.EntireRow.Delete
        End If
    End With
End Sub

Function GetScheduleDate() As Date
    GetScheduleDate = Worksheets(README_NAME).range("NEXT_DATE").Value
End Function

Sub UpdateNextDateScheduled()
    With Worksheets(README_NAME)
        'Update next_date for the following schedule
        .range("NEXT_DATE") = DateAdd(MapFrequencyToVBAFormat(.range("FREQUENCY").Value), 1, .range("NEXT_DATE").Value)
    End With
End Sub

'Receive array of input, iterate through and toggle availability until empty
Sub ToggleUnavailable()
    ToggleAvailability (False)
End Sub

Sub ToggleAvailable()
    ToggleAvailability (True)
End Sub

Sub ToggleAvailability(SwitchOn As Boolean)
    Dim Rng As range
    With Worksheets(README_NAME)
        For Each user In RetrieveVacationUsers()
            Set Rng = .Columns(STARTING_USER_COL).Find(What:=user, LookIn:=xlValues)
            If NotEmpty(Rng) Then
                .Cells(Rng.Row, Rng.Column + 1).Value = SwitchOn
            End If
        Next user
    End With
End Sub
