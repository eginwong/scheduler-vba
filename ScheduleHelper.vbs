Const USER_ROLE_COL = 1
Const ROLE_ROW = 4
Sub InsertBlankSchedule()
    With Worksheets(SCHEDULE_NAME)
        .Rows(3).Insert Shift:=xlShiftUp, CopyOrigin:=1 'don't know how to insert multiple rows
        .Rows(3).Insert Shift:=xlShiftUp, CopyOrigin:=1
        .Rows(3).Insert Shift:=xlShiftUp, CopyOrigin:=1
    End With
End Sub

Sub PrintSchedule()
    Call InsertBlankSchedule
    With Worksheets(SCHEDULE_NAME)
        Dim scheduleDate As Date
        scheduleDate = Worksheets(README_NAME).range("NEXT_DATE").Value
        .Cells(4, 1) = scheduleDate
        
        Dim roleCount As Long
        roleCount = CountRoles
    
        Dim i As Integer
        Dim expectedRow As range
        Set expectedRow = Worksheets(ENGINE_NAME).range("EXPECTED")
        
        For i = 1 To roleCount
            Dim roleName As String
            roleName = Worksheets(ENGINE_NAME).Cells(ROLE_ROW, i + 1).Value
            .Cells(3, i + 1) = roleName 'offset for info columns
            Dim roleLimit As Integer
             'find the expected num of that role required
            roleLimit = Worksheets(ENGINE_NAME).Cells(expectedRow.Row, i + 3 + roleCount).Value
            
            Dim searchIteration As Integer
            Dim Rng As range
            searchIteration = 1
            Set rngLookin = Worksheets(ENGINE_NAME).Columns(i + 3 + roleCount)
            Set rngStart = .Cells(1, i + 3 + roleCount)
            Do
                Set Rng = rngLookin.Find(What:=1, _
                                     After:=rngStart, _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext)
                If NotEmpty(Rng) Then
                    Set rngStart = Rng
                    Dim cellForRole As range
                    Set cellForRole = .Cells(4, i + 1)
                    If cellForRole <> "" Then
                        cellForRole = cellForRole.Value & ", "
                    End If
                    Dim user As String
                    user = Worksheets(ENGINE_NAME).Cells(Rng.Row, USER_ROLE_COL).Value
                    cellForRole = cellForRole.Value & user
                    
                    'Update last run for each user
                    'Need user, role, and date
                    Dim wsUserSearch As range
                    Set wsUserSearch = Worksheets(USERS_NAME).Columns(USER_ROLE_COL).Find(What:=user)
                    
                    If NotEmpty(wsUserSearch) Then
                        Worksheets(USERS_NAME).Cells(wsUserSearch.Row, i + 1) = scheduleDate
                    End If
                    
                    searchIteration = searchIteration + 1
                End If
            Loop Until searchIteration = (roleLimit + 1) Or Rng Is Nothing
        Next i
    End With
End Sub


