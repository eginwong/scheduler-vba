Const STARTING_ROW = 5
Const DATE_COL = 2

Sub ADD_UNAVAILABILITY_RECORD()
    Application.ScreenUpdating = False
    
    Dim userName As String
    Dim vacationDate As Date
    'Retrieve input
    userName = InputWithExit("What is the name of the user with unavailability?", "Add Unavailability Record")
    vacationDate = InputWithExit("When are they unavailable (yyyy-mm-dd)? ", "Add Unavailability Record")

    'Check if user already exists
    If Not UniqueUser(userName) Then
        With Worksheets(UNAVAILABILITY_NAME)
            .Rows(STARTING_ROW).Insert Shift:=xlShiftUp, CopyOrigin:=1
            .range("A" & STARTING_ROW).Value = userName
            .range(Col_Letter(DATE_COL) & STARTING_ROW).Value = vacationDate
            Call SortRecords
        End With
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub REMOVE_OLD_UNAVAILABILITY_RECORDS()
    Application.ScreenUpdating = False
    
    Dim scheduleDate As Date
    scheduleDate = GetScheduleDate
    
    'Sort by oldest
    With Worksheets(UNAVAILABILITY_NAME)
        Call SortRecords
    
        Dim recordPointer As range
        Set recordPointer = .range(Col_Letter(DATE_COL) & STARTING_ROW)
        
        'Remove old rows
        Do While recordPointer.Value < scheduleDate
            recordPointer.EntireRow.Delete
            Set recordPointer = .range(Col_Letter(DATE_COL) & STARTING_ROW)
        Loop
    End With
    
    Application.ScreenUpdating = True
End Sub

'Check to see if any of the vacation records match anyone
'If so, return a collection with their name populated
'If not, return an empty collection
Function RetrieveVacationUsers() As Collection
    Dim coll As New Collection
    
    Dim scheduleDate As Date
    scheduleDate = GetScheduleDate()
    
    Dim Rng As range
    Dim rngStart As range
    With Worksheets(UNAVAILABILITY_NAME)
        Set Rng = .Columns(DATE_COL).Find(What:=scheduleDate, _
                                                LookIn:=xlFormulas, _
                                                LookAt:=xlWhole, _
                                                SearchOrder:=xlByRows, _
                                                SearchDirection:=xlNext)
        If NotEmpty(Rng) Then
            Set rngStart = Rng
            Do
                coll.Add .Cells(Rng.Row, 1).Value
                Set Rng = .Columns(DATE_COL).FindNext(Rng)
            Loop While NotEmpty(Rng) And Rng.Row > rngStart.Row
        End If
    End With
    Set RetrieveVacationUsers = coll

End Function

Private Sub SortRecords()
    With Worksheets(UNAVAILABILITY_NAME)
        .range("A" & STARTING_ROW).CurrentRegion.Sort key1:=.range("B" & STARTING_ROW), _
            order1:=xlAscending, Header:=xlYes
    End With
End Sub
