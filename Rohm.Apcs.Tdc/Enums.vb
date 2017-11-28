Public Enum RunModeType
    Normal = 0
    Separated = 1
    SeparatedEnd = 2
    ReRun = 3
End Enum

Public Enum EndModeType
    Normal = 1
    AbnormalEndReset = 2
    AbnormalEndAccumulate = 3
End Enum

''' <summary>
''' STATUS1
''' </summary>
''' <remarks></remarks>
Public Enum StartingCondition
    WaitForOperation = 0
    InProcess = 1
    InReservation = 2
    Down = 3
    ParallelRun = 4
    ParallelEnd = 5
End Enum

''' <summary>
''' STATUS2
''' </summary>
''' <remarks></remarks>
Public Enum LotCondition
    Normal = 0
    Warning = 2
    FinishShipOut = 8
    LotOut = 9
End Enum