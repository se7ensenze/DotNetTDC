Module ModuleGlobal

    '********* STATUS1 is mean to "Start condition" ***********
    Friend Const CODE_STATUS1_WAIT_FOR_OPERATION As String = "0" 'wait for operation
    Friend Const CODE_STATUS1_IN_PROCESS As String = "1"
    Friend Const CODE_STATUS1_IN_RESERVATION As String = "2"
    Friend Const CODE_STATUS1_DOWN As String = "3"
    Friend Const CODE_STATUS1_PARALLEL_RUN As String = "4"
    Friend Const CODE_STATUS1_PARALLEL_END As String = "5"
    '**********************************************************

    '********* STATUS2 is mean to "Lot condition" ***********
    Friend Const CODE_STATUS2_NORMAL As String = "0" 'wait for operation
    Friend Const CODE_STATUS2_WARNING As String = "2"
    Friend Const CODE_STATUS2_FINISH_SHIPING As String = "8"
    Friend Const CODE_STATUS2_LOT_OUT As String = "9"
    '**********************************************************

    Friend Const CODE_MACHINE_SUB_LOT_RUNNING As String = "0"
    Friend Const CODE_MACHINE_SUB_LOT_END As String = "1"
    Friend Const CODE_MACHINE_SUB_ABNORMAL_RESET As String = "2"
    Friend Const CODE_MACHINE_SUB_ABNORMAL_ACCUM As String = "3"

End Module
