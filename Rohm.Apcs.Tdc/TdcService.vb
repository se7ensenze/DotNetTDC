Imports System.Data.SqlClient
Imports System.IO

Public Class TdcService

    Private Shared m_Instance As TdcService

    Public Shared Function GetInstance() As TdcService
        If m_Instance Is Nothing Then
            m_Instance = New TdcService()
        End If
        Return m_Instance
    End Function

    Private m_TdcDB As TdcDatabase

    Private Sub New()
        m_TdcDB = New TdcDatabase()
        m_Logger = New TdcLoggerTextWriter()
    End Sub

    Private m_Logger As ITdcLogger
    Public Property Logger() As ITdcLogger
        Get
            Return m_Logger
        End Get
        Set(ByVal value As ITdcLogger)
            m_Logger = value
        End Set
    End Property


    Public Property ConnectionString() As String
        Get
            Return m_TdcDB.ConnectionString
        End Get
        Set(ByVal value As String)
            m_TdcDB.ConnectionString = value
        End Set
    End Property


    Private Function DoCommonOperation(ByVal con As SqlConnection, ByVal machineNo As String, ByVal lotNo As String, ByRef mc As APCSDBDataSet.MACHI_TABLERow, ByRef wk As APCSDBDataSet.WorkDataRow, ByVal ret As TdcResponse) As Boolean

        '1.) try to open database
        Try
            con.Open() 'open database
        Catch ex As Exception
            ret.ErrorCode = "70"
            ret.ErrorMessage = "Cannot open connection to database"
            Return False 'return and exists sub
        End Try

        '2.) try to retrieve machine's data
        Dim machineTable As APCSDBDataSet.MACHI_TABLEDataTable

        Try
            machineTable = m_TdcDB.GetMachine(con, machineNo)
        Catch sx As SqlException
            ManageSqlExceptionMessage(sx, ret)
            Return False
        Catch ex As Exception
            ret.ErrorCode = "99"
            ret.ErrorMessage = "Exception:" & ex.Message
            Return False
        End Try

        Dim mcRow As APCSDBDataSet.MACHI_TABLERow = Nothing
        If machineTable.Rows.Count = 1 Then
            mcRow = CType(machineTable.Rows(0), APCSDBDataSet.MACHI_TABLERow)
        Else
            'machine not found
            ret.ErrorCode = "04"
            ret.ErrorMessage = "Machine not found"
            Return False 'return and exists sub
        End If

        '3.) try to retrieve lot data that were combine from LOT1_TABLE and LOT1_DATA and LAYER_TABLE
        Dim workDataTable As APCSDBDataSet.WorkDataDataTable = Nothing
        Try
            workDataTable = m_TdcDB.GetWorkData(con, lotNo)
        Catch sx As SqlException
            ManageSqlExceptionMessage(sx, ret)
            Return False
        Catch ex As Exception
            ret.ErrorCode = "99"
            ret.ErrorMessage = "Exception:" & ex.Message
            Return False
        End Try

        If workDataTable.Rows.Count = 1 Then
            wk = CType(workDataTable.Rows(0), APCSDBDataSet.WorkDataRow)
            If wk.N_OPE_SEQ = 0 Then
                ret.ErrorCode = "99"
                ret.ErrorMessage = "Data Flow error:N_OPE_SEQ is 0"
                Return False
            End If
        Else
            'check row from LOT1_TABLE only
            Dim dt As APCSDBDataSet.LOT1_TABLEDataTable
            Try
                dt = m_TdcDB.GetLot1Table(con, lotNo)
            Catch sx As SqlException
                ManageSqlExceptionMessage(sx, ret)
                Return False
            Catch ex As Exception
                ret.ErrorCode = "99"
                ret.ErrorMessage = "Exception:" & ex.Message
                Return False
            End Try

            Using dt
                If dt.Rows.Count = 1 Then
                    Dim r As APCSDBDataSet.LOT1_TABLERow = CType(dt.Rows(0), APCSDBDataSet.LOT1_TABLERow)
                    If r.OPE_SEQ = 0 Then
                        ret.ErrorCode = "01"
                        ret.ErrorMessage = "Lot " & lotNo & "'s OPE_SEQ = 0"
                        Return False 'return and exists sub
                    End If
                Else
                    'lot no longer exists
                    ret.ErrorCode = "01"
                    ret.ErrorMessage = "Not found " & lotNo & " in LOT1_TABLE"
                    Return False 'return and exists sub
                End If
            End Using
        End If

        '4) check machine's Process_Type and current process of lot data
        If wk.IsOPE_NAMENull() Then
            ret.ErrorCode = "06"
            ret.ErrorMessage = "WORK's OPE_NAME is NULL"
            Return False
        End If

        If mcRow.IsPROCESS_TYPENull() Then
            ret.ErrorCode = "06"
            ret.ErrorMessage = "MACHINE's PROCESS_TYPE is NULL"
            Return False
        End If

        If Not CheckProcessType(mcRow.PROCESS_TYPE, wk.OPE_NAME) Then
            ret.ErrorCode = "06"
            ret.ErrorMessage = "Error process machine process := " & mcRow.PROCESS_TYPE & " but work process := " & wk.OPE_NAME  'error process
            Return False 'return and exists sub
        End If

        Return True 'if the can reach this section it's mean to complete operation

    End Function

    Private SPLITED_STRING_ARRAY_PROCESSTYPE() As String = {"|"}

    Private Function CheckProcessType(ByVal machineProcessType As String, ByVal workProcessType As String) As Boolean
        Dim isCurrectProcess As Boolean = False
        Dim splitedProcesses() As String = machineProcessType.Split(SPLITED_STRING_ARRAY_PROCESSTYPE, StringSplitOptions.RemoveEmptyEntries)
        Dim splitedProcessesCount As Integer = splitedProcesses.GetUpperBound(0)
        If splitedProcessesCount >= 0 Then
            Dim process As String
            For i As Integer = 0 To splitedProcessesCount
                process = splitedProcesses(i).Trim()
                If process = "*" Then
                    isCurrectProcess = True
                    Exit For
                ElseIf process.EndsWith("*") Then
                    'remove last charactor
                    Dim tmp1 As String = Mid$(process, 1, Len(process) - 1)
                    'compare "AUTO" with "AUTO1"
                    If workProcessType.StartsWith(tmp1) Then
                        isCurrectProcess = True
                        Exit For
                    End If
                Else
                    If process = workProcessType Then
                        isCurrectProcess = True
                        Exit For
                    End If
                End If
            Next i
        End If
        Return isCurrectProcess
    End Function

    Private Sub PopulateLOT2_TABLERow(ByVal wk As APCSDBDataSet.WorkDataRow, ByVal row As APCSDBDataSet.LOT2_TABLERow)
        'เอามาเฉพาะฟิลด์ที่ไม่เป้นว่างปล่าว
        row.LOT_NO = wk.LOT_NO
        row.DVI_NO = wk.DVI_NO
        row.PRD_NAME = wk.PRD_NAME
        row.IN_DAY = wk.IN_DAY
        row.OUT_DAY = wk.OUT_DAY
        row.OPE_SEQ = wk.OPE_SEQ
        row.PRD_PIECE = wk.PRD_PIECE
        row.INP_PIECE = wk.INP_PIECE
        row.OUT_PIECE = wk.OUT_PIECE
        If Not wk.IsREAL_DAYNull() Then
            row.REAL_DAY = wk.REAL_DAY
        End If
        If Not row.IsMATER_NAMENull() Then
            row.MATER_NAME = wk.MATER_NAME
        End If
        If Not row.IsMATER_SNAMENull() Then
            row.MATER_SNAME = wk.MATER_SNAME
        End If
        row.Y_LEVEL = wk.Y_LEVEL
        row.STATUS1 = wk.STATUS1
        row.STATUS2 = wk.STATUS2
        row.CYCLE = wk.CYCLE
        If Not wk.IsWAFER_STATUSNull() Then
            row.WAFER_STATUS = wk.WAFER_STATUS
        End If
        row.REP_OPE_SEQ = wk.REP_OPE_SEQ
        row.REPEAT_QNT = wk.REPEAT_QNT
        row.SISAKU_FLG = wk.SISAKU_FLG
        If Not wk.IsROHM_ORDER_MODEL_NAME_ONull() Then
            row.ROHM_ORDER_MODEL_NAME_O = wk.ROHM_ORDER_MODEL_NAME_O
        End If
        If Not wk.IsORDER_NONull() Then
            row.ORDER_NO = wk.ORDER_NO
        End If
        If Not wk.IsFT_MODEL_NAMENull() Then
            row.FT_MODEL_NAME = wk.FT_MODEL_NAME
        End If
        If Not wk.IsTP_RANKNull() Then
            row.TP_RANK = wk.TP_RANK
        End If
        If Not wk.IsWARI_STOCK_KBNNull() Then
            row.WARI_STOCK_KBN = wk.WARI_STOCK_KBN
        End If
        If Not wk.IsWARI_INSTRUCT_KBNNull() Then
            row.WARI_INSTRUCT_KBN = wk.WARI_INSTRUCT_KBN
        End If
        If Not wk.IsFORM_NAMENull() Then
            row.FORM_NAME = wk.FORM_NAME
        End If
        row.GOOD_PIECES = wk.GOOD_PIECES
        row.BAD_PIECES = wk.BAD_PIECES
        If Not wk.IsCREATION_DATENull() Then
            row.CREATION_DATE = wk.CREATION_DATE
        End If
        If wk.IsSEND_FLGNull() Then
            row.SetSEND_FLGNull()
        Else
            row.SEND_FLG = wk.SEND_FLG
        End If
        If wk.IsSEND_DATENull() Then
            row.SetSEND_DATENull()
        Else
            row.SEND_DATE = wk.SEND_DATE
        End If
    End Sub

    Private Sub PopulateLOT2_DATARow(ByVal wk As APCSDBDataSet.WorkDataRow, ByVal row As APCSDBDataSet.LOT2_DATARow)
        'เอามาเฉพาะฟิลด์ที่ไม่เป้นว่างปล่าว
        row.LOT_NO = wk.LOT_NO
        row.OPE_SEQ = wk.OPE_SEQ
        row.N_OPE_SEQ = wk.N_OPE_SEQ
        row.LAY_NO = wk.LAY_NO
        If wk.IsPLAN_DAYNull() Then
            row.SetPLAN_DAYNull()
        Else
            row.PLAN_DAY = wk.PLAN_DAY
        End If
        'row.REAL_START = 
        'row.MACHINE = 
        row.MACHINE_SUB = "0" 'ขอยืมฟิลด์นี้ในการเก็บข้อมูล End Mode 0 = ยังไม่จบ
        If wk.IsSTART_MSGNull() Then
            row.SetSTART_MSGNull()
        Else
            row.START_MSG = wk.START_MSG
        End If
        If wk.IsEND_MSGNull() Then
            row.SetSTART_MSGNull()
        Else
            row.END_MSG = wk.END_MSG
        End If
        'row.OPERATOR1 = 
        'row.OPERATOR2 = 
        row.QUANTITY = 0
        row.LOSS_QTY = 0
        row.DATA_NO = 0
        row.LIMIT_FLG = "1"
        row.LIMIT_TIME1 = 0
        row.LIMIT_FLG1 = "0"
        row.LIMIT_TIME2 = 0
        row.LIMIT_FLG2 = "0"
        row.INTEG1 = 0
        row.INTEG2 = 0
        row.INTEG3 = 0
        row.INTEG4 = 0
        row.REPEAT_FLG = "0"
        row.REPEAT_TIME = 0
        row.WAFER_STATUS = "0000000000000000000000000"
        row.PRD_NAME = wk.PRD_NAME
        row.GOOD_PIECES = 0
        row.BAD_PIECES = 0
        If wk.IsSEND_DATENull() Then
            row.SetSEND_DATENull()
        Else
            row.SEND_DATE = wk.SEND_DATE
        End If

    End Sub

    Public Function LotRequest(ByVal machineNo As String, ByVal lotNo As String, ByVal runMode As RunModeType) As TdcLotRequestResponse

        Dim ret As TdcLotRequestResponse = New TdcLotRequestResponse()
        ret.MCNo = machineNo
        ret.LotNo = lotNo
        ret.HasError = True

        Dim msg As TdcMessage = New TdcMessage("LotRequest", machineNo, lotNo, String.Empty)
        msg.Parameter = "( MCNo:" & machineNo & " LotNo:" & lotNo & " RunMode:" & runMode.ToString() & ")"

        Using con As SqlConnection = m_TdcDB.CreateConnection()
            Dim machineRow As APCSDBDataSet.MACHI_TABLERow = Nothing
            Dim wk As APCSDBDataSet.WorkDataRow = Nothing
            If DoCommonOperation(con, machineNo, lotNo, machineRow, wk, ret) Then

                '1.)reserve or down
                If wk.STATUS1 = CODE_STATUS1_IN_RESERVATION Then
                    ret.ErrorCode = "05"
                    ret.ErrorMessage = "Lot status is Reserve"
                    GoTo LBL_001
                ElseIf wk.STATUS1 = CODE_STATUS1_DOWN Then
                    ret.ErrorCode = "05"
                    ret.ErrorMessage = "Lot Status is Down"
                    GoTo LBL_001
                ElseIf wk.STATUS1 = CODE_STATUS1_IN_PROCESS Then
                    If runMode <> RunModeType.ReRun Then
                        ret.ErrorCode = "02"
                        ret.ErrorMessage = "Running"
                        GoTo LBL_001
                    Else
                        Using tbl2 As APCSDBDataSet.LOT2_DATADataTable = m_TdcDB.GetLot2DataByLotNoOpeSeq(con, wk.LOT_NO, wk.OPE_SEQ)
                            Dim lotEndCount As Integer = 0
                            For Each row As APCSDBDataSet.LOT2_DATARow In tbl2.Rows
                                If Not row.IsREAL_DAYNull() Then
                                    lotEndCount += 1
                                End If
                            Next
                            If lotEndCount <> tbl2.Rows.Count Then
                                ret.ErrorCode = "02"
                                ret.ErrorMessage = "Normal Running is not finish"
                                GoTo LBL_001
                            End If
                        End Using
                    End If
                ElseIf wk.STATUS1 = CODE_STATUS1_PARALLEL_RUN AndAlso (runMode <> RunModeType.Separated AndAlso runMode <> RunModeType.SeparatedEnd) Then 'parallel and parallel end
                    ret.ErrorCode = "02"
                    ret.ErrorMessage = "Running Parallel"
                    GoTo LBL_001
                ElseIf wk.STATUS1 = CODE_STATUS1_PARALLEL_END AndAlso (runMode <> RunModeType.ReRun) Then
                    '2016-10-26 
                    '*** In case of there was separation of lot
                    '*** the abnormal end should be able to run again as Re-Run
                    ret.ErrorCode = "02"
                    ret.ErrorMessage = "Running Parallel End"
                    GoTo LBL_001
                Else
                    If wk.IsSTART_MSGNull() Then
                        ret.Message = ""
                    Else
                        ret.Message = wk.START_MSG
                    End If
                    ret.HasError = False
                    ret.LotNo = wk.LOT_NO
                    ret.GoodPieces = wk.GOOD_PIECES
                    ret.BadPieces = wk.BAD_PIECES
                End If
            End If
LBL_001:
            msg.CodePosition = "LOTREQUEST-0001"
            If wk IsNot Nothing Then
                msg.Status1 = wk.STATUS1
                msg.Status2 = wk.STATUS2
            End If
            If ret.HasError Then
                msg.MessageType = "ERROR"
                msg.MessageText = ret.ErrorCode & ":" & ret.ErrorMessage
            Else
                msg.MessageType = "REPORT"
                msg.MessageText = "Completed"
            End If
            m_Logger.SaveLog(msg)

            Return ret

        End Using

        Return ret
    End Function

    Public Function LotSet(ByVal machineNo As String, ByVal lotNo As String, ByVal startTime As DateTime, ByVal opNo As String, ByVal runMode As RunModeType) As TdcResponse

        'ไม่เอาหน่วยมิลลิวินาที
        startTime = startTime.AddMilliseconds(-startTime.Millisecond)

        Dim ret As TdcResponse = New TdcResponse()
        'by default >> HasError = true
        ret.MCNo = machineNo
        ret.LotNo = lotNo
        ret.HasError = True

        Dim msg As TdcMessage = New TdcMessage("LotSet", machineNo, lotNo, opNo)
        msg.Parameter = "( MCNo:" & machineNo & " LotNo:" & lotNo & " StartTime:" & startTime.ToString("yyyy/MM/dd HH:mm:ss") & " OPNo:" & opNo & " RunMode:" & runMode.ToString() & ")"

        Using con As SqlConnection = m_TdcDB.CreateConnection()

            Dim wk As APCSDBDataSet.WorkDataRow = Nothing
            Dim machineRow As APCSDBDataSet.MACHI_TABLERow = Nothing

            If DoCommonOperation(con, machineNo, lotNo, machineRow, wk, ret) Then

                Dim newStatus1 As String = wk.STATUS1
                Dim newStatus2 As String = wk.STATUS2
                Dim parallelLotCount As Short = wk.CYCLE

                Select Case wk.STATUS1
                    Case CODE_STATUS1_WAIT_FOR_OPERATION
                        If runMode = RunModeType.Normal Then
                            newStatus1 = CODE_STATUS1_IN_PROCESS
                            parallelLotCount = 1S
                        ElseIf runMode = RunModeType.Separated Then
                            newStatus1 = CODE_STATUS1_PARALLEL_RUN
                            parallelLotCount = 1S
                        ElseIf runMode = RunModeType.SeparatedEnd Then
                            ret.ErrorCode = "05"
                            ret.ErrorMessage = "Can not start Parallel End lot before Parallel lot"
                            GoTo LBL_001
                        ElseIf runMode = RunModeType.ReRun Then
                            ret.ErrorCode = "05"
                            ret.ErrorMessage = "Can not start as Re-Run"
                            GoTo LBL_001
                        End If
                    Case CODE_STATUS1_IN_PROCESS
                        If runMode <> RunModeType.ReRun Then
                            ret.ErrorCode = "02"
                            ret.ErrorMessage = "Running"
                            GoTo LBL_001
                        Else
                            Using tbl2 As APCSDBDataSet.LOT2_DATADataTable = m_TdcDB.GetLot2DataByLotNoOpeSeq(con, wk.LOT_NO, wk.OPE_SEQ)
                                Dim lotEndCount As Integer = 0
                                For Each row As APCSDBDataSet.LOT2_DATARow In tbl2.Rows
                                    If Not row.IsREAL_DAYNull() Then
                                        lotEndCount += 1
                                    End If
                                Next
                                If lotEndCount <> tbl2.Rows.Count Then
                                    ret.ErrorCode = "02"
                                    ret.ErrorMessage = "Normal Running is not finish"
                                    GoTo LBL_001
                                End If
                            End Using
                        End If
                    Case CODE_STATUS1_IN_RESERVATION
                        ret.ErrorCode = "05"
                        ret.ErrorMessage = "Lot status is Reserve"
                        GoTo LBL_001
                    Case CODE_STATUS1_DOWN
                        ret.ErrorCode = "05"
                        ret.ErrorMessage = "Lot status is Down"
                        GoTo LBL_001
                    Case CODE_STATUS1_PARALLEL_RUN
                        If runMode = RunModeType.Normal Then
                            ret.ErrorCode = "02"
                            ret.ErrorMessage = "Can not start with Normal mode because Lot status is Parallel running"
                            GoTo LBL_001
                        ElseIf runMode = RunModeType.Separated Then
                            'allow to run parallel and no need to update LOT1_TABLE and LOT2_TABLE
                            'update_LOT1_LOT2_TABLE = False
                            parallelLotCount += 1S
                        ElseIf runMode = RunModeType.SeparatedEnd Then
                            'allow to run parallel end and change status to 5
                            parallelLotCount += 1S
                            newStatus1 = "5"
                        ElseIf runMode = RunModeType.ReRun Then
                            ret.ErrorCode = "02"
                            ret.ErrorMessage = "Can not start with Re-Run mode because Lot status is Parallel running"
                            GoTo LBL_001
                        End If
                    Case CODE_STATUS1_PARALLEL_END  'Parallel End
                        If runMode <> RunModeType.ReRun Then
                            ret.ErrorCode = "02"
                            ret.ErrorMessage = "Can not start with Normal mode because Lot status is Parallel end running"
                            GoTo LBL_001
                        End If
                End Select

                Dim lot1Table As APCSDBDataSet.LOT1_TABLEDataTable
                Try
                    lot1Table = m_TdcDB.GetLot1Table(con, lotNo)
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot1TableRow As APCSDBDataSet.LOT1_TABLERow = CType(lot1Table.Rows(0), APCSDBDataSet.LOT1_TABLERow)

                Dim lot2Table As APCSDBDataSet.LOT2_TABLEDataTable
                Try
                    lot2Table = m_TdcDB.GetLot2Table(con, lotNo)
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot2TableRow As APCSDBDataSet.LOT2_TABLERow = Nothing

                'create new if not exists
                If lot2Table.Rows.Count = 0 Then
                    Dim newLot2TableRow As APCSDBDataSet.LOT2_TABLERow = lot2Table.NewLOT2_TABLERow()
                    PopulateLOT2_TABLERow(wk, newLot2TableRow)
                    lot2Table.Rows.Add(newLot2TableRow)
                End If
                lot2TableRow = CType(lot2Table.Rows(0), APCSDBDataSet.LOT2_TABLERow)

                'if first process update REAL_START time
                Try
                    If m_TdcDB.IsFirstProcess(con, lot1TableRow.LOT_NO, lot1TableRow.OPE_SEQ) Then
                        lot1TableRow.REAL_START = startTime
                        lot2TableRow.REAL_START = startTime
                    End If
                Catch ex As Exception
                    ret.ErrorCode = "01"
                    ret.ErrorMessage = "Could not find first OPE_SEQ of LotNo = " & lot1TableRow.LOT_NO
                    GoTo LBL_001
                End Try

                'try to check LOT2_DATA by PK
                'ต้องเชื่อใจว่า คนละเครื่องกันไม่มีทาง ผลิตงานที่เวลาเดียวกัน PK { REAL_START, LOT_NO }
                Dim lot2Data As APCSDBDataSet.LOT2_DATADataTable
                Try
                    lot2Data = m_TdcDB.GetLot2DataOnLotSet(con, startTime, wk.LOT_NO) 'single row
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot2DataRow As APCSDBDataSet.LOT2_DATARow = Nothing
                'create new if not exists
                If lot2Data.Rows.Count = 0 Then
                    lot2DataRow = lot2Data.NewLOT2_DATARow()
                    lot2DataRow.REAL_START = startTime
                    lot2DataRow.MACHINE = machineNo
                    PopulateLOT2_DATARow(wk, lot2DataRow)
                    lot2Data.Rows.Add(lot2DataRow)
                Else
                    lot2DataRow = CType(lot2Data.Rows(0), APCSDBDataSet.LOT2_DATARow)
                End If

                Dim goodPieces As Integer = 0
                Dim badPieces As Integer = 0

                'pass GOOD_PIECES and BAD_PIECES as refernce type
                Dim tran As SqlTransaction
                Try
                    tran = con.BeginTransaction()
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try


                Try
                    'ต้องลบ ข้อมูลที่เป็นของเครื่องนี้ แต่ไม่มี REAL_DAY (EndTime) 
                    Dim deletedCount As Integer = m_TdcDB.DeleteNotFinishLOT2_DATAOf(con, tran, lotNo, machineNo)
                    If parallelLotCount > 0 AndAlso (runMode = RunModeType.Separated OrElse runMode = RunModeType.SeparatedEnd) Then
                        parallelLotCount -= CShort(deletedCount)
                    End If

                    'ลบข้อมูลของ LOT2_DATA ของโปรเซสปัจจุบันและถัดไปทั้งหมด ... ในกรณีมีการ Move Lot
                    If wk.STATUS1 = CODE_STATUS1_WAIT_FOR_OPERATION Then
                        m_TdcDB.DeleteLOT2_DATACurrentAndAfterOPE_SEQ(con, tran, lotNo, wk.OPE_SEQ)
                    End If

                    lot2DataRow.OPERATOR1 = opNo
                    lot2DataRow.SetOPERATOR2Null() 'OPERATOR2 = NULL
                    lot2DataRow.REAL_START = startTime
                    lot2DataRow.GOOD_PIECES = 0
                    lot2DataRow.BAD_PIECES = 0
                    lot2DataRow.MACHINE_SUB = CODE_MACHINE_SUB_LOT_RUNNING 'ขอยืมฟิลด์เก็บ End Mode
                    m_TdcDB.UpdateLot2Data(con, tran, lot2Data) 'UPDATE LOT2_DATA

                    If m_TdcDB.SumOutputOfLOT2_DATA(con, tran, wk.LOT_NO, wk.LAY_NO, goodPieces, badPieces) Then 'this is TOS additional for correct data in LOT1_TABLE, LOT2_TABLE

                        Dim updateTime As Date = Now
                        updateTime = updateTime.AddMilliseconds(-updateTime.Millisecond)

                        lot1TableRow.GOOD_PIECES = goodPieces 'this is TOS additional for correct data in LOT1_TABLE, LOT2_TABLE
                        lot1TableRow.BAD_PIECES = badPieces 'this is TOS additional for correct data in LOT1_TABLE, LOT2_TABLE
                        lot1TableRow.STATUS1 = newStatus1
                        lot1TableRow.CREATION_DATE = updateTime
                        lot1TableRow.CYCLE = parallelLotCount '2016-10-26 to count parallel lot and check for move lot to next process
                        m_TdcDB.UpdateLot1Table(con, tran, lot1Table) 'UPDATE LOT1_TABLE

                        lot2TableRow.GOOD_PIECES = goodPieces 'this is TOS additional for correct data in LOT1_TABLE, LOT2_TABLE
                        lot2TableRow.BAD_PIECES = badPieces 'this is TOS additional for correct data in LOT1_TABLE, LOT2_TABLE
                        lot2TableRow.STATUS1 = newStatus1
                        lot2TableRow.CREATION_DATE = updateTime
                        lot2TableRow.CYCLE = parallelLotCount '2016-10-26 to count parallel lot and check for move lot to next process

                        m_TdcDB.UpdateLot2Table(con, tran, lot2Table) 'UPDATE LOT2_TABLE

                    End If

                    tran.Commit()

                    If wk.IsSTART_MSGNull() Then
                        ret.Message = ""
                    Else
                        ret.Message = wk.START_MSG
                    End If

                    msg.Status1 = newStatus1
                    msg.Status2 = newStatus2

                    ret.HasError = False

                Catch sx As SqlException
                    Try
                        tran.Rollback()
                        ManageSqlExceptionMessage(sx, ret)
                    Catch ex2 As Exception
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = "Rollback:" & sx.Message & vbNewLine & ", ConnectionState :=" & con.State.ToString()
                    End Try
                Catch ex As Exception
                    Try
                        tran.Rollback()
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = ex.Message
                    Catch ex2 As Exception
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = "Rollback:" & ex.Message & vbNewLine & ", ConnectionState :=" & con.State.ToString()
                    End Try
                Finally
                    lot1Table.Dispose()
                    lot1Table = Nothing
                    lot2Table.Dispose()
                    lot2Table = Nothing
                    lot2Data.Dispose()
                    lot2Data = Nothing
                End Try
            End If

        End Using

LBL_001:
        msg.CodePosition = "LOTSET-0001"

        If ret.HasError Then
            msg.MessageType = "ERROR"
            msg.MessageText = ret.ErrorCode & ":" & ret.ErrorMessage
        Else
            msg.MessageType = "REPORT"
            msg.MessageText = ret.Message
        End If

        m_Logger.SaveLog(msg)

        Return ret
    End Function

    Public Function LotEnd(ByVal machineNo As String, ByVal lotNo As String, ByVal endTime As DateTime, ByVal goodPieces As Integer, ByVal badPieces As Integer, ByVal endMode As EndModeType, ByVal opNo As String) As TdcResponse

        'ไม่เอาหน่วยมิลลิวินาที
        endTime = endTime.AddMilliseconds(-endTime.Millisecond)

        Dim ret As TdcResponse = New TdcResponse()
        ret.MCNo = machineNo
        ret.LotNo = lotNo
        ret.HasError = True

        Dim msg As TdcMessage = New TdcMessage("LotEnd", machineNo, lotNo, opNo)
        msg.Parameter = "( MCNo:" & machineNo & " LotNo:" & lotNo & " EndTime:" & endTime.ToString("yyyy/MM/dd HH:mm:ss") & " GoodPcs:" & goodPieces.ToString() & " NGPcs:" & badPieces.ToString() & " EndMode:" & endMode.ToString() & " OPNo:" & opNo & ")"

        Using con As SqlConnection = m_TdcDB.CreateConnection()

            Dim mc As APCSDBDataSet.MACHI_TABLERow = Nothing
            Dim wk As APCSDBDataSet.WorkDataRow = Nothing

            If DoCommonOperation(con, machineNo, lotNo, mc, wk, ret) Then

                Dim newStatus1 As String = wk.STATUS1
                Dim newStatus2 As String = wk.STATUS2
                Dim machineSub As String = CInt(endMode).ToString() 'ขอยิมฟิลเก็บ EndMode
                Dim newOpeSeq As Short = wk.OPE_SEQ

                Dim lot1Table As APCSDBDataSet.LOT1_TABLEDataTable
                Try
                    lot1Table = m_TdcDB.GetLot1Table(con, lotNo)
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot1TableRow As APCSDBDataSet.LOT1_TABLERow = CType(lot1Table.Rows(0), APCSDBDataSet.LOT1_TABLERow)

                Dim lot2Table As APCSDBDataSet.LOT2_TABLEDataTable
                Try
                    lot2Table = m_TdcDB.GetLot2Table(con, lotNo)
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot2TableRow As APCSDBDataSet.LOT2_TABLERow = Nothing

                If lot2Table.Count = 0 Then
                    'error LOT2_TABLE not found
                    'beacause LOT2_TABLE should be created before LOTSET command or on received LOTSET command
                    ret.ErrorCode = "03" 'Not Run
                    ret.ErrorMessage = "Could not find lot start data of LotNo " & lotNo & " in LOT2_TABLE table"
                    GoTo LBL_001
                End If
                lot2TableRow = CType(lot2Table.Rows(0), APCSDBDataSet.LOT2_TABLERow)

                'try to check LOT2_DATA by 
                Dim lot2Data As APCSDBDataSet.LOT2_DATADataTable
                Try
                    lot2Data = m_TdcDB.GetLot2DataOnLotEnd(con, wk.OPE_SEQ, wk.LOT_NO, machineNo, CODE_MACHINE_SUB_LOT_RUNNING) 'must have single row
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Dim lot2DataRow As APCSDBDataSet.LOT2_DATARow = Nothing
                'create new if not exists
                If lot2Data.Rows.Count = 0 Then
                    'error LOT2_DATA not found
                    'beacause LOT2_DATA should be created in LOTSET command
                    'เป็นไปได้ 2 กรณี คือ มันจบ Lot แล้วเลยหาไม่เจอ กับ มันยังไม่ได้ Run
                    ret.ErrorCode = "03"
                    ret.ErrorMessage = "Not run"
                    GoTo LBL_001
                ElseIf lot2Data.Rows.Count > 1 Then
                    ret.ErrorCode = "99" 'Not Run
                    ret.ErrorMessage = String.Format("There are record in LOT2_TABLE more than one. " & _
                                                     "[LOT_NO:={0}, OPE_SEQ:={1}, MACHINE:={2}, MACHINE_SUB:= {3}]", _
                                                     wk.LOT_NO, wk.OPE_SEQ, machineNo, CODE_MACHINE_SUB_LOT_END)
                    GoTo LBL_001
                Else
                    lot2DataRow = CType(lot2Data.Rows(0), APCSDBDataSet.LOT2_DATARow)
                End If

                Dim sumGoodPieces As Integer = 0
                Dim sumBadPieces As Integer = 0

                'pass GOOD_PIECES and BAD_PIECES as refernce type
                Dim tran As SqlTransaction
                Try
                    tran = con.BeginTransaction()
                Catch sx As SqlException
                    ManageSqlExceptionMessage(sx, ret)
                    GoTo LBL_001
                Catch ex As Exception
                    ret.ErrorCode = "99"
                    ret.ErrorMessage = "Exception:" & ex.Message
                    GoTo LBL_001
                End Try

                Try
                    'update LOT2_DATA
                    lot2DataRow.GOOD_PIECES = goodPieces
                    lot2DataRow.BAD_PIECES = badPieces

                    lot2DataRow.MACHINE_SUB = machineSub 'ขอยืมฟิลด์เก็บ End Mode
                    lot2DataRow.OPERATOR2 = opNo
                    lot2DataRow.REAL_DAY = endTime
                    m_TdcDB.UpdateLot2Data(con, tran, lot2Data) 'UPDATE LOT2_DATA

                    Dim tbl As APCSDBDataSet.LOT2_DATADataTable
                    Try
                        tbl = m_TdcDB.GetLot2DataByLotNoOpeSeq(con, tran, lotNo, lot2DataRow.OPE_SEQ)
                    Catch sx As SqlException
                        ManageSqlExceptionMessage(sx, ret)
                        GoTo LBL_001
                    Catch ex As Exception
                        ret.ErrorCode = "99"
                        ret.ErrorMessage = "Exception:" & ex.Message
                        GoTo LBL_001
                    End Try


                    'if normal end and all lot is end

                    If wk.STATUS1 = CODE_STATUS1_IN_PROCESS Then

                        If endMode = EndModeType.Normal Then

                            newOpeSeq = wk.N_OPE_SEQ
                            newStatus1 = CODE_STATUS1_WAIT_FOR_OPERATION

                            'clear cycle
                            lot1TableRow.CYCLE = 0
                            lot2TableRow.CYCLE = 0

                            If wk.OPE_SEQ = wk.N_OPE_SEQ Then 'have no next process
                                newStatus2 = CODE_STATUS2_FINISH_SHIPING
                            End If

                        ElseIf endMode = EndModeType.AbnormalEndAccumulate Then

                            newOpeSeq = wk.OPE_SEQ
                            newStatus1 = CODE_STATUS1_WAIT_FOR_OPERATION

                        ElseIf endMode = EndModeType.AbnormalEndReset Then

                            newOpeSeq = wk.OPE_SEQ
                            newStatus1 = CODE_STATUS1_WAIT_FOR_OPERATION

                        End If
                    ElseIf wk.STATUS1 = CODE_STATUS1_PARALLEL_END Then

                        Dim isAllParallelEndNormal As Boolean = False

                        Using tbl

                            'old version tdc did not set CYCLE value
                            If wk.CYCLE = 0 Then
                                isAllParallelEndNormal = True
                            Else
                                isAllParallelEndNormal = AllParallelLotAreEndNormal(wk, tbl)
                            End If

                        End Using

                        If isAllParallelEndNormal AndAlso endMode = EndModeType.Normal Then

                            newOpeSeq = wk.N_OPE_SEQ
                            newStatus1 = CODE_STATUS1_WAIT_FOR_OPERATION

                            'clear cycle
                            lot1TableRow.CYCLE = 0
                            lot2TableRow.CYCLE = 0

                            If wk.OPE_SEQ = wk.N_OPE_SEQ Then 'have no next process
                                newStatus2 = CODE_STATUS2_FINISH_SHIPING
                            End If

                        End If

                    Else
                        'error
                        ret.ErrorCode = "02"
                        ret.ErrorMessage = "STATUS1 must be 1 or 5"
                    End If

                    If m_TdcDB.SumOutputOfLOT2_DATA(con, tran, wk.LOT_NO, wk.LAY_NO, sumGoodPieces, sumBadPieces) Then
                        Dim updateTime As Date = Now
                        updateTime = updateTime.AddMilliseconds(-updateTime.Millisecond)
                        If newStatus2 = CODE_STATUS2_FINISH_SHIPING Then
                            m_TdcDB.DeleteLot1TableAndLot1DataByLotNo(con, tran, wk.LOT_NO)
                        Else
                            lot1TableRow.STATUS1 = newStatus1
                            lot1TableRow.STATUS2 = newStatus2
                            lot1TableRow.OPE_SEQ = newOpeSeq
                            lot1TableRow.GOOD_PIECES = sumGoodPieces
                            lot1TableRow.BAD_PIECES = sumBadPieces
                            lot1TableRow.REAL_DAY = endTime
                            lot1TableRow.CREATION_DATE = updateTime
                            Dim lot1Table_updatedRowsCount As Integer = m_TdcDB.UpdateLot1Table(con, tran, lot1Table) 'UPDATE LOT1_TABLE
                            If lot1Table_updatedRowsCount <> 1 Then
                                Throw New Exception("UPDATE LOT1_TABLE expected row count [" & lot1Table_updatedRowsCount.ToString() & "] <> 1")
                            End If

                        End If
                        lot2TableRow.STATUS1 = newStatus1
                        lot2TableRow.STATUS2 = newStatus2
                        lot2TableRow.OPE_SEQ = newOpeSeq
                        lot2TableRow.GOOD_PIECES = sumGoodPieces
                        lot2TableRow.BAD_PIECES = sumBadPieces
                        lot2TableRow.REAL_DAY = endTime
                        lot2TableRow.CREATION_DATE = updateTime
                        Dim lot2Table_updatedRowsCount As Integer = m_TdcDB.UpdateLot2Table(con, tran, lot2Table) 'UPDATE LOT2_TABLE
                        If lot2Table_updatedRowsCount <> 1 Then
                            Throw New Exception("UPDATE LOT2_TABLE expected row count [" & lot2Table_updatedRowsCount.ToString() & "] <> 1")
                        End If
                    End If

                    tran.Commit()

                    If wk.IsEND_MSGNull() Then
                        ret.Message = ""
                    Else
                        ret.Message = wk.END_MSG
                    End If

                    msg.Status1 = newStatus1
                    msg.Status2 = newStatus2

                    ret.HasError = False
                Catch sx As SqlException
                    Try
                        tran.Rollback()
                        ret.ErrorCode = "72"
                        ManageSqlExceptionMessage(sx, ret)
                    Catch ex2 As Exception
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = "Rollback:" & sx.Message & vbNewLine & ", ConnectionState :=" & con.State.ToString()
                    End Try
                Catch ex As Exception
                    Try
                        tran.Rollback()
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = ex.Message
                    Catch ex2 As Exception
                        ret.ErrorCode = "72"
                        ret.ErrorMessage = "Rollback:" & ex.Message & vbNewLine & ", ConnectionState :=" & con.State.ToString()
                    End Try
                Finally
                    lot1Table.Dispose()
                    lot1Table = Nothing
                    lot2Table.Dispose()
                    lot2Table = Nothing
                    lot2Data.Dispose()
                    lot2Data = Nothing
                End Try

            End If

        End Using
LBL_001:
        msg.CodePosition = "LOTEND-0001"
        If ret.HasError Then
            msg.MessageType = "ERROR"
            msg.MessageText = ret.ErrorCode & ":" & ret.ErrorMessage
        Else
            msg.MessageType = "REPORT"
            msg.MessageText = ret.Message
        End If
        m_Logger.SaveLog(msg)

        Return ret
    End Function

    ''' <summary>
    ''' MOVE LOT TO SPECIFIED LAYER_NO
    ''' </summary>
    ''' <param name="lotNo">LOT_NO that you want to be move</param>
    ''' <param name="mcNo">MCNo that execute this command(for keep log only)</param>
    ''' <param name="opNo">OPNo that execute this command(for keep log only)</param>
    ''' <param name="layNo">LAYER_NO for get OPE_SEQ</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MoveLot(ByVal lotNo As String, ByVal mcNo As String, ByVal opNo As String, ByVal layNo As String) As Boolean

        Dim ret As Boolean = False
        Dim msg As TdcMessage = New TdcMessage("MoveLot", mcNo, lotNo, opNo)

        msg.Parameter = "{ MCNo:" & mcNo & ", LotNo:" & lotNo & ", OPNo:" & opNo & ", LAY_NO:" & layNo & "}"

        Using con As SqlConnection = m_TdcDB.CreateConnection()

            con.Open()

            Dim opeSeq As Short = m_TdcDB.GetOpeSeqByLotNoLayNo(con, lotNo, layNo)

            If opeSeq > 0 Then

                Dim lot1Table As APCSDBDataSet.LOT1_TABLEDataTable = m_TdcDB.GetLot1Table(con, lotNo)
                Dim lot1TableRow As APCSDBDataSet.LOT1_TABLERow = CType(lot1Table.Rows(0), APCSDBDataSet.LOT1_TABLERow)

                Dim lot2Table As APCSDBDataSet.LOT2_TABLEDataTable = m_TdcDB.GetLot2Table(con, lotNo)
                Dim lot2TableRow As APCSDBDataSet.LOT2_TABLERow = Nothing

                'create new if not exists
                If lot2Table.Rows.Count = 0 Then
                    Dim newLot2TableRow As APCSDBDataSet.LOT2_TABLERow = lot2Table.NewLOT2_TABLERow()
                    Dim workRowTable As APCSDBDataSet.WorkDataDataTable = m_TdcDB.GetWorkData(con, lotNo)
                    Dim wk As APCSDBDataSet.WorkDataRow = CType(workRowTable.Rows(0), APCSDBDataSet.WorkDataRow)
                    PopulateLOT2_TABLERow(wk, newLot2TableRow)
                    lot2Table.Rows.Add(newLot2TableRow)
                End If

                lot2TableRow = CType(lot2Table.Rows(0), APCSDBDataSet.LOT2_TABLERow)

                Dim tran As SqlTransaction = con.BeginTransaction()

                Try

                    Dim updateTime As Date = Now
                    updateTime = updateTime.AddMilliseconds(-updateTime.Millisecond)

                    lot1TableRow.GOOD_PIECES = 0 'this is TOS additional for correct data in LOT1_TABLE
                    lot1TableRow.BAD_PIECES = 0 'this is TOS additional for correct data in LOT1_TABLE
                    lot1TableRow.STATUS1 = "0"
                    lot1TableRow.STATUS2 = "0"
                    lot1TableRow.OPE_SEQ = opeSeq
                    lot1TableRow.CREATION_DATE = updateTime
                    Dim aff1 As Integer = m_TdcDB.UpdateLot1Table(con, tran, lot1Table) 'UPDATE LOT1_TABLE

                    If aff1 <> 1 Then
                        Throw New Exception("UPDATE LOT1_TABLE expected row count [" & aff1.ToString() & "] <> 1")
                    End If

                    lot2TableRow.GOOD_PIECES = 0 'this is TOS additional for correct data in LOT2_TABLE
                    lot2TableRow.BAD_PIECES = 0 'this is TOS additional for correct data in LOT2_TABLE
                    lot2TableRow.STATUS1 = "0"
                    lot2TableRow.STATUS2 = "0"
                    lot2TableRow.OPE_SEQ = opeSeq
                    lot2TableRow.CREATION_DATE = updateTime
                    Dim aff2 As Integer = m_TdcDB.UpdateLot2Table(con, tran, lot2Table) 'UPDATE LOT2_TABLE

                    If aff2 <> 1 Then
                        Throw New Exception("UPDATE LOT2_TABLE expected row count [" & aff2.ToString() & "] <> 1")
                    End If

                    msg.Status1 = "0"
                    msg.Status2 = "0"

                    tran.Commit()
                    ret = True
                    msg.MessageType = "REPORT"
                    msg.MessageText = "COMPLETED"
                Catch ex As Exception
                    tran.Rollback()
                    msg.MessageType = "ERROR"
                    msg.MessageText = ex.Message
                Finally
                    lot1Table.Dispose()
                    lot1Table = Nothing
                    lot2Table.Dispose()
                    lot2Table = Nothing
                End Try

            End If

        End Using

        m_Logger.SaveLog(msg)

        Return ret

    End Function

    Private Function AllParallelLotAreEnd(ByVal wk As APCSDBDataSet.WorkDataRow, ByVal tbl As APCSDBDataSet.LOT2_DATADataTable) As Boolean

        Dim mcSubEndNormalCount As Integer = 0
        Dim mcNo As String = Nothing
        Dim mcSub As String = Nothing
        For Each r As APCSDBDataSet.LOT2_DATARow In tbl.Rows
            mcNo = r.MACHINE
            mcSub = r.MACHINE_SUB 'ขอยืมฟิลด์เก็บ End Mode      
            If mcSub = CODE_MACHINE_SUB_LOT_END OrElse _
               mcSub = CODE_MACHINE_SUB_ABNORMAL_ACCUM OrElse _
               mcSub = CODE_MACHINE_SUB_ABNORMAL_RESET Then
                mcSubEndNormalCount += 1
            End If
        Next
        Dim ret As Boolean = (mcSubEndNormalCount = wk.CYCLE) '2016-10-24
        Return ret
    End Function

    Private Function AllParallelLotAreEndNormal(ByVal wk As APCSDBDataSet.WorkDataRow, ByVal tbl As APCSDBDataSet.LOT2_DATADataTable) As Boolean

        Dim mcSubEndNormalCount As Integer = 0
        Dim mcNo As String = Nothing
        Dim mcSub As String = Nothing
        For Each r As APCSDBDataSet.LOT2_DATARow In tbl.Rows
            mcNo = r.MACHINE
            mcSub = r.MACHINE_SUB 'ขอยืมฟิลด์เก็บ End Mode      
            If mcSub = CODE_MACHINE_SUB_LOT_END Then
                mcSubEndNormalCount += 1
            End If
        Next
        Dim ret As Boolean = (mcSubEndNormalCount = wk.CYCLE) '2016-10-24
        Return ret
    End Function

    Public Function GetLotInfo(ByVal lotNo As String, ByVal mcNo As String) As LotInfo

        Dim ret As LotInfo = Nothing

        Using con As SqlConnection = m_TdcDB.CreateConnection()
            Using cmd As SqlCommand = con.CreateCommand()
                con.Open()

                'NOTE: 2015-08-13
                '*** ต้องใช้ LOT2_TABLE เพราะไม่มีวันถูกลบ
                '*** ใช้ข้อมูลจาก LOT2_DATA เพราะ มีข้อมูล MACHINE, OPERATOR1, OPERATOR2 
                '    แต่อาจไม่มีแถวที่ลิ้งกับ โปรเซสปัจจุบัน ในกรณี WIP, LOT2_DATA ของโปรเซสนั้นจะยังไม่ถูกสร้างขึ้น
                '*** ใช้ข้อมูลจาก LOT1_DATA เพราะมีข้อมูลของการแพลนสำหรับ ทุกโปรเซสอยู่แล้ว และต้องการค่า LAY_NO
                '    ซึ่งใช้จาก LOT2_DATA ไม่ได้ ในกรณี WIP
                cmd.Parameters.Add("@LOT_NO", SqlDbType.VarChar, 13).Value = lotNo
                cmd.Parameters.Add("@MACHINE", SqlDbType.VarChar, 12).Value = mcNo
                cmd.CommandText = "SELECT C.OPE_NAME, A.OPE_SEQ, A.LOT_NO, A.REAL_START, A.REAL_DAY, " & _
                    "A.GOOD_PIECES, A.BAD_PIECES, A.STATUS1, A.STATUS2" & _
                    ",B.MACHINE, B.OPERATOR1, B.OPERATOR2 FROM LOT2_TABLE AS A " & _
                    "LEFT JOIN LOT2_DATA AS B ON A.LOT_NO = B.LOT_NO AND A.OPE_SEQ = B.OPE_SEQ " & _
                    "LEFT JOIN LOT1_DATA AS D ON A.LOT_NO = D.LOT_NO AND A.OPE_SEQ = D.OPE_SEQ " & _
                    "INNER JOIN LAYER_TABLE AS C ON D.LAY_NO = C.LAY_NO " & _
                    "WHERE (A.LOT_NO = @LOT_NO) AND (B.MACHINE = @MACHINE)"

                Using dt As DataTable = New DataTable()
                    dt.Load(cmd.ExecuteReader())
                    If dt.Rows.Count = 1 Then
                        Dim row As DataRow = dt.Rows(0)

                        ret = New LotInfo()
                        ret.LOT_NO = CStr(row("LOT_NO"))
                        ret.OPE_NAME = CStr(row("OPE_NAME"))
                        If Not Convert.IsDBNull(row("MACHINE")) Then
                            ret.MACHINE = CStr(row("MACHINE")) 'null
                        End If
                        If Not Convert.IsDBNull(row("REAL_START")) Then
                            ret.REAL_START = CDate(row("REAL_START")) 'null
                        End If
                        If Not Convert.IsDBNull(row("REAL_DAY")) Then
                            ret.REAL_DAY = CDate(row("REAL_DAY")) 'null
                        End If
                        ret.GOOD_PIECES = CInt(row("GOOD_PIECES")) 'not null
                        ret.BAD_PIECES = CInt(row("BAD_PIECES")) 'not null
                        If Not Convert.IsDBNull(row("OPERATOR1")) Then
                            ret.OPERATOR1 = CStr(row("OPERATOR1")) 'null
                        End If
                        If Not Convert.IsDBNull(row("OPERATOR2")) Then
                            ret.OPERATOR2 = CStr(row("OPERATOR2")) 'null
                        End If
                        ret.STATUS1 = CType(CInt(row("STATUS1")), StartingCondition) 'not null
                        ret.STATUS2 = CType(CStr(row("STATUS2")), LotCondition) 'not null
                    ElseIf dt.Rows.Count > 1 Then
                        Throw New Exception("Found more than 1 row")
                    Else
                        Return Nothing
                    End If
                End Using

                cmd.CommandText = "SELECT C.OPE_NAME, B.OPE_SEQ, B.LOT_NO, B.MACHINE, B.REAL_START, " & _
                    "B.REAL_DAY, B.GOOD_PIECES, B.BAD_PIECES, B.OPERATOR1, B.OPERATOR2 " & _
                    "FROM LOT2_DATA AS B INNER JOIN LAYER_TABLE AS C ON B.LAY_NO = C.LAY_NO " & _
                    "WHERE (B.LOT_NO = @LOT_NO)"

                Using dt As DataTable = New DataTable()
                    dt.Load(cmd.ExecuteReader())

                    Dim tmp As List(Of LotHistory) = New List(Of LotHistory)
                    Dim hist As LotHistory = Nothing

                    For Each row As DataRow In dt.Rows
                        hist = New LotHistory()

                        hist.OPE_NAME = CStr(row("OPE_NAME"))
                        If Not Convert.IsDBNull(row("MACHINE")) Then
                            hist.MACHINE = CStr(row("MACHINE")) 'null
                        End If
                        hist.REAL_START = CDate(row("REAL_START")) 'not null
                        If Not Convert.IsDBNull(row("REAL_DAY")) Then
                            hist.REAL_DAY = CDate(row("REAL_DAY")) 'null
                        End If
                        hist.GOOD_PIECES = CInt(row("GOOD_PIECES")) 'not null
                        hist.BAD_PIECES = CInt(row("BAD_PIECES")) 'not null
                        If Not Convert.IsDBNull(row("OPERATOR1")) Then
                            hist.OPERATOR1 = CStr(row("OPERATOR1")) 'null
                        End If
                        If Not Convert.IsDBNull(row("OPERATOR2")) Then
                            hist.OPERATOR2 = CStr(row("OPERATOR2")) 'null
                        End If

                        tmp.Add(hist)
                    Next

                    ret.LotHistories = tmp.ToArray()
                    tmp.Clear()
                    tmp = Nothing

                End Using

            End Using
        End Using

        Return ret

    End Function

    Private Sub ManageSqlExceptionMessage(ByVal sx As SqlException, ByVal ret As TdcResponse)
        ret.ErrorCode = "70"
        'https://msdn.microsoft.com/en-us/library/system.data.sqlclient.sqlerror.number.aspx
        If sx.Number = -2 Then
            ret.ErrorMessage = "SqlException: Connection timeout"
        ElseIf sx.Number = 10054 Then
            ret.ErrorMessage = "SqlException: Connection lose"
        Else
            ret.ErrorMessage = "SqlException:" & sx.Message
        End If
    End Sub


End Class
