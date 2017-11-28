Public Class Form1

    '    Public Const defCOLOR_NOMAL = &H80000008
    '    Public Const defCOLOR_ERROR = &HFF&

    '#Region "TOS's declaration"


    '#End Region

    '    Public Function CommnadLotRpt(ByVal strData As String, ByVal strMcnNo As String) As Integer

    '        Dim result As Integer = 0

    '        Dim strArray() As String
    '        strArray = Split(strData, ",")
    '        If UBound(strArray) < 1 Then
    '            CommnadLotRpt = -1
    '            Exit Function
    '        End If

    '        'Read LOT1_DATA
    '        Dim Ret As Integer
    '        Dim strLotNo As String
    '        Dim strRecipe As String
    '        Dim strNotification As String
    '        Dim strGoodPieces As String
    '        Dim strBadPieces As String
    '        Dim strAdd1 As String
    '        Dim strAdd2 As String
    '        Dim strCmd As String
    '        Dim intErrCode As Integer
    '        Dim strMachine As String
    '        Dim strWork As String
    '        Dim iParallel As Integer

    '        Dim strOperator As String 'Declare by TOS
    '        Dim p_ParallelStatus As Integer  'Declare by TOS
    '        Dim p_ParallelMode As Integer  'Declare by TOS

    '        strLotNo = strArray(0)
    '        strOperator = strArray(1)
    '        p_ParallelStatus = 0
    '        p_ParallelMode = 0

    '        '2010.12.03
    '        If UBound(strArray) > 1 Then
    '            iParallel = Val(strArray(2))
    '        Else
    '            iParallel = 0
    '        End If

    '        'machine Search
    '        Ret = Get_MachineName(strMcnNo, strMachine)
    '        If Ret = 0 Then
    '            g_MachineName = strMachine
    '        Else
    '            intErrCode = 4      'machine not found
    '            GoTo G01_CommnadLotRpt
    '        End If

    '        'Lot Search
    '        g_SelfConNowFlag = True
    '        Ret = frmLOT_Input(strLotNo)
    '        g_SelfConNowFlag = False
    '        If Ret = 0 Then
    '            Call frmLOT_Disp()
    '            intErrCode = 0
    '            If (strLotNo = "") Then
    '                intErrCode = 1 'NotFount
    '            ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                intErrCode = 2 'running
    '            End If
    '            '2010.12.13
    '            If (p_ParallelStatus = 4 Or p_ParallelStatus = 5) And iParallel = 0 Then
    '                intErrCode = 5
    '            End If
    '        Else
    '            '### Assy
    '            strNotification = ""
    '            If g_ResCode_frmLOT_Input = 5 Then
    '                strNotification = lblMSG.Caption
    '            End If
    '            intErrCode = ChangeErrorCode(g_ResCode_frmLOT_Input)
    '        End If

    'G01_CommnadLotRpt:
    '        If intErrCode = 0 Then
    '            'From frmLOT Text
    '            With frmLOT
    '                strNotification = frmLOT.txtCmt.Caption
    '                If strNotification = "" And .lblMSG.ForeColor = defCOLOR_ERROR Then
    '                    strNotification = .lblMSG.Caption
    '                End If
    '                strRecipe = .txtRec.Caption
    '                strGoodPieces = .txtKkazu.Caption
    '                strBadPieces = .txtKkazuBad.Caption
    '                '###
    '                strNotification = ConvCrLfToSP(strNotification)
    '                '###
    '                strRecipe = ConvCrLfToSP(strRecipe)
    '                strGoodPieces = ConvCrLfToSP(strGoodPieces)
    '                strBadPieces = ConvCrLfToSP(strBadPieces)
    '            End With
    '        End If

    '        'Display Clear
    '        Call frmLOT.lblCmd_Click(0)

    '        'Send Data Set
    '        If intErrCode = 0 Then                  '-- OK --
    '            strCmd = strLotNo                   'LOT_NO
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strRecipe         'RECIPE
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strNotification   'NOTIFICATION
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strGoodPieces     'Good pieses
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strBadPieces      'Bad pieces
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strAdd1           'ADD1
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strAdd2           'ADD2
    '        Else                                    '-- NG --
    '            strCmd = Format$(intErrCode, "00")
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & MakeErrMsg(intErrCode)
    '            '###
    '            If (strNotification <> "") Then
    '                strCmd = strCmd & strNotification
    '            End If
    '        End If

    '        Dim strSendData As String

    '        strSendData = "@LOTRPT|" & strMcnNo & "|" & strCmd

    '        TimerCount = 0              'Timeout Start
    '        RetryCount = RetryMax
    '        CommnadLotRpt = 0
    '    End Function


    '    Public Function CommnadStrRpt(ByVal strData As String, ByVal strMcnNo As String) As Long
    '        Dim strArray() As String
    '        strArray() = Split(strData, ",")
    '        If UBound(strArray) < 1 Then
    '            'Call PutLogFile(p_LogMode, "ARRAY-END", 2)
    '            CommnadStrRpt = -1
    '            Exit Function
    '        End If

    '        Dim Ret As Integer
    '        Dim strLotNo As String
    '        Dim strCmd As String
    '        Dim strDtMoveIn As String
    '        Dim intErrCode As Integer
    '        Dim strMachine As String
    '        Dim strWork As String

    '        strLotNo = strArray(0)
    '        strDtMoveIn = strArray(1)
    '        p_ParallelStatus = 0
    '        p_ParallelMode = 0

    '        'machine Search
    '        Ret = Get_MachineName(strMcnNo, strMachine)
    '        If Ret = 0 Then
    '            frmLOT.g_MachineName = strMachine
    '        Else
    '            intErrCode = 4      'machine not found
    '            GoTo G01_CommnadStrRpt
    '        End If

    '        g_SelConDtStart = CDate(strDtMoveIn)
    '        g_SelConDateFlag = True

    '        'Input LOT_NO           Lot Search
    '        frmLOT.g_SelfConNowFlag = True
    '        Ret = frmLOT.frmLOT_Input(strLotNo)
    '        frmLOT.g_SelfConNowFlag = False
    '        If Ret = 0 Then
    '            Call frmLOT.frmLOT_Disp()
    '            If gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                'Input 'START'
    '                '2010.12.13
    '                If UBound(strArray) > 1 Then
    '                    p_ParallelMode = Val(strArray(2))
    '                Else
    '                    p_ParallelMode = 0
    '                End If
    '                Call frmLOT.lblCmd_Click(4)     'Start/End
    '                'Input OPERATOR
    '                strWork = strOperator
    '                Ret = frmLOT.frmLOT_Input(strWork)
    '                If Ret = 0 Then
    '                    'Input Eqp (INPMCH)
    '                    Call PutLogFile(p_LogMode, "INPMCH=" & SINI.MCN_FLG7, 2)
    '                    If (SINI.MCN_FLG7 <> 0) Then
    '                        ' EqpInputEnable
    '                        Ret = frmLOT.frmLOT_Input(strMcnNo)
    '                    Else
    '                        ' EqpInputDisable
    '                        Ret = 0
    '                    End If
    '                    If (Ret = 0) Then
    '                        'Input 'Y'
    '                        frmLOT.g_SelfConNowFlag = True
    '                        Ret = frmLOT.frmLOT_Input("Y")
    '                        frmLOT.g_SelfConNowFlag = False
    '                    End If
    '                End If
    '            ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                Ret = 1
    '            Else
    '                Ret = 5
    '            End If
    '        End If
    '        g_SelConDateFlag = False

    '        If Ret = 0 Then
    '            intErrCode = 0
    '        ElseIf Ret = 1 Then
    '            intErrCode = 2 'Running
    '        ElseIf Ret = 5 Then
    '            intErrCode = 5
    '        Else
    '            If (frmLOT.g_ResCode_frmLOT_Input <> 0) Then
    '                '            intErrCode = 99
    '                intErrCode = ChangeErrorCode(frmLOT.g_ResCode_frmLOT_Input)
    '            Else
    '                intErrCode = 1
    '            End If
    '        End If
    'G01_CommnadStrRpt:
    '        Call PutLogFile(p_LogMode, "RE=" & Ret & "/" & intErrCode, 2)

    '        'Display Clear
    '        Call frmLOT.lblCmd_Click(0)

    '        'Send Data Set
    '        If intErrCode = 0 Then                  '-- OK --
    '            strHeadData = "@STRRPT"
    '            strSendData = strHeadData & "|" & strMcnNo & "|"
    '            '### ๓๐•ิท
    '        Else                                    '-- NG --
    '            strCmd = Format$(intErrCode, "00")
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & MakeErrMsg(intErrCode)

    '            strHeadData = "@STRRPT"
    '            strSendData = strHeadData & "|" & strMcnNo & "|"
    '            strSendData = strSendData & strCmd
    '        End If

    '        TimerCount = 0              'Timeout Start
    '        RetryCount = RetryMax
    '        CommnadStrRpt = 0
    '    End Function

    '    '5.End Infomation
    '    '   strData     : [LOT_NO],[Move out],[Good pieses],[Bad pieces],[EndFlg]
    '    '   strSendData : @ENDINF|[LOT_NO],[NOTIFICATION]
    '    Public Function CommnadEndInf(ByVal strData As String, ByVal strMcnNo As String) As Long
    '        Dim strArray() As String
    '        strArray() = Split(strData, ",")
    '        If UBound(strArray) < 3 Then
    '            CommnadEndInf = -1
    '            Exit Function
    '        End If

    '        Dim Ret As Integer
    '        Dim strLotNo As String
    '        Dim strNotification As String
    '        Dim strGoodPieces As String
    '        Dim strBadPieces As String
    '        Dim strCmd As String
    '        Dim intErrCode As Integer
    '        Dim strDtMoveOut As String
    '        Dim strMachine As String
    '        Dim strWork As String
    '        Dim iEndFlg As Integer

    '        strLotNo = strArray(0)
    '        strDtMoveOut = strArray(1)
    '        strGoodPieces = strArray(2)
    '        strBadPieces = strArray(3)

    '        '2010.12.20
    '        If UBound(strArray) > 3 Then
    '            p_AbnormalEnd = Val(strArray(4))
    '        Else
    '            p_AbnormalEnd = 0
    '        End If

    '        'machine Search
    '        Ret = Get_MachineName(strMcnNo, strMachine)
    '        If Ret = 0 Then
    '            frmLOT.g_MachineName = strMachine
    '        Else
    '            intErrCode = 4      'machine not found
    '            GoTo G01_CommnadEndInf
    '        End If

    '        g_SelConDtEnd = CDate(strDtMoveOut)
    '        g_SelConDateFlag = True

    '        'Input LOT_NO           Lot Search
    '        frmLOT.g_SelfConNowFlag = True
    '        Ret = frmLOT.frmLOT_Input(strLotNo)
    '        frmLOT.g_SelfConNowFlag = False

    '        If Ret = 0 Then
    '            Call frmLOT.frmLOT_Disp()
    '            If gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                'Input Operator name
    '                strWork = strOperator
    '                Ret = frmLOT_Input(strWork)
    '                '            Call frmLOT.lblCmd_Click(4)     'Start/End
    '                If Ret = 0 Then
    '                    'Input good PIECES
    '                    Ret = frmLOT_Input(strGoodPieces)
    '                    If Ret = 0 Then
    '                        'Input Bad-Pieces
    '                        Ret = frmLOT_Input(strBadPieces)
    '                        If Ret = 0 Then
    '                            'Input 'Y'
    '                            Ret = frmLOT_Input("Y")
    '                        End If
    '                    End If
    '                End If
    '                If Ret <> 0 Then
    '                    Call PutLogFile(p_LogMode, "ENDINF=" & Ret & "/" & frmLOT.g_ResCode_frmLOT_Input, 2)
    '                    Ret = 99 'Other
    '                End If
    '            ElseIf gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                Ret = 3 'NotRun
    '            Else
    '                Ret = 5
    '            End If
    '        Else
    '            'Call PutLogFile(p_LogMode, "ENDINF=" & Ret & "/" & frmLOT.g_ResCode_frmLOT_Input, 2)
    '            Ret = 1 'NotFound
    '        End If
    '        g_SelConDateFlag = False

    '        If Ret = 0 Then
    '            intErrCode = 0
    '        Else
    '            '        intErrCode = Ret
    '            If (frmLOT.g_ResCode_frmLOT_Input <> 0) Then
    '                intErrCode = ChangeErrorCode(frmLOT.g_ResCode_frmLOT_Input)
    '            Else
    '                intErrCode = Ret
    '            End If
    '        End If

    'G01_CommnadEndInf:
    '        If intErrCode = 0 Then
    '            'From frmLOT Text
    '            With frmLOT
    '                strNotification = frmLOT.txtCmt.Caption
    '                If strNotification = "" And .lblMSG.ForeColor = defCOLOR_ERROR Then
    '                    strNotification = .lblMSG.Caption
    '                    '###
    '                    strNotification = ConvCrLfToSP(strNotification)
    '                End If
    '            End With
    '        End If

    '        'Display Clear
    '        Call frmLOT.lblCmd_Click(0)

    '        'Send Data Set
    '        If intErrCode = 0 Then                  '-- OK --
    '            strCmd = strLotNo                   'LOT_NO
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & strNotification   'NOTIFICATION
    '        Else                                    '-- NG --
    '            strCmd = Format$(intErrCode, "00")
    '            strCmd = strCmd & ","
    '            strCmd = strCmd & MakeErrMsg(intErrCode)
    '        End If

    '        strHeadData = "@ENDINF"
    '        strSendData = strHeadData & "|" & strMcnNo & "|"
    '        strSendData = strSendData & strCmd

    '        TimerCount = 0              'Timeout Start
    '        RetryCount = RetryMax
    '        p_AbnormalEnd = 0
    '        CommnadEndInf = 0
    '    End Function

    '    Function Get_MachineName(ByVal wkDeviceID As String, ByVal wkMachine As String) As Integer
    '        Dim L_sqlDB As SqlDatabase
    '        Dim strSQL As String
    '        On Error Resume Next

    '        Get_MachineName = F_ReadErr
    '        Err.Clear()
    '        If Trim(wkDeviceID) = "" Then GoTo End_Get_MachineName
    '        If pConnectDatabase(L_sqlDB, 0, 1) <> 0& Then
    '            GoTo Skp_Get_MachineName
    '        End If
    '        With L_sqlDB.objRst
    '            strSQL = "SELECT * FROM MACHI_TABLE WHERE MACHI_NAME = '" & wkDeviceID & "'"
    '            .Source = strSQL
    '            .Open()
    '            .MoveFirst()
    '            If Err.Number = cEOFBOF Then
    '                Get_MachineName = F_NotMach
    '            ElseIf Err.Number <> 0 Then
    '            Else
    '                wkMachine = .Fields("MACHI_NAME").Value
    '                Get_MachineName = F_Success
    '            End If
    '            .Close()
    '        End With
    'Skp_Get_MachineName:
    '        Call pDisconnectDatabase(L_sqlDB)
    'End_Get_MachineName:
    '    End Function

    '    Public Const F_Success As Integer = 0
    '    Public Const F_ReadErr As Integer = -1
    '    Public Const F_WriteErr As Integer = -2
    '    Public Const F_ConnectErr As Integer = -3
    '    Public Const F_DeleteErr As Integer = -4
    '    Public Const F_NonExecute As Integer = -5
    '    Public Const F_OverlapLot As Integer = 5
    '    Public Const F_NotArea As Integer = -9
    '    Public Const F_NotMachArea As Integer = -10
    '    Public Const F_NotLot As Integer = -11
    '    Public Const F_NotCarrier As Integer = -12
    '    Public Const F_NoCarrier As Integer = -13
    '    Public Const F_NotLotEnd As Integer = 20
    '    Public Const F_NotMach As Integer = 30
    '    Public Const F_NotLot_ST As Integer = 31
    '    Public Const F_StpLot As Integer = 32
    '    Public Const F_NotFound As Integer = 33
    '    Public Const F_NotInline As Integer = 40
    '    Public Const F_NoMachLot As Integer = 50
    '    Public Const F_NoMachMon As Integer = 51
    '    Public Const F_NoMachMent As Integer = 52
    '    '2007.08.07
    '    Public Const F_UsedLot As Integer = 60
    '    Public Const F_StopLot As Integer = 61
    '    Public Const F_LimitErr As Integer = 62
    '    '2008.08.09
    '    Public Const F_FlowChange As Integer = 70
    '    '2010.08.12
    '    Public Const F_ProcessErr As Integer = 71

    '    Private Function ChangeErrorCode(ByVal intErrCode As Integer) As Integer
    '        Dim Ret As Integer
    '        Select Case intErrCode
    '            Case 1
    '                Ret = 1
    '            Case F_UsedLot
    '                Ret = 5
    '            Case F_StopLot
    '                Ret = 5
    '            Case F_NotLot
    '                Ret = 5
    '            Case F_LimitErr
    '                Ret = 99
    '            Case F_NotArea
    '                Ret = 6
    '            Case F_ProcessErr
    '                Ret = 6
    '            Case F_ConnectErr
    '                Ret = 70
    '            Case F_ReadErr
    '                Ret = 71
    '            Case F_WriteErr
    '                Ret = 72
    '            Case F_DeleteErr
    '                Ret = 73
    '            Case Else
    '                Ret = 99
    '        End Select
    '        ChangeErrorCode = Ret
    '    End Function

    '    Private Const ERRMSG01 As String = "not found"
    '    Private Const ERRMSG02 As String = "running"
    '    Private Const ERRMSG03 As String = "not run"
    '    Private Const ERRMSG04 As String = "machine not found"
    '    Private Const ERRMSG05 As String = "error lot status"
    '    Private Const ERRMSG06 As String = "error process"
    '    Private Const ERRMSG70 As String = "error connect database"
    '    Private Const ERRMSG71 As String = "error read database"
    '    Private Const ERRMSG72 As String = "error write database"
    '    Private Const ERRMSG73 As String = "error delete database"
    '    Private Const ERRMSG99 As String = "other"

    '    Private Function MakeErrMsg(ByVal intErrCode As Integer) As String
    '        On Error Resume Next
    '        Select Case intErrCode
    '            Case 1
    '                MakeErrMsg = ERRMSG01
    '            Case 2
    '                MakeErrMsg = ERRMSG02
    '            Case 3
    '                MakeErrMsg = ERRMSG03
    '            Case 4
    '                MakeErrMsg = ERRMSG04
    '            Case 5
    '                MakeErrMsg = ERRMSG05
    '            Case 6
    '                MakeErrMsg = ERRMSG06
    '            Case 70
    '                MakeErrMsg = ERRMSG70
    '            Case 71
    '                MakeErrMsg = ERRMSG71
    '            Case 72
    '                MakeErrMsg = ERRMSG72
    '            Case 73
    '                MakeErrMsg = ERRMSG73
    '            Case Else
    '                MakeErrMsg = ERRMSG99
    '        End Select
    '    End Function

    '    'Convert Cr,Lf To SPACE
    '    Private Function ConvCrLfToSP(ByVal strData As String) As String
    '        Dim strResult As String
    '        strResult = strData
    '        strResult = Replace(strResult, Chr$(13), " ")
    '        strResult = Replace(strResult, Chr$(10), " ")
    '        strResult = Replace(strResult, ",", " ")
    '        ConvCrLfToSP = strResult
    '    End Function

    '    Public Function frmLOT_Input(ByVal wkInpstr As String) As Integer
    '        Dim Lp As Integer
    '        Dim wkAns As Integer
    '        Dim wkLen As Integer
    '        Dim wkRes As Integer
    '        Dim wkCnt As Integer
    '        Dim bkCnt(6) As Integer
    '        Dim chkFlg As Integer
    '        Dim lpCnt As Integer
    '        Dim wkSTR As String
    '        Dim wkCrNo(2) As String
    '        Dim wkOperator As String
    '        Dim strWork As String
    '        Dim strOpeChk As String

    '        '### Assy
    '        g_ResCode_frmLOT_Input = 0
    '        '###
    '        frmLOT_Input = -1
    '        If wkInpstr = "" Then
    '            If wkINP_No = defINPKazu Or wkINP_No = defINPMcnCD Then
    '            ElseIf wkINP_No = defINPPieceG Or wkINP_No = defINPPieceB Then '###
    '            ElseIf wkINP_No = defINPFrmQtyG Or wkINP_No = defINPFrmQtyB Then '###
    '            ElseIf wkINP_No = defINPBoxNo And wkINP_Flg1 = 1 And Trim(gLot1_HD(p_LotNo).strLotNo) <> "" Then
    '            ElseIf wkINP_No = defINPWaferN Or wkINP_No = defINPWaferL Or wkINP_No = defINPMaker Then        '2003.08.09
    '            Else
    '                Exit Function
    '            End If
    '        End If
    '        '// “—อ
    '        Select Case wkINP_No
    '            Case defINPSetNo

    '                If lblCmd(2).Caption = "" Then GoTo Skp_frmLOT_Input

    '                wkAns = GetSetNoCount(wkInpstr)
    '                If wkAns = 0 Then
    '                    Call PutLogFile(p_LogMode, "SET NG=" & wkInpstr, 2)
    '                    Call ErrMsg(frmLOT, 1011, wkInpstr & " ")
    '                    Exit Function
    '                End If
    '                Call PutLogFile(p_LogMode, "SET Change=" & wkInpstr, 2)

    '                Call Put_initMent(1)

    '                Call frmLOT_Init(0)
    '                Call LotData_ReDim(wkAns + 5, 0)
    '                lblLCnt.Caption = " " & p_LotCount & "/ " & CStr(wkAns)
    '                SMAC.S_MMAX = wkAns
    '                SMAC.S_MAIN = wkInpstr
    '                SMAC.S_MNO = GetSetNo(wkInpstr)
    '                txtSetNo.Caption = SMAC.S_MAIN
    '                p_LotMax = SMAC.S_MMAX
    '                wkINP_Flg3 = 0
    '                Call Get_initMent(1)
    '                Call lblCmd_Click(defFNCCancel)

    '                If frmLOT_AllRead() = False Then
    '                End If

    '                If wkINP_Flg1 = 0 And gLot1_HD(0).intStatus1 = 1 Then
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG
    '                    p_LotNo = 0
    '                    Call frmLOT_Disp2(p_LotNo, 1)
    '                Else
    '                    wkINP_No = defINPBoxNo
    '                End If
    '                gOperator.strOperatorCode = ""
    '                Call ZureDataClear()
    '            Case defINPBatchNo
    '                wkRes = False
    '                If wkInpstr = "" Then
    '                    wkInpstr = gLot1_HD(p_LotNo).strBatchNo
    '                End If
    '                wkLen = Len(wkInpstr)
    '                If wkLen <= 6 Then
    '                    gLot1_HD(p_LotNo).strBatchNo = wkInpstr
    '                    txtBatch.Caption = wkInpstr
    '                    Call PutLogFile(p_LogMode, "BatchNo(" & (p_LotNo + 1) & ") SET=" & wkInpstr, 2)
    '                    wkINP_No = defINPOpeNM
    '                    wkRes = True
    '                End If
    '                If wkRes = False Then
    '                    txtBatch.Caption = wkInpstr
    '                    Call ErrMsg(frmLOT, 1010, "batch no. ")
    '                    Exit Function
    '                End If
    '            Case defINPBoxNo
    '                '// aNoฬ“—อ
    '                wkLen = Len(wkInpstr)
    '                If wkINP_Flg1 = 1 Then
    '                    wkRes = False
    '                    If wkInpstr = "" Then
    '                        '// ๙’่’l“—อ
    '                        wkInpstr = gLot1_HD(p_LotNo).strBoxNo
    '                    End If
    '                    '// ““ฬ—
    '                    If wkLen <= 6 Then
    '                        If LotBoxCheck(wkInpstr) = 1 Then
    '                            '// BOXNOd•ก`FbN
    '                            If StrComp(wkInpstr, gLot1_HD(p_LotNo).strBoxNo, vbTextCompare) <> 0 Then
    '                                '// am“—อ
    '                                If Chk_BoxNo(wkInpstr) = True Then
    '                                    '// “—องภ”๐’ดฆฝ๊
    '                                    txtBoxNo.Caption = wkInpstr
    '                                    Call SetControlColor(txtBoxNo, defErrorMode2)
    '                                    Call ErrMsg(frmLOT, 1013, "")
    '                                    Exit Function
    '                                End If
    '                            End If
    '                            If Get_Carrier_Check(gLot1_DT(p_LotNo).strLayNo, wkCrNo(1), wkCrNo(2)) = F_Success Then
    '                                If (StrComp(wkCrNo(1), wkInpstr, vbTextCompare) <= 0 And StrComp(wkCrNo(2), wkInpstr, vbTextCompare) >= 0) = False Then
    '                                    txtBoxNo.Caption = wkInpstr
    '                                    Call SetControlColor(txtBoxNo, defErrorMode2)
    '                                    Call ErrMsg(frmLOT, 91, "")
    '                                    Exit Function
    '                                End If
    '                            End If
    '                            gLot1_HD(p_LotNo).strBoxNo = wkInpstr
    '                            txtBoxNo.Caption = wkInpstr
    '                            Call PutLogFile(p_LogMode, "BoxNo(" & (p_LotNo + 1) & ") SET=" & wkInpstr, 2)
    '                            wkINP_No = defINPBatchNo
    '                            wkRes = True
    '                        End If
    '                    End If
    '                    If wkRes = False Then
    '                        Call ErrMsg(frmLOT, 1001, "")
    '                        Exit Function
    '                    End If
    '                Else
    '                    If wkLen <= 6 Then
    '                        If LotBoxCheck(wkInpstr) = 1 Then
    '                            If p_LotCount >= SMAC.S_MMAX And SINI.MCN_FLG3 = 0 Then
    '                                Call ErrMsg(frmLOT, 1001, "")
    '                                Exit Function
    '                            ElseIf SINI.MCN_FLG6 = 1 And gLot1_HD(p_LotNo).strLotNo <> "" And gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                                Call ErrMsg(frmLOT, 1021, "")
    '                                Exit Function
    '                            Else
    '                                wkCnt = 0
    '                                wkRes = F_ReadErr
    '                                If SINI.MCN_FLG3 = 1 Then
    '                                    For Lp = 0 To p_LotCount - 1
    '                                        If StrComp(wkInpstr, gLot1_HD(Lp).strBoxNo, vbTextCompare) = 0 Then
    '                                            p_LotNo = Lp
    '                                            wkRes = F_OverlapLot
    '                                            wkCnt = 1
    '                                            Exit For
    '                                        End If
    '                                    Next Lp
    '                                End If
    '                                If wkRes <> F_Success Then
    '                                    If p_LotCount >= SMAC.S_MMAX Then
    '                                        Call ErrMsg(frmLOT, 1002, "")
    '                                        Exit Function
    '                                    End If
    '                                    If wkRes = F_OverlapLot Then
    '                                        wkAns = p_LotNo
    '                                    Else
    '                                        wkAns = p_LotCount
    '                                    End If
    '                                    wkRes = SetLotData(wkInpstr, wkAns)
    '                                    If g_ResCode_frmLOT_Input = 0 Then
    '                                        g_ResCode_frmLOT_Input = wkRes
    '                                    End If
    '                                    If ZureOperationCheck() = 1 And p_LotCount > 0 And gLot1_DT(p_LotCount).strLayNo <> "" Then
    '                                        For Lp = 1 To p_LotCount
    '                                            If gLot1_DT(0).strRecipe <> gLot1_DT(Lp).strRecipe And gLot1_DT(0).strRecipe <> "" Then
    '                                                Call ErrMsg(frmLOT, 9999, "reicipie is wrong")
    '                                                wkRes = 9999
    '                                                Exit For
    '                                            End If
    '                                        Next Lp
    '                                    End If
    '                                    If wkRes = 1 Then
    '                                        '// ““์ฦ
    '                                        wkINP_Flg1 = 1
    '                                    ElseIf wkRes = 2 Then
    '                                        '// oื์ฦ
    '                                        wkINP_Flg1 = 2
    '                                    ElseIf wkRes = 10 Then
    '                                        '// งภิI[o[FฤถII
    '                                        wkINP_Flg4 = 2
    '                                    ElseIf wkRes = 11 Then
    '                                        '// งภิI[o[FxII
    '                                        wkINP_Flg4 = 1
    '                                    ElseIf wkRes = 12 Then
    '                                        '// งภิI[o[FxII
    '                                        wkINP_Flg4 = 3
    '                                    ElseIf wkRes <> F_Success Then
    '                                        If p_LotCount <= 0 Then
    '                                            gLot1_HD(p_LotNo).strLotNo = ""
    '                                            gLot1_HD(p_LotNo).strBoxNo = ""
    '                                            gOpeData(p_LotNo).strMachineName = ""
    '                                        Else
    '                                            gLot1_HD(p_LotNo + 1).strLotNo = ""
    '                                            gLot1_HD(p_LotNo + 1).strBoxNo = ""
    '                                            gOpeData(p_LotNo + 1).strMachineName = ""
    '                                        End If
    '                                        Exit Function
    '                                    End If

    '                                    Call frmLOT_DSet(wkAns, defDSet_MClr)

    '                                    If wkCnt = 0 Then
    '                                        p_LotNo = p_LotCount
    '                                        Call PutLogFile(p_LogMode, "BoxNo" & (p_LotNo + 1) & " Search=" & wkInpstr, 2)  '// O‘
    '                                        p_LotCount = p_LotCount + 1
    '                                    Else
    '                                        Call PutLogFile(p_LogMode, "BoxNo" & (p_LotNo + 1) & " Change=" & wkInpstr, 2)  '// O‘
    '                                    End If
    '                                    Call frmLOT_Disp2(p_LotNo, 1)
    '                                    txtBoxNo.Caption = wkInpstr
    '                                    txtLotNo.Caption = gLot1_HD(p_LotNo).strLotNo

    '                                    If gLot1_DT(p_LotNo).strLayNo = "" And ZureOperationCheck() = 1 Then
    '                                        If NoFpcsLotInpit(Me, wkInpstr, p_LotNo) = 1 Then

    '                                            p_LotCount = p_LotCount - 1
    '                                            p_LotNo = p_LotCount
    '                                            Call frmLOT_Disp2(p_LotNo, 1)
    '                                            Me.txtLotNo.Caption = gLot1_HD(p_LotNo).strLotNo
    '                                            gLot1_HD(p_LotNo).strLotNo = ""
    '                                            Exit Function
    '                                        End If
    '                                        If p_LotNo > 0 Then
    '                                            For Lp = 1 To p_LotCount - 1
    '                                                If gLot1_DT(0).strRecipe <> gLot1_DT(Lp).strRecipe And gLot1_DT(Lp).strRecipe <> "" Then
    '                                                    p_LotCount = p_LotCount - 1
    '                                                    p_LotNo = p_LotCount
    '                                                    Call frmLOT_Disp2(p_LotNo, 1)
    '                                                    Me.txtLotNo.Caption = gLot1_HD(p_LotNo).strLotNo
    '                                                    gLot1_HD(p_LotNo).strLotNo = ""
    '                                                    Call ErrMsg(frmLOT, 9999, "reicipie is wrong")
    '                                                    Exit Function
    '                                                End If
    '                                            Next Lp
    '                                        End If
    '                                    End If
    '                                Else
    '                                    Call PutLogFile(p_LogMode, "BoxNo" & (p_LotNo + 1) & " Change=" & wkInpstr, 2)  '// O‘
    '                                End If
    '                            End If
    '                            wkINP_Flg3 = 0                                  '// H’๖•\ฆช ้๊อI—น
    '                            If wkINP_Flg4 = 1 Then
    '                                wkINP_No = defINPCheck
    '                                frmLOT_Input = 1
    '                                Exit Function
    '                            ElseIf wkINP_Flg4 = 2 Then
    '                                wkINP_No = defINPCheck2
    '                                '                        frmLOT_Input = 1
    '                                '                        Exit Function
    '                            ElseIf wkINP_Flg4 = 3 Then
    '                                wkINP_No = defINPCheck3
    '                                frmLOT_Input = 1
    '                                Exit Function
    '                            ElseIf wkINP_Flg1 = 0 And gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                                '// I—น—พมฝ็
    '                                If SINI.MCN_FLG3 = 1 Then
    '                                    '// ย•สI—นฬ๊
    '                                    wkINP_No = defINPBoxNo
    '                                Else
    '                                    '// A‘ฑI—นฬ๊
    '                                    wkINP_No = defINPKazu
    '                                    wkINP_No = defINPPieceG '###
    '                                End If
    '                            ElseIf wkINP_Flg1 = 1 Then
    '                                '// ““ฬam“o^
    '                                If Trim(gLot1_HD(p_LotNo).strBoxNo) <> "" Then
    '                                    wkINP_No = defINPOpeNM  ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                                    '                            wkINP_No = defINPKazu
    '                                Else
    '                                    wkINP_No = defINPBoxNo
    '                                End If
    '                            ElseIf SINI.MCN_FLG2 = 1 Then
    '                                '// A‘ฑI—นฬ๊
    '                                wkINP_No = defINPOpeNM      ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                                '                        wkINP_No = defINPKazu
    '                            Else
    '                                '// Jn—พมฝ็
    '                                wkINP_Flg2 = 0
    '                            End If
    '                        Else
    '                            '// ์ฦาอXLbv
    '                            '// ์ฦา“—อ
    '                            '                    txtOpeNM.Caption = wkInpStr
    '                            '                    gOperator.strOperatorCode = wkInpStr
    '                            '// Debug
    '                            Call ErrMsg(frmLOT, 1011, wkInpstr & " ")
    '                            Exit Function
    '                        End If
    '                    Else
    '                        If p_LotCount >= SMAC.S_MMAX And SINI.MCN_FLG3 = 0 Then
    '                            '// “—องภ”๐’ดฆฝ๊
    '                            Call ErrMsg(frmLOT, 1002, "")
    '                            Exit Function
    '                        ElseIf SINI.MCN_FLG6 = 1 And gLot1_HD(p_LotNo).strLotNo <> "" And gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                            '// “—องภ”๐’ดฆฝ๊
    '                            Call ErrMsg(frmLOT, 1021, "")
    '                            Exit Function
    '                        Else
    '                            '// ย•สI—นอA
    '                            wkCnt = 0
    '                            wkRes = F_ReadErr
    '                            If SINI.MCN_FLG3 = 1 Then
    '                                For Lp = 0 To p_LotCount - 1
    '                                    If StrComp(wkInpstr, gLot1_HD(Lp).strLotNo, vbTextCompare) = 0 Then
    '                                        p_LotNo = Lp
    '                                        wkRes = F_OverlapLot
    '                                        wkCnt = 1
    '                                        '                                Call frmLOT_Disp2(p_LotNo, 1)
    '                                        '                                wkRes = F_Success
    '                                        Exit For
    '                                    End If
    '                                Next Lp
    '                            End If
    '                            If wkRes <> F_Success Then
    '                                If p_LotCount >= SMAC.S_MMAX Then
    '                                    '// “—องภ”๐’ดฆฝ๊
    '                                    Call ErrMsg(frmLOT, 1002, "")
    '                                    Exit Function
    '                                End If
    '                                If wkRes = F_OverlapLot Then
    '                                    wkAns = p_LotNo
    '                                Else
    '                                    wkAns = p_LotCount
    '                                End If
    '                                wkRes = SetLotData(wkInpstr, wkAns)
    '                                '### Assy
    '                                If g_ResCode_frmLOT_Input = 0 Then
    '                                    g_ResCode_frmLOT_Input = wkRes
    '                                End If
    '                                '### Assy
    '                                '2006.04.19
    '                                If ZureOperationCheck() = 1 And p_LotCount > 0 And gLot1_DT(p_LotCount).strLayNo <> "" Then
    '                                    '`Nฬd•ก`FbN         '2006.04.13
    '                                    For Lp = 1 To p_LotCount
    '                                        If gLot1_DT(0).strRecipe <> gLot1_DT(Lp).strRecipe And gLot1_DT(0).strRecipe <> "" Then
    '                                            Call ErrMsg(frmLOT, 9999, "reicipie is wrong")
    '                                            wkRes = 9999
    '                                            Exit For
    '                                        End If
    '                                    Next Lp
    '                                End If
    '                                If wkRes = 1 Then
    '                                    '// ““์ฦ
    '                                    wkINP_Flg1 = 1
    '                                ElseIf wkRes = 2 Then
    '                                    '// oื์ฦ
    '                                    wkINP_Flg1 = 2
    '                                ElseIf wkRes = 10 Then
    '                                    '// งภิI[o[FฤถII
    '                                    wkINP_Flg4 = 2
    '                                ElseIf wkRes = 11 Then
    '                                    '// งภิI[o[FxII
    '                                    wkINP_Flg4 = 1
    '                                ElseIf wkRes = 12 Then
    '                                    '// งภิI[o[FxII
    '                                    wkINP_Flg4 = 3
    '                                ElseIf wkRes <> F_Success Then
    '                                    If p_LotCount <= 0 Then
    '                                        gLot1_HD(p_LotNo).strLotNo = ""
    '                                        gLot1_HD(p_LotNo).strBoxNo = ""
    '                                        gOpeData(p_LotNo).strMachineName = ""
    '                                    Else
    '                                        gLot1_HD(p_LotNo + 1).strLotNo = ""
    '                                        gLot1_HD(p_LotNo + 1).strBoxNo = ""
    '                                        gOpeData(p_LotNo + 1).strMachineName = ""
    '                                    End If
    '                                    Exit Function
    '                                End If
    '                                '2006.04.19
    '                                '                        If ZureOperationCheck() = 1 And p_LotCount > 0 And gLot1_DT(p_LotCount).strLayNo <> "" Then
    '                                '                            '`Nฬd•ก`FbN         '2006.04.13
    '                                '                            For Lp = 1 To p_LotCount
    '                                '                                If gLot1_DT(0).strRecipe <> gLot1_DT(Lp).strRecipe And gLot1_DT(0).strRecipe <> "" Then
    '                                '                                    Call ErrMsg(frmLOT, 9999, "reicipie is wrong")
    '                                '                                    Exit Function
    '                                '                                End If
    '                                '                            Next Lp
    '                                '                        End If
    '                                Call frmLOT_DSet(wkAns, defDSet_MClr)  '// “—อ๎•๑’วม
    '                                If wkCnt = 0 Then
    '                                    p_LotNo = p_LotCount
    '                                    Call PutLogFile(p_LogMode, "LotNo" & (p_LotNo + 1) & " Search=" & wkInpstr, 2)  '// O‘
    '                                    p_LotCount = p_LotCount + 1
    '                                Else
    '                                    Call PutLogFile(p_LogMode, "BoxNo" & (p_LotNo + 1) & " Change=" & wkInpstr, 2)  '// O‘
    '                                End If
    '                                Call frmLOT_Disp2(p_LotNo, 1)
    '                                txtBoxNo.Caption = gLot1_HD(p_LotNo).strBoxNo
    '                                txtLotNo.Caption = wkInpstr
    '                                '// ************** 2006.04.13 ****************************
    '                                '// FPCS–ข“o^ฬbg
    '                                If gLot1_DT(p_LotNo).strLayNo = "" And ZureOperationCheck() = 1 Then
    '                                    If NoFpcsLotInpit(Me, wkInpstr, p_LotNo) = 1 Then
    '                                        'L“Z
    '                                        p_LotCount = p_LotCount - 1
    '                                        p_LotNo = p_LotCount
    '                                        Call frmLOT_Disp2(p_LotNo, 1)
    '                                        Me.txtLotNo.Caption = gLot1_HD(p_LotNo).strLotNo
    '                                        gLot1_HD(p_LotNo).strLotNo = ""
    '                                        Exit Function
    '                                    End If
    '                                    If p_LotNo > 0 Then
    '                                        '`Nฬd•ก`FbN
    '                                        For Lp = 1 To p_LotCount - 1
    '                                            If gLot1_DT(0).strRecipe <> gLot1_DT(Lp).strRecipe And gLot1_DT(Lp).strRecipe <> "" Then
    '                                                p_LotCount = p_LotCount - 1
    '                                                p_LotNo = p_LotCount
    '                                                Call frmLOT_Disp2(p_LotNo, 1)
    '                                                Me.txtLotNo.Caption = gLot1_HD(p_LotNo).strLotNo
    '                                                gLot1_HD(p_LotNo).strLotNo = ""
    '                                                Call ErrMsg(frmLOT, 9999, "reicipie is wrong")
    '                                                Exit Function
    '                                            End If
    '                                        Next Lp
    '                                    End If
    '                                End If
    '                            Else
    '                                Call PutLogFile(p_LogMode, "LotNo" & (p_LotNo + 1) & " Change=" & wkInpstr, 2)  '// O‘
    '                            End If
    '                        End If
    '                        wkINP_Flg3 = 0                                  '// H’๖•\ฆช ้๊อI—น
    '                        If wkINP_Flg4 = 1 Then
    '                            wkINP_No = defINPCheck
    '                            frmLOT_Input = 1
    '                            Exit Function
    '                        ElseIf wkINP_Flg4 = 2 Then
    '                            wkINP_No = defINPCheck2
    '                            '                    frmLOT_Input = 1
    '                            '                    Exit Function
    '                        ElseIf wkINP_Flg4 = 3 Then
    '                            wkINP_No = defINPCheck3
    '                            frmLOT_Input = 1
    '                            Exit Function
    '                        ElseIf wkINP_Flg1 = 0 And gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                            '// I—น—พมฝ็
    '                            If SINI.MCN_FLG3 = 1 Then
    '                                '// ย•สI—นฬ๊
    '                                wkINP_No = defINPBoxNo
    '                            Else
    '                                wkINP_No = defINPKazu
    '                                wkINP_No = defINPPieceG '###
    '                                wkINP_No = defINPOpeNM      ' 2010.09.09
    '                            End If
    '                        ElseIf wkINP_Flg1 = 1 Then
    '                            '// ““ฬam“o^
    '                            If Trim(gLot1_HD(p_LotNo).strBoxNo) <> "" Then
    '                                wkINP_No = defINPOpeNM      ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                                '                        wkINP_No = defINPKazu
    '                            Else
    '                                wkINP_No = defINPBoxNo
    '                            End If
    '                        ElseIf SINI.MCN_FLG2 = 1 Then
    '                            '// A‘ฑI—นฬ๊
    '                            wkINP_No = defINPOpeNM          ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                            '                    wkINP_No = defINPKazu
    '                        Else
    '                            '// Jn—พมฝ็
    '                            wkINP_Flg2 = 0
    '                        End If
    '                    End If
    '                End If
    '                gOperator.strOperatorCode = ""
    '            Case defINPOpeNM
    '                '// ์ฦา“—อ
    '                txtOpeNM.Caption = wkInpstr
    '                wkLen = Len(wkInpstr)
    '                '// ์ฦา•ถ—๑m”F
    '                '        If (wkLen >= 1 And wkLen <= 16) = False Then
    '                If OperatorCheck(wkInpstr) <> 0 Then                '2006.03.11
    '                    Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Operator=" & wkInpstr & " NG", 2) '// O‘
    '                    Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Operator=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                    '            Call ErrMsg(frmLOT, 1010, "์ฦา“—อ(16…)")
    '                    Call ErrMsg(frmLOT, 9999, "operator name is wrong.")
    '                    Exit Function
    '                End If
    '                gOperator.strOperatorCode = wkInpstr
    '                Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Operator=" & wkInpstr, 2) '// O‘
    '                '---------------------------------------------------------------- 2002.09.26 ADD TDC
    '                '// Jn—ฬ๊ษ`FbN
    '                If gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                    '// I“C“EItC“‘•’u•sย`FbN
    '                    wkInpstr = gOpeData(p_LotNo).strMachineName
    '                    '            If SINI.MCN_FLG7 <> 1 And SINI.MCN_FLG8 = 0 Then
    '                    If (SINI.MCN_FLG7 <> 1 And SINI.MCN_FLG8 = 0) And ZureOperationCheck() = 0 Then         '2006.04.13
    '                        '// I“C“‘•’u•sยฬ๊
    '                        wkRes = Chk_Machine_table(wkInpstr, SINI.MCN_FLG8, "")
    '                        If wkRes = F_ReadErr Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1300, "")
    '                            wkCHK_Flg3 = True
    '                            Exit Function
    '                        ElseIf wkRes = F_NotMach Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1301, "")
    '                            wkCHK_Flg3 = True
    '                            Exit Function
    '                        End If
    '                        '            ElseIf SINI.MCN_FLG7 <> 1 And SINI.MCN_FLG8 = 1 Then
    '                    ElseIf (SINI.MCN_FLG7 <> 1 And SINI.MCN_FLG8 = 1) And ZureOperationCheck() = 0 Then     '2006.04.13
    '                        '// ItC“‘•’u•sยฬ๊
    '                        wkRes = Chk_Machine_table(wkInpstr, SINI.MCN_FLG8, "")
    '                        If wkRes = F_ReadErr Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1300, "")
    '                            wkCHK_Flg3 = True
    '                            Exit Function
    '                        ElseIf wkRes = F_NotMach Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1302, "")
    '                            wkCHK_Flg3 = True
    '                            Exit Function
    '                        End If
    '                    End If
    '                    '2006.01.25
    '                    If ZureOperationCheck() = 1 Then
    '                        '2010.07.22
    '                        '                'Y‘ช’่[h“—อ
    '                        '                If ZureModeInput() <> 0 Then
    '                        '                    Exit Function
    '                        '                End If
    '                    End If
    '                End If
    '                '----------------------------------------------------------------
    '                If SINI.MCN_FLG7 = 1 And gLot1_HD(p_LotNo).intStatus1 = 0 Then
    '                    '// ‘•’u“—อtOช—งมฤขฤJn—พมฝ็
    '                    wkINP_No = defINPMcnCD
    '                ElseIf wkINP_Flg1 = 1 Or wkINP_Flg1 = 2 Then
    '                    '// ““/oื—พมฝ็
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                ElseIf SINI.MCN_FLG2 = 1 Then
    '                    '// A‘ฑI—น—พมฝ็                ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                    '// I—น—พมฝ็
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                Else
    '                    wkINP_No = defINPStart
    '                End If
    '            Case defINPMcnCD
    '                '// ‘•’uฬ“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                        txtMachNM.Caption = gOpeData(p_LotNo).strMachineName
    '                        GoTo Skp_frmLOT_Input_10
    '                    End If
    '                    '// ๙’่’l“—อ(‘ใ•\‘•’u)
    '                    '2007.11.09
    '                    '            wkInpstr = Mid$(gOpeData(p_LotNo).strMachineName, 2, Len(gOpeData(p_LotNo).strMachineName) - 1)
    '                    wkInpstr = gOpeData(p_LotNo).strMachineName
    '                    If Len(wkInpstr) >= 2 Then
    '                        If Mid$(wkInpstr, 1, 1) = "M" Then
    '                            wkInpstr = Mid$(gOpeData(p_LotNo).strMachineName, 2, Len(gOpeData(p_LotNo).strMachineName) - 1)
    '                        End If
    '                    End If
    '                End If
    '                wkCHK_Flg3 = False
    '                txtMachNM.Caption = wkInpstr
    '                wkLen = Len(wkInpstr)
    '                '// ‘•’u•ถ—๑m”F
    '                If (wkLen >= 1 And wkLen <= 12) = False Then
    '                    Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                    Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                    Call ErrMsg(frmLOT, 1010, "input machine(12digit)")
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                '// ‘•’uมชฏธm”F
    '                wkSTR = ""
    '                If ZureOperationCheck() = 1 And ZureModeCheck(gOpeData(p_LotNo).strMode) <> 0 Then
    '                    '‘•’u`FbN๐sํศข         2006.04.13
    '                    If Len(wkInpstr) >= 1 Then
    '                        wkSTR = wkInpstr
    '                        If Mid$(wkInpstr, 1, 1) = "M" Or Mid$(wkInpstr, 1, 1) = "m" Then
    '                        Else
    '                            '2010.09.03
    '                            '                    wkInpstr = "M" & wkInpstr
    '                        End If
    '                    End If
    '                Else
    '                    If SINI.MCN_FLG5 <> 0 Then
    '                        wkRes = Chk_OpeArea_table(gLot1_DT(p_LotNo).strOpeArea, wkInpstr, 1)
    '                        If wkRes <> F_Success Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1012, "")
    '                            '2006.03.24
    '                            If PasswordInput(p_Password) <> 0 Then
    '                                wkCHK_Flg3 = True
    '                                Exit Function
    '                            End If
    '                        End If
    '                    Else
    '                        If Len(wkInpstr) >= 1 Then
    '                            wkSTR = wkInpstr
    '                            If Mid$(wkInpstr, 1, 1) = "M" Or Mid$(wkInpstr, 1, 1) = "m" Then
    '                            Else
    '                                '2010.09.03
    '                                '                        wkInpstr = "M" & wkInpstr
    '                            End If
    '                        End If
    '                    End If
    '                    '// I“C“EItC“‘•’u•sย`FbN
    '                    If SINI.MCN_FLG8 = 0 Then
    '                        '// I“C“‘•’u•sยฬ๊
    '                        wkRes = Chk_Machine_table(wkInpstr, SINI.MCN_FLG8, "")
    '                        If wkRes = F_ReadErr Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1300, "")
    '                            '2006.03.24
    '                            If PasswordInput(p_Password) <> 0 Then
    '                                wkCHK_Flg3 = True
    '                                Exit Function
    '                            End If
    '                        ElseIf wkRes = F_NotMach Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1301, "")
    '                            '2006.03.24
    '                            If PasswordInput(p_Password) <> 0 Then
    '                                wkCHK_Flg3 = True
    '                                Exit Function
    '                            End If
    '                        End If
    '                    ElseIf SINI.MCN_FLG8 = 1 Then
    '                        '// ItC“‘•’u•sยฬ๊
    '                        wkRes = Chk_Machine_table(wkInpstr, SINI.MCN_FLG8, "")
    '                        If wkRes = F_ReadErr Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1300, "")
    '                            '2006.03.24
    '                            If PasswordInput(p_Password) <> 0 Then
    '                                wkCHK_Flg3 = True
    '                                Exit Function
    '                            End If
    '                        ElseIf wkRes = F_NotMach Then
    '                            Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", 2) '// O‘
    '                            Call PutErrLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " NG", p_ErrTblNo, p_ErrTblValue)  '// O‘
    '                            Call ErrMsg(frmLOT, 1302, "")
    '                            '2006.03.24
    '                            If PasswordInput(p_Password) <> 0 Then
    '                                wkCHK_Flg3 = True
    '                                Exit Function
    '                            End If
    '                        End If
    '                    End If
    '                    '// 2010.09.03
    '                    If (SINI.PROCESS_FLG = 1) Then
    '                        strOpeChk = gLayerData(p_LotNo).strCheckOpeName
    '                        If g_SelfConNowFlag = False Then
    '                            wkRes = Chk_Machine_table(wkInpstr, 3, strOpeChk)
    '                        Else
    '                            wkRes = Chk_Machine_table(g_MachineName, 3, strOpeChk)
    '                        End If
    '                        If wkRes <> F_Success Then
    '                            'ปฬH’๖ส’uช•sณ
    '                            Call ErrMsg(frmLOT, 1012, "")
    '                            Exit Function
    '                        End If
    '                    End If
    '                End If
    '                If wkSTR <> "" Then wkInpstr = wkSTR
    '                '// ‘•’u•\ฆ
    '                txtMachNM.Caption = wkInpstr

    '                Call PutLogFile(p_LogMode, IIf(wkINP_Flg2 = 0, "Start", "End") & " Machine=" & wkInpstr & " OK", 2) '// O‘
    '                SMAC.M_BACK = gOpeData(p_LotNo).strMachineName
    '                SMAC.M_DATA = wkInpstr
    '                gOpeData(p_LotNo).strMachineName = wkInpstr
    '                '// “—อ€–ฺ
    '                If wkINP_Flg1 = 1 Or wkINP_Flg1 = 2 Then
    '                    '// ““/oื—พมฝ็
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                ElseIf SINI.MCN_FLG2 = 1 Then
    '                    '// A‘ฑI—น—พมฝ็                ' 2002.1.22 >> ”—ส“—อจ์ฦา“—อึ•ฯX
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                    '// I—น—พมฝ็
    '                    wkINP_No = defINPKazu
    '                    wkINP_No = defINPPieceG '###
    '                Else
    '                    wkINP_No = defINPStart
    '                End If
    'Skp_frmLOT_Input_10:
    '            Case defINPMach
    '                '// ‘•’uฬ‘I‘๐
    '                If lblCmd(2).Caption = "" Then GoTo Skp_frmLOT_Input
    '                '’่t@C๐g—pตศขB
    '                '        If frmLOT_MCNCHK(wkInpStr) = False Then
    '                '            Call PutLogFile(p_LogMode, "Machine NG=" & wkInpStr, 2)  '// O‘
    '                '            Call ErrMsg(frmLOT, 1012, wkInpStr)
    '                '            Exit Function
    '                '        End If
    '                If Len(wkInpstr) >= 1 Then
    '                    If Mid$(wkInpstr, 1, 1) = "M" Or Mid$(wkInpstr, 1, 1) = "m" Then
    '                    Else
    '                        '2010.09.03
    '                        '                wkInpstr = "M" & wkInpstr
    '                    End If
    '                End If
    '                SMAC.M_MAIN = wkInpstr
    '                txtMachNM.Caption = SMAC.M_MAIN
    '                Call frmLOT_Header()
    '                Call PutLogFile(p_LogMode, "Machine Change=" & wkInpstr, 2)  '// O‘
    '                '// H’๖•\ฆช ้๊อ•\ฆ๐
    '                If wkINP_Flg3 <> 0 Then
    '                    Call frmLOT_Disp2(p_LotNo, 3)
    '                    wkINP_Flg3 = 0
    '                Else
    '                    Call frmLOT_Disp2(p_LotNo, 1)
    '                End If
    '                wkINP_No = defINPBoxNo
    '                If p_LotCount <= 0 Then
    '                    '// Zbgmช“—อณ๊ฤศข๊
    '                    If txtSetNo.Caption = "" Then
    '                        wkINP_No = defINPSetNo
    '                    End If
    '                    txtMachNM.Caption = ""
    '                End If
    '                Call Get_IntegXY(SMAC.M_MAIN, SMAC.M_CNT(0), SMAC.M_CNT(1))
    '                gOperator.strOperatorCode = ""
    '            Case defINPStart
    '                '// ์ฦJn
    '                If wkInpstr = "N" Then
    '                    '// ์ฦๆม
    '                    If SMAC.S_MAIN = "" Then
    '                        wkINP_No = defINPSetNo
    '                    Else
    '                        wkINP_No = defINPBoxNo
    '                    End If
    '                    txtOpeNM.Caption = ""
    '                    gOperator.strOperatorCode = ""
    '                    If Trim(SMAC.M_BACK) <> "" Then
    '                        gOpeData(p_LotNo).strMachineName = SMAC.M_BACK
    '                        txtMachNM.Caption = gOpeData(p_LotNo).strMachineName
    '                        SMAC.M_BACK = ""
    '                    End If
    '                    SMAC.M_DATA = ""
    '                    Call frmLOT_Disp()
    '                    Call PutLogFile(p_LogMode, "Start Chancel", 2)  '// O‘
    '                    Exit Function
    '                End If
    '                If wkInpstr <> "Y" Then Exit Function

    '                If ZureOperationCheck() = 1 Then
    '                    '            Call ZureDataClear                              'NA       2006.01.26
    '                End If

    '                '// ์ฦJnf[^ฬ•‘ถ
    '                wkWTFlg = True
    '                wkSTR = ""
    '                If gOpeData(p_LotNo).intCFlg = 1 Then
    '                    wkSTR = "(start)"
    '                ElseIf gOpeData(p_LotNo).intCFlg = 2 Then
    '                    wkSTR = "(shipping)"
    '                ElseIf gOpeData(p_LotNo).intCFlg = 3 Then
    '                    wkSTR = "(flow)"
    '                End If
    '                '*******************
    '                '' ิฬ’ฒฎB
    '                '*******************
    '                Call Get_ConnectTimer(2)
    '                Call PutLogFile(p_LogMode, "Start Save" & wkSTR, 2)       '// O‘
    '                Call DispMsg(lblMSG, frmOKNG, 60, "")
    '                wkRes = PutLotStart(p_LotMax, gLot1_HD(), gLot1_DT(), gOpeData(), gLayerData(), 1)
    '                wkWTFlg = False
    '                If wkRes = F_Success Then
    '                    If ZureOperationCheck() <> 0 Then               '—’Zbg       2006.01.26
    '                        Call ZureDataSet()
    '                        wkOperator = gOperator.strOperatorCode
    '                    End If
    '                    Call PutLogFile(p_LogMode, "Save=OK", 2)   '// O‘
    '                    '// ณํษ—ชฎ—นตฝศ็I—น์ฦึฺ“ฎท้
    '                    '            Call PutLogFile(p_LogMode, "SAVE=OK", 2)
    '                    Call Put_initMent(1)                        '// bg๎•๑•
    '                    Call LotData_ReDim(SMAC.S_MMAX + 5, 0)      '// NA
    '                    Call frmLOT_Init(0)                         '// ๆ–ส๚ป
    '                    p_LotMax = SMAC.S_MMAX                      '// Zbgล‘ๅ”
    '                    Call Get_initMent(1)                        '// bg•๎•๑ๆ“พ
    '                    Call frmLOT_AllRead()                         '// bg•๎•๑ๆ“พ(•\ฆ)
    '                    gOperator.strOperatorCode = ""
    '                    '// bg๎•๑ฬm”F
    '                    '            If wkINP_Flg1 = 0 And gLot1_HD(0).intStatus1 = 1 Then
    '                    If wkINP_Flg1 = 0 And (gLot1_HD(0).intStatus1 = 1 Or ZureData(0).intStatus = 1) Then    '2006.01.26
    '                        '// I—น—ชe—ฬ๊
    '                        If SINI.MCN_FLG3 = 1 Then
    '                            '// ย•ส“o^พมฝ็am^k”m“—อ
    '                            wkINP_No = defINPBoxNo
    '                        Else
    '                            '// ๊“o^พมฝ็์ฦา“—อ
    '                            wkINP_No = defINPOpeNM
    '                        End If
    '                        If ZureOperationCheck() <> 0 Then           '2006.02.03
    '                            gOperator.strOperatorCode = wkOperator
    '                            Me.txtOpeNM.Caption = wkOperator
    '                            For Lp = 0 To p_LotMax
    '                                If gOpeData(Lp).strLotNo <> "" Then
    '                                    gOpeData(Lp).strOperatorCode = wkOperator
    '                                End If
    '                            Next Lp
    '                            wkINP_No = defINPKazu                   '–”“—อึ
    '                            wkINP_No = defINPPieceG '###
    '                        End If
    '                        p_LotNo = 0
    '                        Call frmLOT_Disp2(p_LotNo, 1)
    '                    Else
    '                        wkINP_No = defINPBoxNo
    '                    End If
    '                    Call frmLOT_Disp()
    '                Else
    '                    Call PutLogFile(p_LogMode, "Save=NG", 2)   '// O‘
    '                    Call PutErrLogFile(p_LogMode, "Start Save=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                    If wkRes = F_FlowChange Then
    '                        If g_StatusError = 0 Then
    '                            Call ErrMsg(frmLOT, 1044, "")       '2008.09.12
    '                        Else
    '                            Call ErrMsg(frmLOT, 1045, "")
    '                        End If
    '                    Else
    '                        Call ErrMsg(frmLOT, 1008, "")
    '                    End If
    '                    Exit Function
    '                End If

    '            Case defINPEnd
    '                '// ์ฦI—น

    '                '// 2007.10.31  oืฬm”F
    '                If wkInpstr = "Y" And gLot1_HD(p_LotNo).intStatus2 = 2 And (gLot1_DT(p_LotNo).intOpeSeq = gLot1_DT(p_LotNo).intNOpeSeq) Then
    '                    strWork = lblMSG.Caption
    '                    lblMSG.Caption = "input password for quality failure"
    '                    If PasswordInput(p_QualityPass) <> 0 Then
    '                        wkInpstr = "N"
    '                    End If
    '                    lblMSG.Caption = strWork
    '                End If

    '                '// 2007.08.31
    '                p_OKNG = ""
    '                If wkInpstr = "Y" And p_ChekMode = 1 Then
    '                    strWork = lblMSG.Caption
    '            lblMSG.Caption = "judgement(" & p_CmdGood & "/" & p_CmdNoGood & "/" & p_CmdJump & ")๐“—อตฤญพณขB"
    '                    If InspectionInput(p_OKNG) <> 0 Then
    '                        wkInpstr = "N"
    '                    End If
    '                    lblMSG.Caption = strWork
    '                End If

    '                If wkInpstr = "N" Then
    '                    '2006.02.03
    '                    If ZureOperationCheck() <> 0 Then
    '                        wkINP_No = defINPKazu
    '                        wkINP_No = defINPPieceG '###
    '                        Call frmLOT_Disp()
    '                        Exit Function
    '                    End If

    '                    '// ์ฦๆม
    '                    Call PutLogFile(p_LogMode, "End Cancel", 2)       '// O‘
    '                    If SMAC.S_MAIN = "" Then
    '                        wkINP_No = defINPSetNo
    '                    ElseIf SINI.MCN_FLG2 = 1 Then
    '                        '// A‘ฑI—นฬ๊
    '                        wkINP_No = defINPKazu
    '                        wkINP_No = defINPPieceG '###
    '                    Else
    '                        wkINP_No = defINPBoxNo
    '                    End If
    '                    Call frmLOT_Disp()
    '                    If SINI.MCN_FLG3 = 1 Or SINI.MCN_FLG2 = 1 Then
    '                        gOpeData(p_LotNo).intOpePiece = gOpeData(p_LotNo).intOpePieceMax
    '                        txtKkazu.Caption = gOpeData(p_LotNo).intOpePiece
    '                        txtKkazu.Caption = gOpeData(p_LotNo).intPieceGood '###
    '                        '###
    '                        txtKkazuBad.Caption = gOpeData(p_LotNo).intPieceBad
    '                        txtFKkazu.Caption = gOpeData(p_LotNo).intFrmPieceOK
    '                        txtFKkazuBad.Caption = gOpeData(p_LotNo).intFrmPieceNG
    '                        '###
    '                        '                gOpeData(p_LotNo).intOpePiece = 0
    '                    End If
    '                    Exit Function
    '                End If
    '                If wkInpstr <> "Y" Then Exit Function

    '                '// I—น‘
    '                '*******************
    '                '' ิฬ’ฒฎB
    '                '*******************
    '                Call Get_ConnectTimer(2)
    '                If (gOpeData(p_LotNo).intCFlg >= 1 And gOpeData(p_LotNo).intCFlg <= 3) And gOpeData(p_LotNo).intCFlg2 = 0 And gOpeData(p_LotNo).intStatus1 = 0 Then
    '                    '// ““—(๊A—)
    '                    wkWTFlg = True
    '                    wkAns = 0
    '                    If SINI.MCN_FLG3 = 1 Then
    '                        wkAns = p_LotNo + 1
    '                    End If
    '                    If gOpeData(p_LotNo).intCFlg = 1 Then
    '                        wkSTR = "(start)"
    '                    ElseIf gOpeData(p_LotNo).intCFlg = 2 Then
    '                        wkSTR = "(shipping)"
    '                    ElseIf gOpeData(p_LotNo).intCFlg = 3 Then
    '                        wkSTR = "(flow)"
    '                    End If
    '                    Call PutLogFile(p_LogMode, "Auto Save" & wkSTR, 2)          '// O‘
    '                    wkRes = PutLotEnd_Ichi(p_LotMax, gLot1_HD(), gLot1_DT(), gLot1_DL(), gOpeData(), gLayerData(), wkAns)
    '                    wkWTFlg = False
    '                Else
    '                    '// I—น‘(’สํ—)
    '                    wkWTFlg = True
    '                    wkAns = 0
    '                    If SINI.MCN_FLG3 = 1 Then
    '                        wkAns = p_LotNo + 1
    '                    End If
    '                    wkSTR = ""
    '                    If gOpeData(p_LotNo).intCFlg = 1 Then
    '                        wkSTR = "(start)"
    '                    ElseIf gOpeData(p_LotNo).intCFlg = 2 Then
    '                        wkSTR = "(shipping)"
    '                    ElseIf gOpeData(p_LotNo).intCFlg = 3 Then
    '                        wkSTR = "(flow)"
    '                    End If
    '                    Call PutLogFile(p_LogMode, "End Save" & wkSTR, 2)       '// O‘
    '                    wkRes = PutLotEnd(p_LotMax, gLot1_HD(), gLot1_DT(), gLot1_DL(), gOpeData(), gLayerData(), wkAns)
    '                    wkWTFlg = False
    '                End If
    '                If wkRes = F_Success Then
    '                    '// ณํษ—ชฎ—นตฝ็H’๖ฬ•\ฆ๐sคB
    '                    Call frmLOT_FProcess(defDBStart)    '// H’๖—
    '                    If SMAC.S_MAIN = "" Then
    '                        wkINP_No = defINPSetNo          '// ZbgNoชศฏ๊ฮZbgNo“—อฉ็
    '                    Else
    '                        wkINP_No = defINPBoxNo          '// ’สํamฬ“—อึ
    '                    End If
    '                    Call frmLOT_Disp()
    '                Else
    '                    If wkRes = F_FlowChange Then
    '                        If g_StatusError = 0 Then
    '                            Call ErrMsg(frmLOT, 1044, "")       '2008.09.12
    '                        Else
    '                            Call ErrMsg(frmLOT, 1045, "")
    '                        End If
    '                    ElseIf wkRes <> 99 Then                     'Y‘ชH’๖ศO
    '                        Call ErrMsg(frmLOT, 1008, "")
    '                    End If
    '                    Call PutLogFile(p_LogMode, "Save" & wkSTR & "=NG", 2)  '// O‘
    '                    Call PutErrLogFile(p_LogMode, "Start Save" & wkSTR & "=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                    Exit Function
    '                End If
    '                '###
    '            Case defINPPieceG
    '                ' GoodPiece“—อ
    '                '// –”“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                    End If
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gOpeData(p_LotNo).intPieceGood '###
    '                End If
    '                '2006.04.13
    '                If ZureOperationCheck() = 1 Then
    '                    gOpeData(p_LotNo).intPieceGood = Val(wkInpstr) '###
    '                End If
    '                wkCHK_Flg3 = False
    '                txtKkazu.Caption = wkInpstr

    '                If IsNumeric(wkInpstr) = False Then
    '                    '// “—องภ”๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1009, "")
    '                    Call SetControlColor(txtKkazu, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                    '###ElseIf (Val(wkInpstr) >= 0 And Val(wkInpstr) <= gOpeData(p_LotNo).intPieceGood) = False Then
    '                ElseIf (Val(wkInpstr) >= 0) = False Then
    '                    '// “—อ”ออ๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1010, "")
    '                    Call SetControlColor(txtKkazu, defErrorMode)
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                '###gOpeData(p_LotNo).intLossPiece = gOpeData(p_LotNo).intOpePieceMax - Val(wkInpstr)
    '                gOpeData(p_LotNo).intPieceGood = Val(wkInpstr) '###

    '                If gOpeData(p_LotNo).intRepeatFlg = 1 Or gOpeData(p_LotNo).intRepeatFlg = 3 Or gOpeData(p_LotNo).intRepeatFlg = 5 Then
    '                    wkSTR = "Repeat "
    '                ElseIf gOpeData(p_LotNo).intRepeatFlg = 2 Then
    '                    wkSTR = "Pilot "
    '                End If
    '                Call PutLogFile(p_LogMode, "GOOD-Quantity" & (p_LotNo + 1) & " =" & wkInpstr & "/ " & gOpeData(p_LotNo).intPieceGood, 2)  '// O‘
    '                '// “—อ€–ฺ—
    '                If (True) Then
    '                    wkINP_No = defINPPieceB
    '                End If
    '            Case defINPPieceB
    '                ' BadPiece“—อ
    '                '// –”“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                    End If
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gOpeData(p_LotNo).intPieceBad '###
    '                End If
    '                '2006.04.13
    '                If ZureOperationCheck() = 1 Then
    '                    gOpeData(p_LotNo).intPieceBad = Val(wkInpstr) '###
    '                End If
    '                wkCHK_Flg3 = False
    '                txtKkazuBad.Caption = wkInpstr

    '                If IsNumeric(wkInpstr) = False Then
    '                    '// “—องภ”๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1009, "")
    '                    Call SetControlColor(txtKkazuBad, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                    '###ElseIf (Val(wkInpstr) >= 0 And Val(wkInpstr) <= gOpeData(p_LotNo).intPieceBad) = False Then
    '                ElseIf (Val(wkInpstr) >= 0) = False Then
    '                    '// “—อ”ออ๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1010, "")
    '                    Call SetControlColor(txtKkazuBad, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                '###gOpeData(p_LotNo).intLossPiece = gOpeData(p_LotNo).intPieceBad - Val(wkInpstr)
    '                gOpeData(p_LotNo).intPieceBad = Val(wkInpstr) '###

    '                If gOpeData(p_LotNo).intRepeatFlg = 1 Or gOpeData(p_LotNo).intRepeatFlg = 3 Or gOpeData(p_LotNo).intRepeatFlg = 5 Then
    '                    wkSTR = "Repeat "
    '                ElseIf gOpeData(p_LotNo).intRepeatFlg = 2 Then
    '                    wkSTR = "Pilot "
    '                End If
    '                Call PutLogFile(p_LogMode, "NG-Quantity" & (p_LotNo + 1) & " =" & wkInpstr & "/ " & gOpeData(p_LotNo).intPieceBad, 2)  '// O‘
    '                '// “—อ€–ฺ—
    '                If p_LotCount >= p_LotNo + 1 Then
    '                    If wkINP_Flg1 = 1 Or wkINP_Flg1 = 2 Or SINI.MCN_FLG2 = 1 Then
    '                        If wkINP_Flg1 = 1 Then              '2003.08.01
    '                            '
    '                            '   ““A“—อ€–ฺอAEFn[–ผษท้
    '                            wkINP_No = defINPWaferN
    '                        Else
    '                            '
    '                            '   ““ศOอAEFn[–ผอ•\ฆฬฦศ้B
    '                            wkINP_No = defINPEnd
    '                            gOpeData(p_LotNo).intCFlg = IIf(SINI.MCN_FLG2 = 1, 3, gOpeData(p_LotNo).intCFlg)
    '                            gLot1_HD(p_LotNo).intStatus1 = 1
    '                            wkINP_Flg1 = 0
    '                        End If
    '                    ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                        If gLot1_HD(p_LotNo + 1).strLotNo <> "" And SINI.MCN_FLG3 = 0 Then
    '                            p_LotNo = p_LotNo + 1
    '                            Call frmLOT_Disp2(p_LotNo, 1)
    '                        ElseIf ZureOperationCheck() = 1 Then        '2010.01.06
    '                            If gLot1_DT(p_LotNo).strLayNo = "" Then
    '                                wkINP_No = defINPFILE
    '                            Else
    '                                wkINP_No = defINPFPCS
    '                            End If
    '                        Else
    '                            wkINP_No = defINPEnd
    '                        End If
    '                    Else
    '                        wkINP_No = defINPStart
    '                    End If
    '                ElseIf SINI.MCN_FLG3 = 1 Then
    '                    '// I—น—ลย•สI—นฬ๊A”—ส“—อใI—นm”F
    '                    wkINP_No = defINPEnd
    '                Else
    '                    wkINP_No = defINPBoxNo
    '                    gLot1_HD(p_LotNo).intStatus1 = 1
    '                    wkINP_Flg1 = 0
    '                End If
    '                '### Frame“—อช ้๊GoodFrame“—อึ
    '                If (SINI.FRM_INP = 1) Then
    '                    If (wkINP_No = defINPEnd) Then
    '                        wkINP_No = defINPFrmQtyG
    '                    End If
    '                End If

    '            Case defINPFrmQtyG     '// Goodt[€“—อ
    '                ' GoodPiece“—อ
    '                '// –”“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                    End If
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gOpeData(p_LotNo).intFrmPieceOK '###
    '                End If
    '                '2006.04.13
    '                If ZureOperationCheck() = 1 Then
    '                    gOpeData(p_LotNo).intFrmPieceOK = Val(wkInpstr) '###
    '                End If
    '                wkCHK_Flg3 = False
    '                txtFKkazu.Caption = wkInpstr

    '                If IsNumeric(wkInpstr) = False Then
    '                    '// “—องภ”๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1009, "")
    '                    Call SetControlColor(txtFKkazu, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                    '###ElseIf (Val(wkInpstr) >= 0 And Val(wkInpstr) <= gOpeData(p_LotNo).intPieceGood) = False Then
    '                ElseIf (Val(wkInpstr) >= 0) = False Then
    '                    '// “—อ”ออ๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1010, "")
    '                    Call SetControlColor(txtFKkazu, defErrorMode)
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                '###gOpeData(p_LotNo).intLossPiece = gOpeData(p_LotNo).intOpePieceMax - Val(wkInpstr)
    '                gOpeData(p_LotNo).intFrmPieceOK = Val(wkInpstr) '###

    '                If gOpeData(p_LotNo).intRepeatFlg = 1 Or gOpeData(p_LotNo).intRepeatFlg = 3 Or gOpeData(p_LotNo).intRepeatFlg = 5 Then
    '                    wkSTR = "Repeat "
    '                ElseIf gOpeData(p_LotNo).intRepeatFlg = 2 Then
    '                    wkSTR = "Pilot "
    '                End If
    '                Call PutLogFile(p_LogMode, "TotalProcessFrame" & (p_LotNo + 1) & " =" & wkInpstr & "/ " & gOpeData(p_LotNo).intFrmPieceOK, 2)   '// O‘
    '                '// “—อ€–ฺ—
    '                If (True) Then
    '                    wkINP_No = defINPFrmQtyB
    '                End If
    '            Case defINPFrmQtyB     '// Badt[€“—อ
    '                '// –”“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                    End If
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gOpeData(p_LotNo).intFrmPieceNG '###
    '                End If
    '                '2006.04.13
    '                If ZureOperationCheck() = 1 Then
    '                    gOpeData(p_LotNo).intFrmPieceNG = Val(wkInpstr) '###
    '                End If
    '                wkCHK_Flg3 = False
    '                txtFKkazuBad.Caption = wkInpstr

    '                If IsNumeric(wkInpstr) = False Then
    '                    '// “—องภ”๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1009, "")
    '                    Call SetControlColor(txtFKkazuBad, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                    '###ElseIf (Val(wkInpstr) >= 0 And Val(wkInpstr) <= gOpeData(p_LotNo).intPieceBad) = False Then
    '                ElseIf (Val(wkInpstr) >= 0) = False Then
    '                    '// “—อ”ออ๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1010, "")
    '                    Call SetControlColor(txtFKkazuBad, defErrorMode) '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                '###gOpeData(p_LotNo).intLossPiece = gOpeData(p_LotNo).intPieceBad - Val(wkInpstr)
    '                gOpeData(p_LotNo).intFrmPieceNG = Val(wkInpstr) '###

    '                If gOpeData(p_LotNo).intRepeatFlg = 1 Or gOpeData(p_LotNo).intRepeatFlg = 3 Or gOpeData(p_LotNo).intRepeatFlg = 5 Then
    '                    wkSTR = "Repeat "
    '                ElseIf gOpeData(p_LotNo).intRepeatFlg = 2 Then
    '                    wkSTR = "Pilot "
    '                End If
    '                Call PutLogFile(p_LogMode, "ScrapFrame" & (p_LotNo + 1) & " =" & wkInpstr & "/ " & gOpeData(p_LotNo).intFrmPieceNG, 2)  '// O‘
    '                '// “—อ€–ฺ—
    '                If p_LotCount >= p_LotNo + 1 Then
    '                    If wkINP_Flg1 = 1 Or wkINP_Flg1 = 2 Or SINI.MCN_FLG2 = 1 Then
    '                        If wkINP_Flg1 = 1 Then              '2003.08.01
    '                            '
    '                            '   ““A“—อ€–ฺอAEFn[–ผษท้
    '                            wkINP_No = defINPWaferN
    '                        Else
    '                            '
    '                            '   ““ศOอAEFn[–ผอ•\ฆฬฦศ้B
    '                            wkINP_No = defINPEnd
    '                            gOpeData(p_LotNo).intCFlg = IIf(SINI.MCN_FLG2 = 1, 3, gOpeData(p_LotNo).intCFlg)
    '                            gLot1_HD(p_LotNo).intStatus1 = 1
    '                            wkINP_Flg1 = 0
    '                        End If
    '                    ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                        If gLot1_HD(p_LotNo + 1).strLotNo <> "" And SINI.MCN_FLG3 = 0 Then
    '                            p_LotNo = p_LotNo + 1
    '                            Call frmLOT_Disp2(p_LotNo, 1)
    '                        ElseIf ZureOperationCheck() = 1 Then        '2010.01.06
    '                            If gLot1_DT(p_LotNo).strLayNo = "" Then
    '                                wkINP_No = defINPFILE
    '                            Else
    '                                wkINP_No = defINPFPCS
    '                            End If
    '                        Else
    '                            wkINP_No = defINPEnd
    '                        End If
    '                    Else
    '                        wkINP_No = defINPStart
    '                    End If
    '                ElseIf SINI.MCN_FLG3 = 1 Then
    '                    '// I—น—ลย•สI—นฬ๊A”—ส“—อใI—นm”F
    '                    wkINP_No = defINPEnd
    '                Else
    '                    wkINP_No = defINPBoxNo
    '                    gLot1_HD(p_LotNo).intStatus1 = 1
    '                    wkINP_Flg1 = 0
    '                End If
    '                '###
    '            Case defINPKazu
    '                '// –”“—อ
    '                If wkInpstr = "" Then
    '                    If wkCHK_Flg3 = True Then
    '                        wkCHK_Flg3 = False
    '                        '>> 2002.01.28 “—อNGฤ•\ฆ
    '                        '                txtKkazu.Caption = gOpeData(p_LotNo).intOpePieceMax'###
    '                        '                GoTo Skp_frmLOT_Input_11
    '                    End If
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gOpeData(p_LotNo).intOpePieceMax
    '                End If
    '                '2006.04.13
    '                If ZureOperationCheck() = 1 Then
    '                    gOpeData(p_LotNo).intOpePieceMax = Val(wkInpstr)
    '                End If
    '                wkCHK_Flg3 = False
    '                txtKkazu.Caption = wkInpstr

    '                If IsNumeric(wkInpstr) = False Then
    '                    '// “—องภ”๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1009, "")
    '                    Call SetControlColor(txtKkazu, defErrorMode)
    '                    '###
    '                    '            Call SetControlColor(txtKkazuBad, defErrorMode)
    '                    '            Call SetControlColor(txtFKkazu, defErrorMode)
    '                    '            Call SetControlColor(txtFKkazuBad, defErrorMode)
    '                    '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                ElseIf (Val(wkInpstr) >= 0 And Val(wkInpstr) <= gOpeData(p_LotNo).intOpePieceMax) = False Then
    '                    '// “—อ”ออ๐’ดฆฝ๊
    '                    Call ErrMsg(frmLOT, 1010, "")
    '                    Call SetControlColor(txtKkazu, defErrorMode)
    '                    '###
    '                    '            Call SetControlColor(txtKkazuBad, defErrorMode)
    '                    '            Call SetControlColor(txtFKkazu, defErrorMode)
    '                    '            Call SetControlColor(txtFKkazuBad, defErrorMode)
    '                    '###
    '                    wkCHK_Flg3 = True
    '                    Exit Function
    '                End If
    '                gOpeData(p_LotNo).intLossPiece = gOpeData(p_LotNo).intOpePieceMax - Val(wkInpstr)
    '                gOpeData(p_LotNo).intOpePiece = Val(wkInpstr)

    '                If gOpeData(p_LotNo).intRepeatFlg = 1 Or gOpeData(p_LotNo).intRepeatFlg = 3 Or gOpeData(p_LotNo).intRepeatFlg = 5 Then
    '                    wkSTR = "Repeat "
    '                ElseIf gOpeData(p_LotNo).intRepeatFlg = 2 Then
    '                    wkSTR = "Pilot "
    '                End If
    '                Call PutLogFile(p_LogMode, "Piece" & (p_LotNo + 1) & " =" & wkInpstr & "/ " & gOpeData(p_LotNo).intOpePieceMax, 2) '// O‘
    '                '// “—อ€–ฺ—
    '                If p_LotCount >= p_LotNo + 1 Then
    '                    If wkINP_Flg1 = 1 Or wkINP_Flg1 = 2 Or SINI.MCN_FLG2 = 1 Then
    '                        If wkINP_Flg1 = 1 Then              '2003.08.01
    '                            '
    '                            '   ““A“—อ€–ฺอAEFn[–ผษท้
    '                            wkINP_No = defINPWaferN
    '                        Else
    '                            '
    '                            '   ““ศOอAEFn[–ผอ•\ฆฬฦศ้B
    '                            wkINP_No = defINPEnd
    '                            gOpeData(p_LotNo).intCFlg = IIf(SINI.MCN_FLG2 = 1, 3, gOpeData(p_LotNo).intCFlg)
    '                            gLot1_HD(p_LotNo).intStatus1 = 1
    '                            wkINP_Flg1 = 0
    '                        End If
    '                    ElseIf gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                        If gLot1_HD(p_LotNo + 1).strLotNo <> "" And SINI.MCN_FLG3 = 0 Then
    '                            p_LotNo = p_LotNo + 1
    '                            Call frmLOT_Disp2(p_LotNo, 1)
    '                        ElseIf ZureOperationCheck() = 1 Then        '2010.01.06
    '                            If gLot1_DT(p_LotNo).strLayNo = "" Then
    '                                wkINP_No = defINPFILE
    '                            Else
    '                                wkINP_No = defINPFPCS
    '                            End If
    '                        Else
    '                            wkINP_No = defINPEnd
    '                        End If
    '                    Else
    '                        wkINP_No = defINPStart
    '                    End If
    '                ElseIf SINI.MCN_FLG3 = 1 Then
    '                    '// I—น—ลย•สI—นฬ๊A”—ส“—อใI—นm”F
    '                    wkINP_No = defINPEnd
    '                Else
    '                    wkINP_No = defINPBoxNo
    '                    gLot1_HD(p_LotNo).intStatus1 = 1
    '                    wkINP_Flg1 = 0
    '                End If
    'Skp_frmLOT_Input_11:
    '            Case defINPCheck2
    '                '// งภิตฐสฐJn
    '                '// งภิฤถ—Jn
    '                If wkInpstr = "Y" Then
    '                    wkRes = False
    '                    gRes = False
    '                    If lblCmd(defFNCRepeat).Caption = "" Then
    '                        wkRes = True
    '                        lblCmd(defFNCRepeat).Caption = "rework"
    '                    End If
    '                    Call lblCmd_Click(defFNCRepeat)
    '                    If wkRes = True Then
    '                        lblCmd(defFNCRepeat).Caption = ""
    '                    End If
    '                    If gRes = False Then
    '                        wkInpstr = "N"
    '                    End If
    '                End If
    '                If wkInpstr = "N" Then
    '                    wkINP_Flg4 = 0
    '                    '// ์ฦๆม
    '                    Call PutLogFile(p_LogMode, "End Cancel", 2)       '// O‘
    '                    If SMAC.S_MAIN = "" Then
    '                        wkINP_No = defINPSetNo
    '                    ElseIf SINI.MCN_FLG2 = 1 Then
    '                        '// A‘ฑI—นฬ๊
    '                        wkINP_No = defINPKazu
    '                        wkINP_No = defINPPieceG '###
    '                    Else
    '                        wkINP_No = defINPBoxNo
    '                    End If
    '                    Call frmLOT_DSet(p_LotNo, defDSet_Clear)        '// bg๎•๑NA
    '                    Call frmLOT_DSet(0, defDSet_Move)               '// —bgฺ“ฎ
    '                    Call Put_initMent(1)
    '                    p_LotMax = SMAC.S_MMAX
    '                    p_LotCount = 0
    '                    p_LotNo = 0
    '                    Call frmLOT_AllRead()
    '                    gOperator.strOperatorCode = ""
    '                    Call frmLOT_Disp2(p_LotNo, 1)
    '                    Call frmLOT_Disp()
    '                    Exit Function
    '                End If
    '            Case defINPCheck4
    '                '// ฤถt[’’fJn
    '                If wkInpstr = "N" Then
    '                    '// ทฌพูอณษ–฿ท
    '                    wkINP_No = wkINP_NoBk
    '                    wkINP_NoBk = defINPCancel
    '                    Call frmLOT_Disp()
    '                    Exit Function
    '                ElseIf wkInpstr = "Y" Then
    '                    '// ’’fJn
    '                    If gOpeData(p_LotNo).intRepeatFlg = 3 Then
    '                        wkAns = defDBStartP2
    '                    Else
    '                        wkAns = defDBStart
    '                    End If
    '                    wkRes = PutLotStop_Rep(gLot1_HD(p_LotNo), gLot1_DT(p_LotNo), gLot1_DL(p_LotNo), gOpeData(p_LotNo), gLayerData(p_LotNo), 0, 0)
    '                    If wkRes = F_Success Then
    '                        '// ณํษ—ชฎ—นตฝ็H’๖ฬ•\ฆ๐sคB
    '                        Call frmLOT_FProcess(wkAns)         '// H’๖—
    '                        If SMAC.S_MAIN = "" Then
    '                            wkINP_No = defINPSetNo          '// ZbgNoชศฏ๊ฮZbgNo“—อฉ็
    '                        Else
    '                            wkINP_No = defINPBoxNo          '// ’สํamฬ“—อึ
    '                        End If
    '                        Call frmLOT_Disp()
    '                    Else
    '                        Call ErrMsg(frmLOT, 1008, "")
    '                        Call PutLogFile(p_LogMode, "Save" & wkSTR & "=NG", 2)  '// O‘
    '                        Call PutErrLogFile(p_LogMode, "Start Save" & wkSTR & "=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                        Exit Function
    '                    End If
    '                End If
    '            Case defINPCheck5
    '                '// ฤถt[ฤ“o^Jn
    '                If wkInpstr = "N" Then
    '                    Call lblCmd_Click(defFNCCancel)
    '                    GoTo Skp_frmLOT_Input_01
    '                ElseIf wkInpstr <> "Y" Then
    '                    GoTo Skp_frmLOT_Input_01
    '                End If
    '                '-----------------------------------------
    '                '// Jn’ฬฤถAbg๎•๑I—น—
    '                wkWTFlg = True
    '                wkSTR = "(rework)"
    '                chkFlg = defDBStartR
    '                Call PutLogFile(p_LogMode, "Repeat Save" & wkSTR, 2)       '// O‘
    '                Call DispMsg(lblMSG, frmOKNG, 60, "")
    '                '-----------------------------------------
    '                '// Jn’ษฤถ๐sมฝ๊AI—น—๐sมฤฉ็ฤถ
    '                '// ฝพตA๊A—ลศข๊ฬ
    '                '        If wkCHK_Flg = 0 And SINI.MCN_FLG2 = 0 And gLot1_HD(p_LotNo).intStatus1 <> 0 Then
    '                '2007.08.07
    '                If wkCHK_Flg = 0 And SINI.MCN_FLG2 = 0 And gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                    '*******************
    '                    '// ฤถf[^ฬobNAbv
    '                    '*******************
    '                    Call CopyTOLot1Table(gLot1_HD(p_LotNo), bkLot1_h(1))
    '                    Call CopyTOLot1Data(gLot1_DT(p_LotNo), bkLot1_d(1))
    '                    Call CopyTOOpeData(gOpeData(p_LotNo), bkOpe_d(1))
    '                    '*******************
    '                    '// ฤถ‘Of[^ฬ•
    '                    '*******************
    '                    Call CopyTOLot1Table(bkLot1_h(0), gLot1_HD(p_LotNo))
    '                    Call CopyTOLot1Data(bkLot1_d(0), gLot1_DT(p_LotNo))
    '                    Call CopyTOOpeData(bkOpe_d(0), gOpeData(p_LotNo))
    '                    gOpeData(p_LotNo).strOperatorCode = gOperator.strOperatorCode
    '                    lpCnt = p_LotNo + 1
    '                    Call PutLogFile(p_LogMode, "End Save" & wkSTR, 2)       '// O‘
    '                    wkRes = PutLotEnd(p_LotMax, gLot1_HD(), gLot1_DT(), gLot1_DL(), gOpeData(), gLayerData(), lpCnt)
    '                    'wkRes = F_ReadErr
    '                    '*******************
    '                    '// ฤถf[^ฬ•
    '                    '*******************
    '                    Call CopyTOLot1Table(bkLot1_h(1), gLot1_HD(p_LotNo))
    '                    Call CopyTOLot1Data(bkLot1_d(1), gLot1_DT(p_LotNo))
    '                    Call CopyTOOpeData(bkOpe_d(1), gOpeData(p_LotNo))
    '                    If wkRes <> F_Success Then
    '                        Call PutLogFile(p_LogMode, "Save=NG", 2)            '// O‘
    '                        Call PutErrLogFile(p_LogMode, "Start Save=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                        Call ErrMsg(frmLOT, 1008, "")
    '                        wkWTFlg = False
    '                        Exit Function
    '                    End If
    '                    '// I—น—ใAฤถ
    '                    wkCHK_Flg = 3
    '                End If
    '                '-----------------------------------------
    '                '// ฤถt[Jn
    '                wkRes = PutLotStart_Rep(p_LotMax, gLot1_HD(p_LotNo), gLot1_DT(p_LotNo), gLot1_DL(p_LotNo), gOpeData(p_LotNo), gLayerData(p_LotNo), 0, wkCHK_Flg)
    '                'wkRes = F_ReadErr
    '                wkWTFlg = False
    '                If wkRes = F_Success Then
    '                    Call frmLOT_FProcess(chkFlg)        '// H’๖—
    '                    If SMAC.S_MAIN = "" Then
    '                        wkINP_No = defINPSetNo          '// ZbgNoชศฏ๊ฮZbgNo“—อฉ็
    '                    Else
    '                        wkINP_No = defINPBoxNo          '// ’สํamฬ“—อึ
    '                    End If
    '                    Call frmLOT_Disp()
    '                Else
    '                    Call PutLogFile(p_LogMode, "Save=NG", 2)   '// O‘
    '                    Call PutErrLogFile(p_LogMode, "Start Save=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                    Call ErrMsg(frmLOT, 1008, "")
    '                    Exit Function
    '                End If
    '            Case defINPCheck6
    '                '// pCbg“ฤ“o^Jn
    '                If wkInpstr = "N" Then
    '                    Call lblCmd_Click(defFNCCancel)
    '                ElseIf wkInpstr = "Y" Then
    '                    gRes = False
    '                    wkWTFlg = True
    '                    wkSTR = "(ส฿ฒฏฤ)"
    '                    chkFlg = defDBStartP
    '                    Call PutLogFile(p_LogMode, "Repeat Save" & wkSTR, 2)       '// O‘
    '                    Call DispMsg(lblMSG, frmOKNG, 60, "")
    '                    wkRes = PutLotStart_Plt(p_LotMax, gLot1_HD(p_LotNo), gLot1_DT(p_LotNo), gLot1_DL(p_LotNo), gOpeData(p_LotNo), gLayerData(p_LotNo), 0, 0)
    '                    wkWTFlg = False
    '                    If wkRes = F_Success Then
    '                        Call frmLOT_FProcess(chkFlg)        '// H’๖—
    '                        If SMAC.S_MAIN = "" Then
    '                            wkINP_No = defINPSetNo          '// ZbgNoชศฏ๊ฮZbgNo“—อฉ็
    '                        Else
    '                            wkINP_No = defINPBoxNo          '// ’สํamฬ“—อึ
    '                        End If
    '                        Call frmLOT_Disp()
    '                    Else
    '                        Call PutLogFile(p_LogMode, "Save=NG", 2)   '// O‘
    '                        Call PutErrLogFile(p_LogMode, "Start Save=NG", p_ErrTblNo, p_ErrTblValue)    '// O‘
    '                        Call ErrMsg(frmLOT, 1008, "")
    '                        Exit Function
    '                    End If
    '                End If
    '            Case defINPWaferN                   '2003.08.01
    '                '// EFn[–ผ“—อ
    '                If wkInpstr = "" Then
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gLot1_HD(p_LotNo).strMaterName
    '                End If
    '                gLot1_HD(p_LotNo).strMaterName = Mid$(wkInpstr, 1, 12)

    '                '// “—อ`FbN                2006.08.07
    '                '        wkAns = CheckWaferName(gOpeData(p_LotNo).strDviName, gLot1_HD(p_LotNo).strPrdName, gLot1_HD(p_LotNo).strMaterName)
    '                '### Assy
    '                wkAns = 0                       '2010.07.22
    '                If wkAns = 1 Then
    '                    Call PutLogFile(p_LogMode, "wrong wafer name", 2)
    '                    Call ErrMsg(frmLOT, 9999, "wrong wafer name")
    '                    Exit Function
    '                ElseIf wkAns <> 0 Then
    '                    Call PutLogFile(p_LogMode, "error to get master data", 2)
    '                    Call ErrMsg(frmLOT, 9999, "error to get master data")
    '                    Exit Function
    '                End If

    '                wkINP_No = defINPWaferL         'ฬ“—ออAEFn[bgNo“—อ
    '                Call frmLOT_Disp2(p_LotNo, 1)

    '            Case defINPWaferL                   '2003.08.01
    '                '// EFn[bgNo“—อ
    '                If wkInpstr = "" Then
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gLot1_HD(p_LotNo).strMaterSName
    '                End If
    '                gLot1_HD(p_LotNo).strMaterSName = Mid$(wkInpstr, 1, 15)
    '                wkINP_No = defINPMaker          'ฬ“—ออA[J[–ผ“—อ
    '                Call frmLOT_Disp2(p_LotNo, 1)

    '            Case defINPMaker                    '2003.08.01
    '                '// [J[“—อ
    '                If wkInpstr = "" Then
    '                    '// ๙’่’l“—อ
    '                    wkInpstr = gLot1_HD(p_LotNo).strMaker
    '                End If
    '                gLot1_HD(p_LotNo).strMaker = Mid$(wkInpstr, 1, 6)
    '                wkINP_No = defINPEnd
    '                gOpeData(p_LotNo).intCFlg = IIf(SINI.MCN_FLG2 = 1, 3, gOpeData(p_LotNo).intCFlg)
    '                gLot1_HD(p_LotNo).intStatus1 = 1
    '                Call frmLOT_Disp2(p_LotNo, 1)
    '                '        wkINP_Flg1 = 0                 '2003.08.09

    '            Case defINPFPCS                     '2010.01.06
    '                '// FPCSXVm”F
    '                If wkInpstr = "N" Then
    '                    Call PutLogFile(p_LogMode, "FPCSXVศต", 2)
    '                Else
    '                    Call PutLogFile(p_LogMode, "FPCSXV ่", 2)
    '                End If
    '                For Lp = 0 To 1
    '                    If wkInpstr = "N" Then
    '                        ZureData(Lp).intFPCSFlg = 0
    '                    Else
    '                        ZureData(Lp).intFPCSFlg = 1
    '                    End If
    '                Next Lp
    '                If wkInpstr = "Y" Or wkInpstr = "N" Then
    '                    wkINP_No = defINPFILE
    '                End If

    '            Case defINPFILE                     '2010.01.06
    '                '// ‘ช’่’l“]‘—m”F
    '                If wkInpstr = "N" Then
    '            Call PutLogFile(p_LogMode, "‘ช’่’l“]‘—ศต", 2)
    '                Else
    '            Call PutLogFile(p_LogMode, "‘ช’่’l“]‘— ่", 2)
    '                End If
    '                For Lp = 0 To 1
    '                    If wkInpstr = "N" Then
    '                        ZureData(Lp).intFileFlg = 0
    '                    Else
    '                        ZureData(Lp).intFileFlg = 1
    '                    End If
    '                Next Lp
    '                If wkInpstr = "Y" Or wkInpstr = "N" Then
    '                    '            wkINP_No = defINPStart
    '                    wkINP_No = defINPEnd
    '                End If

    'Skp_frmLOT_Input_01:
    '        End Select
    'Skp_frmLOT_Input:
    '        frmLOT_Input = 0

    '    End Function

    '    Public Sub frmLOT_Disp()
    '        Dim wkRes As Boolean
    '        Dim wkAns As Integer
    '        Dim flg As Boolean          '2007.08.07
    '        flg = False                     '2007.08.07
    '        wkRes = False
    '        Select Case wkINP_No
    '            Case defINPCancel
    '                lblMSG.Caption = ""
    '                wkRes = True
    '            Case defINPMach
    '                Call DispMsg(lblMSG, frmOKNG, 4, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtMachNM, defInputMode)
    '                wkRes = True
    '            Case defINPMcnCD
    '                Call DispMsg(lblMSG, frmOKNG, 4, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtMachNM, defInputMode)
    '                wkRes = True
    '            Case defINPSetNo
    '                Call DispMsg(lblMSG, frmOKNG, 5, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtSetNo, defInputMode)
    '                wkRes = True
    '            Case defINPBatchNo
    '                '// ob`mฬ“—อ
    '                Call DispMsg(lblMSG, frmOKNG, 80, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtBatch, defInputMode)
    '                wkRes = True
    '                flg = True      '2007.08.07
    '            Case defINPBoxNo
    '                '// bg๎•๑“—อ
    '                If wkINP_Flg1 = 1 Then
    '                    If p_LotCount >= SMAC.S_MMAX Then
    '                        '2007.08.07
    '                        If gLot1_HD(p_LotNo).intStatus1 = 1 Then
    '                            Call DispMsg(lblMSG, frmOKNG, 12, "")
    '                            Call SetControlColor(lblCmd(4), defCheckMode3)
    '                            Call SetControlColor(lblMSG, defCheckMode3)
    '                        Else
    '                            Call DispMsg(lblMSG, frmOKNG, 11, "")
    '                            Call SetControlColor(lblCmd(4), defCheckMode2)
    '                            Call SetControlColor(lblMSG, defCheckMode2)
    '                            flg = True      '2007.08.07
    '                        End If
    '                    Else
    '                        Call DispMsg(lblMSG, frmOKNG, 9, "")
    '                        Call frmLOT_AllClear()
    '                        Call SetControlColor(txtBoxNo, defInputMode)
    '                        flg = True      '2007.08.07
    '                    End If
    '                Else
    '                    Call frmLOT_AllClear()
    '                    If p_LotCount >= SMAC.S_MMAX Then
    '                        '2007.08.07
    '                        If (SINI.MCN_FLG3 = 0 And gLot1_HD(0).intStatus1 = 1) Or (SINI.MCN_FLG3 = 1 And gLot1_HD(p_LotNo).intStatus1 = 1) Then
    '                            Call DispMsg(lblMSG, frmOKNG, 12, "")
    '                            Call SetControlColor(lblCmd(4), defCheckMode3)
    '                            Call SetControlColor(lblMSG, defCheckMode3)
    '                        Else
    '                            Call DispMsg(lblMSG, frmOKNG, 11, "")
    '                            Call SetControlColor(lblCmd(4), defCheckMode2)
    '                            Call SetControlColor(lblMSG, defCheckMode2)
    '                            flg = True      '2007.08.07
    '                        End If
    '                        '2007.08.07
    '                    ElseIf p_LotCount > 0 And gLot1_HD(p_LotNo).intStatus1 = 1 And SINI.MCN_FLG3 = 1 Then
    '                        '// I—นย•ส“o^AI—นwฆ
    '                        Call DispMsg(lblMSG, frmOKNG, 12, "")
    '                        Call SetControlColor(lblCmd(4), defCheckMode3)
    '                        Call SetControlColor(lblMSG, defCheckMode3)
    '                    ElseIf p_LotCount > 0 And gLot1_HD(p_LotNo).intStatus1 = 0 And SINI.MCN_FLG6 = 1 And gLot1_HD(p_LotNo).strLotNo <> "" Then
    '                        '// Jnwฆ
    '                        Call DispMsg(lblMSG, frmOKNG, 11, "")
    '                        Call SetControlColor(lblCmd(4), defCheckMode2)
    '                        Call SetControlColor(lblMSG, defCheckMode2)
    '                        flg = True      '2007.08.07
    '                    Else
    '                        '// ’สํ“—อ
    '                        Call DispMsg(lblMSG, frmOKNG, 1, "")
    '                        Call SetControlColor(txtBoxNo, defInputMode)
    '                        Call SetControlColor(txtLotNo, defInputMode)
    '                        Call SetControlColor(txtBoxNo, defInputMode)
    '                        Call SetControlColor(lblCmd(4), defCheckMode4)
    '                        flg = True      '2007.08.07
    '                    End If
    '                End If
    '                wkRes = True
    '            Case defINPStart, defINPEnd
    '                Call DispMsg(lblMSG, frmOKNG, 20, "")
    '                Call frmLOT_AllClear()
    '                wkRes = True
    '            Case defINPCheck2
    '                Call frmLOT_AllClear()
    '                Call DispMsg(lblMSG, frmOKNG, 27, "")
    '                Call SetControlColor(lblMSG, defErrorMode2)
    '                wkRes = True
    '            Case defINPCheck4
    '                Call frmLOT_AllClear()
    '                Call DispMsg(lblMSG, frmOKNG, 28, "")
    '                wkRes = True
    '            Case defINPCheck5
    '                Call frmLOT_AllClear()
    '                Call DispMsg(lblMSG, frmOKNG, 21, "ฤถฬฐฬ")
    '                wkRes = True
    '            Case defINPCheck6
    '                Call frmLOT_AllClear()
    '                Call DispMsg(lblMSG, frmOKNG, 21, "ส฿ฒฏฤืฬ")
    '                wkRes = True
    '            Case defINPOpeNM
    '                Call DispMsg(lblMSG, frmOKNG, 3, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtOpeNM, defInputMode)
    '                flg = True      '2007.08.07
    '                wkRes = True
    '            Case defINPPieceG
    '                If wkINP_Flg1 = 1 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 100, "start")
    '                ElseIf wkINP_Flg1 = 2 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 100, "shipping")
    '                Else
    '                    Call DispMsg(lblMSG, frmOKNG, 100, "")
    '                End If
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtKkazu, defInputMode)
    '                Call SetControlColor(lblCmd(defFNCStEd), defCheckMode4)
    '                wkRes = True
    '            Case defINPPieceB
    '                If wkINP_Flg1 = 1 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 101, "start")
    '                ElseIf wkINP_Flg1 = 2 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 101, "shipping")
    '                Else
    '                    Call DispMsg(lblMSG, frmOKNG, 101, "")
    '                End If
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtKkazuBad, defInputMode)
    '                Call SetControlColor(lblCmd(defFNCStEd), defCheckMode4)
    '                wkRes = True
    '            Case defINPFrmQtyG '// Goodt[€“—อ
    '                If wkINP_Flg1 = 1 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 102, "start")
    '                ElseIf wkINP_Flg1 = 2 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 102, "shipping")
    '                Else
    '                    Call DispMsg(lblMSG, frmOKNG, 102, "")
    '                End If
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtFKkazu, defInputMode)
    '                Call SetControlColor(lblCmd(defFNCStEd), defCheckMode4)
    '                wkRes = True
    '            Case defINPFrmQtyB '// Badt[€“—อ
    '                If wkINP_Flg1 = 1 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 103, "start")
    '                ElseIf wkINP_Flg1 = 2 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 103, "shipping")
    '                Else
    '                    Call DispMsg(lblMSG, frmOKNG, 103, "")
    '                End If
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtFKkazuBad, defInputMode)
    '                Call SetControlColor(lblCmd(defFNCStEd), defCheckMode4)
    '                wkRes = True
    '                '###
    '            Case defINPKazu
    '                '// –”“—อ
    '                If wkINP_Flg1 = 1 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 10, "start")
    '                ElseIf wkINP_Flg1 = 2 Then
    '                    Call DispMsg(lblMSG, frmOKNG, 10, "shipping")
    '                Else
    '                    Call DispMsg(lblMSG, frmOKNG, 10, "")
    '                End If
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtKkazu, defInputMode)
    '                Call SetControlColor(lblCmd(defFNCStEd), defCheckMode4)
    '                wkRes = True
    '            Case defINPWaferN               '2003.08.01
    '                Call DispMsg(lblMSG, frmOKNG, 81, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtWaferN, defInputMode)
    '                wkRes = True
    '            Case defINPWaferL               '2003.08.01
    '                Call DispMsg(lblMSG, frmOKNG, 82, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtWaferL, defInputMode)
    '                wkRes = True
    '            Case defINPMaker                '2003.08.01
    '                Call DispMsg(lblMSG, frmOKNG, 83, "")
    '                Call frmLOT_AllClear()
    '                Call SetControlColor(txtMaker, defInputMode)
    '                wkRes = True
    '            Case defINPFPCS                 '2010.01.06
    '                Call DispMsg(lblMSG, frmOKNG, 92, "")
    '                Call frmLOT_AllClear()
    '                wkRes = True
    '            Case defINPFILE                 '2010.01.06
    '                Call DispMsg(lblMSG, frmOKNG, 93, "")
    '                Call frmLOT_AllClear()
    '                wkRes = True
    '        End Select
    '        If wkRes = True Then
    '            Call frmLOT_Header()
    '            If wkINP_Flg3 <> 0 Then
    '                txtSetNo.Caption = rOpeData(0).strSetNo
    '            ElseIf gLot1_HD(p_LotNo).strLotNo <> "" And SINI.MCN_FLG1 = 1 Then
    '                txtSetNo.Caption = gOpeData(p_LotNo).strSetNo
    '            Else
    '                txtSetNo.Caption = SMAC.S_MAIN
    '            End If
    '            lblLCnt.Caption = " " & p_LotCount & "/ " & CStr(SMAC.S_MMAX)
    '            lblLCnt2.Caption = IIf(p_LotCount <= 0, "", "[" & p_LotNo + 1 & "]")
    '            If wkINP_Flg2 = 1 Then
    '                lblCmd(defFNCCancel).Caption = defFDSP_CLS
    '            Else
    '                lblCmd(defFNCCancel).Caption = defFDSP_CLS
    '            End If
    '            If wkINP_No = defINPStart Or wkINP_No = defINPEnd Or p_LotCount >= 1 Or SINI.MCN_FLG5 = 0 Then
    '                lblCmd(defFNCMach).Caption = ""
    '            ElseIf SINI.MCN_FLG5 <> 0 And SINI.MCN_FLG7 <> 0 Then
    '                lblCmd(defFNCMach).Caption = ""
    '            ElseIf wkINP_No = defINPMach Then
    '                lblCmd(defFNCMach).Caption = defFDSP_INP
    '            Else
    '                lblCmd(defFNCMach).Caption = defFDSP_MAC
    '            End If
    '            If wkINP_No = defINPStart Or wkINP_No = defINPEnd Or SINI.MCN_FLG1 = 1 Or p_LotCount >= 1 Then
    '                lblCmd(defFNCSetNo).Caption = ""
    '            ElseIf wkINP_No = defINPSetNo Then
    '                lblCmd(defFNCSetNo).Caption = defFDSP_INP
    '            Else
    '                lblCmd(defFNCSetNo).Caption = defFDSP_SET
    '            End If
    '            '// ปZbg๎•๑
    '            If (wkINP_No = defINPKazu Or p_LotCount <= 0) Or (wkINP_No = defINPEnd And SINI.MCN_FLG3 = 1) Then
    '                '// ”—ส“—อAย•สI—น
    '                lblCmd(defFNCStEd).Caption = ""
    '            ElseIf (wkINP_No = defINPPieceG) Or (wkINP_No = defINPPieceB) Then '###
    '                '// ”—ส“—อAย•สI—น
    '                lblCmd(defFNCStEd).Caption = ""
    '            ElseIf (wkINP_No = defINPFrmQtyG) Or (wkINP_No = defINPFrmQtyB) Then '###
    '                '// ”—ส“—อAย•สI—น
    '                lblCmd(defFNCStEd).Caption = ""
    '            ElseIf wkINP_No = defINPOpeNM And SINI.MCN_FLG3 = 0 Then
    '                '// ”—ส“—อAย•สI—น
    '                lblCmd(defFNCStEd).Caption = ""
    '            ElseIf wkINP_Flg2 = 0 And SINI.MCN_FLG2 = 0 Then
    '                lblCmd(defFNCStEd).Caption = defFDSP_STA
    '            Else
    '                lblCmd(defFNCStEd).Caption = defFDSP_END
    '            End If
    '            '// —•ฯXiฤถj
    '            If Trim(gLot1_HD(p_LotNo).strLotNo) <> "" Or (gOpeData(p_LotNo).intRepeatFlg = 2 Or gOpeData(p_LotNo).intRepeatFlg = 4) And gOpeData(p_LotNo).intStatus1 = 0 Then
    '                '            lblCmd(defFNCRepeat).Caption = "rework"
    '                lblCmd(defFNCRepeat).Caption = ""           '2010.07.22
    '                '        ElseIf Trim(gLot1_HD(p_LotNo).strLotNo) = "" Or gOpeData(p_LotNo).intRepeatFlg <> 0 Or gOpeData(p_LotNo).intStatus1 <> 0 Then
    '                '2007.08.07
    '            ElseIf Trim(gLot1_HD(p_LotNo).strLotNo) = "" Or gOpeData(p_LotNo).intRepeatFlg <> 0 Or gOpeData(p_LotNo).intStatus1 = 1 Then
    '                lblCmd(defFNCRepeat).Caption = ""
    '            Else
    '                '            lblCmd(defFNCRepeat).Caption = "rework"
    '                lblCmd(defFNCRepeat).Caption = ""           '2010.07.22
    '            End If
    '            '// —•ฯXiส฿ฒฏฤj
    '            '        If Trim(gLot1_HD(p_LotNo).strLotNo) = "" Or gOpeData(p_LotNo).intRepeatFlg <> 0 Or gOpeData(p_LotNo).intStatus1 <> 0 Then
    '            '2007.08.07
    '            If Trim(gLot1_HD(p_LotNo).strLotNo) = "" Or gOpeData(p_LotNo).intRepeatFlg <> 0 Or gOpeData(p_LotNo).intStatus1 = 1 Then
    '                lblCmd(defFNCPilot).Caption = ""
    '            Else
    '                '            lblCmd(defFNCPilot).Caption = "pilot"
    '                lblCmd(defFNCPilot).Caption = ""            '2010.07.22
    '            End If
    '            '// —•ฯXi‘•’uภ’่j
    '            '        If Trim(gLot1_HD(p_LotNo).strLotNo) <> "" And gOpeData(p_LotNo).intStatus1 <> 0 Then
    '            '2007.08.07
    '            If Trim(gLot1_HD(p_LotNo).strLotNo) <> "" And gOpeData(p_LotNo).intStatus1 = 1 Then
    '                '            lblCmd(defFNCMLimit).Caption = "limit"
    '                lblCmd(defFNCMLimit).Caption = ""           '2010.07.22
    '            ElseIf Trim(gLot1_HD(p_LotNo).strLotNo) <> "" And (SINI.MCN_FLG2 <> 0 Or gOpeData(p_LotNo).intCFlg <> 0) = True Then
    '                '            lblCmd(defFNCMLimit).Caption = "limit"
    '                lblCmd(defFNCMLimit).Caption = ""           '2010.07.22
    '            Else
    '                lblCmd(defFNCMLimit).Caption = ""
    '            End If
    '            '// ‘•’u•\ฆ
    '            If Trim(gLot1_HD(p_LotNo).strLotNo) <> "" And gOpeData(p_LotNo).intCFlg_LE = 2 Then
    '                lblMachNM.Caption = "Limited Machine"
    '                Call SetControlColor(lblMachNM, defCheckMode2)
    '            Else
    '                lblMachNM.Caption = "Machine No."
    '                Call SetControlColor(lblMachNM, defCheckMode4)
    '            End If
    '        End If
    '        '// 2007.08.07
    '        '// xbZ[W
    '        If flg = True Then
    '            If gLot1_HD(p_LotNo).strLotNo <> "" And gLot1_HD(p_LotNo).strComment <> "" Then
    '                Call DispMsg(lblMSG, frmOKNG, 9999, gLot1_HD(p_LotNo).strComment)
    '                Call SetControlColor(lblMSG, defErrorMode2)
    '            End If
    '        End If
    '        '// 2010.12.10
    '        If p_ParallelMode = 1 Then
    '            lblParallel.Visible = True
    '            lblParallel.Caption = "Parallel Input"
    '        ElseIf p_ParallelMode = 2 Then
    '            lblParallel.Visible = True
    '            lblParallel.Caption = "Parallel Input"
    '        Else
    '            lblParallel.Visible = False
    '        End If
    '    End Sub

    '    Public gLot1_HD() As Lot1TableFormat

    '    Public Structure Lot1TableFormat
    '        Public strLotNo As String           '// bgm
    '        Public lngDviNo As Long             '// foCXm
    '        Public strPrdName As String           '// @ํ–ผ
    '        Public datInDay As Date             '// ““—\’่“๚
    '        Public datOutDay As Date             '// oื—\’่“๚
    '        Public intOpeSeq As Integer          '// H’๖”ิ
    '        Public intPrdPiece As Long             '// ถY–”
    '        Public intInpPiece As Long             '// ““–”
    '        Public intOutPiece As Long             '// oื–”(H’๖–”)
    '        Public datRealDay As Date             '// ’ส฿“๚
    '        Public strBatchNo As String           '// ob`m
    '        Public strBatchSub As String           '// ob`Tum
    '        Public strMaterName As String           '// —ฟ”ิ
    '        Public strMaterSName As String           '// —ฟ–ผ
    '        Public strMaker As String           '// —ฟ[J[
    '        Public intLevel As Integer          '// —Dๆx
    '        Public intStatus1 As Integer          '// Jn๓‘ิ(0=—–ณตA1=—’)
    '        Public intStatus2 As Integer          '// bg๓‘ิ(0=’สํค1=’โ~ค2=ูํ“ค4=ูํXgbv)
    '        Public intCycle As Integer          '// TCN”
    '        Public strBoxNo As String           '// {bNXm
    '        Public strPrvBoxNo As String           '// ๊—p{bNXm
    '        Public strWaferStatus As String           '// EFnว—m
    '        Public datLimitStrTime As Date             '// งภิ
    '        Public strWagonNo As String           '// S“m
    '        Public intRepOpeSeq As Integer          '// ฤถJnH’๖”ิ
    '        Public intRepeatQnt As Integer          '// ฤถ–”
    '        Public strTotalMask As String           '// g[^}XN
    '        Public intStatus2Save As Integer          '// bg๓‘ิ•‘ถ
    '        Public datRealStart As Date             '// bgJnิ
    '        Public datTroTime As Date             '// ูํ”ญถ“๚
    '        Public intSisakuFlg As Integer          '// ์bgtO(0:’สํA1:์A2:ภฑฤถ)
    '        Public strParentLot As String           '// •ชฬebgm
    '        Public strOpeArea As String           '// ์ฦGA
    '        Public intOpeChgFlg As Integer
    '        Public intPltOpeSeq1 As Integer
    '        Public intPltOpeSeq2 As Integer
    '        Public intPilotQnt As Integer
    '        Public intTotalQnt As Integer     
    '        Public strComment As String        
    '        Public strRohmOrderModename As String
    '        Public strOrderNo As String
    '        Public strFTModelName As String
    '        Public strTPRank As String
    '        Public intWariStockKbn As Integer
    '        Public intWariInstructKbn As Integer
    '        Public strFormName As String
    '        Public intGoodPieces As Long
    '        Public intBadPieces As Long
    '        '###
    '        '    intShipTargetChipCount          As Integer
    '        '    intOfficialChipCount            As Integer
    '        '    intRestLotInitialChipCount      As Integer
    '        '    intRestLotInitialInputMagazine  As Integer
    '        '    intRestLotInitialInputFrame     As Integer
    '        '    strSection                      As String
    '        '    strInvoiceNo                    As String
    '        '    intQCCheckDBNG                  As Integer
    '        '    intQCCheckDBPNashi              As Integer
    '        '    intQCCheckHajikiPullShaer       As Integer
    '        '    intQCCheckWBNG                  As Integer
    '        '    intQCCheckWBInsNG               As Integer
    '        '    intQCCheckOSNG                  As Integer
    '        '    intQCCheckJudge                 As Integer
    '        '###
    '        Public datCreationDate As Date
    '        Public strSendFlg As String
    '        Public datSendDate As Date
    '        Public strDviName As String
    '        Public strDviSub As String
    '    End Structure

End Class
