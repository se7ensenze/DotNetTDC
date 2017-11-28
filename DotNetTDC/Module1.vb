Module Module1

    Public g_MachineName As String
    Public g_SelfConNowFlag As Boolean
    Public p_LotNo As Integer
    Public g_ResCode_frmLOT_Input As Integer

    Public Structure OPE_DATA
        Public strLotNo As String                           '// bgm
        Public strPrdName As String                           '// @ํ–ผ
        Public strLayNo As String                           '// Cm
        Public strDviName As String                           '// foCX–ผ
        '###
        Public strFormName As String                           '// FORM_NAME 2010.08.24.Add
        '###
        Public strBatchNo As String                           '// ob`m >> 2001.08.28 Add
        Public strMachineName As String                           '// ‘•’u–ผ
        Public strMachineSub As String                           '// ‘•’uTu–ผ
        Public strOperatorCode As String                           '// ์ฦาR[h
        Public intRunStatus As Integer                          '// ‘•’uFา“ญ๓‘ิ
        Public intInlineFlg As Integer                          '// ‘•’uFC“C“tO
        Public strMachineIn As String                           '// ‘•’u–ผFC“C“ภ‘•’u
        Public intLimitFlag(1) As Integer                          '// งภI[o[tOP`Q
        '// “e๚
        Public intInteg(4) As Integer                          '// ฯZ’lP`S
        Public intIntegFlg(4) As Integer                          '// ฯZJE“gtO
        Public strRecipeCom As String                           '// ๐R“g
        Public strComment(12) As String                           '// R“g
        '// ฤถ๐
        Public strRepeatNM As String                           '// ฤถ์ฦR[h
        Public strRepeatCom As String                           '// ฤถ์ฦ–ผ()
        Public intRepeatFlg As Integer                          '// ฤถt[tO
        Public intRepeatTime As Integer                          '// J่•ิต๑”
        Public strWaferStatus As String                           '// EFnว—m
        Public strWaferStatusD As String                           '// EFnว—m(ูํ–”—p)
        Public intOpePieceMax As Integer                          '// ล‘ๅถY–”
        Public intOpePiece As Integer                          '// ์ฦถY–”
        Public intOpePieceD As Integer                          '// ์ฦถY–”(ูํ–”—p)
        '###
        Public intPieceGood As Long
        Public intPieceBad As Long
        Public intFrmPieceOK As Long
        Public intFrmPieceNG As Long
        '    intFrmPieceMax  As Long                          '// ล‘ๅFrame–”
        '    intFrmPiece     As Long                          '// ์ฦFrame–”
        '    intFrmPieceD    As Long                          '// ์ฦFrame–”iศใ–”—pj
        '###
        Public intWaferPiece As Integer                          '// EFnฤถ–”
        Public intLossPiece As Integer                          '// ”j‘น–”
        Public datRealDay As Date                             '// Jnิ
        Public intRepOpeSeq As Integer                          '// ฤถH’๖
        Public intPltOpeSeq As Integer                          '// pCbgI—นH’๖
        Public strLayNo_RP As String                           '// ฤถFC[m
        Public strOpe_RP(3) As String                           '// ฤถFH’๖๎•๑
        '// bg๓‘ิ
        Public intErrFlg As Integer                          '// ‘JฺtO(1,11=’โ~๓‘ิA2,22=ูํ๓‘ิj
        Public intErrSyochi As Integer                          '// ูํ’utO
        Public strErrMsg(6) As String                           '// ๓‘ิR“g
        Public intStatus1 As Integer                          '// Jn๓‘ิ
        Public intStatus2 As Integer                          '// bg๓‘ิ
        Public intStatus2Save As Integer                          '// bg๓‘ิ(‘O๎•๑)
        Public intOpeSeqST As Integer                          '// JnH’๖m
        Public intOpeNumber As Integer                          '// ์ฦJnFปฬส’u
        Public intCarryerFlg As Integer                          '// ๊—pLAtO
        Public strPrvBoxNo As String                           '// ๊—pLAm
        '// bg’่`๎•๑
        Public intCFlg As Integer                          '// —tO(0=’สํA1=““A2=oืA3=PA)
        Public intCFlg2 As Integer                          '// JnI—น(0=JnA1=I—น)
        Public intCFlg3 As Integer                          '// ‘—tO(0=–ข—A1=‘ฯ)
        Public intCFlg4 As Integer                          '// JntO(0=’สํA1=JnฯจJn‘Oษ•ฯX)
        Public intCFlg5 As Integer                          '// H’๖tO(0=’สํค1=ฤถJnค2=ฤถI—นค3=ส฿ฒฏฤJn
        '// @@@@@@ค4=’สํH’๖Jnค5=ส฿ฒฏฤฤJค6=‘S–””p
        '// @@@@@@ค7=““ค8=oื)
        Public strSetNo As String                           '// Zbg๎•๑
        '// ภ’่๎•๑
        Public intOpeSeq_LE As Integer                          '// ภ’่‘•’uFH’๖”ิ
        Public strLayNo_LE As String                           '// ภ’่‘•’uFC[m
        Public strOpe_LE(3) As String                           '// ภ’่‘•’uFH’๖๎•๑
        Public intCFlg_LE As Integer                          '// ภ’่‘•’uF0=ศตA1=JnA2=m”F
        '//
        Public intTotalQnt As Integer                          '// ‘–”
        '//
        Public strMode As String                           '// [h (P/N/R)          2006.01.25
    End Structure

End Module
