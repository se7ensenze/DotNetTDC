Public Interface ITdcLogger
    Property Enabled() As Boolean
    Sub SaveLog(ByVal msg As TdcMessage)
End Interface
