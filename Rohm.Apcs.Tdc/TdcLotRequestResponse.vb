Public Class TdcLotRequestResponse
    Inherits TdcResponse

    Public Property BadPieces() As Integer
        Get
            Return m_BadPieces
        End Get
        Set(ByVal value As Integer)
            m_BadPieces = value
        End Set
    End Property
    Private m_BadPieces As Integer

    Public Property GoodPieces() As Integer
        Get
            Return m_GoodPieces
        End Get
        Set(ByVal value As Integer)
            m_GoodPieces = value
        End Set
    End Property
    Private m_GoodPieces As Integer

End Class
