Public Class LotInfo

    Friend Sub New()
    End Sub

    Private m_LOT_NO As String
    Public Property LOT_NO() As String
        Get
            Return m_LOT_NO
        End Get
        Friend Set(ByVal value As String)
            m_LOT_NO = value
        End Set
    End Property

    Private m_STATUS1 As StartingCondition
    Public Property STATUS1() As StartingCondition
        Get
            Return m_STATUS1
        End Get
        Friend Set(ByVal value As StartingCondition)
            m_STATUS1 = value
        End Set
    End Property

    Private m_STATUS2 As LotCondition
    Public Property STATUS2() As LotCondition
        Get
            Return m_STATUS2
        End Get
        Friend Set(ByVal value As LotCondition)
            m_STATUS2 = value
        End Set
    End Property

    Private m_OPE_NAME As String
    Public Property OPE_NAME() As String
        Get
            Return m_OPE_NAME
        End Get
        Friend Set(ByVal value As String)
            m_OPE_NAME = value
        End Set
    End Property

    Private m_MACHINE As String
    Public Property MACHINE() As String
        Get
            Return m_MACHINE
        End Get
        Set(ByVal value As String)
            m_MACHINE = value
        End Set
    End Property

    Private m_REAL_START As Date
    Public Property REAL_START() As Date
        Get
            Return m_REAL_START
        End Get
        Set(ByVal value As Date)
            m_REAL_START = value
        End Set
    End Property

    Private m_REAL_DAY As Date?
    Public Property REAL_DAY() As Date?
        Get
            Return m_REAL_DAY
        End Get
        Set(ByVal value As Date?)
            m_REAL_DAY = value
        End Set
    End Property

    Private m_GOOD_PIECES As Integer
    Public Property GOOD_PIECES() As Integer
        Get
            Return m_GOOD_PIECES
        End Get
        Set(ByVal value As Integer)
            m_GOOD_PIECES = value
        End Set
    End Property

    Private m_BAD_PIECES As Integer
    Public Property BAD_PIECES() As Integer
        Get
            Return m_BAD_PIECES
        End Get
        Set(ByVal value As Integer)
            m_BAD_PIECES = value
        End Set
    End Property

    Private m_OPERATOR1 As String
    Public Property OPERATOR1() As String
        Get
            Return m_OPERATOR1
        End Get
        Set(ByVal value As String)
            m_OPERATOR1 = value
        End Set
    End Property

    Private m_OPERATOR2 As String
    Public Property OPERATOR2() As String
        Get
            Return m_OPERATOR2
        End Get
        Set(ByVal value As String)
            m_OPERATOR2 = value
        End Set
    End Property

    Private m_Histories As LotHistory()
    Public Property LotHistories() As LotHistory()
        Get
            Return m_Histories
        End Get
        Friend Set(ByVal value As LotHistory())
            m_Histories = value
        End Set
    End Property

End Class
