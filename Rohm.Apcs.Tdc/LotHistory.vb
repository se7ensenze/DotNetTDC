
''' <summary>
''' LOT2_DATA
''' </summary>
''' <remarks></remarks>
Public Class LotHistory

    Friend Sub New()
    End Sub

    Private m_MACHINE As String
    Private m_REAL_START As Date
    Private m_REAL_DAY As Date
    Private m_OPE_NAME As String
    Private m_OPERATOR1 As String
    Private m_OPERATOR2 As String
    Private m_GOOD_PIECES As Integer
    Private m_BAD_PIECES As Integer

    Public Property MACHINE() As String
        Get
            Return m_MACHINE
        End Get
        Friend Set(ByVal value As String)
            m_MACHINE = value
        End Set
    End Property

    ''' <summary>
    ''' LOT START TIME
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property REAL_START() As Date
        Get
            Return m_REAL_START
        End Get
        Friend Set(ByVal value As Date)
            m_REAL_START = value
        End Set
    End Property

    ''' <summary>
    ''' LOT END TIME
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property REAL_DAY() As Date
        Get
            Return m_REAL_DAY
        End Get
        Friend Set(ByVal value As Date)
            m_REAL_DAY = value
        End Set
    End Property

    ''' <summary>
    ''' Process Name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OPE_NAME() As String
        Get
            Return m_OPE_NAME
        End Get
        Friend Set(ByVal value As String)
            m_OPE_NAME = value
        End Set
    End Property

    Public Property OPERATOR1() As String
        Get
            Return m_OPERATOR1
        End Get
        Friend Set(ByVal value As String)
            m_OPERATOR1 = value
        End Set
    End Property

    Public Property OPERATOR2() As String
        Get
            Return m_OPERATOR2
        End Get
        Friend Set(ByVal value As String)
            m_OPERATOR2 = value
        End Set
    End Property

    Public Property GOOD_PIECES() As Integer
        Get
            Return m_GOOD_PIECES
        End Get
        Friend Set(ByVal value As Integer)
            m_GOOD_PIECES = value
        End Set
    End Property

    Public Property BAD_PIECES() As Integer
        Get
            Return m_BAD_PIECES
        End Get
        Friend Set(ByVal value As Integer)
            m_BAD_PIECES = value
        End Set
    End Property

End Class
