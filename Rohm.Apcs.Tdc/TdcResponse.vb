Public Class TdcResponse

    Private m_Message As String
    Public Property Message() As String
        Get
            Return m_Message
        End Get
        Set(ByVal value As String)
            m_Message = value
        End Set
    End Property

    Private m_HasError As Boolean
    Public Property HasError() As Boolean
        Get
            Return m_HasError
        End Get
        Set(ByVal value As Boolean)
            m_HasError = value
        End Set
    End Property

    Private m_ErrorCode As String
    Public Property ErrorCode() As String
        Get
            Return m_ErrorCode
        End Get
        Set(ByVal value As String)
            m_ErrorCode = value
        End Set
    End Property

    Private m_ErrorMessage As String
    Public Property ErrorMessage() As String
        Get
            Return m_ErrorMessage
        End Get
        Set(ByVal value As String)
            m_ErrorMessage = value
        End Set
    End Property

    Private m_MCNo As String
    Public Property MCNo() As String
        Get
            Return m_MCNo
        End Get
        Set(ByVal value As String)
            m_MCNo = value
        End Set
    End Property

    Private m_LotNo As String
    Public Property LotNo() As String
        Get
            Return m_LotNo
        End Get
        Set(ByVal value As String)
            m_LotNo = value
        End Set
    End Property


End Class
