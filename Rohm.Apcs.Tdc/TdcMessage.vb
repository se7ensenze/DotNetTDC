Public Structure TdcMessage

    Private m_EventName As String
    Public Property EventName() As String
        Get
            Return m_EventName
        End Get
        Set(ByVal value As String)
            m_EventName = value
        End Set
    End Property

    Private m_Parameter As String
    Public Property Parameter() As String
        Get
            Return m_Parameter
        End Get
        Set(ByVal value As String)
            m_Parameter = value
        End Set
    End Property

    Private m_MessageType As String
    Public Property MessageType() As String
        Get
            Return m_MessageType
        End Get
        Set(ByVal value As String)
            m_MessageType = value
        End Set
    End Property

    Private m_MachineName As String
    Public Property MachineName() As String
        Get
            Return m_MachineName
        End Get
        Set(ByVal value As String)
            m_MachineName = value
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

    Private m_Status1 As String
    Public Property Status1() As String
        Get
            Return m_Status1
        End Get
        Set(ByVal value As String)
            m_Status1 = value
        End Set
    End Property

    Private m_Status2 As String
    Public Property Status2() As String
        Get
            Return m_Status2
        End Get
        Set(ByVal value As String)
            m_Status2 = value
        End Set
    End Property

    Private m_OPNo As String
    Public Property OPNo() As String
        Get
            Return m_OPNo
        End Get
        Set(ByVal value As String)
            m_OPNo = value
        End Set
    End Property

    Private m_MessageText As String
    Public Property MessageText() As String
        Get
            Return m_MessageText
        End Get
        Set(ByVal value As String)
            m_MessageText = value
        End Set
    End Property

    Private m_CodePosition As String
    Public Property CodePosition() As String
        Get
            Return m_CodePosition
        End Get
        Set(ByVal value As String)
            m_CodePosition = value
        End Set
    End Property

    Private m_AppVersion As String
    Public Property AppVersion() As String
        Get
            Return m_AppVersion
        End Get
        Set(ByVal value As String)
            m_AppVersion = value
        End Set
    End Property

    Public Sub New(ByVal eventName As String, ByVal machineName As String, ByVal lotNo As String, ByVal opNo As String)
        m_EventName = eventName
        m_MachineName = machineName
        m_LotNo = lotNo
        m_OPNo = opNo

        m_AppVersion = AssemblyVersionUtil.GetAssemblyDescription()

    End Sub

End Structure
