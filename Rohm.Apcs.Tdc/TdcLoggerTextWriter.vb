Imports System.IO

Public Class TdcLoggerTextWriter
    Implements ITdcLogger

    Private m_Enabled As Boolean = True
    Private m_LogFolder As String
    Private m_LogFileName As String
    Private m_BeforeLogFileName As String

    Private m_LogSize As Integer = 2097152 '2 MB
    Public Property LogSize() As Integer
        Get
            Return m_LogSize
        End Get
        Set(ByVal value As Integer)
            m_LogSize = value
        End Set
    End Property

    Public Property LogFolder() As String
        Get
            Return m_LogFolder
        End Get
        Set(ByVal value As String)
            If Directory.Exists(value) Then
                m_LogFolder = value
                m_LogFileName = Path.Combine(m_LogFolder, "TDC.log")
                m_BeforeLogFileName = Path.Combine(m_LogFolder, "TDC_before.log")
            End If
        End Set
    End Property

    Public Sub New()
        Me.LogFolder = My.Application.Info.DirectoryPath
    End Sub

#Region "ITdcLogger's Members"

    Public Sub SaveLog(ByVal msg As TdcMessage) Implements ITdcLogger.SaveLog
        If Not m_Enabled Then
            Exit Sub
        End If

        Dim msgText As String = "EventName :" & msg.EventName & vbNewLine & vbTab & _
           "Parameter :" & msg.Parameter & vbNewLine & vbTab & _
           "MessageType :" & msg.MessageType & vbNewLine & vbTab & _
           "MachineName :" & msg.MachineName & vbNewLine & vbTab & _
           "LotNo :" & msg.LotNo & vbNewLine & vbTab & _
           "Status1 :" & msg.Status1 & vbNewLine & vbTab & _
           "Status2 :" & msg.Status2 & vbNewLine & vbTab & _
           "OPNo :" & msg.OPNo & vbNewLine & vbTab & _
           "CodePosition :" & msg.CodePosition & vbNewLine & vbTab & _
           "MessageText :" & msg.MessageText & vbNewLine & vbTab & _
           "Version :" & msg.AppVersion

        Dim fi As FileInfo = New FileInfo(m_LogFileName)
        If fi.Exists AndAlso fi.Length > m_LogSize Then

            If File.Exists(m_BeforeLogFileName) Then
                File.Delete(m_BeforeLogFileName)
            End If

            File.Move(m_LogFileName, m_BeforeLogFileName)
            File.Delete(m_LogFileName)
        End If

        Using sw As StreamWriter = New StreamWriter(m_LogFileName, True)
            sw.WriteLine(Now.ToString("yyyy/MM/dd HH:mm:ss") & vbTab & msgText)
        End Using
    End Sub

    Public Property Enabled() As Boolean Implements ITdcLogger.Enabled
        Get
            Return m_Enabled
        End Get
        Set(ByVal value As Boolean)
            m_Enabled = value
        End Set
    End Property

#End Region

End Class
