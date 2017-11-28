Imports Rohm.Apcs.Tdc
Imports System.IO

Public Class TdcLoggerCsvWriter
    Implements ITdcLogger

    Private m_Enabled As Boolean = True
    Private m_LogFolder As String
    Private m_LogFileName As String
    Private m_BeforeLogFileName As String

    Public Property Enabled() As Boolean Implements Rohm.Apcs.Tdc.ITdcLogger.Enabled
        Get
            Return m_Enabled
        End Get
        Set(ByVal value As Boolean)
            m_Enabled = value
        End Set
    End Property


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
                m_LogFileName = Path.Combine(m_LogFolder, "TDC.csv")
                m_BeforeLogFileName = Path.Combine(m_LogFolder, "TDC_before.csv")
            End If
        End Set
    End Property

    Public Sub New()
        Me.LogFolder = My.Application.Info.DirectoryPath
    End Sub

    Public Sub SaveLog(ByVal msg As Rohm.Apcs.Tdc.TdcMessage) Implements Rohm.Apcs.Tdc.ITdcLogger.SaveLog

        If Not m_Enabled Then
            Exit Sub
        End If


        Dim fi As FileInfo = New FileInfo(m_LogFileName)
        If fi.Exists AndAlso fi.Length > m_LogSize Then

            If File.Exists(m_BeforeLogFileName) Then
                File.Delete(m_BeforeLogFileName)
            End If

            File.Move(m_LogFileName, m_BeforeLogFileName)
            File.Delete(m_LogFileName)
        End If

 
        Using sw As StreamWriter = New StreamWriter(m_LogFileName, True)
            sw.WriteLine(Now.ToString("yyyy-MM-dd HH:mm:ss") & "," & _
                    msg.CodePosition & "," & _
                    msg.EventName & "," & _
                    msg.LotNo & "," & _
                    msg.MachineName & "," & _
                    msg.MessageText & "," & _
                    msg.MessageType & "," & _
                    msg.OPNo & "," & _
                    msg.Parameter & "," & _
                    msg.Status1 & "," & _
                    msg.Status2 & "," & _
                    msg.AppVersion)


        End Using

    End Sub
End Class
