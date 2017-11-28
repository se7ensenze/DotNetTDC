Imports Rohm.Apcs.Tdc
Imports System.Data.SqlClient

Public Class Form1

    Private m_TdcService As TdcService

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        m_TdcService = TdcService.GetInstance()
        m_TdcService.ConnectionString = "Data Source=10.28.33.11;Initial Catalog=APCSDB;Persist Security Info=True;User ID=sa;Password='p@$$w0rd'"

        'in case of you want to change format of log file
        m_TdcService.Logger = New TdcLoggerCsvWriter()


    End Sub

    Private Sub RequestButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RequestButton.Click

        Dim sm As RunModeType
        Select Case ComboBox1.Text
            Case "Normal"
                sm = RunModeType.Normal
            Case "Parallel"
                sm = RunModeType.Separated
            Case "Parallel End"
                sm = RunModeType.SeparatedEnd
            Case "ReRun"
                sm = RunModeType.ReRun
        End Select
        Dim res As TdcResponse = m_TdcService.LotRequest(MachineNoTextBox.Text, LotNoTextBox.Text, sm)
        If res.HasError Then
            Dim li As LotInfo = Nothing
            Try
                li = m_TdcService.GetLotInfo(res.LotNo, res.MCNo)
            Catch ex As Exception
            End Try
            Using dlg As TdcAlarmMessageForm = New TdcAlarmMessageForm(res.ErrorCode, res.ErrorMessage, res.LotNo, li)
                dlg.ShowDialog()
            End Using
        End If



    End Sub

    Private Sub StartButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StartButton.Click

        Dim sm As RunModeType = RunModeType.Normal
        Select Case ComboBox1.Text
            Case "Normal"
                sm = RunModeType.Normal
            Case "Parallel"
                sm = RunModeType.Separated
            Case "Parallel End"
                sm = RunModeType.SeparatedEnd
            Case "ReRun"
                sm = RunModeType.ReRun
        End Select

        Dim res As TdcResponse = m_TdcService.LotSet(MachineNoTextBox.Text, LotNoTextBox.Text, Now, OperatorNoTextBox.Text, sm)
        If res.HasError Then
            Dim li As LotInfo = Nothing
            Try
                li = m_TdcService.GetLotInfo(res.LotNo, res.MCNo)
            Catch ex As Exception
            End Try
            Using dlg As TdcAlarmMessageForm = New TdcAlarmMessageForm(res.ErrorCode, res.ErrorMessage, res.LotNo, li)
                dlg.ShowDialog()
            End Using
        End If
    End Sub

    Private Sub EndButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EndButton.Click
        Dim em As EndModeType = EndModeType.Normal

        Select Case ComboBox2.Text
            Case "Normal"
                em = EndModeType.Normal
            Case "Reset"
                em = EndModeType.AbnormalEndReset
            Case "Accumulate"
                em = EndModeType.AbnormalEndAccumulate
        End Select

        Dim res As TdcResponse = m_TdcService.LotEnd(MachineNoTextBox.Text, LotNoTextBox.Text, Now, _
                                            CInt(GoodQtyTextBox.Text), CInt(NGQtyTextBox.Text), em, OperatorNoTextBox.Text)
        If res.HasError Then
            Dim li As LotInfo = Nothing
            Try
                li = m_TdcService.GetLotInfo(res.LotNo, res.MCNo)
            Catch ex As Exception
            End Try
            Using dlg As TdcAlarmMessageForm = New TdcAlarmMessageForm(res.ErrorCode, res.ErrorMessage, res.LotNo, li)
                dlg.ShowDialog()
            End Using
        End If
    End Sub

End Class
