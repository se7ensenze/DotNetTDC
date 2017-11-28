Imports System.Drawing
Public Class TdcAlarmMessageForm
#Region "NameSpace"
    Private m_StartTime As DateTime
    Public Property StartTime() As DateTime
        Get
            Return m_StartTime
        End Get
        Set(ByVal value As DateTime)
            m_StartTime = value
        End Set
    End Property
#End Region
#Region "UI Operation"
    Private Sub TimerDuration_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerDuration.Tick
        Dim timTotal As TimeSpan = Now.Subtract(m_StartTime)
        lbDurationTime.Text = "DurationTime : " & timTotal.Hours.ToString("00") & ":" & _
                                                  timTotal.Minutes.ToString("00") & ":" & _
                                                  timTotal.Seconds.ToString("00")
        If lbTdc.ForeColor = Color.Lime Then
            lbTdc.ForeColor = Color.Black
            lbTdc.BackColor = Color.Lime
        ElseIf lbTdc.ForeColor = Color.Black Then
            lbTdc.ForeColor = Color.WhiteSmoke
            lbTdc.BackColor = Color.Red
        Else
            lbTdc.ForeColor = Color.Lime
            lbTdc.BackColor = Color.Black
        End If

    End Sub
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
    Private Sub ErrorMessageBox(ByVal errorCode As String, ByVal errorMessage As String, ByVal lotNo As String, ByVal info As LotInfo)
        If Not info Is Nothing Then
            lbStatus.Text = info.STATUS1.ToString
            lbProcess.Text = info.OPE_NAME
            dgvHistory.DataSource = info.LotHistories
        Else
            lbStatus.Text = " - "
            lbProcess.Text = " - "
            btnHistory.Visible = False
        End If
        lbLotNo.Text = lotNo
        lbAlarmMessage.Text = errorCode & vbCrLf & errorMessage
        Select Case errorCode
            Case "01"
                lbSolutionEng.Text = "Lot was not register ,Please contact administrator (Tel.83111)"
                lbSolutionThai.Text = "Lot ไม่ได้รับการลงทะเบียน ,กรุณาติดต่อผู้ดูแลระบบโทร.83111"
            Case "02"
                lbSolutionEng.Text = "Please clear Lot before input next lot"
                lbSolutionThai.Text = "กรุณาเคลียLot ก่อนอินพุต Lotต่อไป"
            Case "03"
                lbSolutionEng.Text = "Lot was not start ,Please "
                lbSolutionThai.Text = "Lot ยังไม่ถูกรัน ,กรุณาติดต่อผู้ดูแลระบบ โทร.83111"
            Case "04"
                lbSolutionEng.Text = "Machine was not register ,Please contact administrator (Tel.83111)"
                lbSolutionThai.Text = "เครื่องจักรไม่ได้รับการลงทะเบียน ,กรุณาติดต่อผู้ดูแลระบบ โทร.83111"
            Case "05"
                lbSolutionEng.Text = "Lot was hold by QC,Please contact QC"
                lbSolutionThai.Text = "Lot ถูกระงับโดย QC ,กรุณาติดต่อ QC"
            Case "06"
                lbSolutionEng.Text = "Please check or move this lot before run"
                lbSolutionThai.Text = "กรุณาตรวจสอบหรือย้ายข้อมูลของLotก่อนรัน"
            Case Else  '"70", "71", "72", "73"
                lbSolutionEng.Text = "Please contact Administrator (Tel.83111)"
                lbSolutionThai.Text = "กรุณาติดต่อผู้ดูแลระบบ โทร.83111"
        End Select
    End Sub
#End Region
    Public Sub New(ByVal errorCode As String, ByVal errorMessage As String, ByVal lotNo As String, ByVal info As LotInfo)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        m_StartTime = Now
        ' Add any initialization after the InitializeComponent() call.
        ErrorMessageBox(errorCode, errorMessage, lotNo, info)
    End Sub

    Private m_SmalSize As Size = New Size(483, 294)
    Private m_BigSize As Size = New Size(964, 294)

    Private Sub btnHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHistory.Click
        If Me.Size.Equals(m_SmalSize) Then
            Me.Size = m_BigSize
        Else
            Me.Size = m_SmalSize
        End If
    End Sub

End Class