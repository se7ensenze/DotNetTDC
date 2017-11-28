<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TdcAlarmMessageForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.lbStatus = New System.Windows.Forms.Label
        Me.dgvHistory = New System.Windows.Forms.DataGridView
        Me.Label9 = New System.Windows.Forms.Label
        Me.lbAlarmMessage = New System.Windows.Forms.Label
        Me.TimerDuration = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.lbDurationTime = New System.Windows.Forms.ToolStripStatusLabel
        Me.Label5 = New System.Windows.Forms.Label
        Me.lbProcess = New System.Windows.Forms.Label
        Me.lbSolutionEng = New System.Windows.Forms.Label
        Me.lbLotNo = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.lbSolutionThai = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.lbTdc = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnHistory = New System.Windows.Forms.Button
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.Rohm.Apcs.Tdc.My.Resources.Resources.close
        Me.PictureBox1.Location = New System.Drawing.Point(18, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(42, 39)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 15
        Me.PictureBox1.TabStop = False
        '
        'lbStatus
        '
        Me.lbStatus.AutoSize = True
        Me.lbStatus.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbStatus.ForeColor = System.Drawing.Color.Firebrick
        Me.lbStatus.Location = New System.Drawing.Point(65, 56)
        Me.lbStatus.Name = "lbStatus"
        Me.lbStatus.Size = New System.Drawing.Size(0, 16)
        Me.lbStatus.TabIndex = 0
        '
        'dgvHistory
        '
        Me.dgvHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvHistory.Location = New System.Drawing.Point(498, 53)
        Me.dgvHistory.Name = "dgvHistory"
        Me.dgvHistory.Size = New System.Drawing.Size(442, 205)
        Me.dgvHistory.TabIndex = 14
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(11, 58)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(45, 13)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Status :"
        '
        'lbAlarmMessage
        '
        Me.lbAlarmMessage.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbAlarmMessage.ForeColor = System.Drawing.Color.Firebrick
        Me.lbAlarmMessage.Location = New System.Drawing.Point(65, 81)
        Me.lbAlarmMessage.Name = "lbAlarmMessage"
        Me.lbAlarmMessage.Size = New System.Drawing.Size(363, 54)
        Me.lbAlarmMessage.TabIndex = 0
        '
        'TimerDuration
        '
        Me.TimerDuration.Enabled = True
        Me.TimerDuration.Interval = 300
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lbDurationTime})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 272)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(483, 22)
        Me.StatusStrip1.TabIndex = 16
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lbDurationTime
        '
        Me.lbDurationTime.BackColor = System.Drawing.Color.Transparent
        Me.lbDurationTime.Name = "lbDurationTime"
        Me.lbDurationTime.Size = New System.Drawing.Size(77, 17)
        Me.lbDurationTime.Text = "DurationTime :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.Location = New System.Drawing.Point(5, 33)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 13)
        Me.Label5.TabIndex = 0
        Me.Label5.Text = "Process :"
        '
        'lbProcess
        '
        Me.lbProcess.AutoSize = True
        Me.lbProcess.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbProcess.ForeColor = System.Drawing.Color.Firebrick
        Me.lbProcess.Location = New System.Drawing.Point(65, 31)
        Me.lbProcess.Name = "lbProcess"
        Me.lbProcess.Size = New System.Drawing.Size(0, 16)
        Me.lbProcess.TabIndex = 0
        '
        'lbSolutionEng
        '
        Me.lbSolutionEng.AutoSize = True
        Me.lbSolutionEng.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbSolutionEng.ForeColor = System.Drawing.Color.Blue
        Me.lbSolutionEng.Location = New System.Drawing.Point(63, 9)
        Me.lbSolutionEng.Name = "lbSolutionEng"
        Me.lbSolutionEng.Size = New System.Drawing.Size(0, 16)
        Me.lbSolutionEng.TabIndex = 0
        '
        'lbLotNo
        '
        Me.lbLotNo.AutoSize = True
        Me.lbLotNo.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbLotNo.ForeColor = System.Drawing.Color.Firebrick
        Me.lbLotNo.Location = New System.Drawing.Point(65, 8)
        Me.lbLotNo.Name = "lbLotNo"
        Me.lbLotNo.Size = New System.Drawing.Size(0, 16)
        Me.lbLotNo.TabIndex = 0
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label7.Location = New System.Drawing.Point(5, 9)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(52, 13)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Solution :"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.LightCyan
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.lbSolutionThai)
        Me.Panel2.Controls.Add(Me.lbSolutionEng)
        Me.Panel2.Location = New System.Drawing.Point(18, 198)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(449, 58)
        Me.Panel2.TabIndex = 13
        '
        'lbSolutionThai
        '
        Me.lbSolutionThai.AutoSize = True
        Me.lbSolutionThai.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbSolutionThai.ForeColor = System.Drawing.Color.Blue
        Me.lbSolutionThai.Location = New System.Drawing.Point(63, 31)
        Me.lbSolutionThai.Name = "lbSolutionThai"
        Me.lbSolutionThai.Size = New System.Drawing.Size(0, 16)
        Me.lbSolutionThai.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.PapayaWhip
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lbStatus)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.lbAlarmMessage)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.lbProcess)
        Me.Panel1.Controls.Add(Me.lbLotNo)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(18, 53)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(449, 139)
        Me.Panel1.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 83)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Alarm :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "LotNo :"
        '
        'lbTdc
        '
        Me.lbTdc.AutoSize = True
        Me.lbTdc.BackColor = System.Drawing.Color.Black
        Me.lbTdc.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.lbTdc.ForeColor = System.Drawing.Color.Lime
        Me.lbTdc.Location = New System.Drawing.Point(135, 16)
        Me.lbTdc.Name = "lbTdc"
        Me.lbTdc.Size = New System.Drawing.Size(194, 23)
        Me.lbTdc.TabIndex = 11
        Me.lbTdc.Text = "Tdc Alarm Message"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(683, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 16)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "History"
        '
        'btnHistory
        '
        Me.btnHistory.Location = New System.Drawing.Point(392, 8)
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.Size = New System.Drawing.Size(75, 35)
        Me.btnHistory.TabIndex = 17
        Me.btnHistory.Text = "History -->"
        Me.btnHistory.UseVisualStyleBackColor = True
        '
        'AlarmMessage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(483, 294)
        Me.Controls.Add(Me.btnHistory)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.dgvHistory)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lbTdc)
        Me.Controls.Add(Me.Label3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "AlarmMessage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AlarmMessage"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lbStatus As System.Windows.Forms.Label
    Friend WithEvents dgvHistory As System.Windows.Forms.DataGridView
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lbAlarmMessage As System.Windows.Forms.Label
    Friend WithEvents TimerDuration As System.Windows.Forms.Timer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents lbDurationTime As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lbProcess As System.Windows.Forms.Label
    Friend WithEvents lbSolutionEng As System.Windows.Forms.Label
    Friend WithEvents lbLotNo As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lbTdc As System.Windows.Forms.Label
    Friend WithEvents lbSolutionThai As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnHistory As System.Windows.Forms.Button
End Class
