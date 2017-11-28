<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.RequestButton = New System.Windows.Forms.Button
        Me.StartButton = New System.Windows.Forms.Button
        Me.EndButton = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.MachineNoTextBox = New System.Windows.Forms.TextBox
        Me.LotNoTextBox = New System.Windows.Forms.TextBox
        Me.InputQtyTextBox = New System.Windows.Forms.TextBox
        Me.GoodQtyTextBox = New System.Windows.Forms.TextBox
        Me.NGQtyTextBox = New System.Windows.Forms.TextBox
        Me.OperatorNoTextBox = New System.Windows.Forms.TextBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'RequestButton
        '
        Me.RequestButton.Location = New System.Drawing.Point(148, 317)
        Me.RequestButton.Name = "RequestButton"
        Me.RequestButton.Size = New System.Drawing.Size(75, 23)
        Me.RequestButton.TabIndex = 0
        Me.RequestButton.Text = "Request"
        Me.RequestButton.UseVisualStyleBackColor = True
        '
        'StartButton
        '
        Me.StartButton.Location = New System.Drawing.Point(247, 317)
        Me.StartButton.Name = "StartButton"
        Me.StartButton.Size = New System.Drawing.Size(75, 23)
        Me.StartButton.TabIndex = 1
        Me.StartButton.Text = "Start"
        Me.StartButton.UseVisualStyleBackColor = True
        '
        'EndButton
        '
        Me.EndButton.Location = New System.Drawing.Point(346, 317)
        Me.EndButton.Name = "EndButton"
        Me.EndButton.Size = New System.Drawing.Size(75, 23)
        Me.EndButton.TabIndex = 2
        Me.EndButton.Text = "End"
        Me.EndButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(65, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Machine No"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(55, 66)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Lot No"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(44, 101)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Input Qty"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(38, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 13)
        Me.Label4.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(42, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(52, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Good Qty"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(52, 171)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(42, 13)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "NG Qty"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(29, 206)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(65, 13)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Operator No"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(35, 241)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(59, 13)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Start Mode"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(38, 276)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 13)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "End Mode"
        '
        'MachineNoTextBox
        '
        Me.MachineNoTextBox.Location = New System.Drawing.Point(148, 28)
        Me.MachineNoTextBox.Name = "MachineNoTextBox"
        Me.MachineNoTextBox.Size = New System.Drawing.Size(273, 20)
        Me.MachineNoTextBox.TabIndex = 13
        Me.MachineNoTextBox.Text = "WB-M-155"
        '
        'LotNoTextBox
        '
        Me.LotNoTextBox.Location = New System.Drawing.Point(148, 63)
        Me.LotNoTextBox.Name = "LotNoTextBox"
        Me.LotNoTextBox.Size = New System.Drawing.Size(273, 20)
        Me.LotNoTextBox.TabIndex = 13
        Me.LotNoTextBox.Text = "1718A2199V"
        '
        'InputQtyTextBox
        '
        Me.InputQtyTextBox.Location = New System.Drawing.Point(148, 98)
        Me.InputQtyTextBox.Name = "InputQtyTextBox"
        Me.InputQtyTextBox.Size = New System.Drawing.Size(273, 20)
        Me.InputQtyTextBox.TabIndex = 13
        '
        'GoodQtyTextBox
        '
        Me.GoodQtyTextBox.Location = New System.Drawing.Point(148, 133)
        Me.GoodQtyTextBox.Name = "GoodQtyTextBox"
        Me.GoodQtyTextBox.Size = New System.Drawing.Size(273, 20)
        Me.GoodQtyTextBox.TabIndex = 13
        Me.GoodQtyTextBox.Text = "2000"
        '
        'NGQtyTextBox
        '
        Me.NGQtyTextBox.Location = New System.Drawing.Point(148, 168)
        Me.NGQtyTextBox.Name = "NGQtyTextBox"
        Me.NGQtyTextBox.Size = New System.Drawing.Size(273, 20)
        Me.NGQtyTextBox.TabIndex = 13
        Me.NGQtyTextBox.Text = "20"
        '
        'OperatorNoTextBox
        '
        Me.OperatorNoTextBox.Location = New System.Drawing.Point(148, 203)
        Me.OperatorNoTextBox.Name = "OperatorNoTextBox"
        Me.OperatorNoTextBox.Size = New System.Drawing.Size(273, 20)
        Me.OperatorNoTextBox.TabIndex = 13
        Me.OperatorNoTextBox.Text = "004627"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Normal", "Parallel", "Parallel End", "ReRun"})
        Me.ComboBox1.Location = New System.Drawing.Point(148, 232)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 18
        Me.ComboBox1.Text = "Normal"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"Normal", "Reset", "Accumulate"})
        Me.ComboBox2.Location = New System.Drawing.Point(148, 267)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox2.TabIndex = 19
        Me.ComboBox2.Text = "Normal"
        '
        'Form1
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(468, 382)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.OperatorNoTextBox)
        Me.Controls.Add(Me.NGQtyTextBox)
        Me.Controls.Add(Me.GoodQtyTextBox)
        Me.Controls.Add(Me.InputQtyTextBox)
        Me.Controls.Add(Me.LotNoTextBox)
        Me.Controls.Add(Me.MachineNoTextBox)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.EndButton)
        Me.Controls.Add(Me.StartButton)
        Me.Controls.Add(Me.RequestButton)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RequestButton As System.Windows.Forms.Button
    Friend WithEvents StartButton As System.Windows.Forms.Button
    Friend WithEvents EndButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MachineNoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents LotNoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents InputQtyTextBox As System.Windows.Forms.TextBox
    Friend WithEvents GoodQtyTextBox As System.Windows.Forms.TextBox
    Friend WithEvents NGQtyTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OperatorNoTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox

End Class
