<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormWeeklyMPS
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
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.SSPPivotTable = New System.Windows.Forms.CheckBox()
        Me.WeeklyChart = New System.Windows.Forms.CheckBox()
        Me.MonthlyChart = New System.Windows.Forms.CheckBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.SeriesBottleNeck = New System.Windows.Forms.CheckBox()
        Me.SeriesSupplyPlan = New System.Windows.Forms.CheckBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.SeriesBottleNeckIF = New System.Windows.Forms.CheckBox()
        Me.SeriesSupplyPlanIF = New System.Windows.Forms.CheckBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.SeriesBottleNeckGRP = New System.Windows.Forms.CheckBox()
        Me.SeriesSupplyPlanGRP = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(85, 19)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(91, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(43, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Week"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(562, 312)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(114, 30)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Show Report"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(34, 49)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(45, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Supplier"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckedListBox1)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(502, 177)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Main Selection"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.CheckOnClick = True
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(85, 73)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(389, 64)
        Me.CheckedListBox1.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Family Display"
        '
        'ComboBox2
        '
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(85, 46)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(389, 21)
        Me.ComboBox2.TabIndex = 10
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.SSPPivotTable)
        Me.GroupBox2.Controls.Add(Me.WeeklyChart)
        Me.GroupBox2.Controls.Add(Me.MonthlyChart)
        Me.GroupBox2.Location = New System.Drawing.Point(520, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(158, 177)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Charts && Pivot Table Visible"
        '
        'SSPPivotTable
        '
        Me.SSPPivotTable.AutoSize = True
        Me.SSPPivotTable.Checked = True
        Me.SSPPivotTable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SSPPivotTable.Location = New System.Drawing.Point(6, 64)
        Me.SSPPivotTable.Name = "SSPPivotTable"
        Me.SSPPivotTable.Size = New System.Drawing.Size(104, 17)
        Me.SSPPivotTable.TabIndex = 2
        Me.SSPPivotTable.Text = "SSP Pivot Table"
        Me.SSPPivotTable.UseVisualStyleBackColor = True
        '
        'WeeklyChart
        '
        Me.WeeklyChart.AutoSize = True
        Me.WeeklyChart.Checked = True
        Me.WeeklyChart.CheckState = System.Windows.Forms.CheckState.Checked
        Me.WeeklyChart.Location = New System.Drawing.Point(6, 41)
        Me.WeeklyChart.Name = "WeeklyChart"
        Me.WeeklyChart.Size = New System.Drawing.Size(90, 17)
        Me.WeeklyChart.TabIndex = 1
        Me.WeeklyChart.Text = "Weekly Chart"
        Me.WeeklyChart.UseVisualStyleBackColor = True
        '
        'MonthlyChart
        '
        Me.MonthlyChart.AutoSize = True
        Me.MonthlyChart.Checked = True
        Me.MonthlyChart.CheckState = System.Windows.Forms.CheckState.Checked
        Me.MonthlyChart.Location = New System.Drawing.Point(6, 18)
        Me.MonthlyChart.Name = "MonthlyChart"
        Me.MonthlyChart.Size = New System.Drawing.Size(91, 17)
        Me.MonthlyChart.TabIndex = 0
        Me.MonthlyChart.Text = "Monthly Chart"
        Me.MonthlyChart.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.SeriesBottleNeck)
        Me.GroupBox3.Controls.Add(Me.SeriesSupplyPlan)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 195)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(218, 66)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "All Families - Chart Series Exception"
        '
        'SeriesBottleNeck
        '
        Me.SeriesBottleNeck.AutoSize = True
        Me.SeriesBottleNeck.Checked = True
        Me.SeriesBottleNeck.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesBottleNeck.Location = New System.Drawing.Point(6, 42)
        Me.SeriesBottleNeck.Name = "SeriesBottleNeck"
        Me.SeriesBottleNeck.Size = New System.Drawing.Size(147, 17)
        Me.SeriesBottleNeck.TabIndex = 1
        Me.SeriesBottleNeck.Text = "Include Series Bottleneck"
        Me.SeriesBottleNeck.UseVisualStyleBackColor = True
        '
        'SeriesSupplyPlan
        '
        Me.SeriesSupplyPlan.AutoSize = True
        Me.SeriesSupplyPlan.Checked = True
        Me.SeriesSupplyPlan.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesSupplyPlan.Location = New System.Drawing.Point(6, 19)
        Me.SeriesSupplyPlan.Name = "SeriesSupplyPlan"
        Me.SeriesSupplyPlan.Size = New System.Drawing.Size(151, 17)
        Me.SeriesSupplyPlan.TabIndex = 0
        Me.SeriesSupplyPlan.Text = "Include Series Supply plan"
        Me.SeriesSupplyPlan.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(12, 267)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(502, 13)
        Me.TextBox1.TabIndex = 12
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.SeriesBottleNeckIF)
        Me.GroupBox4.Controls.Add(Me.SeriesSupplyPlanIF)
        Me.GroupBox4.Location = New System.Drawing.Point(236, 195)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(218, 66)
        Me.GroupBox4.TabIndex = 13
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Individual Family - Chart Series Exception"
        '
        'SeriesBottleNeckIF
        '
        Me.SeriesBottleNeckIF.AutoSize = True
        Me.SeriesBottleNeckIF.Checked = True
        Me.SeriesBottleNeckIF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesBottleNeckIF.Location = New System.Drawing.Point(6, 42)
        Me.SeriesBottleNeckIF.Name = "SeriesBottleNeckIF"
        Me.SeriesBottleNeckIF.Size = New System.Drawing.Size(147, 17)
        Me.SeriesBottleNeckIF.TabIndex = 1
        Me.SeriesBottleNeckIF.Text = "Include Series Bottleneck"
        Me.SeriesBottleNeckIF.UseVisualStyleBackColor = True
        '
        'SeriesSupplyPlanIF
        '
        Me.SeriesSupplyPlanIF.AutoSize = True
        Me.SeriesSupplyPlanIF.Checked = True
        Me.SeriesSupplyPlanIF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesSupplyPlanIF.Location = New System.Drawing.Point(6, 19)
        Me.SeriesSupplyPlanIF.Name = "SeriesSupplyPlanIF"
        Me.SeriesSupplyPlanIF.Size = New System.Drawing.Size(151, 17)
        Me.SeriesSupplyPlanIF.TabIndex = 0
        Me.SeriesSupplyPlanIF.Text = "Include Series Supply plan"
        Me.SeriesSupplyPlanIF.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.SeriesBottleNeckGRP)
        Me.GroupBox5.Controls.Add(Me.SeriesSupplyPlanGRP)
        Me.GroupBox5.Location = New System.Drawing.Point(460, 195)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(218, 66)
        Me.GroupBox5.TabIndex = 14
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Group Family - Chart Series Exception"
        '
        'SeriesBottleNeckGRP
        '
        Me.SeriesBottleNeckGRP.AutoSize = True
        Me.SeriesBottleNeckGRP.Checked = True
        Me.SeriesBottleNeckGRP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesBottleNeckGRP.Location = New System.Drawing.Point(6, 42)
        Me.SeriesBottleNeckGRP.Name = "SeriesBottleNeckGRP"
        Me.SeriesBottleNeckGRP.Size = New System.Drawing.Size(147, 17)
        Me.SeriesBottleNeckGRP.TabIndex = 1
        Me.SeriesBottleNeckGRP.Text = "Include Series Bottleneck"
        Me.SeriesBottleNeckGRP.UseVisualStyleBackColor = True
        '
        'SeriesSupplyPlanGRP
        '
        Me.SeriesSupplyPlanGRP.AutoSize = True
        Me.SeriesSupplyPlanGRP.Checked = True
        Me.SeriesSupplyPlanGRP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SeriesSupplyPlanGRP.Location = New System.Drawing.Point(6, 19)
        Me.SeriesSupplyPlanGRP.Name = "SeriesSupplyPlanGRP"
        Me.SeriesSupplyPlanGRP.Size = New System.Drawing.Size(151, 17)
        Me.SeriesSupplyPlanGRP.TabIndex = 0
        Me.SeriesSupplyPlanGRP.Text = "Include Series Supply plan"
        Me.SeriesSupplyPlanGRP.UseVisualStyleBackColor = True
        '
        'FormWeeklyMPS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 356)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FormWeeklyMPS"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Weekly MPS"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents SSPPivotTable As System.Windows.Forms.CheckBox
    Friend WithEvents WeeklyChart As System.Windows.Forms.CheckBox
    Friend WithEvents MonthlyChart As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents SeriesBottleNeck As System.Windows.Forms.CheckBox
    Friend WithEvents SeriesSupplyPlan As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents SeriesBottleNeckIF As System.Windows.Forms.CheckBox
    Friend WithEvents SeriesSupplyPlanIF As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents SeriesBottleNeckGRP As System.Windows.Forms.CheckBox
    Friend WithEvents SeriesSupplyPlanGRP As System.Windows.Forms.CheckBox
End Class
