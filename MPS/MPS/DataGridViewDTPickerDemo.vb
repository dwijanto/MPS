Public Class DataGridViewDTPickerDemo
    Inherits Form

    Private dataGridView1 As New DataGridView()

    <STAThreadAttribute()> _
    Public Shared Sub Main()
        Application.Run(New DataGridViewDTPickerDemo())
    End Sub

    Public Sub New()
        Me.dataGridView1.Dock = DockStyle.Fill
        Me.Controls.Add(Me.dataGridView1)
        Me.Text = "DataGridView calendar column demo"
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) _
        Handles Me.Load

        Dim col As New CalendarColumn()
        Me.dataGridView1.Columns.Add(col)
        Me.dataGridView1.RowCount = 5
        Dim row As DataGridViewRow
        For Each row In Me.dataGridView1.Rows
            row.Cells(0).Value = DateTime.Now
        Next row

    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'DataGridViewDTPickerDemo
        '
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Name = "DataGridViewDTPickerDemo"
        Me.ResumeLayout(False)

    End Sub
End Class


