Imports System.Threading
Imports System.Text
Imports SSP.PublicClass
'Imports SupplierManagement.SharedClass

Public Class FormMonthlyTable
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim WithEvents MonthlyBS As BindingSource
    Dim CBBS As BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    'Dim DBadapter1 As New DBAdapter


    Private Sub FormSupplierCategory_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")

        DS = New DataSet

        Dim mymessage As String = String.Empty
        sb.Clear()

        sb.Append("select id,sm.monthly,yearweek::text as weeklyname,weekly from sspmonthlytable sm left join sspweekly sw on sw.sspweeklyid = sm.weekly;")
        sb.Append("select yearweek,monthly,sspweeklyid from sspweekly order by yearweek desc;")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Monthly"

            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            MonthlyBS = New BindingSource
                            CBBS = New BindingSource
                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns("id").AutoIncrement = True
                            DS.Tables(0).Columns("id").AutoIncrementSeed = 0
                            DS.Tables(0).Columns("id").AutoIncrementStep = -1

                            Dim monthlycol As DataColumn = DS.Tables(0).Columns("monthly")
                            Dim weeklycol As DataColumn = DS.Tables(0).Columns("weekly")


                            DS.Tables(0).Constraints.Add(New UniqueConstraint(monthlycol))
                            DS.Tables(0).Constraints.Add(New UniqueConstraint(weeklycol))


                            MonthlyBS.DataSource = DS.Tables(0)
                            CBBS.DataSource = DS.Tables(1)
                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = MonthlyBS
                            DataGridView1.RowTemplate.Height = 22

                            'TextBox1.DataBindings.Clear()

                            ComboBox1.DataBindings.Clear()

                            ComboBox1.DataSource = DS.Tables(1)
                            ComboBox1.DisplayMember = "yearweek"
                            ComboBox1.ValueMember = "sspweeklyid"
                            ComboBox1.DataBindings.Add(New Binding("SelectedValue", MonthlyBS, "weekly", True, DataSourceUpdateMode.OnPropertyChanged, ""))
                            If IsNothing(MonthlyBS.Current) Then
                                ComboBox1.SelectedIndex = -1
                            End If

                            'TextBox1.DataBindings.Add(New Binding("Text", MonthlyBS, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))


                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        loaddata()
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub MonthlyBS_ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Handles MonthlyBS.ListChanged
        ComboBox1.Enabled = Not IsNothing(MonthlyBS.Current)

    End Sub

    'Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
    '    DataGridView1.Invalidate()
    'End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Try
            MonthlyBS.AddNew()
            ComboBox1.SelectedIndex = -1
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        MonthlyBS.EndEdit()
        If Me.validate Then
            Try
                'get modified rows, send all rows to stored procedure. let the stored procedure create a new record.
                Dim ds2 As DataSet
                ds2 = DS.GetChanges

                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)

                    If Not DBadapter1.MonthlyTableTx(Me, mye) Then
                        MessageBox.Show(mye.message)
                        Exit Sub
                    End If
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    DataGridView1.Invalidate()
                    MessageBox.Show("Saved.")
                End If
            Catch ex As Exception
                MessageBox.Show(" Error:: " & ex.Message)
            End Try
        End If
        DataGridView1.Invalidate()
    End Sub



    Public Overloads Function validate() As Boolean
        Dim myret As Boolean = True
        MyBase.Validate()

        For Each drv As DataRowView In MonthlyBS.List
            If drv.Row.RowState = DataRowState.Modified Or drv.Row.RowState = DataRowState.Added Then
                If Not validaterow(drv) Then
                    myret = False
                End If
            End If
        Next
        Return myret
    End Function

    Private Function validaterow(ByVal drv As DataRowView) As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder
        If IsDBNull(drv.Row.Item("weekly")) Then
            myret = False
            sb.Append("Weekly cannot be blank")
        End If

        drv.Row.RowError = sb.ToString
        Return myret
    End Function

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not IsNothing(MonthlyBS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    MonthlyBS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MonthlyBS.CancelEdit()
    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectionChangeCommitted
        Dim myobj As ComboBox = DirectCast(sender, ComboBox)
        ''Dim bindings = myobj.DataBindings.Cast(Of Binding)().Where(Function(x) x.PropertyName = "SelectedItem" AndAlso x.DataSourceUpdateMode = DataSourceUpdateMode.OnPropertyChanged)

        '1. Force the Combobox to commit the value 
        For Each binding As Binding In myobj.DataBindings
            binding.WriteValue()
            binding.ReadValue()
        Next

        If Not IsNothing(MonthlyBS.Current) Then
            Dim myselected As DataRowView = ComboBox1.SelectedItem
            Dim drv As DataRowView = MonthlyBS.Current

            drv.Row.Item("monthly") = myselected.Item("monthly")
            drv.Row.Item("weeklyname") = myselected.Item("yearweek")
        End If
        DataGridView1.Invalidate()
    End Sub

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        MessageBox.Show(e.Exception.Message)
    End Sub
End Class