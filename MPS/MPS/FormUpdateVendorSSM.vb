Imports SSP.PublicClass
Imports System.Threading
Imports System.Text
Imports System.IO
Public Class FormUpdateVendorSSM


    Public Property Department As Department
    Dim myThread As New Thread(AddressOf DoWork)
    Dim ProgressReportDelegate1 As New ProgressReportDelegate(AddressOf ProgressReport)
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim FileName As String = String.Empty

    Public Sub New()
        InitializeComponent()


    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (myThread.IsAlive) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoWork()
        Dim Result As Boolean = False
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, TextBox2.Text & "Open File..")

        Result = ImportData(FileName, errMsg)
        If Result Then
            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} {3}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString, "Done"))
            ProgressReport(5, "Set to continuous mode again")
            ProgressReport(3, "")
        Else
            errSB.Append(errMsg & vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()

    End Sub

    Sub ProgressReport(ByVal sender As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Me.Invoke(ProgressReportDelegate1, New Object() {sender, message})
        Else
            Select Case sender
                Case 1
                    TextBox1.Text = message
                Case 2
                    'For Progress
                    ToolStripStatusLabel1.Text = message
                Case 3
                    'For Error message
                    TextBox3.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
                    ToolStripStatusLabel1.Text = "Preparing Data .." & myvalue(0) & "/" & myvalue(1)
            End Select
        End If
    End Sub

    Private Sub EndWork()
        ProgressReport(2, "Done")
    End Sub

    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim list As New List(Of String())
        Dim myRecord() As String
        Dim sqlstr As String = String.Empty
        Dim message As String = String.Empty



        Dim period As String = String.Empty

        Dim InsertVendorSB As New StringBuilder
        Dim UpdateVendorSB As New StringBuilder


        Try
            ProgressReport(2, "Preparing Data...")
            ProgressReport(6, "Set Progressbar to Marque")
            Dim Dataset1 = New DataSet


            sqlstr = "select vendorcode from vendor;"


            If Not dbtools1.getDataSet(sqlstr, Dataset1, message) Then
                ProgressReport(3, message & " getdataset failed")
                Return myreturn
            End If
            Dim keys0(0) As DataColumn
            keys0(0) = Dataset1.Tables(0).Columns(0)
            Dataset1.Tables(0).PrimaryKey = keys0

            ProgressReport(3, "Read File...")

            Using objTFParser = New FileIO.TextFieldParser(FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(vbTab)
                    .HasFieldsEnclosedInQuotes = True
                    Do Until .EndOfData
                        myRecord = .ReadFields
                        list.Add(myRecord)
                    Loop
                End With
            End Using

            ProgressReport(5, "Set To continuous")
            Dim myupdate As New Dictionary(Of Long, String)
            For i = 1 To list.Count - 1
                ProgressReport(7, i & "," & list.Count)
                Dim idx0(0) As Object
                idx0(0) = list(i)(0)
                Dim result = Dataset1.Tables(0).Rows.Find(idx0)
                If IsNothing(result) Then
                Else
                   
                        'create for array
                    If UpdateVendorSB.Length > 0 Then
                        UpdateVendorSB.Append(",")
                    End If
                    UpdateVendorSB.Append(String.Format("[{0}::character varying,'{1}']", list(i)(0), list(i)(1)))

                End If

            Next

            'update
            If UpdateVendorSB.Length > 0 Then
                ProgressReport(2, "Update vendor")
                sqlstr = "update vendor set officerid = foo.officerid from (select * from array_to_set2(Array[" & UpdateVendorSB.ToString & "]) as tb (id character varying,officerid character varying))foo where vendorcode = foo.id::bigint;"
                Dim ra As Long

                If Not dbtools1.ExecuteNonQuery(sqlstr, ra, errMessage) Then
                    Return False
                End If
            End If
            myreturn = True
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function

    Private Sub FormUpdateVendorSSM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class