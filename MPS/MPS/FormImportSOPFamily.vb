Imports DJLib
Imports DJLib.Dbtools
Imports System.ComponentModel
Imports System.IO
Imports Npgsql
Imports System.Text

Public Class FormImportSOPFamily
    Private WithEvents BackgroundWorker1 As New BackgroundWorker
    Dim FileName As String
    Dim Status As Boolean = False
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim myconverter As New utf8towin1252
    Dim Dataset1 As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not (BackgroundWorker1.IsBusy) Then
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                FileName = OpenFileDialog1.FileName
                TextBox1.Text = FileName
                Try
                    BackgroundWorker1.WorkerReportsProgress = True
                    BackgroundWorker1.WorkerSupportsCancellation = True
                    BackgroundWorker1.RunWorkerAsync()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(3, TextBox3.Text & "Start")

        Dim errMsg As String = String.Empty
        Status = ImportData(FileName, errMsg)
        If Status Then
            BackgroundWorker1.ReportProgress(2, TextBox2.Text & " Done.")
        Else
            BackgroundWorker1.ReportProgress(3, "Error::" & errMsg)
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Select Case e.ProgressPercentage
            Case 2
                TextBox2.Text = e.UserState
            Case 3
                TextBox3.Text = e.UserState
        End Select
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        FormMenu.setBubbleMessage("Import SOP Families", "Done")
        If Status Then
            If CheckBox1.Checked Then
                Me.Close()
            End If
        End If
    End Sub


    Private Function ImportData(ByVal FileName As String, Optional ByRef errMessage As String = "") As Boolean
        Dim myreturn As Boolean = False
        Dim list As New List(Of String)
        Dim stringBuilder1 As New System.Text.StringBuilder
        Dim stringBuilder2 As New System.Text.StringBuilder
        Dim myRecord() As String
        Dim SOPFamilyId As Long = 0
        Dim RangeId As Long = 0
        Dim SSPCMMFRangeid As Long = 0
        Dim period As Long = 0
        Dim myId As Long = 0
        Try
            BackgroundWorker1.ReportProgress(2, "Preparing Data...")
            Dataset1 = New DataSet
            BackgroundWorker1.ReportProgress(3, "Convert TextFile to Unicode..")
            BackgroundWorker1.ReportProgress(3, "Connect to Db...")
            Dim sqlstr As String = "select sspsopfamilyid,sopfamily from sspsopfamilies;" &
                                   "select cmmf,sopfamilyid from sspcmmfsop;" &
                                   "select cmmf,vendorcode from sspcmmfvendor;"
            dbtools1.getDataSet(sqlstr, Dataset1)
            Dim keys(1) As DataColumn
            keys(0) = Dataset1.Tables(1).Columns(0)
            keys(1) = Dataset1.Tables(1).Columns(1)
            Dataset1.Tables(1).PrimaryKey = keys


            'Dim key2(0) As DataColumn
            'key2(0) = Dataset1.Tables(0).Columns(1)
            'Dataset1.Tables(0).PrimaryKey = key2
            'Dim idstring As String = String.Empty

            Dim Key3(1) As DataColumn
            Key3(0) = Dataset1.Tables(2).Columns(0)
            Key3(1) = Dataset1.Tables(2).Columns(1)
            Dataset1.Tables(2).PrimaryKey = Key3


            Dataset1.Tables(0).TableName = "SSPSOPFamily"
            Dataset1.Tables(1).TableName = "SSPCMMFSOP"
            BackgroundWorker1.ReportProgress(3, "Open File...")
            Try
                Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)
                    Dim line As String = myStream.ReadLine
                    Do While (Not line Is Nothing)
                        list.Add(line)
                        line = myStream.ReadLine
                    Loop
                End Using
            Catch ex As Exception
                errMessage = ex.Message
                Return False
            End Try

            Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
            BackgroundWorker1.ReportProgress(2, "Scanning Data...")
            Dim myhastable As New Hashtable
            Dim cmmfvendorHT As New Hashtable
            Dim SopfamilyHT As New Hashtable
            For i = 1 To list.Count - 1
                Try
                    BackgroundWorker1.ReportProgress(3, "Processing line " & i & " of " & list.Count - 1)
                    myRecord = list(i).Split(vbTab)
                    SOPFamilyId = getSOPFamilyId(myRecord, errMessage)
                    If SOPFamilyId = -1 Then
                        Return False
                    End If

                    'sspcmmfvendor
                    If myRecord(4) <> "" Then
                        Try
                            cmmfvendorHT.Add(myRecord(0).ToString & myRecord(4).ToString, myRecord(0).ToString & "," & myRecord(4).ToString)
                            Dim pkey(1) As Object
                            pkey(0) = myRecord(0).ToString
                            pkey(1) = myRecord(4).ToString
                            Dim DataRow1 As DataRow = Dataset1.Tables(2).Rows.Find(pkey) 'Table sspcmmfvendor
                            If DataRow1 Is Nothing Then
                                stringBuilder2.Append(myRecord(0) & vbTab)
                                stringBuilder2.Append(myRecord(4) & vbCrLf)
                                Try
                                    'SopfamilyHT.Add(SOPFamilyId, myRecord(4))
                                    Dim cmd2 As NpgsqlCommand = New NpgsqlCommand
                                    cmd2.CommandText = "update sspsopfamilies set sopfamily=@sopfamily,sopdescription=@sopdescription,unit=@unit where sspsopfamilyid = @sspsopfamilyid ;"
                                    cmd2.Connection = conn
                                    conn.Open()
                                    cmd2.Parameters.Add("@sopfamily", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(1)
                                    cmd2.Parameters.Add("@sopdescription", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(2)
                                    cmd2.Parameters.Add("@unit", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(3)
                                    cmd2.Parameters.Add("@sspsopfamilyid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = SOPFamilyId
                                    cmd2.ExecuteNonQuery()
                                Catch ex As Exception
                                Finally
                                    conn.Close()
                                End Try
                            Else
                                'Update
                                Try
                                    SopfamilyHT.Add(SOPFamilyId, myRecord(4))
                                    Dim cmd2 As NpgsqlCommand = New NpgsqlCommand
                                    cmd2.CommandText = "update sspsopfamilies set sopfamily=@sopfamily,sopdescription=@sopdescription,unit=@unit where sspsopfamilyid = @sspsopfamilyid ;"
                                    cmd2.Connection = conn
                                    conn.Open()
                                    cmd2.Parameters.Add("@sopfamily", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(1)
                                    cmd2.Parameters.Add("@sopdescription", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(2)
                                    cmd2.Parameters.Add("@unit", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(3)
                                    cmd2.Parameters.Add("@sspsopfamilyid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = SOPFamilyId
                                    cmd2.ExecuteNonQuery()
                                Catch ex As Exception
                                Finally
                                    conn.Close()
                                End Try

                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    'cmmfsop Update Or Add

                    Try
                        'Find cmmfsop if avail then update

                        Dim sspcmmfsopid = getsspcmmfsopId(myRecord, SOPFamilyId, errMessage)
                        If sspcmmfsopid = 0 Then
                            'else Add
                            myhastable.Add(myRecord(0).ToString, myRecord(0))
                            Dim pkey(1) As Object
                            pkey(0) = myRecord(0)
                            pkey(1) = SOPFamilyId
                            Dim DataRow1 As DataRow = Dataset1.Tables(1).Rows.Find(pkey)
                            If DataRow1 Is Nothing Then
                                stringBuilder1.Append(myRecord(0) & vbTab)
                                stringBuilder1.Append(SOPFamilyId & vbCrLf)
                            End If
                        End If

                    Catch ex As Exception
                    End Try
                    Application.DoEvents()


                Catch ex As Exception
                    errMessage = ex.Message
                    Return False
                Finally
                End Try
            Next

            'copy sspcmmfvendor

            sqlstr = "copy sspcmmfvendor(cmmf,vendorcode) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (CMMF Vendor)")
            If stringBuilder2.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, stringBuilder2.ToString, myreturn)
                BackgroundWorker1.ReportProgress(2, "Copy To Db Done.")
            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myreturn = True
            End If
            BackgroundWorker1.ReportProgress(3, "")

            'copy ccpcmmfsop

            sqlstr = "copy sspcmmfsop(cmmf,sopfamilyid) from stdin;"
            BackgroundWorker1.ReportProgress(2, "Copy To Db (CMMF SOPFamilies)")
            If stringBuilder1.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, stringBuilder1.ToString, myreturn)
                BackgroundWorker1.ReportProgress(2, "Copy To Db Done.")
            Else
                BackgroundWorker1.ReportProgress(2, "Nothing to Copy.")
                myreturn = True
            End If
            BackgroundWorker1.ReportProgress(3, "")
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function
   

    Private Sub FormImportSSPSOPFamily_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If BackgroundWorker1.IsBusy Then
            If MessageBox.Show("Are you sure to exit? Your process will continue running in background ", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                e.Cancel = False
            Else
                e.Cancel = True
            End If
            Exit Sub
        End If
    End Sub

    Private Function getSOPFamilyId(ByVal myRecord() As String, Optional ByRef myMessage As String = "") As Long
        Dim Myreturn As Long = -1
        Dim cmd As NpgsqlCommand = New NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select sspsopfamilyid from sspsopfamilies where sopfamily = :sopfamily"
            cmd = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("sopfamily", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(1)
            Myreturn = cmd.ExecuteScalar
            If Myreturn = 0 Then
                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
                cmd1.CommandText = "insert into sspsopfamilies(sopfamily,sopdescription,unit) values(:sopfamily,:sopdescription,:unit);select currval('sspsopfamilies_sspsopfamilyid_seq');"
                cmd1.Connection = conn
                cmd1.Parameters.Add("sopfamily", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(1)
                cmd1.Parameters.Add("sopdescription", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(2)
                cmd1.Parameters.Add("unit", NpgsqlTypes.NpgsqlDbType.Varchar).Value = myRecord(3)
                Myreturn = cmd1.ExecuteScalar           
            End If
        Catch ex As Exception
            myMessage = ex.Message
        Finally
            conn.Close()
        End Try
        Return Myreturn
    End Function

    Private Function getsspcmmfsopId(ByVal myRecord() As String, ByRef SOPFamilyId As Long, Optional ByRef myMessage As String = "") As Long
        Dim Myreturn As Long = -1
        Dim cmd As NpgsqlCommand = New NpgsqlCommand
        Dim conn As NpgsqlConnection = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select sspcmmfsopid from sspcmmfsop where cmmf = :cmmf"
            cmd = New NpgsqlCommand(sqlstr, conn)
            cmd.Parameters.Add("cmmf", NpgsqlTypes.NpgsqlDbType.Bigint).Value = myRecord(0)
            Myreturn = cmd.ExecuteScalar
            If Myreturn > 0 Then
                Dim cmd1 As NpgsqlCommand = New NpgsqlCommand
                cmd1.CommandText = "update sspcmmfsop set sopfamilyid = :sopfamilyid where sspcmmfsopid = :sspcmmfsopid;"
                cmd1.Connection = conn
                cmd1.Parameters.Add("sopfamilyid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = SOPFamilyId
                cmd1.Parameters.Add("sspcmmfsopid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = SOPFamilyId                
                cmd1.ExecuteScalar()
            End If
        Catch ex As Exception
            myMessage = ex.Message
        Finally
            conn.Close()
        End Try
        Return Myreturn
    End Function

End Class