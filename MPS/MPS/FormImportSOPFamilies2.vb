Imports SSP.PublicClass
Imports System.Threading
Imports System.Text
Imports System.IO
Public Class FormImportSOPFamilies2

    Public Property Department As Department

    Dim myThread As New Thread(AddressOf DoWork)
    Dim ProgressReportDelegate1 As New ProgressReportDelegate(AddressOf ProgressReport)
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim FileName As String = String.Empty

    Public Sub New(ByVal mytype As Department)
        InitializeComponent()

        Me.Text = "Import SSP SOP-Families " & mytype.ToString
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
        Dim mylist As New List(Of String)
        Dim myRecord() As String
        Dim sqlstr As String = String.Empty
        Dim message As String = String.Empty

        Dim sspsopfamilyidseq As Long
        Dim sspsopfamilyid As Long
        Dim sspcmmfsopseq As Long
        Dim sspcmmfvendorsopseq As Long
        Dim sspcmmfsopid As Long
        Dim sspcmmfvendorsopid As Long

        Dim sspcmmfvendorseq As Long
        Dim sspcmmfvendorid As Long

        Dim period As String = String.Empty

        Dim sspsopfamilySB As New StringBuilder
        Dim sspcmmfsopidSB As New StringBuilder
        Dim sspcmmfvendorSB As New StringBuilder
        Dim sspcmmfvendorsopidsb As New StringBuilder
        Dim UpdatesopfamilySB As New StringBuilder
        Dim UpdatesopcmmfSB As New StringBuilder

        Try
            ProgressReport(2, "Preparing Data...")
            ProgressReport(6, "Set Progressbar to Marque")
            Dim Dataset1 = New DataSet

            'delete data first
            sqlstr = "delete from sspcmmfvendorsop; select setval('sspcmmfvendorsop_sspcmmfvendorsopid_seq',1,false);"
            dbtools1.ExecuteNonQuery(sqlstr)

            sqlstr = "select max(sspsopfamilyid)as sspsopfamilyid,sopfamily from (select min(sspsopfamilyid) as sspsopfamilyid,upper(trim(sopfamily)) as sopfamily from sspsopfamilies" &
                     " where sopfamily <> ''" &
                     " group by sopfamily order by sopfamily) as foo group by sopfamily;" &
                     "select cmmf,sopfamilyid,sspcmmfsopid from sspcmmfsop;" &
                     "select cmmf,vendorcode,cmmfvendorid from sspcmmfvendor;" &
                     "select sspsopfamilyid from sspsopfamilies order by sspsopfamilyid desc limit 1;" &
                     "select sspcmmfsopid from sspcmmfsop order by sspcmmfsopid desc limit 1;" &
                     "select cmmfvendorid from sspcmmfvendor order by cmmfvendorid desc limit 1;" &
                     "select sspsopfamilyid from sspsopfamilies where sspsopfamilyid = -1; " &
                     "select cmmf,vendorcode,sopfamilyid,sspcmmfvendorsopid from sspcmmfvendorsop;" &
                     "select sspcmmfvendorsopid from sspcmmfvendorsop order by sspcmmfvendorsopid desc limit 1;"



            If Not dbtools1.getDataSet(sqlstr, Dataset1, message) Then
                ProgressReport(3, message & " getdataset failed")
                Return myreturn
            End If
            Dim keys0(0) As DataColumn
            keys0(0) = Dataset1.Tables(0).Columns(1)
            Dataset1.Tables(0).PrimaryKey = keys0

            'Dim keys1(1) As DataColumn
            'keys1(0) = Dataset1.Tables(1).Columns(0)
            'keys1(1) = Dataset1.Tables(1).Columns(1)
            'Dataset1.Tables(1).PrimaryKey = keys1

            Dim keys1(0) As DataColumn
            keys1(0) = Dataset1.Tables(1).Columns(0)
            Dataset1.Tables(1).PrimaryKey = keys1


            Dim Key2(1) As DataColumn
            Key2(0) = Dataset1.Tables(2).Columns(0)
            Key2(1) = Dataset1.Tables(2).Columns(1)
            Dataset1.Tables(2).PrimaryKey = Key2

            sspsopfamilyidseq = Dataset1.Tables(3).Rows(0).Item(0).ToString
            sspcmmfsopseq = Dataset1.Tables(4).Rows(0).Item(0).ToString


            sspcmmfvendorseq = 0
            Try
                'sspcmmfvendorseq = IsDBNull(Dataset1.Tables(5).Rows(0).Item(0).ToString)
                sspcmmfvendorseq = Dataset1.Tables(5).Rows(0).Item(0).ToString
            Catch ex As Exception

            End Try

            
            Dim keys6(0) As DataColumn
            keys6(0) = Dataset1.Tables(6).Columns(0)
            Dataset1.Tables(6).PrimaryKey = keys6


            Dim keys7(1) As DataColumn
            keys7(0) = Dataset1.Tables(7).Columns(0)
            keys7(1) = Dataset1.Tables(7).Columns(1)
            Dataset1.Tables(7).PrimaryKey = keys7

            If Dataset1.Tables(8).Rows.Count > 0 Then
                sspcmmfvendorsopseq = Dataset1.Tables(8).Rows(0).Item(0)
            End If

            ProgressReport(3, "Read File...")

            'Using objTFParser = New FileIO.TextFieldParser(FileName)
            '    With objTFParser
            '        .TextFieldType = FileIO.FieldType.Delimited
            '        .SetDelimiters(vbTab)
            '        .HasFieldsEnclosedInQuotes = True
            '        Do Until .EndOfData
            '            myRecord = .ReadFields
            '            list.Add(myRecord)
            '        Loop
            '    End With
            'End Using

            Try
                Using myStream As StreamReader = New StreamReader(FileName, Encoding.Default)
                    Dim line As String = myStream.ReadLine
                    Do While (Not line Is Nothing)
                        mylist.Add(line)
                        line = myStream.ReadLine
                    Loop
                End Using
            Catch ex As Exception
                errMessage = ex.Message
                Return False
            End Try


            ProgressReport(5, "Set To continuous")
            Dim myupdate As New Dictionary(Of Long, String)
            For i = 1 To mylist.Count - 1
                If i = 6915 Then
                    Debug.Print("hello")
                End If
                ProgressReport(7, i & "," & mylist.Count)
                myRecord = Replace(mylist(i), "'", "''").Split(vbTab)
                'get sopfamilyid
                Dim idx0(0) As Object
                idx0(0) = myRecord(1)
                Dim result = Dataset1.Tables(0).Rows.Find(idx0)
                If IsNothing(result) Then
                    'create
                    sspsopfamilyidseq += 1
                    sspsopfamilyid = sspsopfamilyidseq
                    Dim dr As DataRow = Dataset1.Tables(0).NewRow
                    dr.Item(0) = sspsopfamilyid
                    dr.Item(1) = myRecord(1) 'list(i)(1)
                    Dataset1.Tables(0).Rows.Add(dr)
                    'sspsopfamilySB.Append(list(i)(1) & vbTab & list(i)(2) & vbTab & list(i)(3) & vbCrLf)
                    sspsopfamilySB.Append(myRecord(1) & vbTab & myRecord(2) & vbTab & myRecord(3) & vbCrLf)
                    'table sspsopfamilies : sopfamily,sopdescription,unit
                Else
                    sspsopfamilyid = result.Item("sspsopfamilyid")
                    'Update sop description
                    Dim idx6(0) As Object
                    idx6(0) = sspsopfamilyid
                    result = Dataset1.Tables(6).Rows.Find(idx6)
                    If IsNothing(result) Then
                        'create for array
                        If UpdatesopfamilySB.Length > 0 Then
                            UpdatesopfamilySB.Append(",")
                        End If
                        UpdatesopfamilySB.Append(String.Format("[{0}::character varying,'{1}','{2}']", sspsopfamilyid, myRecord(2), myRecord(3)))
                    End If
                End If

                'get cmmf,sopfamilyid  1 cmmf has only 1 sopfamilyid
                'Dim idx1(1) As Object
                'idx1(0) = list(i)(0)
                'idx1(1) = sspsopfamilyid
                Dim idx1(0) As Object
                idx1(0) = myRecord(0)
                result = Dataset1.Tables(1).Rows.Find(idx1)
                If IsNothing(result) Then
                    sspcmmfsopseq += 1
                    sspcmmfsopid = sspcmmfsopseq
                    Dim dr As DataRow = Dataset1.Tables(1).NewRow
                    dr.Item(0) = myRecord(0)
                    dr.Item(1) = sspsopfamilyid
                    dr.Item(2) = sspcmmfsopid
                    Dataset1.Tables(1).Rows.Add(dr)
                    sspcmmfsopidSB.Append(myRecord(0) & vbTab & sspsopfamilyid & vbCrLf)
                    'table sspcmmfsop : cmmf,sopfamilyid
                Else
                    sspcmmfsopid = result.Item(2).ToString

                    'Update with the latest sspsopfamilyid
                    If result.Item(1) <> sspsopfamilyid Then
                        'update
                        If UpdatesopcmmfSB.Length > 0 Then
                            UpdatesopcmmfSB.Append(",")
                        End If
                        UpdatesopcmmfSB.Append(String.Format("[{0}::character varying,'{1}']", sspcmmfsopid, sspsopfamilyid))
                    End If

                End If

                'New 1 cmmf,vendor has 1 sopfamilyid
                'Dim idx7(1) As Object
                'idx7(0) = myRecord(0)
                'idx7(1) = myRecord(4)
                'result = Dataset1.Tables(7).Rows.Find(idx7)
                'If IsNothing(result) Then
                '    sspcmmfvendorsopseq += 1
                '    sspcmmfvendorsopid = sspcmmfvendorsopseq
                '    Dim dr As DataRow = Dataset1.Tables(7).NewRow
                '    dr.Item(0) = myRecord(0)
                '    dr.Item(1) = myRecord(4)
                '    dr.Item(2) = sspsopfamilyid
                '    dr.Item(3) = sspcmmfvendorsopid
                '    Dataset1.Tables(7).Rows.Add(dr)
                '    sspcmmfvendorsopidsb.Append(myRecord(0) & vbTab & myRecord(4) & vbTab & sspsopfamilyid & vbCrLf)
                '    'table sspcmmfsop : cmmf,sopfamilyid
                'Else
                '    sspcmmfvendorsopid = result.Item(3).ToString
                'End If



                sspcmmfvendorsopseq += 1
                sspcmmfvendorsopid = sspcmmfvendorsopseq
                Dim dr1 As DataRow = Dataset1.Tables(7).NewRow
                dr1.Item(0) = myRecord(0)
                dr1.Item(1) = myRecord(4)
                dr1.Item(2) = sspsopfamilyid
                dr1.Item(3) = sspcmmfvendorsopid
                Dataset1.Tables(7).Rows.Add(dr1)
                sspcmmfvendorsopidsb.Append(myRecord(0) & vbTab & myRecord(4) & vbTab & sspsopfamilyid & vbCrLf)
                'table sspcmmfsop : cmmf,sopfamilyid


                'get cmmf vendor
                Dim idx2(1) As Object
                idx2(0) = myRecord(0)
                idx2(1) = myRecord(4)
                result = Dataset1.Tables(2).Rows.Find(idx2)
                If IsNothing(result) Then
                    sspcmmfvendorseq += 1
                    sspcmmfvendorid = sspcmmfvendorseq
                    Dim dr As DataRow = Dataset1.Tables(2).NewRow
                    dr.Item(0) = myRecord(0)
                    dr.Item(1) = myRecord(4)
                    dr.Item(2) = sspcmmfvendorid
                    Dataset1.Tables(2).Rows.Add(dr)
                    sspcmmfvendorSB.Append(myRecord(0) & vbTab & myRecord(4) & vbCrLf)
                    'table sspcmmfvendor : cmmf,vendor
                Else
                    sspcmmfvendorid = result.Item(2)
                End If

            Next

            If sspsopfamilySB.Length > 0 Then
                'table sspsopfamilies : sopfamily,sopdescription,unit
                sqlstr = "copy sspsopfamilies(sopfamily,sopdescription,unit) from stdin;"
                ProgressReport(2, "Copy sspsopfamilies")
                errMessage = dbtools1.copy(sqlstr, sspsopfamilySB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            End If

            If sspcmmfsopidSB.Length > 0 Then
                'table sspcmmfsop : cmmf,sopfamilyid
                sqlstr = "copy sspcmmfsop(cmmf,sopfamilyid) from stdin with null as 'Null';"
                ProgressReport(2, "Copy sspcmmfsop")
                errMessage = dbtools1.copy(sqlstr, sspcmmfsopidSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            End If

            ' *** new added 03 nov 2014 ***
            If sspcmmfvendorsopidSB.Length > 0 Then
                'table sspcmmfsop : cmmf,sopfamilyid
                sqlstr = "copy sspcmmfvendorsop(cmmf,vendorcode,sopfamilyid) from stdin with null as 'Null';"
                ProgressReport(2, "Copy sspcmmfvendorsop")
                errMessage = dbtools1.copy(sqlstr, sspcmmfvendorsopidSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            End If

            If sspcmmfvendorSB.Length > 0 Then
                'table sspcmmfvendor : cmmf,vendor
                sqlstr = "copy sspcmmfvendor(cmmf,vendorcode) from stdin;"
                ProgressReport(2, "Copy CMMFVendor")

                errMessage = dbtools1.copy(sqlstr, sspcmmfvendorSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If

            End If

            'update
            If UpdatesopfamilySB.Length > 0 Then
                ProgressReport(2, "Update sspsopfamilies")
                sqlstr = "update sspsopfamilies set sopdescription = foo.description,unit = foo.unit from (select * from array_to_set3(Array[" & UpdatesopfamilySB.ToString & "]) as tb (id character varying,description character varying,unit character varying))foo where sspsopfamilyid = foo.id::bigint;"
                Dim ra As Long

                If Not dbtools1.ExecuteNonQuery(sqlstr, ra, errMessage) Then
                    Return False
                End If
            End If

            If UpdatesopcmmfSB.Length > 0 Then
                ProgressReport(2, "Update sspcmmfsop")
                sqlstr = "update sspcmmfsop set sopfamilyid = foo.sspsopfamilyid::bigint from (select * from array_to_set2(Array[" & UpdatesopcmmfSB.ToString & "]) as tb (id character varying,sspsopfamilyid character varying))foo where sspcmmfsopid = foo.id::bigint;"
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


End Class