Imports SSP.PublicClass
Imports System.Threading
Imports System.Text
Imports System.IO

Delegate Sub ProgressReportDelegate(ByVal sender As Integer, ByVal message As String)

Public Class FormImportFTYCapV2
    Public Property sspftycap As String
    Public Property sspftycapdata As String
    Public Property sspftycapseqname As String
    Public Property sspftycapdataseqname As String
    Public Property Department As Department

    Dim myThread As New Thread(AddressOf DoWork)
    Dim ProgressReportDelegate1 As New ProgressReportDelegate(AddressOf ProgressReport)
    Dim OpenFileDialog1 As New OpenFileDialog
    Dim FileName As String = String.Empty
    Dim sspsopfamiliesSB As New System.Text.StringBuilder
    Dim sspftycapSB As New System.Text.StringBuilder
    Dim sspftycapDataSB As New System.Text.StringBuilder

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
                    ToolStripStatusLabel1.Text = message
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

        Dim sopfamilyid As Long
        Dim sspftycapseq As Long
        Dim ftycapid As Long
        Dim period As String = String.Empty
        Dim sspsopfamilyidseq As Long
        Dim sspsopfamilyid As Long
        Dim sspsopfamilysb As New StringBuilder

        sspftycapSB.Clear()
        sspftycapDataSB.Clear()

        Try
            ProgressReport(2, "Preparing Data...")
            ProgressReport(6, "Set Progressbar to Marque")
            Dim Dataset1 = New DataSet

            sqlstr = "select max(sspsopfamilyid)as sspsopfamilyid,sopfamily from (select min(sspsopfamilyid) as sspsopfamilyid,upper(trim(sopfamily)) as sopfamily from sspsopfamilies" &
                     " where sopfamily <> ''" &
                     " group by sopfamily order by sopfamily) as foo group by sopfamily;" &
                     "select sspsopfamilyid from sspsopfamilies order by sspsopfamilyid desc limit 1;"
            If Not dbtools1.getDataSet(sqlstr, Dataset1, message) Then
                ProgressReport(3, message)
                Return myreturn
            End If

            Dim keys(0) As DataColumn
            keys(0) = Dataset1.Tables(0).Columns(1)
            Dataset1.Tables(0).PrimaryKey = keys

            sspsopfamilyidseq = Dataset1.Tables(1).Rows(0).Item(0).ToString

            ProgressReport(3, "Read File...")

            Using objTFParser = New FileIO.TextFieldParser(FileName)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = False
                    Do Until .EndOfData
                        myRecord = .ReadFields
                        list.Add(myRecord)
                    Loop
                End With
            End Using

            sqlstr = "Delete from " & sspftycap & " where period = " & list(1)(0) & ";" &
                     "select setval('" & sspftycapseqname & "',getlastid('ftycapid','" & sspftycap & "') + 1,false);" &
                     "select setval('" & sspftycapdataseqname & "',getlastid('ftycapdataid','" & sspftycapdata & "' ) + 1,false);" &
                     "select ftycapid from " & sspftycap & " order by ftycapid desc limit 1;"


            Dim ra As Long
            If Not dbtools1.ExecuteScalar(sqlstr, ra, message) Then
                errMessage = message
                Return False
            End If
            sspftycapseq = ra

            ProgressReport(5, "Set To continuous")
            For i = 1 To list.Count - 1
                ProgressReport(7, i & "," & list.Count)
                If i = 1012 Then
                    Debug.Print("debugmode")
                End If
                'get sopfamily
                Dim pk0(0) As Object
                pk0(0) = list(i)(3)
                Dim result = Dataset1.Tables(0).Rows.Find(pk0)
                If IsNothing(result) Then
                    'ProgressReport(3, list(i)(3) & " SopFamily not found. Please Import SOP Families Global First.")
                    'errMessage = list(i)(3) & " SopFamily not found. Please Import SOP Families Global First."
                    'Return myreturn
                    'create
                    sspsopfamilyidseq += 1
                    sspsopfamilyid = sspsopfamilyidseq
                    Dim dr As DataRow = Dataset1.Tables(0).NewRow
                    dr.Item(0) = sspsopfamilyid
                    dr.Item(1) = list(i)(3)
                    Dataset1.Tables(0).Rows.Add(dr)
                    'sspsopfamilySB.Append(list(i)(1) & vbTab & list(i)(2) & vbTab & list(i)(3) & vbCrLf)
                    sspsopfamilysb.Append(list(i)(3) & vbTab & list(i)(4) & vbCrLf)
                Else
                    sopfamilyid = result.Item("sspsopfamilyid")
                End If


                ftycapid = sspftycapseq
                sspftycapseq += 1
                'period,typeofinfoid,vendorcode,sopfamilyid
                sspftycapSB.Append(list(i)(0) & vbTab &
                                   "1" & vbTab &
                                   list(i)(1) & vbTab &
                                   sopfamilyid & vbCrLf)

                'read data monthly
                Dim mymonth As String = String.Empty
                For j = 6 To list(i).Count - 1
                    
                    'check for blank value
                    If Not list(i)(j) = "" Then
                        'check for valid int
                        If Not IsNumeric(list(i)(j)) Then
                            Err.Raise(1, Description:=String.Format("The line Number {0} contains invalid number '{1}'", i, list(i)(j)))
                        End If
                        sspftycapDataSB.Append(ftycapid & vbTab)
                        sspftycapDataSB.Append(1 & vbTab) 'Periodtype 1 for monthly 2 for weekly
                        sspftycapDataSB.Append(HelperClass1.getdate(j - 6, list(i)(0)) & vbTab) 'calculation starting from 0 as current month
                        If j = 6 Then
                            mymonth = "M"
                        Else
                            mymonth = "M+" & Format(j - 6, "00")
                        End If
                        sspftycapDataSB.Append(mymonth & vbTab)

                        sspftycapDataSB.Append(list(i)(j) & vbCrLf)
                        'ftycapid,periodetypeid,startingdate,datalabel,datavalue

                    End If
                Next
            Next

            'Copy Fty Cap

            If sspsopfamilysb.Length > 0 Then
                'table sspsopfamilies : sopfamily,sopdescription,unit
                sqlstr = "copy sspsopfamilies(sopfamily,sopdescription) from stdin;"
                ProgressReport(2, "Copy sspsopfamilies")
                errMessage = dbtools1.copy(sqlstr, sspsopfamilysb.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            End If

            sqlstr = "copy " & sspftycap & "(period,typeofinfoid,vendorcode,sopfamilyid) from stdin;"
            ProgressReport(2, "Copy To Db (Fty Cap)")
            'Debug.Print(sspftycapSB.ToString)

            If sspftycapSB.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, sspftycapSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
            Else
                myreturn = True
            End If

            'copy Fty Cap Data

            sqlstr = "copy " & sspftycapdata & "(ftycapid,periodtypeid,startingdate,datalabel,datavalue) from stdin;"
            ProgressReport(2, "Copy To Db (Fty Cap Data)")
            Debug.Print(sspftycapDataSB.ToString)
            If sspftycapDataSB.ToString <> "" Then
                errMessage = dbtools1.copy(sqlstr, sspftycapDataSB.ToString, myreturn)
                If Not myreturn Then
                    Return myreturn
                End If
                ProgressReport(2, "Copy To Db Done.")
            Else
                ProgressReport(2, "Nothing to Copy.")
                myreturn = True
            End If
            ProgressReport(3, "")
        Catch ex As Exception
            errMessage = ex.Message
        End Try
        Return myreturn
    End Function

    Public Sub New(ByVal department As Department, ByVal TableCapacity As String, ByVal TableCapacitySeq As String, ByVal TableCapacityData As String, ByVal TableCapacityDataSeq As String)

        ' This call is required by the designer.
        InitializeComponent()
        Me.sspftycap = TableCapacity
        Me.sspftycapseqname = TableCapacitySeq
        Me.sspftycapdata = TableCapacityData
        Me.sspftycapdataseqname = TableCapacityDataSeq
        Me.Text = "Import Factory Capacity (" & department.ToString & ")"

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class

