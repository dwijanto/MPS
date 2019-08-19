Imports System
Imports System.Data
Imports System.IO
Imports Npgsql
Imports System.Xml
Imports System.Configuration
Imports System.Windows.Forms

Public Class Dbtools
    Implements IDisposable
    Private mHost As String
    Private mPort As String
    Private mDatabase As String
    Private mUserId As String
    Private mPassword As String
    Private mGroup As String
    Private mConnectString As String
    Private conn As NpgsqlConnection
    Private command As NpgsqlCommand
    Private CopyIn1 As NpgsqlCopyIn

    Public Property Userid As String
        Get
            Return mUserId
        End Get
        Set(ByVal value As String)
            mUserId = value
        End Set
    End Property
    Public Property Password As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal host As String, ByVal port As String, ByVal database As String, ByVal userid As String, ByVal password As String)
        mHost = host
        mPort = port
        mDatabase = database
        mUserId = userid
        mPassword = password
    End Sub
    Public Sub New(ByVal userid As String, ByVal password As String)
        mUserId = userid
        mPassword = password
    End Sub

    ReadOnly Property Host() As String
        Get
            Return mHost
        End Get
    End Property
    ReadOnly Property Port() As String
        Get
            Return mPort
        End Get
    End Property
    ReadOnly Property Database() As String
        Get
            Return mDatabase
        End Get
    End Property

    'ReadOnly Property Userid() As String
    '    Get
    '        Return mUserId
    '    End Get
    'End Property

    'ReadOnly Property password() As String
    '    Get
    '        Return mPassword
    '    End Get
    'End Property

    ReadOnly Property Group() As String
        Get
            Return mGroup
        End Get
    End Property

    Public Function getConnectionString() As String
        Dim builder As New NpgsqlConnectionStringBuilder()
        builder.ConnectionString = My.Settings.connectionstring1 '"Host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;" 'My.Settings.connectionstring1 ' "Host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;" 'My.Settings.ConnectionString2
        builder.Add("User Id", mUserId)
        builder.Add("password", mPassword)
        Return builder.ConnectionString
    End Function

    Public Function getVer() As String
        Dim m_xmld = New XmlDocument()
        m_xmld.Load(Application.ExecutablePath & ".manifest")
        Return m_xmld.ChildNodes.Item(1).ChildNodes.Item(0).Attributes.GetNamedItem("version").Value
    End Function
    Public Function getData(ByVal sqlstr As String, Optional ByVal errMessage As String = "") As DataTable
        Dim Table As DataTable = New DataTable
        Try
            Using npDataAdapter = New NpgsqlDataAdapter(sqlstr, getConnectionString)
                Dim cb As New NpgsqlCommandBuilder(npDataAdapter)
                npDataAdapter.Fill(Table)
            End Using
        Catch ex As NpgsqlException
            errMessage = ex.Message
        End Try
        Return Table
    End Function
    Public Function getData(ByVal sqlStr As String, ByRef Table As DataTable, ByRef errMsg As String) As Boolean
        Dim myRet As Boolean = False
        Try
            Using npDataAdapter = New NpgsqlDataAdapter(sqlStr, getConnectionString)
                Dim cb As New NpgsqlCommandBuilder(npDataAdapter)
                npDataAdapter.Fill(Table)
            End Using
            myRet = True
        Catch ex As NpgsqlException
            errMsg = ex.Message
        End Try
        Return myRet
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    recordAffected = command.ExecuteNonQuery
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function
    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myRet As Boolean = False
        Dim myresult As Object
        Using conn = New NpgsqlConnection(getConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    myresult = command.ExecuteScalar
                    recordAffected = IIf(IsDBNull(myresult), 0, myresult)
                    myRet = True
                Catch ex As NpgsqlException
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myRet
    End Function

    Public Function getGroupName(ByRef GroupName As String) As Boolean
        Dim myret As String = False
        Dim sqlstr As String = Nothing
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            sqlstr = "select _groupname from _groupuser where _membername = " & escapeString(mUserId)
            Using command = New NpgsqlCommand(sqlstr, conn)
                Try
                    GroupName = command.ExecuteScalar
                    myret = True
                Catch ex As NpgsqlException
                    MsgBox(ex.Message)
                End Try
            End Using
        End Using

        Return myret
    End Function

    Public Function canLogin(ByRef errorMessage As String) As Boolean
        Dim myret As String = False
        Dim sqlstr As String = Nothing
        Try
            Using conn = New NpgsqlConnection(getConnectionString())
                conn.Open()
                sqlstr = "select _groupname from _groupuser where _membername = " & escapeString(mUserId)
                Using command = New NpgsqlCommand(sqlstr, conn)
                    Try
                        mGroup = command.ExecuteScalar
                        myret = True
                    Catch ex As NpgsqlException
                        errorMessage = ex.Message
                    End Try
                End Using
            End Using
        Catch ex As NpgsqlException
            errorMessage = ex.Message
        End Try
        Return myret
    End Function
    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(getConnectionString())
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function
    Public Function Copy(ByVal sqlstr As String, ByVal myfile As String, ByRef message As String) As Boolean
        Dim myret As Boolean = False
        Dim copyinstream As Stream
        Dim buf(9) As Byte
        Dim i As Long

        Using FileStream1 As FileStream = File.OpenRead(myfile)
            Using conn = New NpgsqlConnection(getConnectionString())
                conn.Open()
                Using command = New NpgsqlCommand(sqlstr, conn)
                    CopyIn1 = New NpgsqlCopyIn(command, conn)
                    Try
                        CopyIn1.Start()
                        copyinstream = CopyIn1.CopyStream
                        i = FileStream1.Read(buf, 0, buf.Length)
                        While i > 0
                            copyinstream.Write(buf, 0, i)
                            i = FileStream1.Read(buf, 0, buf.Length)
                        End While
                        copyinstream.Close()
                        myret = True
                    Catch ex As NpgsqlException
                        Try
                            CopyIn1.Cancel("Undo Copy")
                            message = ex.Message & " ex"
                        Catch ex2 As NpgsqlException
                            If ex2.Message.Contains("Undo Copy") Then
                                message = ex2.Message & " ex2"
                            End If
                        End Try
                    End Try

                End Using
            End Using
        End Using
        Return myret
    End Function


    Public Shared Function DateFormatDDMMYYYY(ByRef DateInput As String) As String
        Dim myRet As String
        Dim arrDate(2) As String
        Dim arrTmp As String()
        Try
            arrTmp = DateInput.Split("/")
            arrDate(0) = arrTmp(2)
            arrDate(1) = arrTmp(1)
            arrDate(2) = arrTmp(0)
            myRet = "'" & String.Join("-", arrDate) & "'"
        Catch ex As Exception
            arrTmp = DateInput.Split("-")
            arrDate(0) = arrTmp(2)
            arrDate(1) = arrTmp(1)
            arrDate(2) = arrTmp(0)
            myRet = "'" & String.Join("-", arrDate) & "'"
        End Try        
        Return myRet
    End Function

    Public Shared Function DateFormatyyyyMMdd(ByRef DateInput As Object) As String
        Dim myRet As String
        myRet = "NULL"
        If Not DateInput Is Nothing Then
            Dim arrDate(2) As String
            arrDate(0) = DateInput.Year
            arrDate(1) = DateInput.Month
            arrDate(2) = DateInput.Day
            myRet = "'" & String.Join("-", arrDate) & "'"
        End If
        Return myRet
    End Function
    Public Shared Function DateFormatDotNet(ByRef DateInput As String) As Date
        Dim myRet As Date = CDate(DateInput)
        Return myRet
    End Function
    Public Function ValidNum(ByRef NumInput As String) As String
        Dim myRet As String
        myRet = NumInput.Replace(",", ".")
        myRet = myRet.Replace("ND", "\N")
        myRet = myRet.Replace("NA", "\N")
        Return myRet
    End Function
    Public Shared Function ValidNumber(ByRef myNumber As Object) As Object
        Dim myRet As Object = myNumber

        If Not IsNumeric(myNumber) Then
            myRet = DBNull.Value
        End If
        Return myRet
    End Function
    Public Function getDataSet(ByVal sqlstr As String, ByVal DataAdapter As NpgsqlDataAdapter, ByRef DataSet As DataSet, ByVal Table As String, Optional ByRef message As String = "") As Boolean

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataSet, Table)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Public Function getDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Public Function getDataSet(ByVal sqlstr As String, ByVal Table As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataSet, Table)
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
    Public Function getDataReader(ByVal sqlstr As String, ByRef DataReader As NpgsqlDataReader, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                Dim command As New NpgsqlCommand
                command.Connection = conn
                command.CommandText = sqlstr
                command.CommandType = CommandType.Text
                DataReader = command.ExecuteReader
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Public Function GetDataCollection(ByVal sqlstr As String, ByRef DataCollection As Collection, ByRef errorMessage As String) As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim DataTable As New DataTable
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataTable)
            End Using
            For Each DataRow In DataTable.Rows
                DataCollection.Add(DataRow(0), DataRow(1).ToString) '0-> Value, 1-> Key
            Next
            myret = True
        Catch ex As NpgsqlException
            errorMessage = ex.Message
        End Try
        Return myret
    End Function

    Public Function DeleteRecord(ByVal sqlstr As String, ByVal Id As String, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                Using command As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
                    command.Parameters.Add(New NpgsqlParameter("@mykey", NpgsqlTypes.NpgsqlDbType.Text))
                    command.CommandType = CommandType.Text
                    command.Parameters(0).Value = Id
                    command.ExecuteNonQuery()
                End Using
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function
    Public Function DeleteRecord(ByVal sqlstr As String, ByVal Id As Int64, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(getConnectionString)
                conn.Open()
                Using command As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
                    command.Parameters.Add(New NpgsqlParameter("@mykey", NpgsqlTypes.NpgsqlDbType.Bigint))
                    command.CommandType = CommandType.Text
                    command.Parameters(0).Value = Id
                    command.ExecuteNonQuery()
                End Using
            End Using
            myret = True
        Catch ex As NpgsqlException
            message = ex.Message
        End Try
        Return myret
    End Function

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try

        'Try
        '    Do Until System.Runtime.InteropServices.Marshal.ReleaseComObject(o) <= 0
        '    Loop


        'Catch ex As Exception
        'Finally
        '    o = Nothing
        'End Try
    End Sub
    Public Shared Sub fadeout(ByVal o As System.Windows.Forms.Form)
        Dim i As Double
        For i = -1 To 0 Step 0.01
            o.Opacity = System.Math.Abs(i)
            o.Refresh()
        Next
    End Sub
    Public Shared Sub fadein(ByVal o As System.Windows.Forms.Form)
        Dim i As Double
        For i = 0 To 1 Step 0.01
            o.Opacity = System.Math.Abs(i)
            o.Refresh()
        Next
    End Sub

    Public Shared Function MoveToFolder(ByVal Foldername As String, ByVal source As String, ByRef Message As String) As Boolean
        Dim i As Integer = 0
        Dim myCopy As String = ""
        Dim Destination As String = ""
        Dim FileName As String = Path.GetFileName(source)
        Dim myRet As Boolean = False
        While File.Exists(Foldername & "\" & FileName)
            i += 1
            FileName = "Copy(" & i & ") of " & Path.GetFileName(source)
        End While

        Destination = Foldername & "\" & FileName
        Try
            If File.Exists(source) Then
                Directory.Move(source, Destination)
            End If
            myRet = True
        Catch ex As Exception
            Message = ex.Message
        End Try
        Return myRet
    End Function

    Public Shared Function ValidateFileName(ByVal foldername As String, ByVal source As String) As String
        ValidateFileName = source
        Dim FileName = Path.GetFileName(source)
        Dim i As Integer = 0
        While File.Exists(foldername & "\" & FileName)
            i += 1
            FileName = "Copy(" & i & ") of " & Path.GetFileName(source)
            ValidateFileName = foldername & "\" & FileName
        End While
        Return ValidateFileName
    End Function

    Public Shared Function getUsedRange(ByVal osheet As Object, ByRef row As Long, ByRef col As Long) As Boolean
        Dim myRet As Boolean = False
        Dim myRanges As Object = osheet.UsedRange
        col = myRanges.Columns.Count
        row = myRanges.Rows.Count
        Return myRet
        releaseComObject(osheet)
    End Function
    Public Shared Function MinMaxAvg(ByVal myColl As Object) As Double()
        Dim myResult(2) As Double
        Dim min As Double = vbNull
        Dim max As Double = vbNull
        Dim avg As Double = vbNull
        For Each myvalue As Object In myColl
            If min = vbNull Then
                min = myvalue
                max = myvalue
            End If
            If myvalue < min Then min = myvalue
            If myvalue > max Then max = myvalue
            avg += myvalue
        Next
        avg = avg / myColl.Count
        myResult(0) = min
        myResult(1) = max
        myResult(2) = avg
        Return myResult
    End Function
    Public Shared Function CheckChar(ByVal mystring As String) As Boolean
        Dim myarr(mystring.Length) As Integer
        For i = 0 To mystring.Length - 1
            myarr(i) = Asc(mystring.Substring(i, 1))
        Next
        Return True
    End Function
    Public Shared Function CrossTab(ByVal DataTable As DataTable, ByVal LeftColumn As String, ByVal TopHeader As String, ByVal DataValue As String, Optional ByVal pFix As String = "F_") As DataTable
        If DataTable Is Nothing Then
            Return Nothing
        End If
        Dim DataTableOut As New DataTable
        Dim DataLeftColumn As New DataTable
        Dim DataColHeader As New DataTable

        DataLeftColumn = DataTable.DefaultView.ToTable(True, DataTable.Columns(LeftColumn).ColumnName)
        DataColHeader = DataTable.DefaultView.ToTable(True, DataTable.Columns(TopHeader).ColumnName)

        Dim ColumnLeft As New DataColumn
        ColumnLeft.ColumnName = LeftColumn
        ColumnLeft.Caption = LeftColumn
        ColumnLeft.DataType = System.Type.GetType("System.DateTime")
        DataTableOut.Columns.Add(ColumnLeft)

        'Create Columns
        For Each rowCol As DataRow In DataColHeader.Rows
            Dim col As New DataColumn
            col.ColumnName = pFix & rowCol.Item(TopHeader).ToString.Trim
            DataTableOut.Columns.Add(col)
        Next

        'Create Rows
        Dim DataRow As DataRow
        For Each drow As DataRow In DataLeftColumn.Rows
            DataRow = DataTableOut.NewRow
            DataRow.Item(0) = drow.Item(LeftColumn)
            DataTableOut.Rows.Add(DataRow)
        Next

        Dim xval As Int32 = 0
        Dim yval As Int32 = 0

        'Populate data
        For Each mRow As DataRow In DataTable.Rows
            Dim xrowVal As DateTime = mRow.Item(LeftColumn)
            Dim dataval As String = mRow.Item(DataValue).ToString
            Dim ycolVal As String = mRow.Item(TopHeader).ToString.Trim

            For Each nrow As DataRow In DataTableOut.Rows


                If xrowVal = nrow.Item(0) Then
                    For xval = 0 To nrow.Table.Columns.Count() - 1
                        If nrow.Table.Columns(xval).ColumnName = pFix & ycolVal Then
                            Dim rindex As Int32 = DataTableOut.Rows.IndexOf(nrow)
                            DataTableOut.Rows(rindex).Item(xval) = dataval
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        Next
        'DataTableOut.DefaultView.Sort = DataTableOut.Columns(0).ColumnName

        Return DataTableOut
    End Function


    Public Shared Function escapeString(ByVal vValue As Object) As String
        vValue = Trim(vValue)
        If vValue Is Nothing Or vValue = "" Then
            escapeString = "Null"
        Else
            escapeString = "'" & Replace(Replace(vValue, "", ""), "'", "''") & "'"
        End If

    End Function

    Public Sub FillCombobox(ByRef ComboBox As ComboBox, ByVal Sqlstr As String)
        Using conn As New NpgsqlConnection(getConnectionString)
            conn.Open()
            Dim command As New NpgsqlCommand
            command.Connection = conn
            command.CommandText = Sqlstr
            command.CommandType = CommandType.Text
            Using DataReader As NpgsqlDataReader = command.ExecuteReader
                While DataReader.HasRows
                    DataReader.Read()
                    ComboBox.Items.Add(DataReader.Item(0).ToString)
                End While
            End Using
        End Using
    End Sub

    Public Sub FillComboboxDataSource(ByRef ComboBox As ComboBox, ByVal Sqlstr As String)
        Using conn As New NpgsqlConnection(getConnectionString)
            conn.Open()
            Dim command As New NpgsqlCommand
            command.Connection = conn
            command.CommandText = Sqlstr
            command.CommandType = CommandType.Text
            Dim Dataset As New DataSet
            If getDataSet(Sqlstr, Dataset) Then
                ComboBox.DataSource = Dataset.Tables(0).DefaultView
                ComboBox.ValueMember = Dataset.Tables(0).Columns(0).ToString
                ComboBox.DisplayMember = Dataset.Tables(0).Columns(1).ToString
            End If
        End Using
    End Sub

    Public Sub FillCheckedListBoxDataSource(ByRef CheckedListBox As CheckedListBox, ByVal Sqlstr As String)
        Using conn As New NpgsqlConnection(getConnectionString)
            conn.Open()
            Dim command As New NpgsqlCommand
            command.Connection = conn
            command.CommandText = Sqlstr
            command.CommandType = CommandType.Text
            Dim Dataset As New DataSet
            If getDataSet(Sqlstr, Dataset) Then
                CheckedListBox.DataSource = Dataset.Tables(0).DefaultView
                CheckedListBox.ValueMember = Dataset.Tables(0).Columns(0).ToString
                CheckedListBox.DisplayMember = Dataset.Tables(0).Columns(1).ToString
            End If
        End Using
    End Sub
    Public Function getweekperiod(ByVal StartDate As Date) As Integer
        Dim myResult As Integer
        conn = New NpgsqlConnection(getConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "select date_part('week'," & DateFormatyyyyMMdd(StartDate) & "::date)::integer"
            Dim command As NpgsqlCommand = New NpgsqlCommand(sqlstr, conn)
            myResult = command.ExecuteScalar
        Catch ex As Exception
        Finally
            conn.Close()
        End Try

        Return myResult

    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
