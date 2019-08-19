Imports Npgsql
Imports NpgsqlTypes
Imports System.IO
Imports SSP.PublicClass

Public Class DBAdapter
    Implements IDisposable
    Private connectionstring
    Dim mytransaction As NpgsqlTransaction

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

    Public Sub New()
        connectionstring = dbtools1.getConnectionString
    End Sub

    Private Sub onRowUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If

        End If
    End Sub

    Private Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub

    Function SupplierCategoryTx(ByVal formSupplierCategory As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "doc.sp_updatesuppliercategory"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_insertsuppliercategory"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "category").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "myorder").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "suppliercategoryid").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "doc.sp_deletesuppliercategory"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "supplierscategoryid").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Public Overloads Function TbgetDataSet(ByVal sqlstr As String, ByRef DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New NpgsqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New NpgsqlConnection(connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                'DataAdapter.MissingSchemaAction = MissingSchemaAction.Add
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As NpgsqlException
            Dim obj = TryCast(ex.Errors(0), NpgsqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.InternalQuery
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function

    Function MonthlyTableTx(ByVal formMonthlyTable As FormMonthlyTable, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowUpdate)
        Try
            Using conn As New NpgsqlConnection(connectionstring)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                sqlstr = "sp_updatemonthlytabletx"
                DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "monthly").SourceVersion = DataRowVersion.Current
                DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "weekly").SourceVersion = DataRowVersion.Current

                DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_insertmontlytabletx"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "monthly").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "weekly").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").Direction = ParameterDirection.InputOutput
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sp_deletemonthlytabletx"
                DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "id").SourceVersion = DataRowVersion.Original
                DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                DataAdapter.UpdateCommand.Transaction = mytransaction
                DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function
    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New NpgsqlConnection(connectionstring)
            conn.Open()
            Dim cmd As NpgsqlCommand = New NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function

End Class

Public Class ContentBaseEventArgs
    Inherits EventArgs
    Public Property dataset As DataSet
    Public Property message As String
    Public Property hasChanges As Boolean
    Public Property ra As Integer
    Public Property continueonerror As Boolean

    Public Sub New(ByVal dataset As DataSet, ByRef haschanges As Boolean, ByRef message As String, ByRef recordaffected As Integer, ByVal continueonerror As Boolean)
        Me.dataset = dataset
        Me.message = message
        Me.ra = ra
        Me.continueonerror = continueonerror
    End Sub
End Class