Imports Npgsql
Imports DJLib
Imports DJLib.Dbtools
Namespace classes
    Public Class DBClass
        Shared conn As NpgsqlConnection
        Shared cmd As NpgsqlCommand
        Shared Sqlstr As String
        Shared dbtools1 As New Dbtools(myUserid, myPassword)
        Shared ConnectionString = dbtools1.getConnectionString


        Public Class WeeklyModel
            'number -> string. except Date
            Public weeklyid As String 'Long
            Public yearweek As String 'Integer
            Public startdate As Date
            Public monthly As Date
            Public label As String 'String
            Public crossmonth As Boolean

            Public Shared Sub InsertWeekly(ByVal weekly As WeeklyModel)
                conn = New NpgsqlConnection(ConnectionString)
                Sqlstr = "Insert into sspweekly(yearweek,startdate,monthly,label,crossmonth) values (@yearweek,@startdate,@monthly,@label,@crossmonth)"
                Try
                    conn.Open()
                    cmd = New NpgsqlCommand(Sqlstr, conn)
                    cmd.Parameters.Add("@yearweek", NpgsqlTypes.NpgsqlDbType.Integer).Value = weekly.yearweek
                    cmd.Parameters.Add("@startdate", NpgsqlTypes.NpgsqlDbType.Date).Value = weekly.startdate
                    cmd.Parameters.Add("@monthly", NpgsqlTypes.NpgsqlDbType.Date).Value = weekly.monthly
                    cmd.Parameters.Add("@label", NpgsqlTypes.NpgsqlDbType.Varchar).Value = weekly.label
                    cmd.Parameters.Add("@crossmonth", NpgsqlTypes.NpgsqlDbType.Boolean).Value = weekly.crossmonth
                    cmd.ExecuteNonQuery()
                Catch ex As Exception

                Finally
                    conn.Close()
                End Try
            End Sub

            Public Shared Sub UpdateWeekly(ByVal weekly As WeeklyModel)
                conn = New NpgsqlConnection(ConnectionString)

                Try
                    conn.Open()
                    Sqlstr = "Update sspweekly set label=@label,crossmonth=@crossmonth where sspweeklyid=@sspweeklyid"
                    cmd = New Npgsql.NpgsqlCommand(Sqlstr, conn)
                    cmd.Parameters.Add("@sspweeklyid", NpgsqlTypes.NpgsqlDbType.Bigint).Value = weekly.weeklyid
                    cmd.Parameters.Add("@label", NpgsqlTypes.NpgsqlDbType.Varchar).Value = weekly.label
                    cmd.Parameters.Add("@crossmonth", NpgsqlTypes.NpgsqlDbType.Boolean).Value = weekly.crossmonth
                    Dim lnewid As Long = cmd.ExecuteScalar
                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    conn.Close()
                End Try
            End Sub


        End Class

        Public Class MonthlyModel
            Public monthlyid As String 'Long
            Public period As String 'Integer
            Public mydate As Date


            Public Shared Sub InsertMonthly(ByVal monthly As MonthlyModel)
                conn = New NpgsqlConnection(ConnectionString)
                Sqlstr = "Insert into sspmonthly(period,mydate) values (@period,@mydate)"
                Try
                    conn.Open()
                    cmd = New NpgsqlCommand(Sqlstr, conn)
                    cmd.Parameters.Add("@period", NpgsqlTypes.NpgsqlDbType.Integer).Value = monthly.period
                    cmd.Parameters.Add("@mydate", NpgsqlTypes.NpgsqlDbType.Date).Value = monthly.mydate
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                Finally
                    conn.Close()
                End Try
            End Sub

        End Class

        Public Class DailyModel
            Public Property dailyid As String
            Public Property weekpriod As String
            Public Property description As String
            Public Property monthperiod As Date
            Public Property holiday As String 'Integer
            Public Property dailydate As Date
            Public Property isholiday As Boolean

            Public Shared Sub UpdateCmd(ByVal dailymodel As DailyModel)
                conn = New NpgsqlConnection(ConnectionString)
                Try
                    conn.Open()
                    Dim sqlstr As String = "Update sspdaily set holiday=@holiday,description=@description,isholiday=@isholiday where dailyid=@dailyid"
                    Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
                    command.Parameters.Add(New Npgsql.NpgsqlParameter("@dailyid", NpgsqlTypes.NpgsqlDbType.Integer)).Value = dailymodel.dailyid
                    command.Parameters.Add(New Npgsql.NpgsqlParameter("@description", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = If(dailymodel.description = "", DBNull.Value, dailymodel.description)
                    command.Parameters.Add(New Npgsql.NpgsqlParameter("@holiday", NpgsqlTypes.NpgsqlDbType.Integer)).Value = If(dailymodel.isholiday, 1, 0)
                    command.Parameters.Add(New Npgsql.NpgsqlParameter("@isholiday", NpgsqlTypes.NpgsqlDbType.Boolean)).Value = dailymodel.isholiday

                    Dim lnewid As Long = command.ExecuteScalar

                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    conn.Close()
                End Try
            End Sub
        End Class
        Public Class Reports
            Public Property YearWeek As Integer
            Public Function getStartingDate() As Date
                Dim sqlstr As String = "select startdate from weektomonth where yearweek = " & YearWeek
                Dim mydate As Date
                Using conn As New NpgsqlConnection(ConnectionString)
                    conn.Open()
                    Dim command As New NpgsqlCommand(sqlstr, conn)
                    mydate = command.ExecuteScalar
                End Using
                Return (mydate)
            End Function
        End Class
    End Class
End Namespace