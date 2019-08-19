Public Class HelperClass

    Public Function getdate(ByVal index As Integer, ByVal period As String) As String
        Dim myresult As String = String.Empty

        Dim myyear = CInt(period.Substring(0, 4))
        Dim mymonth = CInt(period.Substring(4, 2))
        mymonth = mymonth + index
        If mymonth > 12 Then
            mymonth = mymonth - 12
            myyear = myyear + 1
        End If
        myresult = "'" & myyear & "-" & mymonth & "-1'"
        Return myresult

    End Function
    Public Function DateFormatDDMMYYYY(ByRef DateInput As String) As String
        Dim myRet As String = "Null"
        Dim arrDate(2) As String
        Dim arrTmp As String()

        Try
            If DateInput.Contains("/") Then
                arrTmp = DateInput.Split("/")
                arrDate(0) = arrTmp(2)
                arrDate(1) = arrTmp(1)
                arrDate(2) = arrTmp(0)
                myRet = "'" & String.Join("-", arrDate) & "'"
            ElseIf DateInput.Contains("-") Then
                arrTmp = DateInput.Split("-")
                arrDate(0) = arrTmp(2)
                arrDate(1) = arrTmp(1)
                arrDate(2) = arrTmp(0)
                myRet = "'" & String.Join("-", arrDate) & "'"
            End If
        Catch ex As Exception
            
        End Try
        Return myRet
    End Function

End Class
