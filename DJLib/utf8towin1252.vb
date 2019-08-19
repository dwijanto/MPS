Imports System.Text
Imports System.IO
Public Class utf8towin1252


    Dim myutf8win1252 As New Collection

    Public Sub New()
        myutf8win1252.Add(218, "239")
        myutf8win1252.Add(32, "191")
        myutf8win1252.Add(32, "189")

    End Sub

    Public Sub convert(ByRef myString As String)
        Dim memoryStream As New MemoryStream
        Dim utf8bytes As Byte() = Encoding.UTF8.GetBytes(myString)
        Dim myresult As Byte = Nothing
        Dim stringbuilder1 As New StringBuilder
        For i = 0 To utf8bytes.Length - 1
            Try
                utf8bytes(i) = myutf8win1252(utf8bytes(i).ToString)
            Catch ex As Exception
            End Try
        Next

        For i = 0 To utf8bytes.Length - 1
            Try
                stringbuilder1.Append(Chr(utf8bytes(i)))
            Catch ex As Exception
            End Try
        Next
        myString = stringbuilder1.ToString

        'memoryStream.Write(utf8bytes, 0, utf8bytes.Length)
        'Dim sr As StreamReader = New StreamReader(memoryStream)
        'myString = sr.ReadToEnd()
    End Sub
End Class
