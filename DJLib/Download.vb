Imports System.Net
Imports System.IO
Imports System.Text
Imports System.ComponentModel
Public Class Download
    Implements IDisposable

    Public Event StateChange(ByVal progress As Int64, ByVal obj As Object)
    Public Filesize As Long
    Private m_url As String
    Private m_saveto As String
    Private m_useproxy As Boolean
    Private m_proxyserver As String
    Private m_username As String
    Private m_password As String
    Private m_server As String
    Private m_Result As Boolean


    Public Sub New(ByVal url As String, ByVal saveto As String, ByVal useproxy As Boolean, ByVal proxyserver As String, ByVal username As String, ByVal password As String, ByVal server As String)
        m_url = url
        m_saveto = saveto
        m_useproxy = useproxy
        m_proxyserver = proxyserver
        m_username = username
        m_password = password
        m_server = server
    End Sub
    Public Property getSaveTo()
        Get
            Return m_saveto
        End Get
        Set(ByVal value)
            m_saveto = value
        End Set
    End Property
    Public ReadOnly Property getResult()
        Get
            Return m_Result
        End Get
    End Property
    Public ReadOnly Property getUrl()
        Get
            Return m_url
        End Get
    End Property
    Public Function Start(ByRef message As String, ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As Boolean
        Dim myhttpwebrequest As HttpWebRequest = Nothing
        Dim myhttpWebResponse As HttpWebResponse = Nothing
        Dim receiveStream As Stream
        Dim progress As Long
        Dim FS As FileStream = Nothing
        Dim downBuffer(2048) As Byte
        Dim iCount As Integer

        Try
            RaiseEvent StateChange(7, "Preparing Connection..")
            myhttpwebrequest = CType(WebRequest.Create(m_url), HttpWebRequest)


            If m_useproxy Then
                myhttpwebrequest.Proxy = New WebProxy(m_proxyserver, True)
                myhttpwebrequest.Proxy.Credentials = New NetworkCredential(m_username, m_password, m_server)
            End If
            myhttpwebrequest.UserAgent = "Raw Material Request"
            Debug.Print(myhttpwebrequest.Timeout)
            myhttpwebrequest.Timeout = 300000
            'System.Windows.Forms.MessageBox.Show(myhttpwebrequest.Timeout, "New Timeout", Windows.Forms.MessageBoxButtons.OK)
            RaiseEvent StateChange(1, "Waiting for Response..." & myhttpwebrequest.Timeout)
            myhttpWebResponse = CType(myhttpwebrequest.GetResponse, HttpWebResponse)
            If worker.CancellationPending Then
                e.Cancel = True
            Else

                Filesize = myhttpWebResponse.ContentLength
                RaiseEvent StateChange(1, "Receive Stream..")
                receiveStream = myhttpWebResponse.GetResponseStream()
                RaiseEvent StateChange(1, "Preparing File..")
                m_url = "2.csv"
                Dim filename = m_saveto & "\" & Path.GetFileName(m_url)
                RaiseEvent StateChange(1, filename)
                'FS = New FileStream(m_saveto & "\" & Path.GetFileName(m_url), FileMode.Create, FileAccess.Write, FileShare.None)
                FS = New FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None)
                Dim count As Integer = receiveStream.Read(downBuffer, 0, downBuffer.Length)
                progress = count

                While count > 0
                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit While
                    Else

                        FS.Write(downBuffer, 0, count)
                        count = receiveStream.Read(downBuffer, 0, downBuffer.Length)
                        progress += count
                        iCount += 1
                        If iCount = 5 Then
                            'RaiseEvent StateChange(progress, "download")
                            iCount = 0
                        End If
                    End If
                End While

                FS.Close()
                'RaiseEvent StateChange(progress, "download")
                RaiseEvent StateChange(6, "download completed.")
                m_Result = True
            End If
        
        Catch ew As WebException
            If ew.Status = WebExceptionStatus.ProtocolError Then
                Dim resp As WebResponse = ew.Response
                Using sr As StreamReader = New StreamReader(resp.GetResponseStream)
                    Dim myFullPath = System.AppDomain.CurrentDomain.BaseDirectory & "responseError.txt"
                    Using writer As StreamWriter = New StreamWriter(myFullPath)
                        writer.WriteLine(sr.ReadToEnd())
                    End Using                   
                End Using
            End If
        Catch ex As Exception
            message = ex.Message
            m_Result = False
        Finally
            If e.Cancel = True Then
                m_Result = False
            End If
            If Not FS Is Nothing Then FS.Close()
            If Not myhttpWebResponse Is Nothing Then myhttpWebResponse.Close()
        End Try
        Return m_Result
    End Function

    Public Function bytesToLabel(ByVal value As Long) As String
        Dim myret As String = vbEmpty
        Dim desc() As String = {"B", "KB", "MB", "GB", "TB"}
        Dim ratio As Long = 1
        Dim index As Integer = 0
        While (value / ratio) > 1024
            index += 1
            ratio *= 1024
        End While
        myret = Format(value / ratio, "0.0") & " " & desc(index)
        Return myret
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
