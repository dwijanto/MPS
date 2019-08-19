Imports System.Configuration
Public Class SecureConnectionString
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Shared Function ProtectConnectionString(ByRef protect As Boolean, ByRef message As String) As Boolean
        Dim boolchanged As Boolean = False
        Dim strProvider As String = "DataProtectionConfigurationProvider"
        Try
            Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath)
            Dim section As ConnectionStringsSection = config.GetSection("connectionStrings")
            If protect Then
                section.SectionInformation.ProtectSection(strProvider)
            Else
                section.SectionInformation.UnprotectSection()
            End If
            section.SectionInformation.ForceSave = True
            config.Save()

            boolchanged = True
        Catch ex As Exception
            message = ex.Message
        End Try
        Return boolchanged
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
