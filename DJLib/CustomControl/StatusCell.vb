Imports System.Windows.Forms
Imports System.Reflection
Imports System.Drawing


Public Enum StatusImage
    Green
    Yellow
    Red
End Enum

Public Class StatusCell
    Inherits DataGridViewImageCell

    Public Sub New()
        Me.ImageLayout = DataGridViewImageCellLayout.Zoom
    End Sub

    Protected Overrides Function GetFormattedValue(ByVal value As Object, ByVal rowIndex As Integer, ByRef cellStyle As System.Windows.Forms.DataGridViewCellStyle, ByVal valueTypeConverter As System.ComponentModel.TypeConverter, ByVal formattedValueTypeConverter As System.ComponentModel.TypeConverter, ByVal context As System.Windows.Forms.DataGridViewDataErrorContexts) As Object
        Dim resource As String = "FirstBindingApp.Red.bmp"
        Dim status As StatusImage = StatusImage.Red
        Dim owningCol As StatusColumn = TryCast(OwningColumn, StatusColumn)

        If owningCol IsNot Nothing Then
            status = owningCol.defaultStatus
        End If
        If TypeOf value Is StatusImage OrElse TypeOf value Is Integer Then
            status = DirectCast(value, StatusImage)
        End If
        Dim img As Image = Nothing
        Select Case status
            Case StatusImage.Green
                resource = "DataGridViewAutoFilter.Green.bmp"
                img = My.Resources.Green
                Exit Select
            Case StatusImage.Yellow
                resource = "DataGridViewAutoFilter.Yellow.bmp"
                img = My.Resources.Yellow
                Exit Select
            Case StatusImage.Red
                resource = "DataGridViewAutoFilter.Red.bmp"
                img = My.Resources.Red
                Exit Select
            Case Else
                Exit Select
        End Select
        'Dim loadedassembly As Assembly = Assembly.GetExecutingAssembly
        'Dim stream As IO.Stream = loadedassembly.GetManifestResourceStream(resource)
        'Dim img As Image = Image.FromStream(stream)
        cellStyle.Alignment = DataGridViewContentAlignment.TopCenter
        Return img
    End Function
End Class
