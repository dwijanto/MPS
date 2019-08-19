Imports System.Windows.Forms



Public Class StatusColumn
    Inherits DataGridViewColumn

    Dim _defaultStatus As StatusImage = StatusImage.Red

    Public Sub New()
        MyBase.New(New StatusCell())
    End Sub

    Public Property defaultStatus() As StatusImage
        Get
            Return _defaultStatus

        End Get
        Set(ByVal value As StatusImage)
            _defaultStatus = value
        End Set
    End Property

    Public Overrides Function Clone() As Object
        Dim col As StatusColumn = TryCast(MyBase.Clone(), StatusColumn)
        col.defaultStatus = _defaultStatus
        Return col
    End Function

    Public Overrides Property CellTemplate As System.Windows.Forms.DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As System.Windows.Forms.DataGridViewCell)
            If (value Is Nothing) OrElse Not (TypeOf value Is StatusCell) Then
                Throw New ArgumentException("Invalid cell Type, StatusColumns can only contain StatusCells")
            End If
            MyBase.CellTemplate = value
        End Set
    End Property
End Class
