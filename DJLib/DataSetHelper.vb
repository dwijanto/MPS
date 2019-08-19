Public Class DataSetHelper
    Public DataSetHelper As DataSet
    Private m_FieldInfo As ArrayList, m_FieldList As String
    Public Sub New(ByVal DataSet As DataSet)
        DataSetHelper = DataSet
    End Sub
    Public Sub New()
        DataSetHelper = Nothing
    End Sub


    Private Class FieldInfo
        Public RelationName As String
        Public Fieldname As String      'source table field name
        Public FieldAlias As String     'destination table field name
        Public Aggregate As String
    End Class

    Private Sub parseFieldList(ByVal FieldList As String, Optional ByVal AllowRelation As Boolean = False)
        '
        ' Parses FieldList into FieldInfo objects and adds them to the m_FieldInfo private member
        '
        ' FieldList syntax: [relationname.]fieldname[ alias],...
        '
        If m_fieldList = FieldList Then Exit Sub
        m_FieldInfo = New ArrayList()
        m_FieldList = FieldList
        Dim field As FieldInfo, FieldParts() As String, Fields() As String = FieldList.Split(",")
        Dim i As Integer
        For i = 0 To Fields.Length - 1
            field = New FieldInfo
            '
            'Parse FieldAlias
            '
            FieldParts = Fields(i).Trim().Split(" ")
            Select Case FieldParts.Length
                Case 1
                    ' to be set at the end of the loop
                Case 2
                    field.FieldAlias = FieldParts(1)
                Case Else
                    Throw New ArgumentException("Too many spaces in field definition: '" & Fields(i) & "'.")
            End Select
            '
            'Parse Fieldname and RelationName
            '
            FieldParts = FieldParts(0).Split(".")
            Select Case FieldParts.Length
                Case 1
                    field.Fieldname = FieldParts(0)
                Case 2
                    If Not AllowRelation Then
                        Throw New ArgumentException("Relation specifiers not allowed in field list: '" & Fields(i) & "'.")
                    End If
                    field.RelationName = FieldParts(0).Trim()
                    field.Fieldname = FieldParts(1).TrimEnd
                Case Else
                    Throw New ArgumentException("Invalid field definition: '" & Fields(i) & "'.")
            End Select
            If field.FieldAlias = "" Then field.FieldAlias = field.Fieldname
            m_FieldInfo.Add(field)
        Next

    End Sub
    Public Function CreateJoinTable(ByVal TableName As String, _
                                    ByVal SourceTable As DataTable, _
                                    ByVal FieldList As String) As DataTable
        '
        ' Creates a table based on fields of another table and related parent tables
        '
        ' FieldList syntax : [relationname.]fieldname[ alias][,[relationname.]fieldname[ alias]]....
        '
        ' dt = dsHelper.CreateJoinTable("TestTable", ds.Tables!Employees, "FirstName FName,LastName LName,DepartmentEmployee.DepartmentName Department")

        If FieldList = "" Then
            Throw New ArgumentException("You must specify at least one field in the field list.")
            'return CrateTable(Tablename,SourceTable)
        Else
            Dim DataTable As New DataTable(TableName)
            parseFieldList(FieldList, True)
            Dim Field As FieldInfo, DataColumn As DataColumn
            For Each Field In m_FieldInfo
                If Field.RelationName = "" Then
                    DataColumn = SourceTable.Columns(Field.Fieldname)
                    DataTable.Columns.Add(Field.FieldAlias, DataColumn.DataType)
                Else
                    DataColumn = SourceTable.ParentRelations(Field.RelationName).ParentTable.Columns(Field.Fieldname)
                    DataTable.Columns.Add(Field.FieldAlias, DataColumn.DataType)
                End If
            Next
            If Not DataSetHelper Is Nothing Then DataSetHelper.Tables.Add(DataTable)
            Return DataTable
        End If
    End Function
    Public Sub InsertJoinInto(ByVal DestTable As DataTable, _
                          ByVal SourceTable As DataTable, _
                          ByVal FieldList As String, _
                          Optional ByVal RowFilter As String = "", _
                          Optional ByVal Sort As String = "")
        '
        ' Copies the selected rows and columns from SourceTable and then inserts them to DestTable
        ' FieldList has same format as CreateJoinTable
        '
        If FieldList = "" Then
            Throw New ArgumentException("You must specify at least one field in the field list.")
            ' InsertInto(DestTable, SourceTable, RowFilter, Sort)
        Else
            ParseFieldList(FieldList, True)
            Dim Field As FieldInfo
            Dim Rows() As DataRow = SourceTable.Select(RowFilter, Sort)
            Dim SourceRow, ParentRow, DestRow As DataRow
            For Each SourceRow In Rows
                DestRow = DestTable.NewRow()
                For Each Field In m_FieldInfo
                    If Field.RelationName = "" Then
                        DestRow(Field.FieldAlias) = SourceRow(Field.FieldName)
                    Else
                        Try
                            ParentRow = SourceRow.GetParentRow(Field.RelationName)
                            DestRow(Field.FieldAlias) = ParentRow(Field.Fieldname)
                        Catch ex As Exception
                        End Try
                    End If
                Next
                DestTable.Rows.Add(DestRow)
            Next
        End If
    End Sub

    Public Function SelectJoinInto(ByVal TableName As String, _
                                   ByVal SourceTable As DataTable, _
                                   ByVal FieldList As String, _
                                   Optional ByVal RowFilter As String = "", _
                                   Optional ByVal Sort As String = "") As DataTable
        Dim dt As DataTable = CreateJoinTable(TableName, SourceTable, FieldList)
        InsertJoinInto(dt, SourceTable, FieldList, RowFilter, Sort)

        Return dt
    End Function
End Class
