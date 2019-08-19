Imports DJLib
Imports DJLib.Dbtools
Imports Npgsql
Public Class FormGroupFamily
    Dim DataSet As DataSet
    Dim DataSetMaster As DataSet
    Dim dbtools1 As New Dbtools(myUserid, myPassword)
    Dim ConnectionString As String = dbtools1.getConnectionString
    Dim Masterbindingsource As BindingSource
    Dim Detailsbindingsource As BindingSource
    Dim MasterVendor As BindingSource
    Dim DetailsGroup As BindingSource
    Dim DetailsFamily As BindingSource
    Dim currencymanagerGroup As CurrencyManager
    Dim currencymanagerFamily As CurrencyManager
    Dim conn As NpgsqlConnection
    Dim allowValidate As Boolean = False
    Dim allowValidate2 As Boolean = False
    Dim groupId As Long
    Dim FamilyTxId As Long
    Enum Group
        sspsopfamilygroupid
        sopfamilygroup
    End Enum
    Enum GroupTx
        sspsopfamilygrouptxid
        sspsopfamilygroupid
        sspsopfamilyid
    End Enum

    Enum Vendor
        VendorCode
        VendorName
    End Enum

    Enum VendorGroup
        Vendorgroup
        VendorCode
        sspsopfamilygroupid
        GroupName
    End Enum

    Enum VendorFamily
        vendorgroup
        sspsopfamilygroupid
        sspSOPFamilyId
        SOPFamilyName
        sspsopfamilygrouptxid
    End Enum

    Private Class GroupModel
        Public Property sspsofamilygroupid() As Long
        Public Property sopfamilygroup() As String
    End Class
    Private Class GroupTxModel
        Public Property sspsopfamilygrouptxid() As Long
        Public Property sspsopfamilygroupid() As Long
        Public Property sspsopfamilyid() As Long
    End Class

    Private Sub FormGroupFamily_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'clear group record
        Dim sqlstr As String = "delete from sspsopfamilygroup where sspsopfamilygroupid in (select g.sspsopfamilygroupid from sspsopfamilygroup g" & _
                 " left join sspsopfamilygrouptx gtx on gtx.sspsopfamilygroupid = g.sspsopfamilygroupid" & _
                 " where sspsopfamilygrouptxid is null)"
        dbtools1.ExecuteNonQuery(sqlstr)
        MasterVendor.DataSource = Nothing
        DetailsGroup.DataSource = Nothing
        DetailsFamily.DataSource = Nothing
        DataGridView1.DataSource = Nothing
        DataGridView2.DataSource = Nothing
    End Sub
    Private Sub FormGroupFamily_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call loadDataMaster()   
        ComboBox1.DataSource = MasterVendor
        ComboBox1.DisplayMember = "Vendorname"
        ComboBox1.ValueMember = "vendorcode"
        'Call bindDataGrid()
    End Sub
    Private Sub loadDataMaster()
        Try
            DataSetMaster = New DataSet
            Dim stringbuilder1 As New System.Text.StringBuilder
            stringbuilder1.Append("select 1 as vendorcode,'All Vendor' as vendorname union all (select  ssp.vendorcode,v.vendorname from ( select ssp.vendorcode from ssp" & _
                                  " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                                  " left join sspcmmfsop scs on scs.cmmf = scr.cmmf" & _
                                  " left join sspsopfamilies ssf on ssf.sspsopfamilyid = scs.sopfamilyid" & _
                                  " left join sspsopfamilygrouptx gtx on gtx.sspsopfamilyid = ssf.sspsopfamilyid" & _
                                  " left join sspsopfamilygroup g on g.sspsopfamilygroupid = gtx.sspsopfamilygroupid" & _
                                  " group by vendorcode) as ssp" & _
                                  " left join vendor v on v.vendorcode = ssp.vendorcode" & _
                                  " group by ssp.vendorcode,v.vendorname order by vendorname);") 'Vendor
            stringbuilder1.Append("(select '1' || ssp.sspsopfamilygroupid::text as vendorgroup, 1 as vendorcode,ssp.sspsopfamilygroupid,g.sopfamilygroup from (select sspsopfamilygroupid from ssp left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid left join sspcmmfsop scs on scs.cmmf = scr.cmmf left join sspsopfamilies ssf on ssf.sspsopfamilyid = scs.sopfamilyid left join sspsopfamilygrouptx gtx on gtx.sspsopfamilyid = ssf.sspsopfamilyid where" & _
                                  " (Not sspsopfamilygroupid Is null) group by sspsopfamilygroupid) as ssp left join sspsopfamilygroup g on g.sspsopfamilygroupid = ssp.sspsopfamilygroupid) " & _
                                  " union all " & _
                                  " (select ssp.vendorcode::text || ssp.sspsopfamilygroupid::text as vendorgroup, ssp.vendorcode,ssp.sspsopfamilygroupid,g.sopfamilygroup from (select ssp.vendorcode, sspsopfamilygroupid from ssp" & _
                                  " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                                  " left join sspcmmfsop scs on scs.cmmf = scr.cmmf" & _
                                  " left join sspsopfamilies ssf on ssf.sspsopfamilyid = scs.sopfamilyid" & _
                                  " left join sspsopfamilygrouptx gtx on gtx.sspsopfamilyid = ssf.sspsopfamilyid" & _
                                  " where(Not sspsopfamilygroupid Is null)" & _
                                  " group by ssp.vendorcode,sspsopfamilygroupid) as ssp" & _
                                  " left join sspsopfamilygroup g on g.sspsopfamilygroupid = ssp.sspsopfamilygroupid);") 'group
            stringbuilder1.Append("(select  '1' || ssp.sspsopfamilygroupid::text as vendorgroup,ssp.sspsopfamilygroupid,ssp.sspsopfamilyid,ssp.sopdescription,sspsopfamilygrouptxid from " & _
                                   " (select  sspsopfamilygroupid ,ssf.sspsopfamilyid,ssf.sopdescription,sspsopfamilygrouptxid" & _
                                   " from SSP " & _
                                   " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid " & _
                                   " left join sspcmmfsop scs on scs.cmmf = scr.cmmf " & _
                                   " left join sspsopfamilies ssf on ssf.sspsopfamilyid = scs.sopfamilyid " & _
                                   " left join sspsopfamilygrouptx gtx on gtx.sspsopfamilyid = ssf.sspsopfamilyid " & _
                                   "  where(Not sspsopfamilygroupid Is null)" & _
                                   " group by sspsopfamilygroupid,ssf.sspsopfamilyid,ssf.sopdescription,sspsopfamilygrouptxid) as ssp " & _
                                   " ) union all " & _
                                  " (select  ssp.vendorcode::text || ssp.sspsopfamilygroupid::text ,ssp.sspsopfamilygroupid,ssp.sspsopfamilyid,ssp.sopdescription,sspsopfamilygrouptxid from " & _
                                  " (select ssp.vendorcode, sspsopfamilygroupid ,ssf.sspsopfamilyid,ssf.sopdescription,sspsopfamilygrouptxid" & _
                                  " from SSP" & _
                                  " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid " & _
                                  " left join sspcmmfsop scs on scs.cmmf = scr.cmmf " & _
                                  " left join sspsopfamilies ssf on ssf.sspsopfamilyid = scs.sopfamilyid " & _
                                  " left join sspsopfamilygrouptx gtx on gtx.sspsopfamilyid = ssf.sspsopfamilyid " & _
                                  " where(Not sspsopfamilygroupid Is null)" & _
                                  " group by ssp.vendorcode,sspsopfamilygroupid,ssf.sspsopfamilyid,ssf.sopdescription,sspsopfamilygrouptxid) as ssp " & _
                                  " );") 'familytx
            stringbuilder1.Append("(select 1 as vendorcode, foo.sopfamilyid, ssf.sopdescription from (select scs.sopfamilyid from (select ssp.sspcmmfrangeid from ssp" & _
                                  " group by ssp.sspcmmfrangeid) as ssp" & _
                                  " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                                  " left join sspcmmfsop scs on scs.cmmf = scr.cmmf" & _
                                  " where(Not sopfamilyid Is null)" & _
                                  " group by scs.sopfamilyid) as foo" & _
                                  " left join sspsopfamilies ssf on ssf.sspsopfamilyid = foo.sopfamilyid" & _
                                  " order by sopdescription)" & _
                                  " union all" & _
                                  " (select v.vendorcode, foo.sopfamilyid, ssf.sopdescription from (select ssp.vendorcode,scs.sopfamilyid from (select ssp.vendorcode,ssp.sspcmmfrangeid from ssp" & _
                                  " group by ssp.vendorcode,ssp.sspcmmfrangeid) as ssp " & _
                                  " left join sspcmmfrange scr on scr.sspcmmfrangeid = ssp.sspcmmfrangeid" & _
                                  " left join sspcmmfsop scs on scs.cmmf = scr.cmmf" & _
                                  " where(Not sopfamilyid Is null)" & _
                                  " group by ssp.vendorcode,scs.sopfamilyid) as foo" & _
                                  " left join vendor v on v.vendorcode = foo.vendorcode" & _
                                  " left join sspsopfamilies ssf on ssf.sspsopfamilyid = foo.sopfamilyid" & _
                                  " order by vendorcode,sopdescription" & _
                                  " );") 'groupTX for all and each vendor

            dbtools1.getDataSet(stringbuilder1.ToString, DataSetMaster)
            DataSetMaster.Locale = System.Globalization.CultureInfo.InvariantCulture

            DataSetMaster.Tables(0).TableName = "Vendor"
            Dim pkey(0) As DataColumn
            pkey(0) = DataSetMaster.Tables("Vendor").Columns(Vendor.VendorCode)
            DataSetMaster.Tables("Vendor").PrimaryKey = pkey

            DataSetMaster.Tables(1).TableName = "VendorGroup"
            Dim pkey1(1) As DataColumn
            pkey1(0) = DataSetMaster.Tables("VendorGroup").Columns(VendorGroup.VendorCode)
            pkey1(1) = DataSetMaster.Tables("VendorGroup").Columns(VendorGroup.sspsopfamilygroupid)
            DataSetMaster.Tables("VendorGroup").PrimaryKey = pkey1

            DataSetMaster.Tables(2).TableName = "GroupFamily"
            Dim pkey2(2) As DataColumn

            pkey2(0) = DataSetMaster.Tables("GroupFamily").Columns(VendorFamily.vendorgroup)
            pkey2(1) = DataSetMaster.Tables("GroupFamily").Columns(VendorFamily.sspsopfamilygroupid)
            pkey2(2) = DataSetMaster.Tables("GroupFamily").Columns(VendorFamily.sspsopfamilygrouptxid)
            DataSetMaster.Tables("GroupFamily").PrimaryKey = pkey2

            Dim DataRelation1 As New DataRelation("VendorGroupRel", DataSetMaster.Tables(0).Columns(0), DataSetMaster.Tables(1).Columns(1))
            Dim DataRelation2 As New DataRelation("GroupFamilyRel", DataSetMaster.Tables(1).Columns(0), DataSetMaster.Tables(2).Columns(0))
            Dim DataRelation3 As New DataRelation("VendorFamilyRel", DataSetMaster.Tables(0).Columns(0), DataSetMaster.Tables(3).Columns(0))

            MasterVendor = New BindingSource
            DetailsGroup = New BindingSource
            DetailsFamily = New BindingSource

            DataSetMaster.Relations.Add(DataRelation1)
            DataSetMaster.Relations.Add(DataRelation2)
            DataSetMaster.Relations.Add(DataRelation3)
            MasterVendor.DataSource = DataSetMaster
            MasterVendor.DataMember = "Vendor"
            DetailsGroup.DataSource = MasterVendor
            DetailsGroup.DataMember = "VendorGroupRel"
            DetailsFamily.DataSource = DetailsGroup
            DetailsFamily.DataMember = "GroupFamilyRel"
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error LoadDataMaster", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        
    End Sub
   

    Private Sub bindDataGrid()

        Try
            DataGridView1.DataSource = Nothing
            DataGridView2.DataSource = Nothing
            DataGridView1.DataSource = DetailsGroup
            DataGridView2.DataSource = DetailsFamily
            With DataGridView1
                .AutoSize = True
                .TopLeftHeaderCell.Value = "Group"
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With
            With DataGridView2
                .AutoSize = True
                .TopLeftHeaderCell.Value = "Family"
                .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                .RowsDefaultCellStyle.BackColor = Color.White
                .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            End With

            DataGridView1.Columns.Item(VendorGroup.sspsopfamilygroupid).Visible = False
            DataGridView1.Columns.Item(VendorGroup.VendorCode).Visible = False
            DataGridView1.Columns.Item(VendorGroup.Vendorgroup).Visible = False
            DataGridView1.Columns.Item(VendorGroup.GroupName).HeaderText = "Group Description"
            DataGridView1.Columns.Item(VendorGroup.GroupName).Width = 400

            DataGridView2.Columns.Item(VendorFamily.SOPFamilyName).Visible = False
            DataGridView2.Columns.Item(VendorFamily.sspsopfamilygroupid).Visible = False
            DataGridView2.Columns.Item(VendorFamily.sspsopfamilygrouptxid).Visible = False
            DataGridView2.Columns.Item(VendorFamily.sspSOPFamilyId).Visible = False
            DataGridView2.Columns.Item(VendorFamily.vendorgroup).Visible = False

            'datagridview2 addCombobox
            AddComboBoxColumns()
            currencymanagerGroup = CType(Me.BindingContext(DetailsGroup), CurrencyManager)
            currencymanagerFamily = CType(Me.BindingContext(DetailsFamily), CurrencyManager)
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AddComboBoxColumns()
        Dim comboboxColumn As DataGridViewComboBoxColumn
        comboboxColumn = CreateComboBoxColumn()
        SetAlternateChoicesUsingDataSource(comboboxColumn)
        comboboxColumn.HeaderText = "Family Description"
        'DataGridView2.Columns.Insert(1, comboboxColumn)
        DataGridView2.Columns.Add(comboboxColumn)
    End Sub
    Private Function CreateComboBoxColumn() As DataGridViewComboBoxColumn
        Dim column As New DataGridViewComboBoxColumn()
        With column
            .DataPropertyName = "sspsopfamilyid"
            .HeaderText = "Family Description"
            .DropDownWidth = 400
            .Width = 400
            .MaxDropDownItems = 10
            .FlatStyle = FlatStyle.Flat
            .SortMode = DataGridViewColumnSortMode.Automatic
        End With
        Return column
    End Function
    Private Sub SetAlternateChoicesUsingDataSource(ByVal comboboxColumn As DataGridViewComboBoxColumn)

        'Dim mytable As DataTable = DataSet.Tables(2)
        Dim mytable As New DataTable
        Dim myfilter As String
        Try
            Try
                Dim myrow As DataRowView = ComboBox1.SelectedValue
                myfilter = "vendorcode=" & myrow.Item(0)
            Catch ex As Exception
                myfilter = "vendorcode=" & ComboBox1.SelectedValue
            End Try

            DataSetMaster.Tables(3).DefaultView.RowFilter = myfilter
            Dim DataTable As DataTable = DataSetMaster.Tables(3).DefaultView.ToTable
            For Each e As Object In DataTable.Rows
                comboboxColumn.Items.Add(New Family With {.id = e.item(1),
                                                           .Description = e.item(2).ToString})
            Next
            comboboxColumn.DataPropertyName = "sspsopfamilyid"
            comboboxColumn.DisplayMember = "Description"
            comboboxColumn.ValueMember = "id"
            'comboboxColumn.Items.Add(DBNull.Value)
            'comboboxColumn.DefaultCellStyle.NullValue = ""
        Catch ex As Exception

        End Try

    End Sub
    Private Class Family
        Public Property id As Long
        Public Property Description As String
    End Class

#Region "DataGridView1"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DetailsGroup.AddNew()
        Dim myIdx As Long = InsertcmdGroup()
        DataGridView1.Item(VendorGroup.VendorCode.ToString, currencymanagerGroup.Position).Value = ComboBox1.SelectedValue
        DataGridView1.Item(VendorGroup.Vendorgroup, currencymanagerGroup.Position).Value = ComboBox1.SelectedValue.ToString & myIdx.ToString
        DataGridView1.Item(VendorGroup.sspsopfamilygroupid.ToString, currencymanagerGroup.Position).Value = myIdx
        'Button5_Click(Me, e)
    End Sub
    Private Sub DeleteCmdGroup(ByVal id As Long)
        conn = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "delete from  sspsopfamilygroup  where sspsopfamilygroupid=@id"

            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = id
            Dim lnewid As Long = command.ExecuteNonQuery

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
    Private Function InsertcmdGroup() As Long
        conn = New NpgsqlConnection(ConnectionString)
        Dim lnewid As Long
        Try
            conn.Open()
            Dim sqlstr As String = "insert into sspsopfamilygroup(sopfamilygroup) values(null);select currval('sspsopfamilygroup_sspsopfamilygroupid_seq');"
            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            lnewid = command.ExecuteScalar
            'Call InsertcmdGroupTx()
            allowValidate = True
        Catch ex As Exception
        Finally
            conn.Close()
        End Try

        Return lnewid
    End Function
    Private Sub UpdateCmdGroup(ByVal Group As GroupModel)
        conn = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "Update sspsopfamilygroup set sopfamilygroup=@sopfamilygroup  where sspsopfamilygroupid=@sspsopfamilygroupid"
            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspsopfamilygroupid", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = Group.sspsofamilygroupid
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sopfamilygroup", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = IIf(Group.sopfamilygroup = "", DBNull.Value, Group.sopfamilygroup)
            Dim lnewid As Long = command.ExecuteScalar
            allowValidate = False
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        allowValidate = True
    End Sub

    Private Sub DataGridView1_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowValidated
        'Update record here
        If allowValidate Then
            Dim DataRow1 As DataRow = Nothing
            If CheckRowStateChanged(DataRow1) Then
                Call updateRowGroup(DataRow1)
                allowValidate = False
            End If
        End If

    End Sub
    Public Function CheckRowStateChanged(ByRef DataRow As DataRow) As Boolean
        If allowValidate Then
            If Not currencymanagerGroup Is Nothing Then
                Try
                    Dim pkey(1) As Object
                    pkey(0) = DataGridView1.Item(VendorGroup.VendorCode.ToString, currencymanagerGroup.Position).Value
                    pkey(1) = DataGridView1.Item(2, currencymanagerGroup.Position).Value
                    DataRow = DataSetMaster.Tables(1).Rows.Find(pkey)
                    'check any rowchanges
                    If Not (DataRow.RowState = DataRowState.Unchanged) Then
                        Return True
                    End If
                Catch ex As Exception

                End Try

            End If
        End If
        Return False
    End Function
    Public Sub updateRowGroup(ByVal DataRow As DataRow)
        If allowValidate Then
            If Not currencymanagerGroup Is Nothing Then
                Try
                    Call UpdateCmdGroup(New GroupModel With {.sspsofamilygroupid = DataRow.Item(VendorGroup.sspsopfamilygroupid.ToString).ToString,
                                                             .sopfamilygroup = DataRow.Item(VendorGroup.GroupName).ToString})
                    'Commit transaction
                    DataRow.AcceptChanges()
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If MessageBox.Show("Are you sure you wish to delete the selected row?", "Delete Row?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim pkey(1) As Object
            pkey(0) = DataGridView1.Item(VendorGroup.VendorCode, currencymanagerGroup.Position).Value
            pkey(1) = DataGridView1.Item(VendorGroup.sspsopfamilygroupid, currencymanagerGroup.Position).Value

            Dim DataRow As DataRow = DataSetMaster.Tables(1).Rows.Find(pkey)
            Try
                DeleteCmdGroup(DataRow.Item(VendorGroup.sspsopfamilygroupid))
                DataRow.Delete()
            Catch ex As Exception

            End Try


        End If
    End Sub
#End Region

#Region "DataGridView2"
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        DetailsFamily.AddNew()
        FamilyTxId = InsertcmdGroupTx()
        DataGridView2.Item(0, currencymanagerFamily.Position).Value = ComboBox1.SelectedValue.ToString & DataGridView1.Item(VendorGroup.sspsopfamilygroupid, currencymanagerGroup.Position).Value.ToString

        DataGridView2.Item(1, currencymanagerFamily.Position).Value = DataGridView1.Item(VendorGroup.sspsopfamilygroupid, currencymanagerGroup.Position).Value
        DataGridView2.Item(4, currencymanagerFamily.Position).Value = FamilyTxId
        allowValidate2 = True
    End Sub
    Private Sub DeleteCmdGroupTx(ByVal id As Long)
        conn = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "delete from  sspsopfamilygrouptx  where sspsopfamilygrouptxid=@id"

            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@id", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = id
            Dim lnewid As Long = command.ExecuteNonQuery

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub
    Private Function InsertcmdGroupTx() As Long
        conn = New NpgsqlConnection(ConnectionString)
        Dim lnewid As Long
        Try
            conn.Open()
            Dim sqlstr As String = "insert into sspsopfamilygrouptx(sspsopfamilygroupid) values(@sspsopfamilygroupid);select currval('sspsopfamilygrouptx_sspsopfamilygrouptxid_seq');"
            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspsopfamilygroupid", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = DataGridView1.Item(Group.sspsopfamilygroupid.ToString, currencymanagerGroup.Position).Value
            lnewid = command.ExecuteScalar
            allowValidate2 = True
        Catch ex As Exception
        Finally
            conn.Close()
        End Try

        Return lnewid
    End Function
    Private Sub UpdateCmdGroupTx(ByVal Group As GroupTxModel)
        conn = New NpgsqlConnection(ConnectionString)
        Try
            conn.Open()
            Dim sqlstr As String = "Update sspsopfamilygrouptx set sspsopfamilygroupid=@sspsopfamilygroupid, sspsopfamilyid=@sspsopfamilyid where sspsopfamilygrouptxid=@sspsopfamilygrouptxid"
            Dim command As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspsopfamilygrouptxid", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = Group.sspsopfamilygrouptxid
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspsopfamilygroupid", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = Group.sspsopfamilygroupid
            command.Parameters.Add(New Npgsql.NpgsqlParameter("@sspsopfamilyid", NpgsqlTypes.NpgsqlDbType.Bigint)).Value = Group.sspsopfamilyid
            Dim lnewid As Long = command.ExecuteScalar

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            allowValidate = False
            conn.Close()
        End Try
    End Sub


    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If MessageBox.Show("Are you sure you wish to delete the selected row?", "Delete Row?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Dim pkey(2) As Object
            pkey(0) = DataGridView2.Item(VendorFamily.vendorgroup, currencymanagerFamily.Position).Value
            pkey(1) = DataGridView2.Item(VendorFamily.sspsopfamilygroupid, currencymanagerFamily.Position).Value
            pkey(2) = DataGridView2.Item(VendorFamily.sspsopfamilygrouptxid, currencymanagerFamily.Position).Value
            Dim DataRow = DataSetMaster.Tables(2).Rows.Find(pkey)
            Try
                DeleteCmdGroupTx(DataGridView2.Item(VendorFamily.sspsopfamilygrouptxid, currencymanagerFamily.Position).Value)
                DataRow.Delete()
            Catch ex As Exception

            End Try


        End If
    End Sub

    Private Sub DataGridView2_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView2.EditingControlShowing
        allowValidate2 = True
    End Sub

    Private Sub DataGridView2_RowValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.RowValidated
        'Update record here
        If allowValidate2 Then
            Dim DataRow1 As DataRow = Nothing
            If CheckRowStateChangedTx(DataRow1) Then
                Call updateRowGroupTx(DataRow1)
                allowValidate2 = False
            End If
        End If
        'AllowValidate = False
    End Sub
    Public Function CheckRowStateChangedTx(ByRef DataRow As DataRow) As Boolean
        If allowValidate2 Then
            If Not currencymanagerFamily Is Nothing Then
                Try
                    Dim pkey(2) As Object
                    pkey(0) = DataGridView2.Item(VendorFamily.vendorgroup, currencymanagerFamily.Position).Value
                    pkey(1) = DataGridView2.Item(VendorFamily.sspsopfamilygroupid.ToString, currencymanagerFamily.Position).Value
                    pkey(2) = DataGridView2.Item(VendorFamily.sspsopfamilygrouptxid, currencymanagerFamily.Position).Value
                    DataRow = DataSetMaster.Tables(2).Rows.Find(pkey)
                    'check any rowchanges
                    If Not (DataRow.RowState = DataRowState.Unchanged) Then
                        Return True
                    End If
                Catch ex As Exception

                End Try

            End If
        End If
        Return False
    End Function
    Public Sub updateRowGroupTx(ByVal DataRow As DataRow)
        If allowValidate2 Then
            If Not currencymanagerFamily Is Nothing Then
                Try
                    Call UpdateCmdGroupTx(New GroupTxModel With {.sspsopfamilygrouptxid = DataRow.Item(VendorFamily.sspsopfamilygrouptxid.ToString).ToString,
                                                               .sspsopfamilyid = DataRow.Item(VendorFamily.sspSOPFamilyId.ToString).ToString,
                                                               .sspsopfamilygroupid = DataRow.Item(VendorFamily.sspsopfamilygroupid.ToString).ToString})
                    'Commit transaction
                    DataRow.AcceptChanges()
                Catch ex As Exception
                    'MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

#End Region

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call bindDataGrid()
    End Sub



End Class