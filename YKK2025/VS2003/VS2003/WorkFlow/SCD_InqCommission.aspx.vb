Imports System.Data
Imports System.Data.OleDb

Public Class SCD_InqCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DProgress As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DRno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCodeNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKeepData As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     �ۭq�[���w
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SCD_InqCommission.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���
        pFormNo = Request.QueryString("pFormNo")    '��渹�X
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim wSts As String
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '002001' And FormNo <= '002099' "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (InqUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or InqUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '�̿ೡ��
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

        '���
        DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue

        Search_Item_Attribute()

    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            If pFormNo = "002002" Then
                wTableName = "V_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1") + "_02"
            Else
                wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
            End If
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then

            DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
            DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
            DataGrid1.Columns.Item(2).HeaderText = "�̿��"
            DataGrid1.Columns.Item(3).HeaderText = "�̿��"
            SQL = "SELECT "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
            If pFormNo >= "002002" And pFormNo >= "002003" Then
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����' When '2' Then '�}�oNG' Else '�}�o����' End As Field2, "
            Else
                If pFormNo = "002002" Then
                    SQL &= "Case when " + wTableName + ".Sts = 0 Then '�}�o��' When " + wTableName + ".sts = 1 and  " + wTableName + ".sts_norder =1 then '�}�o����(������`)'  When " + wTableName + ".sts = 1  Then '�}�o����(OK)' When " + wTableName + ".sts = 2 Then '�}�o����(NG)'  Else '����/����'  End As Field2, "
                Else
                    SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����' Else '�}�o����' End As Field2, "
                End If
            End If
            SQL = SQL + "V_WaitHandle_01.Division + '--' + "
            SQL = SQL + "V_WaitHandle_01.ApplyName as Field3, "
            SQL = SQL + "V_WaitHandle_01.ApplyTime as Field4, "

            If pFormNo = "002001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(6).HeaderText = "�}�oNo"
                DataGrid1.Columns.Item(7).HeaderText = "Code"

                SQL = SQL + wTableName + ".AppBuyer   as Field5, "
                SQL = SQL + wTableName + ".Rno        as Field6, "
                SQL = SQL + wTableName + ".DevNo      as Field7, "
                SQL = SQL + wTableName + ".CodeNo     as Field8, "
            End If
            If pFormNo >= "002002" And pFormNo <= "002003" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(6).HeaderText = "�}�oNo"
                DataGrid1.Columns.Item(7).HeaderText = "Code"

                SQL = SQL + wTableName + ".AppBuyer   as Field5, "
                SQL = SQL + wTableName + ".No         as Field6, "
                SQL = SQL + wTableName + ".DevNo      as Field7, "
                SQL = SQL + wTableName + ".CodeNo     as Field8, "
            End If

            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_01.FormSno,Len(V_WaitHandle_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_01.Step,Len(V_WaitHandle_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_01.SeqNo,Len(V_WaitHandle_01.SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_01.ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "
            '------------------------------------
            SQL = SQL + "Where V_WaitHandle_01.Step  < '10' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If
            '�}�o���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If
            '�}�o�������A
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If
            'No
            If DNo.Text <> "�e�U��No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".DivName = '" + DDivision.SelectedValue + "'"
            End If
            '�̿�H
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                If pFormNo = "002002" Then
                    SQL = SQL + " And " + wTableName + ".AppPer Like '%" + DPerson.Text + "%'"
                Else
                    SQL = SQL + " And " + wTableName + ".WF1 Like '%" + DPerson.Text + "%'"
                End If
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".AppBuyer Like '%" + DBuyer.Text + "%'"
            End If
            '�e�UNo
            If DRno.Text <> "�e�UNo" And DRno.Text <> "" Then
                If pFormNo = "002002" Then
                    SQL = SQL + " And " + wTableName + ".No Like '%" + DRno.Text + "%'"
                Else
                    SQL = SQL + " And " + wTableName + ".Rno Like '%" + DRno.Text + "%'"
                End If
            End If
            '�}�oNo
            If DDevNo.Text <> "�}�oNo" And DDevNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".DevNo Like '%" + DDevNo.Text + "%'"
            End If
            'Code-No
            If DCodeNo.Text <> "Code-No" And DCodeNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".CodeNo Like '%" + DCodeNo.Text + "%'"
            End If
            '������
            SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"

            'Sort
            SQL = SQL + " Order by " + wTableName + ".CreateTime desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()

            OleDbConnection1.Close()
        End If

    End Sub

    '****************************************************
    '  �w�ʦs���
    '****************************************************
    Sub KeepDataList()
        Dim i As Integer = 0
        Dim wTableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Flow")
        If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
            If pFormNo = "002002" Then
                wTableName = "V_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1") + "_02"
            Else
                wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
            End If
        End If
        OleDbConnection1.Close()

        If wTableName <> "" Then

            DataGrid1.Columns.Item(0).HeaderText = "�e�UNo"
            DataGrid1.Columns.Item(1).HeaderText = "�}�o���A"
            DataGrid1.Columns.Item(2).HeaderText = "�̿��"
            DataGrid1.Columns.Item(3).HeaderText = "�̿��"
            SQL = "SELECT "
            SQL = SQL + "Case " + wTableName + ".No When '' Then '���s��' Else " + wTableName + ".No End As Field1, "
            If pFormNo >= "002002" And pFormNo >= "002003" Then
                SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����' When '2' Then '�}�oNG' Else '�}�o����' End As Field2, "
            Else
                If pFormNo = "002002" Then
                    SQL &= "Case when " + wTableName + ".Sts = 0 Then '�}�o��' When " + wTableName + ".sts = 1 and  " + wTableName + ".sts_norder =1 then '�}�o����(������`)'  When " + wTableName + ".sts = 1  Then '�}�o����(OK)' When " + wTableName + ".sts = 2 Then '�}�o����(NG)'  Else '����/����'  End As Field2, "
                Else
                    SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����' Else '�}�o����' End As Field2, "
                End If

            End If
            SQL = SQL + "V_WaitHandle_OLD_01.Division + '--' + "
            SQL = SQL + "V_WaitHandle_OLD_01.ApplyName as Field3, "
            SQL = SQL + "V_WaitHandle_OLD_01.ApplyTime as Field4, "

            If pFormNo = "002001" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(6).HeaderText = "�}�oNo"
                DataGrid1.Columns.Item(7).HeaderText = "Code"

                SQL = SQL + wTableName + ".AppBuyer   as Field5, "
                SQL = SQL + wTableName + ".Rno        as Field6, "
                SQL = SQL + wTableName + ".DevNo      as Field7, "
                SQL = SQL + wTableName + ".CodeNo     as Field8, "
            End If
            If pFormNo >= "002002" And pFormNo <= "002003" Then
                DataGrid1.Columns.Item(4).HeaderText = "Buyer"
                DataGrid1.Columns.Item(5).HeaderText = "�e�UNo"
                DataGrid1.Columns.Item(6).HeaderText = "�}�oNo"
                DataGrid1.Columns.Item(7).HeaderText = "Code"

                SQL = SQL + wTableName + ".AppBuyer   as Field5, "
                SQL = SQL + wTableName + ".No         as Field6, "
                SQL = SQL + wTableName + ".DevNo      as Field7, "
                SQL = SQL + wTableName + ".CodeNo     as Field8, "
            End If

            SQL = SQL + " '....' as WorkFlow, ViewURL, "
            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + V_WaitHandle_OLD_01.FormNo + "
            SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_OLD_01.FormSno,Len(V_WaitHandle_OLD_01.FormSno)) + "
            SQL = SQL + "'&pStep='    + str(V_WaitHandle_OLD_01.Step,Len(V_WaitHandle_OLD_01.Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_OLD_01.SeqNo,Len(V_WaitHandle_OLD_01.SeqNo)) + "
            SQL = SQL + "'&pKeepdata='   + rtrim(ltrim(str(1))) + "
            SQL = SQL + "'&pApplyID=' + V_WaitHandle_OLD_01.ApplyID "
            SQL = SQL + " As OPURL "
            SQL = SQL + " FROM " + wTableName + " "
            SQL = SQL + " Left Outer Join V_WaitHandle_OLD_01 ON " + wTableName + ".FormNo=V_WaitHandle_OLD_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_OLD_01.FormSno "
            '------------------------------------
            SQL = SQL + "Where V_WaitHandle_OLD_01.Step  < '10' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".FormNo = '" + DFormName.SelectedValue + "'"
            End If
            '�}�o���A
            If DProgress.SelectedValue <> "ALL" Then
                If DProgress.SelectedValue = "1" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '0'  "
                End If
                If DProgress.SelectedValue = "2" Then
                    SQL = SQL + " And " + wTableName + ".Sts =  '1'  "
                End If
            End If
            '�}�o�������A
            If DSts.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".Sts = '" + DSts.SelectedValue + "'"
            End If
            'No
            If DNo.Text <> "�e�U��No." And DNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".No Like '%" + DNo.Text + "%'"
            End If
            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And " + wTableName + ".DivName = '" + DDivision.SelectedValue + "'"
            End If
            '�̿�H
            If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
                If pFormNo = "002002" Then
                    SQL = SQL + " And " + wTableName + ".AppPer Like '%" + DPerson.Text + "%'"
                Else
                    SQL = SQL + " And " + wTableName + ".WF1 Like '%" + DPerson.Text + "%'"
                End If
            End If
            'Buyer
            If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".AppBuyer Like '%" + DBuyer.Text + "%'"
            End If
            '�e�UNo
            If DRno.Text <> "�e�UNo" And DRno.Text <> "" Then
                If pFormNo = "002002" Then
                    SQL = SQL + " And " + wTableName + ".No Like '%" + DRno.Text + "%'"
                Else
                    SQL = SQL + " And " + wTableName + ".Rno Like '%" + DRno.Text + "%'"
                End If
            End If

            '�}�oNo
            If DDevNo.Text <> "�}�oNo" And DDevNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".DevNo Like '%" + DDevNo.Text + "%'"
            End If
            'Code-No
            If DCodeNo.Text <> "Code-No" And DCodeNo.Text <> "" Then
                SQL = SQL + " And " + wTableName + ".CodeNo Like '%" + DCodeNo.Text + "%'"
            End If
            '������
            SQL = SQL + " And " + wTableName + ".CompletedTime between '" + DSDate.Text + " 00:00:00' and '" + DEDate.Text + " 23:59:59'"

            'Sort
            SQL = SQL + " Order by " + wTableName + ".CreateTime desc "

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "WaitHandle")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()


            OleDbConnection1.Close()
        End If
    End Sub

    Sub Search_Item_Attribute()
        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DRno.ReadOnly = True
        DDevNo.ReadOnly = True
        DCodeNo.ReadOnly = True

        'No
        DNo.Text = "�e�U��No."
        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"
        '�e�UNo
        DRno.Text = "�e�UNo"
        '�}�oNo
        DDevNo.Text = "�}�oNo"
        'Code-No
        DCodeNo.Text = "Code-No"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        If pFormNo = "002001" Then
            DBuyer.ReadOnly = False
            DRno.ReadOnly = False
            DDevNo.ReadOnly = False
            DCodeNo.ReadOnly = False
        End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
 
        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--���ʦs���
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--�w�ʦs���

            End If
        End If


    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue

        DNo.ReadOnly = True
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True
        DRno.ReadOnly = True
        DDevNo.ReadOnly = True
        DCodeNo.ReadOnly = True

        DProgress.Enabled = True    ''�w�]���Ұ�
        DProgress.SelectedIndex = 0 ''�w�]��ALL
        DDivision.Enabled = True    ''�w�]���Ұ�
        DSts.SelectedIndex = 0      ''�w�]��ALL

        DataGrid1.DataBind()    ''�M��DataGrid

        'No
        DNo.Text = "�e�U��No."
        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"
        '�e�UNo
        DRno.Text = "�e�UNo"
        '�}�oNo
        DDevNo.Text = "�}�oNo"
        'Code-No
        DCodeNo.Text = "Code-No"

        DNo.ReadOnly = False
        DPerson.ReadOnly = False
        If pFormNo = "002001" Then
            DBuyer.ReadOnly = False
            DRno.ReadOnly = False
            DDevNo.ReadOnly = False
            DCodeNo.ReadOnly = False
        End If
        If pFormNo = "002002" Then
            DBuyer.ReadOnly = False
            DDevNo.ReadOnly = False
            DCodeNo.ReadOnly = False
        End If
        If pFormNo = "002003" Then
            DDivision.Enabled = False
            DPerson.ReadOnly = True
            DBuyer.ReadOnly = False
            DDevNo.ReadOnly = False
            DCodeNo.ReadOnly = False
        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue

        '
        If DKeepData.SelectedValue = "0" Then
            DataList()      '--���ʦs���
        Else
            If DKeepData.SelectedValue = "1" Then
                KeepDataList()  '--�w�ʦs���

            End If
        End If
    End Sub


    '*****************************************************************
    '**
    '**     ��Excel�@�ε{��
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        pFormNo = DFormName.SelectedValue
        DataGrid1.AllowPaging = False   '�{���O���P
        DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=SCD_InqCommission.xls")     '�{���O���P
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        DataGrid1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        DataGrid1.AllowPaging = wAllowPaging        '�{���O���P
    End Sub


End Class


