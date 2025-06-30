Imports System.Data
Imports System.Data.OleDb

Public Class MapManagement
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents Datagrid2 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BExcel1 As System.Web.UI.WebControls.ImageButton

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

    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim wMapNo As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "MapManagement.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetKey()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetKey()
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '����
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

        'Buyer
        DBuyer.Text = "Buyer"
        '�ϭ�
        DMapNo.Text = "�ϸ�"
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapFilePath")
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "   '���oDB���
        SQL = SQL & "No, MapNo, FormSno, CompletedTime, Buyer, "
        SQL = SQL & "Case Sts When 0 Then '�}�o��' When 1 Then '�}�oOK' When 2 Then '�}�oNG' ELSE '����' END AS StsDesc, "
        SQL = SQL & "'" & Path & "' + MapFile As URL "
        SQL = SQL & "FROM F_MapSheet "
        SQL = SQL + "Where Sts <> '0' And Sts <> '3' "
        '����
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And   Division = '" + DDivision.SelectedValue + "'"
        End If
        'Buyer
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And   Buyer Like '%" + DBuyer.Text + "%'"
        End If
        '�ϭ�
        If DMapNo.Text <> "�ϸ�" And DMapNo.Text <> "" Then
            SQL = SQL + "And MapNo Like '%" & DMapNo.Text & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
        Datagrid2.DataSource = DBDataSet1
        Datagrid2.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        Datagrid2.CurrentPageIndex = 0
        DataList()

        DataGrid1.DataSource = Nothing
        DataGrid1.DataBind()
    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")

        If e.CommandSource.CommandName = "ShowInformation" Then  '�I�����ˬd
            Dim Key As String = Datagrid2.DataKeys(e.Item.ItemIndex)  '�ҿ����Map No
            wMapNo = Datagrid2.DataKeys(e.Item.ItemIndex)
            Dim SQL As String

            DataGrid1.CurrentPageIndex = 0

            'DB�s���]�w
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            SQL = "SELECT "   '���oDB���
            SQL = SQL & "Case Sts When 0 then '�}�o��' When 1 then '�}�o����-OK' When 2 then '�}�o����-NG' Else '�}�o����' End As StsDesc, "
            SQL = SQL & "No, MapNo, FormSno, CompletedTime, ModReasonCode, ModContent, Buyer, "
            SQL = SQL & "'" & Path & "' + MapFile As URL "
            SQL = SQL & "FROM F_MapModSheet "
            'SQL = SQL + "Where Sts = '1' "
            'SQL = SQL + "And OriMapNo Like '%" & Key & "%' "
            SQL = SQL + "Where OriMapNo Like '%" & Key & "%' "
            SQL = SQL + "Order by CompletedTime Desc "

            Dim DBDataSet2 As New DataSet
            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "F_MapModSheet")
            DataGrid1.DataSource = DBDataSet2
            DataGrid1.DataBind()

            If DBDataSet2.Tables("F_MapModSheet").Rows.Count > 0 Then
                BExcel1.Visible = True
            Else
                BExcel1.Visible = False
            End If

            'DB�s������
            OleDbConnection1.Close()
        End If
    End Sub

    Private Sub Datagrid2_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles Datagrid2.PageIndexChanged
        Datagrid2.CurrentPageIndex = e.NewPageIndex   'DataGrid���W�U��
        DataList()

        DataGrid1.DataSource = Nothing
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        Dim SQL As String
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid���W�U��

        SQL = "SELECT "   '���oDB���
        SQL = SQL & "No, MapNo, FormSno, CompletedTime, ModReason, ModReasonDesc, Buyer, "
        SQL = SQL & "'" & Path & "' + MapFile As URL "
        SQL = SQL & "FROM F_MapModSheet "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "And MapNo Like '%" & wMapNo & "%' "
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet2 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "F_MapModSheet")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()

        OleDbConnection1.Close()
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
        Dim wAllowPaging As Boolean = Datagrid2.AllowPaging   '�{���O���P

        'Datagrid2.AllowPaging = False   '�{���O���P
        'DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=Map_Management_List.xls")     '�{���O���P
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        Datagrid2.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        Datagrid2.AllowPaging = wAllowPaging        '�{���O���P
    End Sub

    Private Sub BExcel1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click, BExcel1.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        'DataGrid1.AllowPaging = False   '�{���O���P
        'DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=Map_Management_Mod_List.xls")     '�{���O���P
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
