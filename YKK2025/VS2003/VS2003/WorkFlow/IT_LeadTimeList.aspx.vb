Imports System.Data
Imports System.Data.OleDb

Public Class IT_LeadTimeList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_LeadTimeList.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '001101' And FormNo <= '001199' "
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
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '�u�{
        SQL = "SELECT Level FROM M_LeadTime "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        SQL = SQL + "Group by Level "
        SQL = SQL + "Order by Level "

        OleDbConnection1.Open()
        DLevel.Items.Clear()
        'DLevel.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_LeadTime")
        DBTable1 = DBDataSet1.Tables("M_LeadTime")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Level")
            ListItem1.Value = DBTable1.Rows(i).Item("Level")
            DLevel.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
        DLevel.Visible = False
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "FormName, FormNo, StepName, Step, Level, SUM(Hours) as Hours "
        'SQL = SQL + "StepName As StepNameDesc,  "
        'SQL = SQL + "Case SeqNo When 0 Then '�e�ǳ�' When 1 Then '�s�{' When 2 Then '��ǳ�' Else '' End As SeqNoDesc, "
        'SQL = SQL + "'�@ �� �̡G[' + CreateUser + '], ' + "
        'SQL = SQL + "'�@���ɶ��G[' + Convert(VarChar, CreateTime, 20) + '], ' + "
        'SQL = SQL + "'�� �� �̡G[' + ModifyUser + '], ' + "
        'SQL = SQL + "'�ק�ɶ��G[' + Convert(VarChar, ModifyTime, 20) + '], ' "
        'SQL = SQL + " As Reference "
        SQL = SQL + "FROM M_LeadTime "
        SQL = SQL + "Where Active = '1' "
        '���
        SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        '�u�{
        If DLevel.SelectedValue <> "ALL" Then
            SQL = SQL + " And   Level = '" + DLevel.SelectedValue + "'"
        End If
        'Sort
        SQL = SQL + " Group by FormNo, FormName, Step, StepName, Level "
        SQL = SQL + " Order by FormNo, FormName, Step, StepName, Level "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_LeadTime")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub


    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        '�u�{
        SQL = "SELECT Level FROM M_LeadTime "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        SQL = SQL + "Group by Level "
        SQL = SQL + "Order by Level "

        OleDbConnection1.Open()
        DLevel.Items.Clear()
        'DLevel.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_LeadTime")
        DBTable1 = DBDataSet1.Tables("M_LeadTime")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Level")
            ListItem1.Value = DBTable1.Rows(i).Item("Level")
            DLevel.Items.Add(ListItem1)
        Next
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
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        DataGrid1.AllowPaging = False   '�{���O���P
        DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=LeadTime_List.xls")     '�{���O���P
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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
