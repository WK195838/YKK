Imports System.Data
Imports System.Data.OleDb

Public Class AgentList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DUserName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAllForm As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAgentName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton

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
        Response.Cookies("PGM").Value = "AgentList.aspx"

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
        SQL = SQL + "  And FormNo <= '900000' "
        SQL = SQL + "  And (InqAuthority = '0' "
        SQL = SQL + "       Or (InqAuthority = '1' "
        SQL = SQL + "           And (   InqUser     like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
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

        '����
        SQL = "SELECT UserName, UserID FROM M_Agent "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "Group by UserName, UserID "
        SQL = SQL + "Order by UserName, UserID "

        OleDbConnection1.Open()
        DUserName.Items.Clear()
        DUserName.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Agent")
        DBTable1 = DBDataSet1.Tables("M_Agent")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("UserName")
            ListItem1.Value = DBTable1.Rows(i).Item("UserID")
            DUserName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '�N�z�H
        SQL = "SELECT AgentName, AgentID FROM M_Agent "
        SQL = SQL + " Where Active = '1' "
        SQL = SQL + "Group by AgentName, AgentID "
        SQL = SQL + "Order by AgentName, AgentID "

        OleDbConnection1.Open()
        DAgentName.Items.Clear()
        DAgentName.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Agent1")
        DBTable1 = DBDataSet1.Tables("M_Agent1")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("AgentName")
            ListItem1.Value = DBTable1.Rows(i).Item("AgentID")
            DAgentName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "FormName, FormNo, Description, UserName, AgentName, "
        SQL = SQL + "Case AllForm When 0 Then '���e�U��' Else '��@�e�U��' End As AllFormDesc, "
        SQL = SQL + "Convert(VarChar, StartDate, 20) + '~' + Convert(VarChar, EndDate, 20) As RangeDesc "
        SQL = SQL + "FROM M_Agent "
        SQL = SQL + "Where Active = '1' "
        '�N�z����
        If DAllForm.SelectedValue <> "ALL" Then
            SQL = SQL + " And   AllForm = '" + DAllForm.SelectedValue + "'"
        End If
        '���
        If DAllForm.SelectedValue <> "0" And DAllForm.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '����
        If DUserName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   UserID = '" + DUserName.SelectedValue + "'"
        End If
        '�N�z�H
        If DAgentName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   AgentID = '" + DAgentName.SelectedValue + "'"
        End If

        'Sort
        SQL = SQL + " Order by StartDate, EndDate "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Agent")
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=Agent_List.xls")     '�{���O���P
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
