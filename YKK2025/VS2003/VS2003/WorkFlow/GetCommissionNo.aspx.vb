Imports System.Data
Imports System.Data.OleDb

Public Class GetCommissionNo
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DMsg As System.Web.UI.WebControls.TextBox
    Protected WithEvents Go As System.Web.UI.WebControls.Button
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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "GetCommissionNo.aspx"

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
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �z�ﶵ��
    '**
    '*****************************************************************
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
        SQL = SQL + "  And ( (FormNo >= '000001' And FormNo <= '001000') or FormNo='800001' ) "
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

        '�̿ೡ��
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("DivName")
            ListItem1.Value = DBTable1.Rows(i).Item("DivName")
            DDivision.Items.Add(ListItem1)
        Next

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataList()
    End Sub

    Sub DataList()
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
            wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        End If
        OleDbConnection1.Close()

        SQL = "SELECT Top 20 "
        SQL = SQL + "V_CommissionNo_01.No As No, "
        SQL = SQL + "V_CommissionNo_01.UserName As CreateUser, "
        SQL = SQL + "V_CommissionNo_01.FormSno As FormSno, "
        SQL = SQL + "Convert(VarChar, V_CommissionNo_01.CreateTime, 20) As CreateTime, "
        SQL = SQL + "Case " + wTableName + ".Sts When '0' Then '�}�o��' When '1' Then '�}�o����(OK)' When '2' Then '�}�o����(NG)' Else '�}�o����' End As Sts "
        SQL = SQL + "FROM " + wTableName + " "
        SQL = SQL + "Left Outer Join V_CommissionNo_01 ON " + wTableName + ".FormNo =V_CommissionNo_01.FormNo  "
        SQL = SQL + "                                 And " + wTableName + ".FormSno=V_CommissionNo_01.FormSno "

        SQL = SQL + "Where V_CommissionNo_01.No <> 'zzzzzzzzzzzzzzzzzzzz' "
        '���
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   V_CommissionNo_01.FormNo = '" + DFormName.SelectedValue + "'"
        End If
        '����
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And   V_CommissionNo_01.Division = '" + DDivision.SelectedValue + "'"
        End If
        'Sort
        SQL = SQL + " Order by V_CommissionNo_01.CreateTime Desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "CommissionNo")
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
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        DataGrid1.AllowPaging = False   '�{���O���P
        DataList()                      '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=Last_Commission_No.xls")     '�{���O���P
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
