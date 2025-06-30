Imports System.Data
Imports System.Data.OleDb

Public Class HR_DelayList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DDays As System.Web.UI.WebControls.DropDownList

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
        Response.Cookies("PGM").Value = "HR_DelayList.aspx"

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
        SQL = SQL + "  And FormNo >= '001001' And FormNo <= '001050' "
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

        '�ӽг���
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

        pFormNo = DFormName.SelectedValue

        Search_Item_Attribute()

    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        DataGrid1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        DataGrid1.Columns.Item(0).HeaderText = "No."
        DataGrid1.Columns.Item(1).HeaderText = "���A"
        DataGrid1.Columns.Item(2).HeaderText = "�ӽЪ�"
        DataGrid1.Columns.Item(3).HeaderText = "�ӽФ�"
        DataGrid1.Columns.Item(4).HeaderText = "�֩w��"
        DataGrid1.Columns.Item(5).HeaderText = "�w�w������"

        SQL = "SELECT "
        SQL = SQL + "Case No When '' Then '���s��' Else No End As Field1, "
        SQL = SQL + "Case V_WaitHandle_01.Delay When '1' Then '����' Else '���`' End As Field2, "
        SQL = SQL + "Division + '--' + ApplyName as Field3, "
        SQL = SQL + "ApplyTime as Field4, "
        SQL = SQL + "DivName + '--' + DecideName as Field5, "
        SQL = SQL + "BEndTime as Field6, "

        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + V_WaitHandle_01.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_01.FormSno,Len(V_WaitHandle_01.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(V_WaitHandle_01.Step,Len(V_WaitHandle_01.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_01.SeqNo,Len(V_WaitHandle_01.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + V_WaitHandle_01.ApplyID "
        SQL = SQL + " As OPURL "

        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where V_WaitHandle_01.Active = '1'    "
        SQL = SQL + "  And V_WaitHandle_01.FlowType = '1'  "
        SQL = SQL + "  And Delay = '1' "

        '���
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And FormNo = '" + DFormName.SelectedValue + "'"
        End If

        '����
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
        End If
        '�ӽЪ�
        If DName.Text <> "�ӽЪ�" And DName.Text <> "" Then
            SQL = SQL + " And ApplyName Like '%" + DName.Text + "%'"
        End If
        'No
        If DNo.Text <> "No." And DNo.Text <> "" Then
            SQL = SQL + " And No Like '%" + DNo.Text + "%'"
        End If
        '�i�ת��A
        If DDays.SelectedValue <> "ALL" Then
            If DDays.SelectedValue = "3" Then SQL = SQL + "   And Datediff(n, BEndTime, getdate()) <= '4320' "
            If DDays.SelectedValue = "6" Then SQL = SQL + "   And Datediff(n, BEndTime, getdate()) <= '8640' "
            If DDays.SelectedValue = "10" Then SQL = SQL + "  And Datediff(n, BEndTime, getdate()) <= '14400' "
        End If
        'Sort
        SQL = SQL + " Order by BEndTime "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()

    End Sub

    Sub Search_Item_Attribute()
        DNo.Text = "No."
        DName.Text = "�ӽЪ�"

        DNo.ReadOnly = False
        DName.ReadOnly = False
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub


    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        'No
        DNo.Text = "No."
        '�ӽЪ�
        DName.Text = "�ӽЪ�"

        DNo.ReadOnly = False
        DName.ReadOnly = False
        DDivision.Enabled = True    ''�w�]���Ұ�
        pFormNo = DFormName.SelectedValue
        DataGrid1.DataBind()    ''�M��DataGrid

    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        pFormNo = DFormName.SelectedValue
        DataList()
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

        Response.AppendHeader("Content-Disposition", "attachment;filename=OverTimeSheet.xls")     '�{���O���P
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
