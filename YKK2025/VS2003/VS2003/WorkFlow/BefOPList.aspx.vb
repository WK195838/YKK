Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class BefOPList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BRead As System.Web.UI.HtmlControls.HtmlInputButton
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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As String          '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�
    Dim wKeepData As String         '�ʦs

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            ShowReadButton()
            SetOPDataList()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�
        wKeepData = Request.QueryString("pKeepData")  '�ʦs
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�\Ū���s
    '**
    '*****************************************************************
    Sub ShowReadButton()
        If wApplyID = "" Then
            BRead.Visible = True
        Else
            BRead.Visible = False
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�u�{���
    '**
    '*****************************************************************
    Sub SetOPDataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "'�w�w�}�l�G[' + BStartTimeDesc + '], ' + "
        SQL = SQL + "'�w�w�����G[' + BEndTimeDesc + '], ' + "
        SQL = SQL + "'��ڶ}�l�G[' + AStartTimeDesc + '], ' + "
        SQL = SQL + "'��ڧ����G[' + AEndTimeDesc + '] ' As Description, "
        SQL = SQL + "URL, "
        'POP-ADD-Start
        If wFormNo = "000003" Then
            SQL = SQL + "Case Step When 80 Then 'POP' Else '' End As POP, "
            SQL = SQL + "Case Step When 80 Then 'http://10.245.1.10/WorkFlowSub/POPList.aspx?pSys=SPD&pNo=' + No + '&pFormNo=000003&pFormSno=' + LTrim(str(FormSno)) + '&pStep=' + LTrim(str(Step)) Else '' End As POPURL "

        Else
            SQL = SQL + "'' As POP, "
            SQL = SQL + "'' As POPURL "
        End If
        'POP-ADD-End
        If wKeepData = "1" Then
            SQL = SQL + "FROM V_WaitHandle_OLD_01 "
        Else
            SQL = SQL + "FROM V_WaitHandle_01 "
        End If
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & wFormSno & "' "
        'SQL = SQL + "Order by Unique_ID Desc "
        SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub


    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid���W�U��
        SetOPDataList()

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �\Ū����
    '**
    '*****************************************************************
    Private Sub BRead_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BRead.ServerClick
        Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DOPReady", "�w�\Ū"))

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
        SetOPDataList()                 '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=OP_Detail_List.xls")     '�{���O���P
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
