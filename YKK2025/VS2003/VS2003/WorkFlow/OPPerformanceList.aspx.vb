Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPPerformanceList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
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
    Dim pFormSno As Integer
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim pTable As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPPerformanceList.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            DataList()
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
        pFormSno = Request.QueryString("pFormSno")    '��渹�X
        pTable = Request.QueryString("pTable")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()

        SQL = "SELECT No, FormNo, FormSno, Step, SStep, EStep, Division, Buyer, Finished, Suppiler, Level, BTime, ATime, TimeRange, "
        SQL = SQL + "Case Sts When 0 Then '�}�o��' When 1 Then 'OK' When 2 Then 'NG' Else '����' End As Sts, "
        SQL = SQL + "BTime-ATime As Balance, "

        SQL = SQL + "'OPPerformanceWorkID.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pStep=' + str(Step,Len(Step)) + "
        SQL = SQL + "'&pSeqNo=' + str(SeqNo,Len(SeqNo)) + "
        SQL = SQL + "'&pWorkID=' + WorkID "
        SQL = SQL + " As URL "

        SQL = SQL + "FROM " + pTable + " "
        SQL = SQL + "Where RecordType = '0' "
        SQL = SQL + "  and FormNo = '" & pFormNo & "' "
        SQL = SQL + "  and FormSno = '" & CStr(pFormSno) & "' "
        SQL = SQL + "Order by Step "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Header")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()

    End Sub


End Class
