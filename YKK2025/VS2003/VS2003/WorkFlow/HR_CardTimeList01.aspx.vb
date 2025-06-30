Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_CardTimeList01
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
    Dim wMonth As String        '�u�@��
    Dim wEmpID As String        'Emp-ID
    Dim wDepoID As String           '���q�O

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            CardData()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        wMonth = Request.QueryString("pMonth")    '�u�@��
        wEmpID = Request.QueryString("pEmpID")    'Emp-ID
        wDepoID = Request.QueryString("pDepoID")        '���q�O
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �z���d���
    '**
    '*****************************************************************
    Sub CardData()
        Dim wStartDate As String = wMonth + "/1"
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()

        SQL = "SELECT "
        SQL = SQL + "Convert(VARCHAR(10), CDate, 111) as CDateDesc, "
        SQL = SQL + "EmpID, TimeA, TimeB "
        SQL = SQL + "FROM HR_WorkTime "
        SQL = SQL + "Where EmpID  = '" & wEmpID & "' "
        SQL = SQL + "  And DepoID = '" & wDepoID & "' "
        SQL = SQL + "  And CDate >= '" & wStartDate & "' "
        SQL = SQL + "Order by CDate Desc "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WorkTime")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

End Class
