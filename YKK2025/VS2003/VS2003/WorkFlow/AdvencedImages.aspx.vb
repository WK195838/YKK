Imports System.Data
Imports System.Data.OleDb

Public Class AdvencedImages
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BFind As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKStart As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox11 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKEnd As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataList1 As System.Web.UI.WebControls.DataList

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
        If Not Me.IsPostBack Then
        End If
    End Sub

    Private Sub BFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BFind.Click

        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        'MsgBox(DKStart.Text & "~" & DKEnd.Text)

        SQL = "select TOP 10 "
        SQL = SQL & "case when formno='000003' then [SliderGRCode] + '  ���s' "
        SQL = SQL & "     else [SliderGRCode] + '  �~�`' "
        SQL = SQL & "end as Code, "
        SQL = SQL & "No As No, '' as NoUrl, "
        SQL = SQL & "ImagePath As ImagePath "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where formno in ('000003','000007') "
        ' Start
        If DKStart.Text <> "" Then
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKStart.Text) & ") >= '" & DKStart.Text & "' "
        End If
        ' End
        If DKEnd.Text <> "" Then
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKEnd.Text) & ") >= '" & DKEnd.Text & "' "
        End If
        '
        SQL = SQL & "Order by substring([SliderGRCode],1," & Len(DKEnd.Text) & ") "
        '
        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "IMAGE")
        '
        DataList1.DataSource = DBDataSet1.Tables("IMAGE")
        DataList1.DataBind()

        OleDbConnection1.Close()
    End Sub

End Class
