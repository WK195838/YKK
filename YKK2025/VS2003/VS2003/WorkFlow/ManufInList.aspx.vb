Imports System.Data
Imports System.Data.OleDb

Public Class OPContractList
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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wFun As String              'Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wFun = Request.QueryString("pFun") 'Function
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "Case Sts When 0 then '����' When 1 then '�w��OK' When 2 then '�w��NG' else '���' end as StsDesc, "
        SQL = SQL + "No, Convert(Varchar, Date,111) as DateDesc, "
        If wFun = "SF" Then         ''�P�_�O�_�����B�z
            SQL = SQL + "'' as SliderCode, '' as MapNo, "
        Else
            SQL = SQL + "SliderCode, MapNo, "
        End If
        SQL = SQL + "OFormNo+'-'+str(OFormSno,Len(OFormSno)) as OFormDesc, "

        If wFun = "OP" Then
            SQL = SQL + "Target as NFormDesc, "
            'SQL = SQL + "'' as NFormDesc, "
            SQL = SQL + "'AppendSpecSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL,Target "
            SQL = SQL + "FROM F_AppendSpecSheet "
        End If
        If wFun = "CT" Then
            SQL = SQL + "NFormNo+'-'+str(NFormSno,Len(NFormSno)) as NFormDesc, "
            SQL = SQL + "'ManufInCTSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL "
            SQL = SQL + "FROM F_ManufInCTSheet "
        End If
        If wFun = "SD" Then
            SQL = SQL + "NFormNo+'-'+str(NFormSno,Len(NFormSno)) as NFormDesc, "
            SQL = SQL + "'ManufInSDSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL "
            SQL = SQL + "FROM F_ManufInSDSheet "
        End If
        If wFun = "SF" Then
            SQL = SQL + "'' as NFormDesc, "
            SQL = SQL + "'SufaceSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL "
            SQL = SQL + "FROM F_SufaceSheet "
        End If
        If wFun = "CA" Then
            SQL = SQL + "ColorItem as NFormDesc, "
            'SQL = SQL + "'' as NFormDesc, "
            SQL = SQL + "'ColorAppendSheet_02.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) "
            SQL = SQL + " As URL "
            SQL = SQL + "FROM F_ColorAppendSheet "
        End If

        SQL = SQL + "Where OFormNo = '" & wFormNo & "'"
        SQL = SQL + "  and OFormSno = '" & CStr(wFormSno) & "'"
        'SQL = SQL + "  and (Sts = '2' or Sts = '0') "
        SQL = SQL + "Order by Date Desc "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Sheet")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

End Class
