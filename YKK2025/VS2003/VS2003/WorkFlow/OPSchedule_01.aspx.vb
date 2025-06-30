Imports System.Data
Imports System.Data.OleDb

Public Class OPSchedule_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList
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
    Dim pUserID As String
    Dim pWorkID As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPSchedule_01.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            DataList()
        End If
    End Sub

    Sub SetParameter()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet

        pFormNo = Request.QueryString("pFormNo")
        pUserID = Request.QueryString("pUserID")

        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        If pWorkID = "" Then
            SQL = "Select * From M_Users "
            SQL = SQL & " Where UserID =  '" & pUserID & "'"
            SQL = SQL & "   And Active = '1' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                pWorkID = DBDataSet1.Tables("M_Users").Rows(0).Item("WorkID")
            Else
                pWorkID = ""
            End If
        End If

        'DB�s������
        OleDbConnection1.Close()
    End Sub

    Sub DataList()
        If pWorkID <> "" Then
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, ApplyName, DecideName, StepNameDesc, "
            SQL = SQL + "'�y�{��T' As WorkFlow, "
            SQL = SQL + "Convert(VarChar, BStartTime, 20) + '~' + Convert(VarChar, BEndTime, 20) As Description, "
            SQL = SQL + "ViewURL, "

            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + FormNo + "
            SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
            SQL = SQL + "'&pStep='    + str(Step,Len(Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + ApplyID "
            SQL = SQL + " As OPURL "

            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where Active = '1' "
            SQL = SQL + " And   WorkID = '" + pWorkID + "'"
            'FlowType
            If DFlowType.SelectedValue <> "ALL" Then
                SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
            End If
            SQL = SQL + " Order by BStartTime "

            OleDbConnection1.Open()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "WaitHandle")
            DataGrid1.DataSource = DBDataSet1
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub


End Class
