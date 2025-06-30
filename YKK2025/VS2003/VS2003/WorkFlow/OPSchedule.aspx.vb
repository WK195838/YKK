Imports System.Data
Imports System.Data.OleDb

Public Class OPScheduleList
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList

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
    Dim pWorkID As String
    Dim pFormNo As String
    Dim pType As String
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "OPSchedule.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            SetSearchItem()
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
        pWorkID = Request.QueryString("pWorkID")
        pFormNo = Request.QueryString("pFormNo")    '��渹�X
        pType = Request.QueryString("pType")
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        For i = 0 To DFlowType.Items.Count - 1
            If DFlowType.Items.Item(i).Value = pType Then
                DFlowType.Items.Item(i).Selected = True
            End If
        Next
    End Sub

    Sub DataList()
        Dim wTableName As String = ""
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        'OleDbConnection1.Open()
        'SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        'SQL = SQL + " Where Active  =  '1' "
        'SQL = SQL + "   And FormNo = '" + pFormNo + "'"
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "Flow")
        'If DBDataSet1.Tables("Flow").Rows.Count > 0 Then
        'wTableName = "F_" + DBDataSet1.Tables("Flow").Rows(0).Item("TableName1")
        'End If
        'OleDbConnection1.Close()

        'If wTableName <> "" Then
        SQL = "SELECT "
        SQL = SQL + "FormName + '(' + str(FormSno, Len(FormSno)) +')' as No, "
        SQL = SQL + "ApplyName as Person, "
        SQL = SQL + "StepName as StepName, "
        SQL = SQL + "DecideName as StepPerson, "
        SQL = SQL + "Convert(VarChar, BStartTime, 20) + '~' + Convert(VarChar, BEndTime, 20) As BOPTime, "
        SQL = SQL + "Convert(VarChar, AStartTime, 20) as AStartTime, "
        SQL = SQL + "Convert(VarChar, ApplyTime, 20)  as ApplyTime, "
        SQL = SQL + "'����ɶ��G[' + Convert(VarChar, ReceiptTime, 20) + '], ' +  "
        SQL = SQL + "'�\Ū�����G[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'�����\Ū�G[' + Convert(VarChar, FirstReadTime, 20) + '], ' + "
        SQL = SQL + "'�̫�\Ū�G[' + Convert(VarChar, LastReadTime, 20) + ']  ' as ReadTime "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And WorkID = '" + pWorkID + "' "
        'FlowType
        If DFlowType.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FlowType = '" + DFlowType.SelectedValue + "'"
        End If
        'Sort
        SQL = SQL + " Order by BStartTime "

        'SQL = "SELECT "
        'SQL = SQL + wTableName + ".No as No, "
        'SQL = SQL + wTableName + ".Person as Person, "
        'SQL = SQL + "V_WaitHandle_01.StepName as StepName, "
        'SQL = SQL + "V_WaitHandle_01.DecideName as StepPerson, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.BStartTime, 20) + '~' + Convert(VarChar, V_WaitHandle_01.BEndTime, 20) As BOPTime, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.AStartTime, 20) as AStartTime, "
        'SQL = SQL + "Convert(VarChar, V_WaitHandle_01.ApplyTime, 20)  as ApplyTime, "
        'SQL = SQL + "'����ɶ��G[' + Convert(VarChar, V_WaitHandle_01.ReceiptTime, 20) + '], ' +  "
        'SQL = SQL + "'�\Ū�����G[' + Convert(VarChar, V_WaitHandle_01.ReadTimeLimit, 20) + '], ' + "
        'SQL = SQL + "'�����\Ū�G[' + Convert(VarChar, V_WaitHandle_01.FirstReadTime, 20) + '], ' + "
        'SQL = SQL + "'�̫�\Ū�G[' + Convert(VarChar, V_WaitHandle_01.LastReadTime, 20) + ']  ' as ReadTime "
        'SQL = SQL + "FROM " + wTableName + " "
        'SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
        'SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "

        'SQL = SQL + "Where " + wTableName + ".Sts = '0' "
        'SQL = SQL + "  And V_WaitHandle_01.Active = '1' "
        'SQL = SQL + "  And V_WaitHandle_01.WorkID = '" + pWorkID + "' "
        'FlowType
        'If DFlowType.SelectedValue <> "ALL" Then
        'SQL = SQL + " And   V_WaitHandle_01.FlowType = '" + DFlowType.SelectedValue + "'"
        'End If
        'Sort
        'SQL = SQL + " Order by BStartTime "

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()

        'End If
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataList()
    End Sub

End Class
