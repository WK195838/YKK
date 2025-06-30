Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class OPInqWaitHandle
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DDecide As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "OPInqWaitHandle.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        If Not Me.IsPostBack Then
            pFormNo = "000001"
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
        NowDateTime = CStr(DateTime.Now.Year) + _
                      CStr(DateTime.Now.Month) + _
                      CStr(DateTime.Now.Day) + _
                      CStr(DateTime.Now.Hour) + _
                      CStr(DateTime.Now.Minute) + _
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
            If ListItem1.Value = pFormNo Then ListItem1.Selected = True
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        'Step
        DBDataSet1.Clear()
        SQL = "Select Step, str(Step) + ':' + StepName As StepName From M_Flow "
        SQL = SQL + "Where Step     <  '900' "
        SQL = SQL + "  and FormNo   =  '" + pFormNo + "' "
        SQL = SQL + "  and FlowType <> '0'   "
        SQL = SQL + "Group by Step, StepName "
        SQL = SQL + "Order by Step, StepName "
        DStep.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        DBTable1 = DBDataSet1.Tables("M_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DStep.Items.Add(ListItem1)
        Next

        '�̿ೡ��
        DBDataSet1.Clear()
        SQL = "Select DivName From M_Users Group by DivName Order by DivName "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Users")
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

    Sub Search_Item_Attribute()
        DPerson.ReadOnly = True
        DDecide.ReadOnly = True
        DBuyer.ReadOnly = True

        '�̿��
        DPerson.Text = "�̿��"
        '�֩w��
        DDecide.Text = "�֩w��"
        'Buyer
        DBuyer.Text = "Buyer"
        DPerson.ReadOnly = False
        DDecide.ReadOnly = False
        DBuyer.ReadOnly = False

        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or + _
           pFormNo = "000007" Or pFormNo = "000008" Or pFormNo = "000009" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000010" Then
        End If
        If pFormNo = "000011" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000012" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000013" Then
        End If
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim wTableName As String = ""
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
        SQL = SQL + " Where Active  =  '1' "
        SQL = SQL + "   And FormNo = '" + DFormName.SelectedValue + "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Form")
        If DBDataSet1.Tables("Form").Rows.Count > 0 Then
            wTableName = "V_" + DBDataSet1.Tables("Form").Rows(0).Item("TableName1") + "_01"
        End If

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, T_Step, T_SeqNo, T_DecideName, Buyer, Division,  "
        SQL = SQL + "Case  When No='' or No IS NULL Then '���s��' Else No End As No, "
        SQL = SQL + "T_StsDesc, T_FormName, T_FlowTypeDesc, T_ApplyName, T_StepNameDesc, "
        SQL = SQL + "'�ӽЮɶ��G[' + Convert(VarChar, T_ApplyTime, 20) + '], ' + "
        SQL = SQL + "'����ɶ��G[' + Convert(VarChar, T_ReceiptTime, 20) + '], ' + "
        SQL = SQL + "'�\Ū�����G[' + Convert(VarChar, T_ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'�����\Ū�G[' + T_FirstReadTimeDesc + '], ' + "
        SQL = SQL + "'�̫�\Ū�G[' + T_LastReadTimeDesc  + '], ' + "
        SQL = SQL + "'�w�w�}�l�G[' + Convert(VarChar, T_BStartTime, 20) + '], ' + "
        SQL = SQL + "'�w�w�����G[' + Convert(VarChar, T_BEndTime, 20) + '], ' + "
        SQL = SQL + "'��ڶ}�l�G[' + Convert(VarChar, T_AStartTime, 20) + '] '  "
        SQL = SQL + " As Description, T_ViewURL "
        SQL = SQL + "FROM " + wTableName + " "
        SQL = SQL + "Where T_Active = '1' "
        SQL = SQL + "  And T_FlowType <> '0' "
        SQL = SQL + "  And Action = '0' "

        '���
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   FormNo = '" + DFormName.SelectedValue + "'"
        End If
        'Step
        SQL = SQL + " And   T_Step = '" + DStep.SelectedValue + "'"
        '�̿ೡ��
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
        End If
        '�̿��
        If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
            SQL = SQL + " And Person Like '%" + DPerson.Text + "%'"
        End If
        '�֩w��
        If DDecide.Text <> "�֩w��" And DDecide.Text <> "" Then
            SQL = SQL + " And T_DecideName Like '%" + DDecide.Text + "%'"
        End If
        'Buyer
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
        End If
        SQL = SQL + " Order by T_ApplyTime "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        SetSearchItem()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        pFormNo = DFormName.SelectedValue
        DataList()
    End Sub

End Class
