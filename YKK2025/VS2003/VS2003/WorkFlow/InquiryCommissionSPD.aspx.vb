Imports System.Data
Imports System.Data.OleDb

Public Class InquiryCommissionSPD
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DSort As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSortKey As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "InquiryCommissionSPD.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo <= '100000' "
        'SQL = SQL + "  And (Authority = '0' "
        'SQL = SQL + "       Or (Authority = '1' "
        'SQL = SQL + "           And (Users1 like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        'SQL = SQL + "                Or Users2 like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        'SQL = SQL + "               ) "
        'SQL = SQL + "          ) "
        'SQL = SQL + "      ) "
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

        '����
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp Where Cat='100' and DKey='DIVISION' Order by Data "
        DDivision.Items.Clear()
        DDivision.Items.Add("ALL")
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DDivision.Items.Add(ListItem1)
        Next

        'Buyer
        DBuyer.Text = ""

        '���
        DSDate.Text = CStr(DateAdd("d", -30, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
    End Sub

    Sub DataList()
        Dim i As Integer = 0
        Dim TableName As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT TableName1 FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo = '" & DFormName.SelectedValue & "' "
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet2, "M_FORM")
        If DBDataSet2.Tables("M_FORM").Rows.Count > 0 Then
            TableName = "V_" & DBDataSet2.Tables("M_FORM").Rows(0).Item("TableName1") & "_01"
        End If
        OleDbConnection1.Close()

        If TableName <> "" Then
            SQL = "SELECT "
            SQL = SQL + "T_FormNo, T_FormSno, T_Step, T_SeqNo, "
            SQL = SQL + "T_StsDesc, T_FormName, T_FlowTypeDesc, T_ApplyName, T_DecideName, T_StepNameDesc, T_ApplyTime, "
            SQL = SQL + "'�y�{��T' As T_WorkFlow, "
            SQL = SQL + "'�ӽЮɶ��G[' + Convert(VarChar, T_ApplyTime, 20) + '], ' + "
            SQL = SQL + "'����ɶ��G[' + Convert(VarChar, T_ReceiptTime, 20) + '], ' + "
            SQL = SQL + "'�\Ū�����G[' + Convert(VarChar, T_ReadTimeLimit, 20) + '], ' + "
            SQL = SQL + "'�����\Ū�G[' + T_FirstReadTimeDesc + '], ' + "
            SQL = SQL + "'�̫�\Ū�G[' + T_LastReadTimeDesc  + '], ' + "
            SQL = SQL + "'�w�w�}�l�G[' + Convert(VarChar, T_BStartTime, 20) + '], ' + "
            SQL = SQL + "'�w�w�����G[' + Convert(VarChar, T_BEndTime, 20) + '], ' + "
            SQL = SQL + "'��ڶ}�l�G[' + Convert(VarChar, T_AStartTime, 20) + '], ' + "
            SQL = SQL + "'��ڧ����G[' + T_AEndTimeDesc + '] '  "
            SQL = SQL + " As Description, T_ViewURL, "

            SQL = SQL + "'BefOPList.aspx?' + "
            SQL = SQL + "'pFormNo='   + T_FormNo + "
            SQL = SQL + "'&pFormSno=' + str(T_FormSno,Len(T_FormSno)) + "
            SQL = SQL + "'&pStep='    + str(T_Step,Len(T_Step)) + "
            SQL = SQL + "'&pSeqNo='   + str(T_SeqNo,Len(T_SeqNo)) + "
            SQL = SQL + "'&pApplyID=' + T_ApplyID "
            SQL = SQL + " As T_OPURL "

            SQL = SQL + "FROM " & TableName & " "
            SQL = SQL + "Where T_Active = '0' "
            SQL = SQL + "  And T_Step   = '1' "
            '���
            If DFormName.SelectedValue <> "ALL" Then
                SQL = SQL + " And   T_FormNo = '" + DFormName.SelectedValue + "'"
            End If
            '����
            If DDivision.SelectedValue <> "ALL" Then
                SQL = SQL + " And   Division = '" + DDivision.SelectedValue + "'"
            End If
            'Buyer
            If DBuyer.Text <> "ALL" And DBuyer.Text <> "" Then
                SQL = SQL + " And   Buyer Like '%" + DBuyer.Text + "%'"
            End If
            '���
            If DSDate.Text = "" Then DSDate.Text = CStr(DateAdd("d", -30, DateTime.Now.Today))
            If DEDate.Text = "" Then DEDate.Text = CStr(DateTime.Now.Today)
            SQL = SQL + " And   T_ApplyTime >= '" + DSDate.Text + " 00:00:00" + "'"
            SQL = SQL + " And   T_ApplyTime <= '" + DEDate.Text + " 23:59:59" + "'"
            'Sort
            SQL = SQL + " Order by " + DSortKey.SelectedValue + " " + DSort.SelectedValue

            OleDbConnection1.Open()
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "Sheet")
            DataGrid1.DataSource = DBDataSet1
            DataGrid1.DataBind()
            OleDbConnection1.Close()
        End If

    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

End Class
