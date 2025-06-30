Imports System.Data
Imports System.Data.OleDb

Public Class TrackCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSort As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents DFlowType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DSortKey As System.Web.UI.WebControls.DropDownList

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
    Dim oCommon As New Common.CommonService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '���c��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wDepo As String = "CL1"      '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
    'Modify-End

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "TrackCommission.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            Track_DataList()
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
        SQL = "SELECT FormName, FormNo FROM V_WaitHandle_01 "
        SQL = SQL + " Where FormNo <= '900000' "
        SQL = SQL + "   And DecideHistory Like '%" + Request.QueryString("pUserID") + "%'"
        SQL = SQL + "   And Active = '1' "
        SQL = SQL + "   And FSts = '0' "
        SQL = SQL + " Group by FormName, FormNo "
        SQL = SQL + " Order by FormName, FormNo "

        OleDbConnection1.Open()
        DFormName.Items.Clear()
        DFormName.Items.Add("ALL")
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "WaitHandle")
        DBTable1 = DBDataSet1.Tables("WaitHandle")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("FormName")
            ListItem1.Value = DBTable1.Rows(i).Item("FormNo")
            DFormName.Items.Add(ListItem1)
        Next
        OleDbConnection1.Close()

        '���
        DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        DEDate.Text = CStr(DateTime.Now.Today)
        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"
    End Sub

    Sub Track_DataList()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        SQL = "SELECT "
        SQL = SQL + "Case  When No='' or No IS NULL Then '���s��' Else No End As No, "
        SQL = SQL + "ApplyName As Person, "
        SQL = SQL + "Unique_ID As ID, StepName, DecideName as StepPerson, FlowTypeDesc, "

        SQL = SQL + "'�}�l�w�w�G[' + Convert(VarChar, V_WaitHandle_01.BStartTime, 20) + '], ' + "
        SQL = SQL + "'�����w�w�G[' + Convert(VarChar, V_WaitHandle_01.BEndTime  , 20) + '], ' + "
        SQL = SQL + "'����ɶ��G[' + Convert(VarChar, V_WaitHandle_01.ReceiptTime, 20) + '], ' +  "
        SQL = SQL + "'�\Ū�����G[' + Convert(VarChar, V_WaitHandle_01.ReadTimeLimit, 20) + '], ' + "
        SQL = SQL + "'�����\Ū�G[' + Convert(VarChar, V_WaitHandle_01.FirstReadTimeDesc, 20) + '], ' + "
        SQL = SQL + "'�̫�\Ū�G[' + Convert(VarChar, V_WaitHandle_01.LastReadTimeDesc, 20) + ']  ' as OPTime, "
        SQL = SQL + "Convert(VarChar, V_WaitHandle_01.AStartTime, 20) as AStartTime, "
        SQL = SQL + "Convert(VarChar, V_WaitHandle_01.ApplyTime, 20)  as ApplyTime, "

        SQL = SQL + " '....' as WorkFlow, ViewURL, "
        SQL = SQL + "'BefOPList.aspx?' + "
        SQL = SQL + "'pFormNo='   + V_WaitHandle_01.FormNo + "
        SQL = SQL + "'&pFormSno=' + str(V_WaitHandle_01.FormSno,Len(V_WaitHandle_01.FormSno)) + "
        SQL = SQL + "'&pStep='    + str(V_WaitHandle_01.Step,Len(V_WaitHandle_01.Step)) + "
        SQL = SQL + "'&pSeqNo='   + str(V_WaitHandle_01.SeqNo,Len(V_WaitHandle_01.SeqNo)) + "
        SQL = SQL + "'&pApplyID=' + V_WaitHandle_01.ApplyID "
        SQL = SQL + " As OPURL "
        SQL = SQL + "FROM V_WaitHandle_01 "

        SQL = SQL + "Where FormNo <= '900000' "
        SQL = SQL + "  And DecideHistory Like '%" + Request.QueryString("pUserID") + "%'"
        SQL = SQL + "  And FSts = '0' "
        SQL = SQL + "  And Active = '1' "

        '���
        If DFormName.SelectedValue <> "ALL" Then
            SQL = SQL + " And   V_WaitHandle_01.FormNo = '" + DFormName.SelectedValue + "'"
        End If

        '���O
        If DFlowType.SelectedValue <> "ALL" Then
            SQL = SQL + " And   V_WaitHandle_01.FlowType = '" + DFlowType.SelectedValue + "'"
        End If

        '���
        If DSDate.Text = "" Then DSDate.Text = CStr(DateAdd("d", -180, DateTime.Now.Today))
        If DEDate.Text = "" Then DEDate.Text = CStr(DateTime.Now.Today)
        SQL = SQL + " And   V_WaitHandle_01.ApplyTime >= '" + DSDate.Text + " 00:00:00" + "'"
        SQL = SQL + " And   V_WaitHandle_01.ApplyTime <= '" + DEDate.Text + " 23:59:59" + "'"

        'Sort--���渹
        If DSortKey.SelectedValue <> "DFormSno" Then
            SQL = SQL + " Order by V_WaitHandle_01." + DSortKey.SelectedValue + " " + DSort.SelectedValue
        Else
            SQL = SQL + " Order by V_WaitHandle_01.FormNo " + DSort.SelectedValue + ", V_WaitHandle_01.FormSno " + DSort.SelectedValue
        End If

        OleDbConnection1.Open()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet2, "WaitHandle")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()
        OleDbConnection1.Close()

    End Sub


    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged

        DataGrid1.CurrentPageIndex = e.NewPageIndex
        Track_DataList()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        Track_DataList()
    End Sub


    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand

        If e.CommandSource.CommandName = "Select" Then  '�I�����ˬd
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '�ҿ����Map No

            Dim pFormNo As String = ""
            Dim pFormSno As Integer = 0
            Dim pStep As Integer = 0
            Dim pLevel As String = ""
            Dim pMakeMap As String = ""
            Dim pMakeMapName As String = ""
            Dim pUserID As String = ""

            Dim wFormNo As String = ""
            Dim wTableName As String = ""
            Dim URL As String = ""

            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim DBDataSet2 As New DataSet
            Dim DBDataSet3 As New DataSet
            Dim DBDataSet4 As New DataSet

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'OleDbConnection1.Open()
            'SQL = "SELECT FormNo, Depo FROM V_WaitHandle_01 "
            'SQL = SQL + " Where Unique_ID = '" + Key + "'"
            'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            'DBAdapter1.Fill(DBDataSet1, "WaitHandle")
            'If DBDataSet1.Tables("WaitHandle").Rows.Count > 0 Then
            'wFormNo = DBDataSet1.Tables("WaitHandle").Rows(0).Item("FormNo")
            '
            'SPD�e�U��
            'If wFormNo >= "000001" And wFormNo <= "001000" Then wDepo = "CL"
            'HRW�e�U��
            'If wFormNo >= "001001" And wFormNo <= "001050" Then wDepo = "TP"
            'HRWCL�e�U��
            'If wFormNo >= "001051" And wFormNo <= "001099" Then wDepo = "CL"
            'SCD�e�U��
            'If wFormNo >= "002001" And wFormNo <= "002099" Then wDepo = "CL"
            '
            'End If
            'OleDbConnection1.Close()
            '
            wDepo = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
            'Modify-End

            OleDbConnection1.Open()
            SQL = "SELECT FormNo, FormName, TableName1 FROM M_Form "
            SQL = SQL + " Where Active  =  '1' "
            SQL = SQL + "   And FormNo = '" + wFormNo + "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet2, "Flow")
            If DBDataSet2.Tables("Flow").Rows.Count > 0 Then
                wTableName = "F_" + DBDataSet2.Tables("Flow").Rows(0).Item("TableName1")
            End If
            OleDbConnection1.Close()

            OleDbConnection1.Open()
            SQL = "SELECT "
            SQL = SQL + "V_WaitHandle_01.FormNo as FormNo, "
            SQL = SQL + "V_WaitHandle_01.FormSno as FormSno, "
            SQL = SQL + "V_WaitHandle_01.Step as Step, "
            SQL = SQL + "V_WaitHandle_01.DecideID as UserID, "

            'SPD�e�U��
            If wFormNo >= "000001" And wFormNo <= "001000" Then
                If wFormNo = "000001" Or wFormNo = "000002" Then
                    SQL = SQL + wTableName + ".Level as Level, "
                    SQL = SQL + wTableName + ".MakeMap as MakeMap "
                Else
                    If wFormNo = "000011" Or wFormNo = "000012" Or wFormNo = "000014" Or wFormNo = "000015" Then
                        SQL = SQL + "'0' as Level "
                    Else
                        SQL = SQL + wTableName + ".Level as Level "
                    End If
                End If
            End If
            'HRW�e�U��
            If wFormNo >= "001001" And wFormNo <= "001050" Then
                SQL = SQL + "'0' as Level "
            End If
            'HRWCL�e�U��
            If wFormNo >= "001051" And wFormNo <= "001099" Then
                SQL = SQL + "'0' as Level "
            End If
            'SCD�e�U��
            If wFormNo >= "002001" And wFormNo <= "002099" Then
                SQL = SQL + "'0' as Level "
            End If

            SQL = SQL + "FROM " + wTableName + " "
            SQL = SQL + "Left Outer Join V_WaitHandle_01 ON " + wTableName + ".FormNo=V_WaitHandle_01.FormNo "
            SQL = SQL + "                               And " + wTableName + ".FormSno=V_WaitHandle_01.FormSno "
            SQL = SQL + "Where V_WaitHandle_01.Unique_ID = '" + Key + "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet3, "WaitData")
            If DBDataSet3.Tables("WaitData").Rows.Count > 0 Then
                pFormNo = DBDataSet3.Tables("WaitData").Rows(0).Item("FormNo")
                pFormSno = DBDataSet3.Tables("WaitData").Rows(0).Item("FormSno")
                pStep = DBDataSet3.Tables("WaitData").Rows(0).Item("Step")
                pUserID = DBDataSet3.Tables("WaitData").Rows(0).Item("UserID")
                pLevel = DBDataSet3.Tables("WaitData").Rows(0).Item("Level")
                If wFormNo = "000001" Or wFormNo = "000002" Then
                    pMakeMapName = DBDataSet3.Tables("WaitData").Rows(0).Item("MakeMap")
                End If
            End If
            OleDbConnection1.Close()

            If pMakeMapName <> "" Then
                OleDbConnection1.Open()
                SQL = "Select UserID From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserName =  '" & pMakeMapName & "'"
                Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter4.Fill(DBDataSet4, "M_Users")
                If DBDataSet4.Tables("M_Users").Rows.Count > 0 Then pMakeMap = DBDataSet4.Tables("M_Users").Rows(0).Item("UserID")
                OleDbConnection1.Close()
            End If

            If wFormNo = "000001" Or wFormNo = "000002" Then
                URL = "Simulation.aspx?pFormNo=" & pFormNo & _
                                      "&pFormSno=" & pFormSno & _
                                      "&pStep=" & pStep & _
                                      "&pUserID=" & pUserID & _
                                      "&pLevel=" & pLevel & _
                                      "&pAllocateID=" & pMakeMap & _
                                      "&pDepo=" & wDepo
            Else
                URL = "Simulation.aspx?pFormNo=" & pFormNo & _
                                      "&pFormSno=" & pFormSno & _
                                      "&pStep=" & pStep & _
                                      "&pUserID=" & pUserID & _
                                      "&pLevel=" & pLevel & _
                                      "&pDepo=" & wDepo
            End If

            'Call JavaScript
            Dim scriptString As String = "<script language=JavaScript> "
            scriptString = scriptString & "OpenSimulationSheet('" & URL & "'); "
            scriptString = scriptString & "</script>"
            If (Not Me.IsStartupScriptRegistered("Startup")) Then
                Me.RegisterStartupScript("Startup", scriptString)
            End If
        End If
    End Sub

End Class
