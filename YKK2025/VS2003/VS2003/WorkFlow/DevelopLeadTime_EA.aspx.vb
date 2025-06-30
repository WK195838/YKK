Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class DevelopLeadTime_EA
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BExcel As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DSts As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DEStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSStep As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DSDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSDate As System.Web.UI.WebControls.Button
    Protected WithEvents DEDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BEDate As System.Web.UI.WebControls.Button
    Protected WithEvents Go As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox

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
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim NowDateTime As String       '�{�b����ɶ�
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "CL"      '���c��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wDepo As String = "CL1"      '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
    'Modify-End

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "DevelopLeadTime_EA.aspx"
        Server.ScriptTimeout = 1200   '900��

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
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
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
        DSStep.Items.Clear()
        DEStep.Items.Clear()
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        DBTable1 = DBDataSet1.Tables("M_Flow")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("StepName")
            ListItem1.Value = DBTable1.Rows(i).Item("Step")
            DSStep.Items.Add(ListItem1)
            DEStep.Items.Add(ListItem1)
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

        '���
        DSDate.Text = CStr(DateTime.Now.Year) & "/" & CStr(DateTime.Now.Month) & "/1"
        DEDate.Text = CStr(DateTime.Now.Today)

        BSDate.Attributes("onclick") = "calendarPicker('Form1.DSDate');"
        BEDate.Attributes("onclick") = "calendarPicker('Form1.DEDate');"

        pFormNo = DFormName.SelectedValue
        Search_Item_Attribute()

    End Sub

    Sub Search_Item_Attribute()
        DPerson.ReadOnly = True
        DBuyer.ReadOnly = True

        '�̿��
        DPerson.Text = "�̿��"
        'Buyer
        DBuyer.Text = "Buyer"

        DPerson.ReadOnly = False
        If pFormNo = "000001" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000002" Then
            DBuyer.ReadOnly = False
        End If
        If pFormNo = "000003" Or pFormNo = "000004" Or pFormNo = "000005" Or _
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

    Private Sub DFormName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFormName.SelectedIndexChanged
        pFormNo = DFormName.SelectedValue
        SetSearchItem()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        pFormNo = DFormName.SelectedValue
        CreateWorkTable()
    End Sub

    Sub CreateWorkTable()
        Dim i As Integer
        Dim SQL As String
        Dim wTime As Integer
        Dim wSStep As Integer = 0
        Dim wEStep As Integer = 0
        Dim wSDate As String = DSDate.Text & " 00:00:00"
        Dim wEDate As String = DEDate.Text & " 23:59:59"
        Dim WorkTable As String = "Temp_DevelopLeadTime_" & Request.Cookies("UserID").Value

        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBTable1 As DataTable

        Dim OleDBCommand1 As New OleDbCommand
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()

        SQL = "SELECT FormNo, FormSno, No, FSts as Sts, Division, Buyer, AStartTime, AEndTime, "
        SQL = SQL + "Step, StepName, SeqNo, CompletedTime, ReceiptTime, ApplyTime,ApplyName, AWorkTime "

        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where Active = '0' "
        SQL = SQL + "  And FormNo = '" + DFormName.SelectedValue + "' "
        SQL = SQL + "  And Step   >= '" + DSStep.SelectedValue + "' "
        SQL = SQL + "  And Step   <= '" + DEStep.SelectedValue + "' "
        SQL = SQL + "  And AEndTime >= '" + wSDate + "' "
        SQL = SQL + "  And AEndTime <= '" + wEDate + "' "
        SQL = SQL + "  And FormSno  > '2000' "
        SQL = SQL + "  And FlowType <> '0' "

        If DSts.SelectedValue <> "ALL" Then
            SQL = SQL + " And FSts = '" + DSts.SelectedValue + "'"
        End If
        If DDivision.SelectedValue <> "ALL" Then
            SQL = SQL + " And Division = '" + DDivision.SelectedValue + "'"
        End If
        If DPerson.Text <> "�̿��" And DPerson.Text <> "" Then
            SQL = SQL + " And ApplyName Like '%" + DPerson.Text + "%'"
        End If
        If DBuyer.Text <> "Buyer" And DBuyer.Text <> "" Then
            SQL = SQL + " And Buyer Like '%" + DBuyer.Text + "%'"
        End If
        SQL = SQL + "Order by No "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Sheet")
        DBTable1 = DBDataSet1.Tables("Sheet")

        'Call Stored Procedure Create WorkTable
        SQL = "Exec sp_Temp_DevelopLeadTime '" & WorkTable & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQL
        OleDBCommand1.ExecuteNonQuery()

        '�p��u�{�ɶ�-Object
        For i = 0 To DBTable1.Rows.Count - 1
            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "RecordType, FormNo, FormSno, Step, SStep, EStep, CompletedTime, "     '1~7
            SQL = SQL + "Buyer, No, Sts, Division, ATimeDesc, ATime, BTimeDesc, BTime, CTimeDesc, CTime, "  '8~17
            SQL = SQL + "CreateUser, CreateTime "  '18~19
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("FormNo") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("FormSno")) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("Step")) + "', "
            SQL = SQL + " '" + CStr(DSStep.SelectedValue) + "', "
            SQL = SQL + " '" + CStr(DEStep.SelectedValue) + "', "

            Dim Str As String = CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("CompletedTime").Date) + " " + _
                                CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("CompletedTime").Hour) + ":" + _
                                CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("CompletedTime").Minute) + ":" + _
                                CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("CompletedTime").Second)
            SQL = SQL + " '" + Str + "', "

            SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Buyer") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("No") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("Sts")) + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Sheet").Rows(i).Item("Division") + "-" + DBDataSet1.Tables("Sheet").Rows(i).Item("ApplyName") + "', "
            '(a).��ڶ}�l~��ڧ���
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AStartTime")) + " ~ " + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime")) + "', "
            If DBDataSet1.Tables("Sheet").Rows(i).Item("AWorkTime") < 0 Then
                SQL = SQL + " '" + "0" + "', "
            Else
                SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AWorkTime")) + "', "
            End If
            '(b).����}�l~��ڧ���
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("ReceiptTime")) + " ~ " + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime")) + "', "

            oSchedule.GetDevelopTime(DBDataSet1.Tables("Sheet").Rows(i).Item("ReceiptTime"), DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime"), wTime, wDepo)

            If wTime < 0 Then
                SQL = SQL + " '" + "0" + "', "
            Else
                SQL = SQL + " '" + CStr(wTime) + "', "
            End If
            '(c).�e�U�}�l~��ڧ���
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("ApplyTime")) + " ~ " + CStr(DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime")) + "', "

            oSchedule.GetDevelopTime(DBDataSet1.Tables("Sheet").Rows(i).Item("ApplyTime"), DBDataSet1.Tables("Sheet").Rows(i).Item("AEndTime"), wTime, wDepo)

            If wTime < 0 Then
                SQL = SQL + " '" + "0" + "', "
            Else
                SQL = SQL + " '" + CStr(wTime) + "', "
            End If

            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "' "       '�@���ɶ�
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next

        DBDataSet1.Clear()

        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division, "
        SQL = SQL + "sum(ATime) as ATime, sum(BTime) as BTime "
        SQL = SQL + "FROM " & WorkTable & " "
        SQL = SQL + "group by FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division "
        SQL = SQL + "order by FormNo, FormSno, SStep, EStep, Buyer, No, Sts, Division "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Group")
        DBTable1 = DBDataSet1.Tables("Group")

        For i = 0 To DBTable1.Rows.Count - 1
            SQL = "Insert into " & WorkTable & " "
            SQL = SQL + "( "
            SQL = SQL + "RecordType, FormNo, FormSno, Step, SStep, EStep, CompletedTime, "     '1~7
            SQL = SQL + "Buyer, No, Sts, Division, ATimeDesc, Atime, BTimeDesc, BTime, CTimeDesc, CTime, "  '8~16
            SQL = SQL + "CreateUser, CreateTime "  '17~18
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '1', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("FormNo") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("FormSno")) + "', "
            SQL = SQL + " '0', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("SStep")) + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("EStep")) + "', "
            SQL = SQL + " '" + NowDateTime + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Buyer") + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("No") + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("Sts")) + "', "
            SQL = SQL + " '" + DBDataSet1.Tables("Group").Rows(i).Item("Division") + "', "
            '(a).��ڶ}�l~��ڧ���
            SQL = SQL + " '" + "" + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("ATime")) + "', "
            '(b).����}�l~��ڧ���
            SQL = SQL + " '" + "" + "', "
            SQL = SQL + " '" + CStr(DBDataSet1.Tables("Group").Rows(i).Item("BTime")) + "', "
            '(c).�e�U�}�l~��ڧ���
            SQL = SQL + " '" + "" + "', "
            SQL = SQL + " '" + "0" + "', "

            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '�@����
            SQL = SQL + " '" + NowDateTime + "' "       '�@���ɶ�
            SQL = SQL + " ) "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Next

        SQL = "SELECT No, FormNo, FormSno, SStep, EStep, Division, Buyer, ATime, BTime, "
        SQL = SQL + "str(SStep)+' ~ '+ str(EStep) As Step, "
        SQL = SQL + "Case Sts When 0 Then '�}�o��' When 1 Then 'OK' When 2 Then 'NG' Else '����' End As Sts, "

        SQL = SQL + "'DevelopLeadTime_EAList.aspx?' + "
        SQL = SQL + "'pFormNo='   + FormNo + "
        SQL = SQL + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        SQL = SQL + "'&pStart=' + str(SStep,Len(SStep)) + "
        SQL = SQL + "'&pEnd=' + str(EStep,Len(EStep)) + "
        SQL = SQL + "'&pStartTime=' + '" & wSDate + "' + "
        SQL = SQL + "'&pEndTime='   + '" & wEDate + "' + "
        SQL = SQL + "'&pTable=' + '" & WorkTable + "'"
        SQL = SQL + " As URL "

        SQL = SQL + "FROM " + WorkTable + " "
        SQL = SQL + "Where RecordType = '1' "
        SQL = SQL + "Order by No "

        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet2, "Header")
        DataGrid1.DataSource = DBDataSet2
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = DataGrid1.AllowPaging   '�{���O���P

        DataGrid1.AllowPaging = False   '�{���O���P

        Response.AppendHeader("Content-Disposition", "attachment;filename=DevelopLeadTime_List.xls")     '�{���O���P
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
