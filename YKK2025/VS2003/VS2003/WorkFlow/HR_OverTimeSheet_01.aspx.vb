Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_OverTimeSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOverTimeDate As System.Web.UI.WebControls.Button
    Protected WithEvents DOverTimeDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFood As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DFAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOverTimeSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDateType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTraffic As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DCVacation As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOverTime As System.Web.UI.WebControls.Button
    Protected WithEvents D46Hours As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAgentID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX15H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX167H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX20H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX267H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX15M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX167M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX20M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX267M As System.Web.UI.WebControls.TextBox

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
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As Integer              '�ʺA����Top��m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wAgentID As String          '�Q�N�z�HID
    Dim NowDateTime As String       '�{�b����ɶ�
    Dim wNextGate As String         '�U�@���H
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String = "TP"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End
    Dim wUserName As String = ""    '�m�W�N�z��

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '�]�w�@�ΰѼ�
        TopPosition()           '���s��RequestedField��Top��m

        If Not Me.IsPostBack Then   '���OPostBack
            ShowSheetField("New")   '��������ܤ�����J�ˬd
            ShowSheetFunction()     '���\����s���
            If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
                ShowFormData()      '��ܪ����
                UpdateTranFile()    '��s������
            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�
        Else
            ShowSheetField("Post")  '��������ܤ�����J�ˬd
            ShowMessage()           '�W�Ǹ���ˬd����ܰT��
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
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        ' PAYROLL FIELD-START
        DFPRAH1.Style("left") = -500 & "px"
        DFPRAM1.Style("left") = -500 & "px"
        DFPRAH2.Style("left") = -500 & "px"
        DFPRAM2.Style("left") = -500 & "px"
        DFPRBH.Style("left") = -500 & "px"
        DFPRBM.Style("left") = -500 & "px"
        DFPRCH.Style("left") = -500 & "px"
        DFPRCM.Style("left") = -500 & "px"

        DFPRTAX15H.Style("left") = -500 & "px"
        DFPRTAX15M.Style("left") = -500 & "px"
        DFPRTAX167H.Style("left") = -500 & "px"
        DFPRTAX167M.Style("left") = -500 & "px"
        DFPRTAX20H.Style("left") = -500 & "px"
        DFPRTAX20M.Style("left") = -500 & "px"
        DFPRTAX267H.Style("left") = -500 & "px"
        DFPRTAX267M.Style("left") = -500 & "px"

        DAgentID.Style("left") = -500 & "px"
        DSalaryYM1.Style("left") = -500 & "px"
        ' END

        Response.Cookies("PGM").Value = "HR_OverTimeSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '�W�Ǹ���ˬd����ܰT��
    '**
    '*****************************************************************
    Sub ShowMessage()
        'Dim Message As String = ""

        'Check������
        'If DContactFile.Visible Then
        'If DContactFile.PostedFile.FileName <> "" Then
        'If Message = "" Then
        'Message = "������"
        'Else
        '    Message = Message & ", " & "������"
        'End If
        'End If
        'End If

        'If Message <> "" Then
        'Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
        'Response.Write(YKK.ShowMessage(Message))
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��s������
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�
        'Modify-End
        '��渹�X,���y����,�u�{���d���X,�Ǹ�,�x�_,ñ�֪�
    End Sub
    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("OverTimeFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_OverTimeSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_OverTimeSheet")
        If DBDataSet1.Tables("F_OverTimeSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("No")                     'No
            DOverTimeDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("OverTimeDate")   '�[�Z���
            DDateType.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DateType")           '�[�Z�O
            DSalaryYM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")           '���ݦ~��
            DSalaryYM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")          '���ݦ~��-�ˬd�ϥ�
            DDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Date")                   '�ӽФ��
            SetFieldData("Name", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Name"))          '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DivisionCode")   '�����N�X
            SetFieldData("CVacation", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("CVacation"))    '�ե�
            SetFieldData("Food", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Food"))          '�뭹
            SetFieldData("Traffic", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Traffic"))    '��q

            SetFieldData("BStartH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartH").ToString)   '�w�w�}�l-��
            SetFieldData("BStartM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartM").ToString)   '�w�w�}�l-��
            SetFieldData("BEndH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndH").ToString)       '�w�w�פ�-��
            SetFieldData("BEndM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndM").ToString)       '�w�w�פ�-��
            DBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BH").ToString                      '�p�⵲�G-��
            DBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BM").ToString                      '�p�⵲�G-��

            SetFieldData("AStartH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartH").ToString)   '��ڶ}�l-��
            SetFieldData("AStartM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartM").ToString)   '��ڶ}�l-��
            SetFieldData("AEndH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndH").ToString)       '��ڲפ�-��
            SetFieldData("AEndM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndM").ToString)       '��ڲפ�-��
            DAH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AH").ToString                      '�p�⵲�G-��
            DAM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AM").ToString                      '�p�⵲�G-��

            DFAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH1").ToString                  '�֩w����2��-��
            DFAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM1").ToString                  '�֩w����2��-��
            DFAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH2").ToString                  '�֩w����2�~-��
            DFAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM2").ToString                  '�֩w����2�~-��

            DFBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBH").ToString                    '�֩w����-��
            DFBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBM").ToString                    '�֩w����-��
            DFCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCH").ToString                    '�֩w��w����-��
            DFCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCM").ToString                    '�֩w��w����-��

            ' PAYROLL FIELD-START
            DFPRAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAH1").ToString              'PAYROLL�֩w����2��-��
            DFPRAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAM1").ToString              'PAYROLL�֩w����2��-��
            DFPRAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAH2").ToString              'PAYROLL�֩w����2�~-��
            DFPRAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAM2").ToString              'PAYROLL�֩w����2�~-��

            DFPRBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRBH").ToString                'PAYROLL�֩w����-��
            DFPRBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRBM").ToString                'PAYROLL�֩w����-��
            DFPRCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRCH").ToString                'PAYROLL�֩w��w����-��
            DFPRCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRCM").ToString                'PAYROLL�֩w��w����-��

            ' JOY BY 2017/6/19
            DFPRTAX15H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX15H").ToString        'PAYROLL�֩w���|1.5
            DFPRTAX15M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX15M").ToString        '
            DFPRTAX167H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX167H").ToString      'PAYROLL�֩w���|1.67
            DFPRTAX167M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX167M").ToString      '
            DFPRTAX20H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX20H").ToString        'PAYROLL�֩w���|2.0
            DFPRTAX20M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX20M").ToString        '
            DFPRTAX267H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX267H").ToString      'PAYROLL�֩w���|2.67
            DFPRTAX267M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX267M").ToString      '

            ' END
            DFReason.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FReason")                     '�[�Z�z��
            '
            '�W�L46H/92H
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' ����w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' ���饼��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' �`�w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' �`����
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "�`�ɼơG[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "����ɼơG[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-�L�Ǧ�
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-���
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-������
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' �զ�
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
            '������
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '����

                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '����z�ѥN�X
                        If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
                            DReasonDesc.Text = ""   '�����L����
                        Else
                            DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '����z��
                            DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '�����L����
                        End If

                    End If
                End If
            End If
            '�֩w�i�����
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet9, "DecideHistory")
            DataGrid9.DataSource = DBDataSet9
            DataGrid9.DataBind()
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����z�ѥN�X�I���ƥ�
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     ���\����s���
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '�q�lñ�����ϥ�
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '���[�ɮץ��ϥ�(��FormField���]�w)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '�x�s���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '�s
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
                '--
            Else
                BSAVE.Visible = False
            End If
            'NG-1���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                '�s
                BNG1.Attributes("onclick") = "this.disabled = true;" & "Button('NG1', '" + BNG1.Value + "');"
                '--
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                '�s
                BNG2.Attributes("onclick") = "this.disabled = true;" & "Button('NG2', '" + BNG2.Value + "');"
                '--
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                '-- ������s���G�� JOY 16/08/17
                '��
                'BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                '�s
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '" + BOK.Value + "');"
                '--
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '��Ǻ޲z
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet���
                DOverTimeSheet1.Visible = True   '���Sheet-1
                DDescSheet.Visible = True        '����Sheet

                '��Ǻ޲z
                If wDelay = 1 Then
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '����Sheet
                    Else
                        DDelay.Visible = False  '����Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '����z�ѥN�X
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "���`�G�ݿ�J����z��")
                    DReason.Visible = True         '����z��
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "���`�G�ݿ�J����z��")
                    DReasonDesc.Visible = True     '�����L����
                    Top = 896
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 782
                End If

                '������
                DDecideDesc.Visible = True      '����
                '�����ݿ�J
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "���`�G�ݿ�J����")
                End If

                '���s��m
                BNG1.Style.Add("Top", Top)     'NG1���s
                BNG2.Style.Add("Top", Top)     'NG2���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                '�֩w�i��
                DHistoryLabel.Style.Add("Top", Top + 24)  '�֩w�i��
                DataGrid9.Style.Add("Top", Top + 48)     '�֩w�i��
            End If
        Else
            Top = 704
            'Sheet����
            DDescSheet.Visible = False  '����Sheet
            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False    '����
            DReasonCode.Visible = False    '����z�ѥN�X
            DReason.Visible = False        '����z��
            DReasonDesc.Visible = False    '�����L����
            DHistoryLabel.Visible = False  '�֩w�i��
            '���s��m
            BNG1.Style.Add("Top", Top)     'NG1���s
            BNG2.Style.Add("Top", Top)     'NG2���s
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      'OK���s
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     �]�w�u�X�����ƥ�
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'BOverTimeDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DOverTimeDate', 'Form1.DDateType', 'Form1.DSalaryYM');"  '�[�Z���
        '
        BOverTimeDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DOverTimeDate', 'Form1.DDateType', 'Form1.DSalaryYM');"  '�[�Z���
        'Modify-End
        BCardTime.Attributes("onclick") = "ShowCardTime();"    '��d�O��
        BOverTime.Attributes("onclick") = "ShowOverTime();"    '�[�Z�O��
        ' ��ܳq���T��
        Dim Cmd As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From W_ReadMessageHistory "
        SQL = SQL & " Where FormNo = '" & wFormNo & "' "
        SQL = SQL & "   And ReadUser = '" & Request.QueryString("pUserID") & "' "
        SQL = SQL & "   And ReadFlag = '" & "1" & "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MessageHistory")
        If DBDataSet1.Tables("MessageHistory").Rows.Count <= 0 Then
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/WorkFlowSub/NoticeMessageOverTime.aspx?pUserID=" + Request.QueryString("pUserID") & "&pFormNo=" + wFormNo & "','NoticeMessage','status=0,toolbar=0,width=550,height=700,resizable=no,scrollbars=no');" + _
                  "</script>"
        End If

        OleDbConnection1.Close()
        '
        If Cmd <> "" Then
            Response.Write(Cmd)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        '���s��RequestedField��Top��m
        If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 782
                Else
                    If DDelay.Visible = True Then
                        Top = 896
                    Else
                        Top = 782
                    End If
                End If
            End If
        Else
            Top = 704
        End If
        '----
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     ��������ܤ�����J�ˬd
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '�ظm�����ݩʰ}�C
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '��渹�X,�u�{���d���X,���W,����ݩ�

        SetFieldAttribute(pPost)     '���U����ݩʤ�����J�ˬd���]�w
    End Sub
    '*****************************************************************
    '**(SetFieldAttribute)
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim wEmpID, wJobTitle, wJobCode, wDivision, wDivisionCode, wDepoName, wDepoCode, wSalaryYM As String
        Dim wDateType As String = "����"

        OleDbConnection1.Open()
        '���o�ӽЪ̸�T
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
        SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            wDepoCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            wDepoName = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            'Delete-Start by Joy 2009/11/20(2010��ƾ����)
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "01" Then wDepo = "CL"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "10" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "11" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "51" Then wDepo = "YA"
            'Delete-End
            wUserName = DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")
            wEmpID = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            wJobTitle = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            wJobCode = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            wDivision = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            wDivisionCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        '���o���[�Z�O
        DBDataSet1.Clear()
        SQL = "Select VacationType From M_Vacation "
        SQL = SQL + "Where Active = '1' "
        'Modify-Start by Joy 2009/11/20(2010��ƾ����)
        'SQL = SQL + "  and Depo = '" + wDepo + "' "
        '
        SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
        'Modify-End
        SQL = SQL + "  and YMD  = '" + CStr(DateTime.Now.Today) + "' "
        SQL = SQL + "Order by YMD "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Vacation")
        If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
            If DBDataSet1.Tables("Vacation").Rows(0).Item("VacationType") = 0 Then
                wDateType = "����"
            Else
                If DBDataSet1.Tables("Vacation").Rows(0).Item("VacationType") = 1 Then
                    wDateType = "��w����"
                Else
                    wDateType = "���𰲤�"
                End If
            End If
        End If
        '���o���ݦ~��
        If DateTime.Now.Month < 10 Then
            wSalaryYM = CStr(DateTime.Now.Year) + "/0" + CStr(DateTime.Now.Month)
        Else
            wSalaryYM = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month)
        End If
        '------------------------------------------------------------------------------------------
        'No
        Select Case FindFieldInf("No")
            Case 0  '���
                'DNo.BackColor = Color.LightGray
                DNo.BackColor = Color.White
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '�ק�+�ˬd
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "���`�G�ݿ�J�ܢ�")
                DNo.Visible = True
            Case 2  '�ק�
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '����
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        '�[�Z���
        Select Case FindFieldInf("OverTimeDate")
            Case 0  '���
                DOverTimeDate.BackColor = Color.LightGray
                DOverTimeDate.ReadOnly = True
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = False
            Case 1  '�ק�+�ˬd
                DOverTimeDate.BackColor = Color.GreenYellow
                DOverTimeDate.ReadOnly = True
                ShowRequiredFieldValidator("DOverTimeDateRqd", "DOverTimeDate", "���`�G�ݿ�J�[�Z���")
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = True
            Case 2  '�ק�
                DOverTimeDate.BackColor = Color.Yellow
                DOverTimeDate.ReadOnly = True
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = True
            Case Else   '����
                DOverTimeDate.Visible = False
                BOverTimeDate.Visible = False
        End Select
        If pPost = "New" Then DOverTimeDate.Text = CStr(DateTime.Now.Today)
        '�ӽФ��
        Select Case FindFieldInf("Date")
            Case 0  '���
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
            Case 1  '�ק�+�ˬd
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = True
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J�ӽФ��")
                DDate.Visible = True
            Case 2  '�ק�
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = True
                DDate.Visible = True
            Case Else   '����
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '�[�Z�O
        Select Case FindFieldInf("DateType")
            Case 0  '���
                DDateType.BackColor = Color.LightGray
                DDateType.Visible = True
            Case 1  '�ק�+�ˬd
                DDateType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateTypeRqd", "DDateType", "���`�G�ݿ�J�[�Z�O")
                DDateType.Visible = True
            Case 2  '�ק�
                DDateType.BackColor = Color.Yellow
                DDateType.Visible = True
            Case Else   '����
                DDateType.Visible = False
        End Select
        If pPost = "New" Then DDateType.Text = wDateType
        '���ݦ~��
        Select Case FindFieldInf("SalaryYM")
            Case 0  '���
                DSalaryYM.BackColor = Color.LightGray
                DSalaryYM.ReadOnly = True
                DSalaryYM.Visible = True
            Case 1  '�ק�+�ˬd
                DSalaryYM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSalaryYMRqd", "DSalaryYM", "���`�G�ݿ�J���ݦ~��")
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case 2  '�ק�
                DSalaryYM.BackColor = Color.Yellow
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case Else   '����
                DSalaryYM.Visible = False
        End Select
        If pPost = "New" Then DSalaryYM.Text = wSalaryYM
        '�m�W
        Select Case FindFieldInf("Name")
            Case 0  '���
                DName.BackColor = Color.LightGray
                DName.Visible = True
            Case 1  '�ק�+�ˬd
                DName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNameRqd", "DName", "���`�G�ݿ�J�m�W")
                DName.Visible = True
            Case 2  '�ק�
                DName.BackColor = Color.Yellow
                DName.Visible = True
            Case Else   '����
                DName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Name", "ZZZZZZ")
        'EmpID
        Select Case FindFieldInf("EmpID")
            Case 0  '���
                DEmpID.BackColor = Color.LightGray
                DEmpID.ReadOnly = True
                DEmpID.Visible = True
            Case 1  '�ק�+�ˬd
                DEmpID.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEmpIDRqd", "DEmpID", "���`�G�ݿ�J�d��")
                DEmpID.Visible = True
            Case 2  '�ק�
                DEmpID.BackColor = Color.Yellow
                DEmpID.Visible = True
            Case Else   '����
                DEmpID.Visible = False
        End Select
        If pPost = "New" Then DEmpID.Text = wEmpID
        '¾��
        Select Case FindFieldInf("JobTitle")
            Case 0  '���
                DJobTitle.BackColor = Color.LightGray
                DJobTitle.ReadOnly = True
                DJobTitle.Visible = True
            Case 1  '�ק�+�ˬd
                DJobTitle.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobTitleRqd", "DJobTitle", "���`�G�ݿ�J¾��")
                DJobTitle.Visible = True
            Case 2  '�ק�
                DJobTitle.BackColor = Color.Yellow
                DJobTitle.Visible = True
            Case Else   '����
                DJobTitle.Visible = False
        End Select
        If pPost = "New" Then DJobTitle.Text = wJobTitle
        '¾�٥N�X
        Select Case FindFieldInf("JobCode")
            Case 0  '���
                DJobCode.BackColor = Color.LightGray
                DJobCode.ReadOnly = True
                DJobCode.Visible = True
            Case 1  '�ק�+�ˬd
                DJobCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobCodeRqd", "DJobCode", "���`�G�ݿ�J¾�٥N�X")
                DJobCode.Visible = True
            Case 2  '�ק�
                DJobCode.BackColor = Color.Yellow
                DJobCode.Visible = True
            Case Else   '����
                DJobCode.Visible = False
        End Select
        If pPost = "New" Then DJobCode.Text = wJobCode
        'Depo Name
        Select Case FindFieldInf("DepoName")
            Case 0  '���
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '�ק�+�ˬd
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "���`�G�ݿ�J���q")
                DDepoName.Visible = True
            Case 2  '�ק�
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '����
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then DDepoName.Text = wDepoName
        'Depo Code
        Select Case FindFieldInf("DepoCode")
            Case 0  '���
                DDepoCode.BackColor = Color.LightGray
                DDepoCode.ReadOnly = True
                DDepoCode.Visible = True
            Case 1  '�ק�+�ˬd
                DDepoCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoCodeRqd", "DDepoCode", "���`�G�ݿ�J���q�N�X")
                DDepoCode.Visible = True
            Case 2  '�ק�
                DDepoCode.BackColor = Color.Yellow
                DDepoCode.Visible = True
            Case Else   '����
                DDepoCode.Visible = False
        End Select
        If pPost = "New" Then DDepoCode.Text = wDepoCode
        '����
        Select Case FindFieldInf("Division")
            Case 0  '���
                DDivision.BackColor = Color.LightGray
                DDivision.ReadOnly = True
                DDivision.Visible = True
            Case 1  '�ק�+�ˬd
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "���`�G�ݿ�J����")
                DDivision.Visible = True
            Case 2  '�ק�
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '����
                DDivision.Visible = False
        End Select
        If pPost = "New" Then DDivision.Text = wDivision
        '�����N�X
        Select Case FindFieldInf("DivisionCode")
            Case 0  '���
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.ReadOnly = True
                DDivisionCode.Visible = True
            Case 1  '�ק�+�ˬd
                DDivisionCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "���`�G�ݿ�J�����N�X")
                DDivisionCode.Visible = True
            Case 2  '�ק�
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.Visible = True
            Case Else   '����
                DDivisionCode.Visible = False
        End Select
        If pPost = "New" Then DDivisionCode.Text = wDivisionCode
        '�ե�
        Select Case FindFieldInf("CVacation")
            Case 0  '���
                DCVacation.BackColor = Color.LightGray
                DCVacation.Visible = True
            Case 1  '�ק�+�ˬd
                DCVacation.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCVacationRqd", "DCVacation", "���`�G�ݿ�J�ե�")
                DCVacation.Visible = True
            Case 2  '�ק�
                DCVacation.BackColor = Color.Yellow
                DCVacation.Visible = True
            Case Else   '����
                DCVacation.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CVacation", "ZZZZZZ")
        '�뭹
        Select Case FindFieldInf("Food")
            Case 0  '���
                DFood.BackColor = Color.LightGray
                DFood.Visible = True
            Case 1  '�ק�+�ˬd
                DFood.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFoodRqd", "DFood", "���`�G�ݿ�J�뭹")
                DFood.Visible = True
            Case 2  '�ק�
                DFood.BackColor = Color.Yellow
                DFood.Visible = True
            Case Else   '����
                DFood.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Food", "ZZZZZZ")
        '��q
        Select Case FindFieldInf("Traffic")
            Case 0  '���
                DTraffic.BackColor = Color.LightGray
                DTraffic.Visible = True
            Case 1  '�ק�+�ˬd
                DTraffic.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTrafficRqd", "DTraffic", "���`�G�ݿ�J��q")
                DTraffic.Visible = True
            Case 2  '�ק�
                DTraffic.BackColor = Color.Yellow
                DTraffic.Visible = True
            Case Else   '����
                DTraffic.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Traffic", "ZZZZZZ")
        '�w�w�}�l-��
        Select Case FindFieldInf("BStartH")
            Case 0  '���
                DBStartH.BackColor = Color.LightGray
                DBStartH.Visible = True
            Case 1  '�ק�+�ˬd
                DBStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartHRqd", "DBStartH", "���`�G�ݿ�J�w�w�}�l-��")
                DBStartH.Visible = True
            Case 2  '�ק�
                DBStartH.BackColor = Color.Yellow
                DBStartH.Visible = True
            Case Else   '����
                DBStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartH", "ZZZZZZ")
        '�w�w�}�l-��
        Select Case FindFieldInf("BStartM")
            Case 0  '���
                DBStartM.BackColor = Color.LightGray
                DBStartM.Visible = True
            Case 1  '�ק�+�ˬd
                DBStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartMRqd", "DBStartM", "���`�G�ݿ�J�w�w�}�l-��")
                DBStartM.Visible = True
            Case 2  '�ק�
                DBStartM.BackColor = Color.Yellow
                DBStartM.Visible = True
            Case Else   '����
                DBStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartM", "ZZZZZZ")
        '�w�w�פ�-��
        Select Case FindFieldInf("BEndH")
            Case 0  '���
                DBEndH.BackColor = Color.LightGray
                DBEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DBEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndHRqd", "DBEndH", "���`�G�ݿ�J�w�w�פ�-��")
                DBEndH.Visible = True
            Case 2  '�ק�
                DBEndH.BackColor = Color.Yellow
                DBEndH.Visible = True
            Case Else   '����
                DBEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndH", "ZZZZZZ")
        '�w�w�פ�-��
        Select Case FindFieldInf("BEndM")
            Case 0  '���
                DBEndM.BackColor = Color.LightGray
                DBEndM.Visible = True
            Case 1  '�ק�+�ˬd
                DBEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndMRqd", "DBEndM", "���`�G�ݿ�J�w�w�פ�-��")
                DBEndM.Visible = True
            Case 2  '�ק�
                DBEndM.BackColor = Color.Yellow
                DBEndM.Visible = True
            Case Else   '����
                DBEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndM", "ZZZZZZ")
        '�w�w�p��-��
        Select Case FindFieldInf("BH")
            Case 0  '���
                DBH.BackColor = Color.LightGray
                DBH.ReadOnly = True
                DBH.Visible = True
            Case 1  '�ק�+�ˬd
                DBH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBHRqd", "DBH", "���`�G�ݿ�J�w�w�p��-��")
                DBH.Visible = True
            Case 2  '�ק�
                DBH.BackColor = Color.Yellow
                DBH.Visible = True
            Case Else   '����
                DBH.Visible = False
        End Select
        If pPost = "New" Then DBH.Text = "0"
        '�w�w�p��-��
        Select Case FindFieldInf("BM")
            Case 0  '���
                DBM.BackColor = Color.LightGray
                DBM.ReadOnly = True
                DBM.Visible = True
            Case 1  '�ק�+�ˬd
                DBM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBMRqd", "DBM", "���`�G�ݿ�J�w�w�p��-��")
                DBM.Visible = True
            Case 2  '�ק�
                DBM.BackColor = Color.Yellow
                DBM.Visible = True
            Case Else   '����
                DBM.Visible = False
        End Select
        If pPost = "New" Then DBM.Text = "0"
        '��ڶ}�l-��
        Select Case FindFieldInf("AStartH")
            Case 0  '���
                DAStartH.BackColor = Color.LightGray
                DAStartH.Visible = True
            Case 1  '�ק�+�ˬd
                DAStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartHRqd", "DAStartH", "���`�G�ݿ�J��ڶ}�l-��")
                DAStartH.Visible = True
            Case 2  '�ק�
                DAStartH.BackColor = Color.Yellow
                DAStartH.Visible = True
            Case Else   '����
                DAStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartH", "ZZZZZZ")
        '��ڶ}�l-��
        Select Case FindFieldInf("AStartM")
            Case 0  '���
                DAStartM.BackColor = Color.LightGray
                DAStartM.Visible = True
            Case 1  '�ק�+�ˬd
                DAStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartMRqd", "DAStartM", "���`�G�ݿ�J��ڶ}�l-��")
                DAStartM.Visible = True
            Case 2  '�ק�
                DAStartM.BackColor = Color.Yellow
                DAStartM.Visible = True
            Case Else   '����
                DAStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartM", "ZZZZZZ")
        '��ڲפ�-��
        Select Case FindFieldInf("AEndH")
            Case 0  '���
                DAEndH.BackColor = Color.LightGray
                DAEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DAEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndHRqd", "DAEndH", "���`�G�ݿ�J��ڲפ�-��")
                DAEndH.Visible = True
            Case 2  '�ק�
                DAEndH.BackColor = Color.Yellow
                DAEndH.Visible = True
            Case Else   '����
                DAEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndH", "ZZZZZZ")
        '��ڲפ�-��
        Select Case FindFieldInf("AEndM")
            Case 0  '���
                DAEndM.BackColor = Color.LightGray
                DAEndM.Visible = True
            Case 1  '�ק�+�ˬd
                DAEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndMRqd", "DAEndM", "���`�G�ݿ�J��ڲפ�-��")
                DAEndM.Visible = True
            Case 2  '�ק�
                DAEndM.BackColor = Color.Yellow
                DAEndM.Visible = True
            Case Else   '����
                DAEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndM", "ZZZZZZ")
        '��ڭp��-��
        Select Case FindFieldInf("AH")
            Case 0  '���
                DAH.BackColor = Color.LightGray
                DAH.ReadOnly = True
                DAH.Visible = True
            Case 1  '�ק�+�ˬd
                DAH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAHRqd", "DAH", "���`�G�ݿ�J��ڭp��-��")
                DAH.Visible = True
            Case 2  '�ק�
                DAH.BackColor = Color.Yellow
                DAH.Visible = True
            Case Else   '����
                DAH.Visible = False
        End Select
        If pPost = "New" Then DAH.Text = "0"
        '��ڭp��-��
        Select Case FindFieldInf("AM")
            Case 0  '���
                DAM.BackColor = Color.LightGray
                DAM.ReadOnly = True
                DAM.Visible = True
            Case 1  '�ק�+�ˬd
                DAM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAMRqd", "DAM", "���`�G�ݿ�J��ڭp��-��")
                DAM.Visible = True
            Case 2  '�ק�
                DAM.BackColor = Color.Yellow
                DAM.Visible = True
            Case Else   '����
                DAM.Visible = False
        End Select
        If pPost = "New" Then DAM.Text = "0"
        '�֩w����2��-��
        Select Case FindFieldInf("FAH1")
            Case 0  '���
                DFAH1.BackColor = Color.LightGray
                DFAH1.ReadOnly = True
                DFAH1.Visible = True
            Case 1  '�ק�+�ˬd
                DFAH1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAH1Rqd", "DFAH1", "���`�G�ݿ�J�֩w����2��-��")
                DFAH1.Visible = True
            Case 2  '�ק�
                DFAH1.BackColor = Color.Yellow
                DFAH1.Visible = True
            Case Else   '����
                DFAH1.Visible = False
        End Select
        If pPost = "New" Then DFAH1.Text = "0"
        '�֩w����2��-��
        Select Case FindFieldInf("FAM1")
            Case 0  '���
                DFAM1.BackColor = Color.LightGray
                DFAM1.ReadOnly = True
                DFAM1.Visible = True
            Case 1  '�ק�+�ˬd
                DFAM1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAM1Rqd", "DFAM1", "���`�G�ݿ�J�֩w����2��-��")
                DFAM1.Visible = True
            Case 2  '�ק�
                DFAM1.BackColor = Color.Yellow
                DFAM1.Visible = True
            Case Else   '����
                DFAM1.Visible = False
        End Select
        If pPost = "New" Then DFAM1.Text = "0"
        '�֩w����2�~-��
        Select Case FindFieldInf("FAH2")
            Case 0  '���
                DFAH2.BackColor = Color.LightGray
                DFAH2.ReadOnly = True
                DFAH2.Visible = True
            Case 1  '�ק�+�ˬd
                DFAH2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAH2Rqd", "DFAH2", "���`�G�ݿ�J�֩w����2�~-��")
                DFAH2.Visible = True
            Case 2  '�ק�
                DFAH2.BackColor = Color.Yellow
                DFAH2.Visible = True
            Case Else   '����
                DFAH2.Visible = False
        End Select
        If pPost = "New" Then DFAH2.Text = "0"
        '�֩w����2�~-��
        Select Case FindFieldInf("FAM2")
            Case 0  '���
                DFAM2.BackColor = Color.LightGray
                DFAM2.ReadOnly = True
                DFAM2.Visible = True
            Case 1  '�ק�+�ˬd
                DFAM2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAM2Rqd", "DFAM2", "���`�G�ݿ�J�֩w����2�~-��")
                DFAM2.Visible = True
            Case 2  '�ק�
                DFAM2.BackColor = Color.Yellow
                DFAM2.Visible = True
            Case Else   '����
                DFAM2.Visible = False
        End Select
        If pPost = "New" Then DFAM2.Text = "0"
        '�֩w����-��
        Select Case FindFieldInf("FBH")
            Case 0  '���
                DFBH.BackColor = Color.LightGray
                DFBH.ReadOnly = True
                DFBH.Visible = True
            Case 1  '�ק�+�ˬd
                DFBH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFBHRqd", "DFBH", "���`�G�ݿ�J�֩w����-��")
                DFBH.Visible = True
            Case 2  '�ק�
                DFBH.BackColor = Color.Yellow
                DFBH.Visible = True
            Case Else   '����
                DFBH.Visible = False
        End Select
        If pPost = "New" Then DFBH.Text = "0"
        '�֩w����-��
        Select Case FindFieldInf("FBM")
            Case 0  '���
                DFBM.BackColor = Color.LightGray
                DFBM.ReadOnly = True
                DFBM.Visible = True
            Case 1  '�ק�+�ˬd
                DFBM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFBMRqd", "DFBM", "���`�G�ݿ�J�֩w����-��")
                DFBM.Visible = True
            Case 2  '�ק�
                DFBM.BackColor = Color.Yellow
                DFBM.Visible = True
            Case Else   '����
                DFBM.Visible = False
        End Select
        If pPost = "New" Then DFBM.Text = "0"
        '�֩w��w����-��
        Select Case FindFieldInf("FCH")
            Case 0  '���
                DFCH.BackColor = Color.LightGray
                DFCH.ReadOnly = True
                DFCH.Visible = True
            Case 1  '�ק�+�ˬd
                DFCH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCHRqd", "DFCH", "���`�G�ݿ�J�֩w��w����-��")
                DFCH.Visible = True
            Case 2  '�ק�
                DFCH.BackColor = Color.Yellow
                DFCH.Visible = True
            Case Else   '����
                DFCH.Visible = False
        End Select
        If pPost = "New" Then DFCH.Text = "0"
        '�֩w��w����-��
        Select Case FindFieldInf("FCM")
            Case 0  '���
                DFCM.BackColor = Color.LightGray
                DFCM.ReadOnly = True
                DFCM.Visible = True
            Case 1  '�ק�+�ˬd
                DFCM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCMRqd", "DFCM", "���`�G�ݿ�J�֩w��w����-��")
                DFCM.Visible = True
            Case 2  '�ק�
                DFCM.BackColor = Color.Yellow
                DFCM.Visible = True
            Case Else   '����
                DFCM.Visible = False
        End Select
        If pPost = "New" Then DFCM.Text = "0"
        '�[�Z�z��
        Select Case FindFieldInf("FReason")
            Case 0  '���
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '�ק�+�ˬd
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "���`�G�ݿ�J�[�Z�z��")
                DFReason.Visible = True
            Case 2  '�ק�
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '����
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""
        '
        ' �W�L46H
        ' Step=1 �_��
        If wStep = 1 Then
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' ����w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' ���饼��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' �`�w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' �`����
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "�`�ɼơG[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "����ɼơG[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-�L�Ǧ�
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-���
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-������
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' �զ�
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
        End If
        '
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     �إߪ��ݿ�J���
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        idx = FindFieldInf(pFieldName)

        '�m�W
        If pFieldName = "Name" Then
            DName.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DName.Items.Add(ListItem1)
                End If
            Else
                '�n�J��
                Dim ListItem1 As New ListItem
                ListItem1.Text = wUserName
                ListItem1.Value = Request.QueryString("pUserID")
                ListItem1.Selected = True
                DName.Items.Add(ListItem1)
                '�����N�z
                SQL = "Select UserName, UserID From M_Agent  "
                SQL = SQL + "Where Active = '1' "
                SQL = SQL + "  And AllForm = '0' "
                SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Agent")
                DBTable1 = DBDataSet1.Tables("M_Agent")
                If DBTable1.Rows.Count <= 0 Then
                    '��@���N�z
                    DBDataSet1.Clear()
                    SQL = "Select UserName, UserID From M_Agent  "
                    SQL = SQL + "Where Active = '1' "
                    SQL = SQL + "  And AllForm = '1' "
                    SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                    SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                    SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                    SQL = SQL + "  And FormNo = '001001' "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Agent")
                    DBTable1 = DBDataSet1.Tables("M_Agent")
                    If DBTable1.Rows.Count <= 0 Then
                    Else
                        For i = 0 To DBTable1.Rows.Count - 1
                            Dim ListItem2 As New ListItem
                            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                            ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                            DName.Items.Add(ListItem2)
                        Next
                    End If
                Else
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        '�ե�
        If pFieldName = "CVacation" Then
            DCVacation.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCVacation.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='CVACATION' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCVacation.Items.Add(ListItem1)
                Next
            End If
        End If
        '�뭹
        If pFieldName = "Food" Then
            DFood.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFood.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='FOOD' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFood.Items.Add(ListItem1)
                Next
            End If
        End If
        '��q
        If pFieldName = "Traffic" Then
            DTraffic.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTraffic.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='TRAFFIC' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTraffic.Items.Add(ListItem1)
                Next
            End If
        End If
        '�w�w�}�l-��
        If pFieldName = "BStartH" Then
            DBStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '�w�w�}�l-��
        If pFieldName = "BStartM" Then
            DBStartM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBStartM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBStartM.Items.Add(ListItem1)
                Next
            End If
        End If
        '�w�w�פ�-��
        If pFieldName = "BEndH" Then
            DBEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '�w�w�פ�-��
        If pFieldName = "BEndM" Then
            DBEndM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBEndM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBEndM.Items.Add(ListItem1)
                Next
            End If
        End If
        '��ڶ}�l-��
        If pFieldName = "AStartH" Then
            DAStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '��ڶ}�l-��
        If pFieldName = "AStartM" Then
            DAStartM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAStartM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartM.Items.Add(ListItem1)
                Next
            End If
        End If
        '��ڲפ�-��
        If pFieldName = "AEndH" Then
            DAEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '��ڲפ�-��
        If pFieldName = "AEndM" Then
            DAEndM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAEndM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndM.Items.Add(ListItem1)
                Next
            End If
        End If

        '����z�ѥN�X
        If pFieldName = "ReasonCode" Then
            DReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReasonCode.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DKey")
                    ListItem1.Value = DBTable1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '����z��
        If pFieldName = "Reason" Then
            SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                DReason.Text = DBTable1.Rows(i).Item("Data")
            Next
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FindFieldInf)
    '**     �M��������ݩ�
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
                Dim Message As String = ""

                'Check�z��
                If ErrCode = 0 Then
                    If DReasonCode.Visible = True Then
                        If DReasonCode.SelectedValue = "99" Then
                            If DReasonDesc.Text = "" Then ErrCode = 9210
                        End If
                    End If
                End If
                '�x�s���
                If ErrCode = 0 Then
                    ModifyData("SAVE", "0")           '��s����� Sts=0(����)
                    ModifyTranData("SAVE", "0")       '��s������
                Else
                    If ErrCode = 9210 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '�W���ɮ�ErrCode=0

                If ErrCode = 0 Then
                    Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("OK", 0, "1")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("NG1", 1, "2")
            End If
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '����Button�B�@
                FlowControl("NG2", 2, "3")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     �y�{����
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   �U�@���d�ɨϥ� 
    '**        pSts=0:���B�z, 1:OK, 2:NG1, 3:NG2, 4:�w�\Ū, 5:�Q���  ��sT_Waithandle���A
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '�ŧi�@��Error�ΨӧP�O�O�_��ƥ��`�A�w�]��False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        'Check����z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If
        'Check�~����ݦ~���ܧ�
        If ErrCode = 0 Then
            If wStep = 50 Then
                If DSalaryYM.Text <> DSalaryYM1.Text Then
                    ErrCode = 9041
                End If
            End If
        End If
        'Check�w�w�p�⵲�G
        If ErrCode = 0 Then
            If wStep = 1 Or wStep = 500 Then
                If DBH.Text = "0" And DBM.Text = "0" Then ErrCode = 9050
            End If
        End If
        'Check��ڭp�⵲�G
        If ErrCode = 0 Then
            If wStep = 30 Then
                If pFun = "OK" Then
                    If DAH.Text = "0" And DAM.Text = "0" Then ErrCode = 9050
                End If
            End If
        End If
        'Check�ե�
        If ErrCode = 0 Then
            If DCVacation.SelectedValue = "2.�n" Then
                ' �w�w�ե�
                If wStep = 1 Then
                    If CInt(DBH.Text) < 4 Then
                        ErrCode = 9051
                    End If
                End If
                If wStep = 500 Then
                    If CInt(DBH.Text) < 4 Then
                        ErrCode = 9051
                    End If
                End If
                ' ��ڽե�
                If wStep = 30 Then
                    If pFun = "OK" Then
                        If CInt(DAH.Text) < 4 Then
                            ErrCode = 9051
                        End If
                    End If
                End If
            End If
        End If
        'Check�[�Z�ɼ�
        If ErrCode = 0 Then
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim DBDataSet1 As New DataSet
            Dim SQL As String
            '�W�L46H/92H
            Dim wHours46 As Boolean = False '46H-Flag
            Dim wHours92 As Boolean = False '92H-Flag
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            Dim xHours As Integer = 0
            Dim yHours As Integer = 0
            OleDbConnection1.Open()
            '���o�ӽЪ�-ID
            DAgentID.Text = ""
            SQL = "Select UserID From M_Users "
            SQL = SQL & " Where UserID = '" & DName.SelectedValue & "'"
            SQL = SQL & "   And Active = '1' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                ' �Q�N�z�HID
                If UCase(DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")) <> UCase(Request.QueryString("pUserID")) Then
                    DAgentID.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")
                End If
            End If
            ' ����w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' ���饼��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' �`�w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' �`����
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            xHours = CDbl(DFAH1.Text) * 60 + CDbl(DFAH2.Text) * 60 + CDbl(DFBH.Text) * 60 + CDbl(DFCH.Text) * 60 + CDbl(DFAM1.Text) + CDbl(DFAM2.Text) + CDbl(DFBM.Text) + CDbl(DFCM.Text)
            yHours = CDbl(DFAH1.Text) * 60 + CDbl(DFAH2.Text) * 60 + CDbl(DFAM1.Text) + CDbl(DFAM2.Text)
            If aHours + bHours + yHours > 2760 Then '46H
                wHours46 = True
            End If
            If cHours + dHours + xHours > 5520 Then   '92H
                wHours92 = True
            End If
            OleDbConnection1.Close()
            'Check 46, 92H
            If wStep = 500 And pFun = "NG2" Then
                ' [500]���s�ץ� and ������ �� ��Check    
            Else
                If wHours46 = True Then
                    If wStep = 1 Or wStep = 500 Then
                        If UCase(Request.QueryString("pUserID")) <> "GAS007" Then
                            ErrCode = 9052
                        End If
                    End If
                End If
                If ErrCode = 0 Then
                    If wHours92 = True Then
                        If wStep = 1 Or wStep = 500 Then
                            If UCase(Request.QueryString("pUserID")) <> "GAS007" Then
                                ErrCode = 9053
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' 2016/12/23 �Ұ�k�ܧ����
        ' @@@@
        ' �[�Z���()
        ' ���𰲤�: 4H, 8H, 12H
        ' �Ұ���,�갲��,�갲��(ykk): 8H, 12H

        ' 2018/3/1 �Ұ�k�ܧ����
        ' �]03/01�Ұ�k�תk, �нЭק�[�Z��]�w,����!
        '   -->�𮧤�[�Z���u�ɡA�אּ�ֹ�p��(��>12H�ɤ��ХX�{�T�{�T��,�Ȥ��}��ӽ�)
        If ErrCode = 0 Then
            ' 2018/3/1 �Ұ�k�ܧ����
            If DDateType.Text = "���𰲤�" Or DDateType.Text = "����" Or DDateType.Text = "��w����" Then
                If CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) > 12 Then
                    ErrCode = 9059
                Else
                    If CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) > 0 Then
                        ErrCode = 9059
                    End If
                End If
            End If

            ' 2016/12/23 �Ұ�k�ܧ����
            'If CDate(DOverTimeDate.Text) >= CDate("2016/12/23") Then
            '    If DDateType.Text = "���𰲤�" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 4 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9056
            '        End If
            '    End If
            '    If DDateType.Text = "����" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9057
            '        End If
            '    End If
            '    If DDateType.Text = "��w����" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9058
            '        End If
            '    End If
            'End If
        End If
        '
        '�ˬd�e�U��No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then

                ErrCode = oCommon.CommissionNo("001001", wFormSno, wStep, DNo.Text)     '��渹�X, ���y����, �u�{, �e�U��No

                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '�O�_����
            Dim RepeatRun As Boolean = False    '�O�_���а���
            Dim MultiJob As Integer = 0         '�h�H�֩w
            Dim wLevel As String = ""           '������

            While Run = True
                Run = False     '����Flag=������
                '--���o�U�@���ѼƳ]�w---------
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=�q��
                Dim pCount As Integer
                '--���oLeadTime�ѼƳ]�w---------
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--���o�u�{�t���ѼƳ]�w---------
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--��L�ѼƳ]�w---------
                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '���y����
                Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '�P�_�O�_�_��
                        '���o���y����
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '��渹�X, ���y����
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '�ӽЬy�{��ƫظm
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                            'Modify-End

                            '�]�w�e�UNo
                            DNo.Text = SetNo(NewFormSno)
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '���O�q�������а���
                            '��s������
                            ModifyTranData(pFun, pSts)

                            '�y�{��Ƶ���
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '��渹�X,���y����,�u�{���d���X,��ƾ�,ñ�֪�, �y�{�����_(�|ñ)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '�O�q�������а���
                        End If
                    End If
                End If

                '���o�U�@��
                If ErrCode = 0 And pRunNextStep = 1 Then
                    Dim wAllocateID As String = ""

                    If DAgentID.Text <> "" Then
                        RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, DAgentID.Text, wAllocateID, MultiJob, _
                                                      pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    Else
                        RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                      pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    End If
                    '��渹�X,�u�{���d���X,ñ�֪�,�ӽЪ�,�Q�N�z��,�Q���w��,�h�H�֩w�u�{No,
                    '�U�@�u�{��, ���X, ����, �Q�N�z��, �H��, �B�z��k, �ʧ@(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '�ظm�y�{���
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '���o�U�@���H��(�T���ɨϥ�)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '���o�֩w��-�s�զ�O��
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '���o�u�{�t���̫���
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '�֩w��, ��渹�X, �u�{���X, ���O(0:�q��,1:�֩w), �}�l���, �̫���, ���

                            '���o�w�w�}�l,������{�p��
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '��渹�X,�u�{���X,������,QC-L/T,�{�b�ɶ�, �w�w�}�l���, �w�w�������, ��ƾ�

                            '�ظm�y�{���
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,�ӽЪ�ID,��ƾ�,�ظm��, ñ�֪�, �Q�N�z��, �ӽЪ�, �w�w�}�l��, �w�w������, ���n��
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '��u�{��{�վ�
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        'ñ�֪�,��渹�X,���y����,�u�{���d���X,�Ǹ�,�{�b���,������,��ƾ�
                        'Modify-End
                    End If
                End If
                '�x�s�����
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If
                    Else
                        If pNextStep = 999 Then     '�u�{������?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '��s�����
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else
                            ModifyData(pFun, "0")         '��s����� Sts=0(����)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '�ǰe�l��
                    If pNextStep <> 999 Then
                        For i = 1 To pCount

                            oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O

                        Next i
                    Else

                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O

                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If

                    wStep = pNextStep     '�U�@�u�{���d���X
                    wFormSno = NewFormSno '�U�@�u�{���y����
                Else
                    EnabledButton()   '�_��Button�B�@
                    If ErrCode = 9110 Then Message = "���o���y�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9120 Then Message = "�y�{��Ƨ�s���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9130 Then Message = "�U�u�{�p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9131 Then Message = "�L�U�u�{�޲z�H,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9140 Then Message = "�u�{�w�w�}�l�Χ�����p�ⲧ�`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9150 Then Message = "�U�@�u�{��ƫظm���`,�нT�{�γs���t�ΤH��!"
                    If ErrCode = 9160 Then Message = "�u�{������ƫظm���`(999),�нT�{�γs���t�ΤH��!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '�x�s���ErrCode=0
            End While       '���а���

            If ErrCode = 0 Then
                '--�l��ǰe---------
                oCommon.SendMail()

                Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '�_��Button�B�@
            If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
            If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
            If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
            If ErrCode = 9040 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9041 Then Message = "�~����ݦ~�뤣�i�ܧ�,�нT�{!"
            If ErrCode = 9050 Then Message = "�w�w�ι�ڥ[�Z�ɶ��ݶ�g,�нT�{!"

            If ErrCode = 9051 Then Message = "�[�Z�ɶ��ݶW�L [4 H] �H�W�~�i�ե�,�нT�{!"
            If ErrCode = 9052 Then Message = "�W�L����[46H]�k�wĵ�ɽu, �L�k��g�[�Z�ӽг�,�нT�{!"
            If ErrCode = 9053 Then Message = "�W�L���[92H]�k�wĵ�ɽu, �L�k��g�[�Z�ӽг�,�нT�{!"

            If ErrCode = 9056 Then Message = "�𰲤�(���𰲤�)�[�Z�ɼƻݲŦX [4H]/[8H]/[12H],�нT�{!"
            If ErrCode = 9057 Then Message = "����(�Ұ���)�[�Z�ɼƻݲŦX [8H]/[12H],�нT�{!"
            If ErrCode = 9058 Then Message = "��w����Z�ɼƻݲŦX [8H]/[12H],�нT�{!"

            If ErrCode = 9059 Then Message = "�𰲤�[�Z�ɼƤ��i�W�L [12H],�нT�{!"

            If ErrCode = 9060 Then Message = "�e�U��No.����,�нT�{�e�U��No.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     ���o���i�ת��A
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            'NG-1���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            End If
            'NG-2���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            End If
            'OK���s
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("OverTimeFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_OverTimeSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "            '1~5
        SQl = SQl + "Date, OverTimeDate, DateType, Name, EmpID, "          '6~10
        SQl = SQl + "JobTitle, JobCode, Division, DivisionCode, CVacation, Food, Traffic, "    '11~17
        SQl = SQl + "BStartH, BStartM, BEndH, BEndM, BH, BM, "             '17~22
        SQl = SQl + "AStartH, AStartM, AEndH, AEndM, AH, AM, "             '23~28
        SQl = SQl + "FAH1, FAM1, FAH2, FAM2, "                             '29~32
        SQl = SQl + "FBH, FBM, FCH, FCM, "                                 '33~37
        ' PAYROLL FIELD-START
        SQl = SQl + "FPRAH1, FPRAM1, FPRAH2, FPRAM2, "                     '
        SQl = SQl + "FPRBH, FPRBM, FPRCH, FPRCM, "                         '

        ' JOY BY 2017/6/19
        SQl = SQl + "FPRTAX15H, FPRTAX15M, FPRTAX167H, FPRTAX167M, "
        SQl = SQl + "FPRTAX20H, FPRTAX20M, FPRTAX267H, FPRTAX267M, "

        ' END
        SQl = SQl + "FReason, SalaryYM, DepoName, DepoCode, "              '37~40
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '41~44
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '001001', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~10
        SQl = SQl + " '" + DDate.Text + "', "                   '�ӽФ��
        SQl = SQl + " '" + DOverTimeDate.Text + "', "           '�[�Z���
        SQl = SQl + " N'" + DDateType.Text + "', "              '�[�Z�O
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', "     '�m�W
        SQl = SQl + " '" + DEmpID.Text + "', "                  'EMP-ID
        '11~15
        SQl = SQl + " N'" + DJobTitle.Text + "', "              '¾��
        SQl = SQl + " '" + DJobCode.Text + "', "                '¾�٥N�X
        SQl = SQl + " N'" + DDivision.Text + "', "              '����
        SQl = SQl + " '" + DDivisionCode.Text + "', "           '�����N�X
        SQl = SQl + " N'" + DCVacation.SelectedValue + "', "    '�ե�
        SQl = SQl + " N'" + DFood.SelectedValue + "', "         '�뭹
        SQl = SQl + " N'" + DTraffic.SelectedValue + "', "      '��q
        '16~21
        SQl = SQl + " '" + CStr(CInt(DBStartH.SelectedValue)) + "', "  '�w�w�}�l-��
        SQl = SQl + " '" + CStr(CInt(DBStartM.SelectedValue)) + "', "  '�w�w�}�l-��
        SQl = SQl + " '" + CStr(CInt(DBEndH.SelectedValue)) + "', "    '�w�w�פ�-��
        SQl = SQl + " '" + CStr(CInt(DBEndM.SelectedValue)) + "', "    '�w�w�פ�-��
        SQl = SQl + " '" + CStr(CInt(DBH.Text)) + "', "                '�w�w�p��-��
        SQl = SQl + " '" + CStr(CInt(DBM.Text)) + "', "                '�w�w�p��-��
        '22~27
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "  '��ڶ}�l-��
        SQl = SQl + " '" + CStr(CInt(DAStartM.SelectedValue)) + "', "  '��ڶ}�l-��
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "    '��ڲפ�-��
        SQl = SQl + " '" + CStr(CInt(DAEndM.SelectedValue)) + "', "    '��ڲפ�-��
        SQl = SQl + " '" + CStr(CInt(DAH.Text)) + "', "                '��ڭp��-��
        SQl = SQl + " '" + CStr(CInt(DAM.Text)) + "', "                '��ڭp��-��
        '28~31
        SQl = SQl + " '" + CStr(CInt(DFAH1.Text)) + "', "              '�֩w����2��-��
        SQl = SQl + " '" + CStr(CInt(DFAM1.Text)) + "', "              '�֩w����2��-��
        SQl = SQl + " '" + CStr(CInt(DFAH2.Text)) + "', "              '�֩w����2�~-��
        SQl = SQl + " '" + CStr(CInt(DFAM2.Text)) + "', "              '�֩w����2�~-��
        '32~36
        SQl = SQl + " '" + CStr(CInt(DFBH.Text)) + "', "               '�֩w����-��
        SQl = SQl + " '" + CStr(CInt(DFBM.Text)) + "', "               '�֩w����-��
        SQl = SQl + " '" + CStr(CInt(DFCH.Text)) + "', "               '�֩w��w����-��
        SQl = SQl + " '" + CStr(CInt(DFCM.Text)) + "', "               '�֩w��w����-��
        ' PAYROLL FIELD-START
        SQl = SQl + " '" + CStr(CInt(DFPRAH1.Text)) + "', "            'PAYROLL�֩w����2��-��
        SQl = SQl + " '" + CStr(CInt(DFPRAM1.Text)) + "', "            'PAYROLL�֩w����2��-��
        SQl = SQl + " '" + CStr(CInt(DFPRAH2.Text)) + "', "            'PAYROLL�֩w����2�~-��
        SQl = SQl + " '" + CStr(CInt(DFPRAM2.Text)) + "', "            'PAYROLL�֩w����2�~-��

        SQl = SQl + " '" + CStr(CInt(DFPRBH.Text)) + "', "             'PAYROLL�֩w����-��
        SQl = SQl + " '" + CStr(CInt(DFPRBM.Text)) + "', "             'PAYROLL�֩w����-��
        SQl = SQl + " '" + CStr(CInt(DFPRCH.Text)) + "', "             'PAYROLL�֩w��w����-��
        SQl = SQl + " '" + CStr(CInt(DFPRCM.Text)) + "', "             'PAYROLL�֩w��w����-��

        ' JOY BY 2017/6/19
        SQl = SQl + " '" + CStr(CInt(DFPRTAX15H.Text)) + "', "         'PAYROLL�֩w���| 1.5
        SQl = SQl + " '" + CStr(CInt(DFPRTAX15M.Text)) + "', "         '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX167H.Text)) + "', "        'PAYROLL�֩w���| 1.67
        SQl = SQl + " '" + CStr(CInt(DFPRTAX167M.Text)) + "', "        '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX20H.Text)) + "', "         'PAYROLL�֩w���| 2.0
        SQl = SQl + " '" + CStr(CInt(DFPRTAX20M.Text)) + "', "         '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX267H.Text)) + "', "        'PAYROLL�֩w���| 2.67
        SQl = SQl + " '" + CStr(CInt(DFPRTAX267M.Text)) + "', "        '

        ' END
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "   '�[�Z�z��
        '37~39
        SQl = SQl + " '" + DSalaryYM.Text + "', "                      '���ݦ~��
        SQl = SQl + " N'" + DDepoName.Text + "', "                     '���q�O
        SQl = SQl + " '" + DDepoCode.Text + "', "                      '���q�OCode
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '�@����
        SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                       '�ק��
        SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     ��s�����
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("OverTimeFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_OverTimeSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " OverTimeDate = '" & DOverTimeDate.Text & "',"
        SQl = SQl + " DateType = N'" & DDateType.Text & "',"
        SQl = SQl + " SalaryYM = '" & DSalaryYM.Text & "',"
        SQl = SQl + " Name = N'" & DName.SelectedItem.Text & "',"
        SQl = SQl + " EmpID = '" & DEmpID.Text & "',"
        SQl = SQl + " JobTitle = N'" & DJobTitle.Text & "',"
        SQl = SQl + " JobCode = '" & DJobCode.Text & "',"
        SQl = SQl + " DepoName = N'" & DDepoName.Text & "',"
        SQl = SQl + " DepoCode = '" & DDepoCode.Text & "',"
        SQl = SQl + " Division = N'" & DDivision.Text & "',"
        SQl = SQl + " DivisionCode = '" & DDivisionCode.Text & "',"
        SQl = SQl + " CVacation = N'" & DCVacation.SelectedValue & "',"
        SQl = SQl + " Food = N'" & DFood.SelectedValue & "',"
        SQl = SQl + " Traffic = N'" & DTraffic.SelectedValue & "',"

        SQl = SQl + " BStartH = '" & CStr(CInt(DBStartH.SelectedValue)) & "',"
        SQl = SQl + " BStartM = '" & CStr(CInt(DBStartM.SelectedValue)) & "',"
        SQl = SQl + " BEndH = '" & CStr(CInt(DBEndH.SelectedValue)) & "',"
        SQl = SQl + " BEndM = '" & CStr(CInt(DBEndM.SelectedValue)) & "',"
        SQl = SQl + " BH = '" & CStr(CInt(DBH.Text)) & "',"
        SQl = SQl + " BM = '" & CStr(CInt(DBM.Text)) & "',"

        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AStartM = '" & CStr(CInt(DAStartM.SelectedValue)) & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " AEndM = '" & CStr(CInt(DAEndM.SelectedValue)) & "',"
        SQl = SQl + " AH = '" & CStr(CInt(DAH.Text)) & "',"
        SQl = SQl + " AM = '" & CStr(CInt(DAM.Text)) & "',"

        SQl = SQl + " FAH1 = '" & CStr(CInt(DFAH1.Text)) & "',"
        SQl = SQl + " FAM1 = '" & CStr(CInt(DFAM1.Text)) & "',"
        SQl = SQl + " FAH2 = '" & CStr(CInt(DFAH2.Text)) & "',"
        SQl = SQl + " FAM2 = '" & CStr(CInt(DFAM2.Text)) & "',"
        SQl = SQl + " FBH = '" & CStr(CInt(DFBH.Text)) & "',"
        SQl = SQl + " FBM = '" & CStr(CInt(DFBM.Text)) & "',"
        SQl = SQl + " FCH = '" & CStr(CInt(DFCH.Text)) & "',"
        SQl = SQl + " FCM = '" & CStr(CInt(DFCM.Text)) & "',"

        ' PAYROLL FIELD-START
        SQl = SQl + " FPRAH1 = '" & CStr(CInt(DFPRAH1.Text)) & "',"
        SQl = SQl + " FPRAM1 = '" & CStr(CInt(DFPRAM1.Text)) & "',"
        SQl = SQl + " FPRAH2 = '" & CStr(CInt(DFPRAH2.Text)) & "',"
        SQl = SQl + " FPRAM2 = '" & CStr(CInt(DFPRAM2.Text)) & "',"
        SQl = SQl + " FPRBH = '" & CStr(CInt(DFPRBH.Text)) & "',"
        SQl = SQl + " FPRBM = '" & CStr(CInt(DFPRBM.Text)) & "',"
        SQl = SQl + " FPRCH = '" & CStr(CInt(DFPRCH.Text)) & "',"
        SQl = SQl + " FPRCM = '" & CStr(CInt(DFPRCM.Text)) & "',"

        ' JOY BY 2017/6/19
        SQl = SQl + " FPRTAX15H = '" & CStr(CInt(DFPRTAX15H.Text)) & "',"
        SQl = SQl + " FPRTAX15M = '" & CStr(CInt(DFPRTAX15M.Text)) & "',"
        SQl = SQl + " FPRTAX167H = '" & CStr(CInt(DFPRTAX167H.Text)) & "',"
        SQl = SQl + " FPRTAX167M = '" & CStr(CInt(DFPRTAX167M.Text)) & "',"
        SQl = SQl + " FPRTAX20H = '" & CStr(CInt(DFPRTAX20H.Text)) & "',"
        SQl = SQl + " FPRTAX20M = '" & CStr(CInt(DFPRTAX20M.Text)) & "',"
        SQl = SQl + " FPRTAX267H = '" & CStr(CInt(DFPRTAX267H.Text)) & "',"
        SQl = SQl + " FPRTAX267M = '" & CStr(CInt(DFPRTAX267M.Text)) & "',"

        ' END
        SQl = SQl + " FReason = N'" & YKK.ReplaceString(DFReason.Text) & "',"
        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     �l�[�����ƩM�e�U���Ӫ�
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        Dim DBDataSet1 As New DataSet

        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "T_CommissionNo")

        If DBDataSet1.Tables("T_CommissionNo").Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDBCommand1.ExecuteNonQuery()
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBDataSet1.Tables("T_CommissionNo").Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQl
                    OleDBCommand1.ExecuteNonQuery()
                End If
            End If
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �m�W�ܧ� 
    '**
    '*****************************************************************
    Private Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        '���o�ӽЪ̸�T
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName, UserID From M_Users "
        SQL = SQL & " Where UserID = '" & DName.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            DDepoCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            DDepoName.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            DEmpID.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            DJobCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            DDivision.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            DDivisionCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        '
        ' �W�L46H
        ' Step=1 �_��
        If wStep = 1 Then
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' ����w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' ���饼��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' �`�w��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' �`����
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "�`�ɼơG[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "����ɼơG[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-�L�Ǧ�
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-���
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-������
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' �զ�
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
        End If
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����w�w�פ�-�� 
    '**
    '*****************************************************************
    Private Sub DBEndH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBEndH.SelectedIndexChanged
        Dim StartTime As String = DBStartH.SelectedValue + DBStartM.SelectedValue
        Dim EndTime As String = DBEndH.SelectedValue + DBEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DBStartH.SelectedValue, DBStartM.SelectedValue, DBEndH.SelectedValue, DBEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DBH.Text = CStr(Hour)
            DBM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DBH.Text = "0"
            DBM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����w�w�פ�-�� 
    '**
    '*****************************************************************
    Private Sub DBEndM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBEndM.SelectedIndexChanged
        Dim StartTime As String = DBStartH.SelectedValue + DBStartM.SelectedValue
        Dim EndTime As String = DBEndH.SelectedValue + DBEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DBStartH.SelectedValue, DBStartM.SelectedValue, DBEndH.SelectedValue, DBEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DBH.Text = CStr(Hour)
            DBM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DBH.Text = "0"
            DBM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ڲפ�-�� 
    '**
    '*****************************************************************
    Private Sub DAEndH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAEndH.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ڲפ�-�� 
    '**
    '*****************************************************************
    Private Sub DAEndM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAEndM.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �p��֩w�ɶ� 
    '**
    '*****************************************************************
    Sub CalApproveOverTime(ByVal sHour As Integer, ByVal sMinute As Integer)
        Dim xHour As Integer = sHour
        Dim xMinute As Integer = sMinute
        '�Ҷԭp�� -----------------------------------------
        '  �M���
        DFAH1.Text = "0"
        DFAM1.Text = "0"
        DFAH2.Text = "0"
        DFAM2.Text = "0"
        DFBH.Text = "0"
        DFBM.Text = "0"
        DFCH.Text = "0"
        DFCM.Text = "0"
        '  �p��
        Select Case DDateType.Text
            Case "����"
                If sHour > 8 Then
                    DFBH.Text = "8"
                    DFBM.Text = "0"
                    sHour = sHour - 8
                    If sHour > 2 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        If sHour = 2 And sMinute > 0 Then
                            DFAH1.Text = "2"
                            DFAM1.Text = "0"
                            DFAH2.Text = CStr(sHour - 2)
                            DFAM2.Text = CStr(sMinute)
                        Else
                            DFAH1.Text = CStr(sHour)
                            DFAM1.Text = CStr(sMinute)
                        End If
                    End If
                Else
                    DFBH.Text = CStr(sHour)
                    DFBM.Text = CStr(sMinute)
                End If
            Case "��w����"
                If sHour > 8 Then
                    DFCH.Text = "8"
                    DFCM.Text = "0"
                    sHour = sHour - 8
                    If sHour > 2 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        If sHour = 2 And sMinute > 0 Then
                            DFAH1.Text = "2"
                            DFAM1.Text = "0"
                            DFAH2.Text = CStr(sHour - 2)
                            DFAM2.Text = CStr(sMinute)
                        Else
                            DFAH1.Text = CStr(sHour)
                            DFAM1.Text = CStr(sMinute)
                        End If
                    End If
                Else
                    DFCH.Text = CStr(sHour)
                    DFCM.Text = CStr(sMinute)
                End If
            Case "���𰲤�"
                If sHour > 2 Then
                    DFAH1.Text = "2"
                    DFAM1.Text = "0"
                    DFAH2.Text = CStr(sHour - 2)
                    DFAM2.Text = CStr(sMinute)
                Else
                    If sHour = 2 And sMinute > 0 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        DFAH1.Text = CStr(sHour)
                        DFAM1.Text = CStr(sMinute)
                    End If
                End If
            Case Else   '����
                If sHour > 2 Then
                    DFAH1.Text = "2"
                    DFAM1.Text = "0"
                    DFAH2.Text = CStr(sHour - 2)
                    DFAM2.Text = CStr(sMinute)
                Else
                    If sHour = 2 And sMinute > 0 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        DFAH1.Text = CStr(sHour)
                        DFAM1.Text = CStr(sMinute)
                    End If
                End If
        End Select
        '�~��p�� -----------------------------------------
        '�M���
        DFPRAH1.Text = "0"
        DFPRAM1.Text = "0"
        DFPRAH2.Text = "0"
        DFPRAM2.Text = "0"
        DFPRBH.Text = "0"
        DFPRBM.Text = "0"
        DFPRCH.Text = "0"
        DFPRCM.Text = "0"

        DFPRTAX15H.Text = "0"
        DFPRTAX15M.Text = "0"
        DFPRTAX167H.Text = "0"
        DFPRTAX167M.Text = "0"
        DFPRTAX20H.Text = "0"
        DFPRTAX20M.Text = "0"
        DFPRTAX267H.Text = "0"
        DFPRTAX267M.Text = "0"

        '  �M�w�[�Z��-�[�Z���O (����,����,��w����)
        Dim SQL, xDateType As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        xDateType = DDateType.Text
        SQL = "Select *  From M_Vacation "      ' �~��M�Φ�ƾ�(�u����w����)
        SQL = SQL & " Where Active = '1' "
        SQL = SQL & " And Depo = '" & "PR1" & "'"
        SQL = SQL & " And YMD = '" & DOverTimeDate.Text & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Vacation")
        If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
            xDateType = "��w����"
        Else
            If xDateType = "��w����" Then
                xDateType = "����"
            End If
        End If
        OleDbConnection1.Close()
        '  �p��[�Z�O�v  JOY BY 2017/6/19
        Select Case xDateType
            Case "���𰲤�"
                ' ���𰲤�: 1~2=���|1.5, 3~8=���|1.67, 9~ =���|2.67
                Select Case xHour
                    Case 1 To 2
                        DFPRTAX15H.Text = CStr(xHour)
                        If xHour = 2 Then
                            DFPRTAX167M.Text = CStr(xMinute)
                        Else
                            DFPRTAX15M.Text = CStr(xMinute)
                        End If
                    Case 3 To 8
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(xHour - 2)
                        If xHour = 8 Then
                            DFPRTAX267M.Text = CStr(xMinute)
                        Else
                            DFPRTAX167M.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(6)
                        DFPRTAX267H.Text = CStr(xHour - 8)
                        DFPRTAX267M.Text = CStr(xMinute)
                End Select
            Case "����"
                ' �Ұ���: 1~8=�K�|1.5, 9~10=���|1.5, 11~ =���|1.67
                ' �갲��: 1~8=�K�|1.5, 9~10=���|1.5, 11~ =���|1.67
                Select Case xHour
                    Case 1 To 8
                        DFPRBH.Text = CStr(xHour)
                        If xHour = 8 Then
                            DFPRTAX15M.Text = CStr(xMinute)
                        Else
                            DFPRBM.Text = CStr(xMinute)
                        End If
                    Case 9 To 10
                        DFPRBH.Text = CStr(8)
                        DFPRTAX15H.Text = CStr(xHour - 8)
                        If xHour = 10 Then
                            DFPRTAX167M.Text = CStr(xMinute)
                        Else
                            DFPRTAX15M.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRBH.Text = CStr(8)
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(xHour - 10)
                        DFPRTAX167M.Text = CStr(xMinute)
                End Select
            Case "��w����"
                ' �갲��(ykk): 1~8=�K�|2.0, , 9~ =���|2.0
                Select Case xHour
                    Case 1 To 8
                        DFPRCH.Text = CStr(xHour)
                        If xHour = 8 Then
                            DFPRTAX20M.Text = CStr(xMinute)
                        Else
                            DFPRCM.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRCH.Text = CStr(8)
                        DFPRTAX20H.Text = CStr(xHour - 8)
                        DFPRTAX20M.Text =  CStr(xMinute)
                End Select
            Case Else
                ' ����: 9~10=���|1.34, 11~ =���|1.67
                DFPRAH1.Text = DFAH1.Text
                DFPRAM1.Text = DFAM1.Text
                DFPRAH2.Text = DFAH2.Text
                DFPRAM2.Text = DFAM2.Text
        End Select

        'Select Case xDateType
        '    Case "���𰲤�"     '(���𰲤�,�𰲤�)

        '        ' MOD-Start 2016/12/28
        '        If CDate(DOverTimeDate.Text) >= CDate("2016/12/23") Then
        '            If xHour >= 2 Then
        '                DFPRBH.Text = CStr(2)           '2��=1.5
        '                DFPRAH2.Text = CStr(xHour - 2)  '2�~=1.67
        '                DFPRAM2.Text = CStr(xMinute)
        '            Else
        '                DFPRBH.Text = CStr(xHour)
        '                DFPRBM.Text = CStr(xMinute)
        '            End If
        '        Else
        '            DFPRBH.Text = CStr(xHour)
        '            DFPRBM.Text = CStr(xMinute)
        '        End If
        '        'DFPRBH.Text = CStr(xHour)
        '        'DFPRBM.Text = CStr(xMinute)
        '        ' MOD-End

        '    Case "����"         '(�Ұ���,�갲��)
        '        DFPRBH.Text = CStr(xHour)
        '        DFPRBM.Text = CStr(xMinute)
        '    Case "��w����"     '(�갲��ykk)
        '        DFPRCH.Text = CStr(xHour)
        '        DFPRCM.Text = CStr(xMinute)
        '    Case Else   '����
        '        DFPRAH1.Text = DFAH1.Text
        '        DFPRAM1.Text = DFAM1.Text
        '        DFPRAH2.Text = DFAH2.Text
        '        DFPRAM2.Text = DFAM2.Text
        'End Select

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �p��[�Z�_���ɶ� 
    '**
    '*****************************************************************
    Function CalOverTime(ByVal sHH As String, ByVal sMM As String, ByVal eHH As String, ByVal eMM As String) As Integer
        Dim StartTime As Integer = CInt(sHH) * 60 + CInt(sMM)
        Dim EndTime As Integer = CInt(eHH) * 60 + CInt(eMM)
        CalOverTime = EndTime - StartTime
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �s�s�e�UNo 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        'Set�����
        Str2 = CStr(DateTime.Now.Month)  '��
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '��
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str + Str2
        'Set�渹
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�xx
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '�w�q���\���ɮ׮榡
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '���o�ɮ׮榡
        For i = 0 To allowedExtensions.Length - 1           '�v�@�ˬd���\���榡���O�_���W�Ǫ��榡
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check�W���ɮ�Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 2000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
        'Check�W���ɮ׮榡
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �_��Button(OK, Ng1, NG2, Save)�B�@ 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function InputCheck() As Integer
        InputCheck = 0
        'No
        If InputCheck = 0 Then
            If FindFieldInf("No") = 1 Then
                If DNo.Text = "" Then InputCheck = 1
            End If
        End If
        '���
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�m�W
        If InputCheck = 0 Then
            If FindFieldInf("Name") = 1 Then
                If DName.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'EmpID
        If InputCheck = 0 Then
            If FindFieldInf("EmpID") = 1 Then
                If DEmpID.Text = "" Then InputCheck = 1
            End If
        End If
        '¾��
        If InputCheck = 0 Then
            If FindFieldInf("JobTitle") = 1 Then
                If DJobTitle.Text = "" Then InputCheck = 1
            End If
        End If
        '¾�٥N�X
        If InputCheck = 0 Then
            If FindFieldInf("JobCode") = 1 Then
                If DJobCode.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Name
        If InputCheck = 0 Then
            If FindFieldInf("DepoName") = 1 Then
                If DDepoName.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Code
        If InputCheck = 0 Then
            If FindFieldInf("DepoCode") = 1 Then
                If DDepoCode.Text = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.Text = "" Then InputCheck = 1
            End If
        End If
        '�����N�X
        If InputCheck = 0 Then
            If FindFieldInf("DivisionCode") = 1 Then
                If DDivisionCode.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�Z���
        If InputCheck = 0 Then
            If FindFieldInf("OverTimeDate") = 1 Then
                If DOverTimeDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�Z�O
        If InputCheck = 0 Then
            If FindFieldInf("DateType") = 1 Then
                If DDateType.Text = "" Then InputCheck = 1
            End If
        End If
        '���ݦ~��
        If InputCheck = 0 Then
            If FindFieldInf("SalaryYM") = 1 Then
                If DSalaryYM.Text = "" Then InputCheck = 1
            End If
        End If
        '�ե�
        If InputCheck = 0 Then
            If FindFieldInf("CVacation") = 1 Then
                If DCVacation.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�뭹
        If InputCheck = 0 Then
            If FindFieldInf("Food") = 1 Then
                If DFood.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��q
        If InputCheck = 0 Then
            If FindFieldInf("Traffic") = 1 Then
                If DTraffic.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("BStartH") = 1 Then
                If DBStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("BStartM") = 1 Then
                If DBStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�פ�-��
        If InputCheck = 0 Then
            If FindFieldInf("BEndH") = 1 Then
                If DBEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�פ�-��
        If InputCheck = 0 Then
            If FindFieldInf("BEndM") = 1 Then
                If DBEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�p��-��
        If InputCheck = 0 Then
            If FindFieldInf("BH") = 1 Then
                If DBH.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w�p��-��
        If InputCheck = 0 Then
            If FindFieldInf("BM") = 1 Then
                If DBM.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڶ}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("AStartH") = 1 Then
                If DAStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڶ}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("AStartM") = 1 Then
                If DAStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڲפ�-��
        If InputCheck = 0 Then
            If FindFieldInf("AEndH") = 1 Then
                If DAEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڲפ�-��
        If InputCheck = 0 Then
            If FindFieldInf("AEndM") = 1 Then
                If DAEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڭp��-��
        If InputCheck = 0 Then
            If FindFieldInf("AH") = 1 Then
                If DAH.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڭp��-��
        If InputCheck = 0 Then
            If FindFieldInf("AM") = 1 Then
                If DAM.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����2��-��
        If InputCheck = 0 Then
            If FindFieldInf("FAH1") = 1 Then
                If DFAH1.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����2��-��
        If InputCheck = 0 Then
            If FindFieldInf("FAM1") = 1 Then
                If DFAM1.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����2�~-��
        If InputCheck = 0 Then
            If FindFieldInf("FAH2") = 1 Then
                If DFAH2.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����2�~-��
        If InputCheck = 0 Then
            If FindFieldInf("FAM2") = 1 Then
                If DFAM2.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����-��
        If InputCheck = 0 Then
            If FindFieldInf("FBH") = 1 Then
                If DFBH.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w����-��
        If InputCheck = 0 Then
            If FindFieldInf("FBM") = 1 Then
                If DFBM.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w��w����-��
        If InputCheck = 0 Then
            If FindFieldInf("FCH") = 1 Then
                If DFCH.Text = "" Then InputCheck = 1
            End If
        End If
        '�֩w��w����-��
        If InputCheck = 0 Then
            If FindFieldInf("FCM") = 1 Then
                If DFCM.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�Z�z��
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
