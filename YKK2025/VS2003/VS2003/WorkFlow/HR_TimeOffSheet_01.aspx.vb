Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_TimeOffSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DTimeOffAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDieType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DEvidence As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalary As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DBStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DADays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVacation As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DOTHours As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTNo1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTNo2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTNo3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTNo5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTNo4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents BVARecord As System.Web.UI.WebControls.Button
    Protected WithEvents BOTRecord As System.Web.UI.WebControls.Button
    Protected WithEvents LOTNo1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo5 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAfter As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BAEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents BAStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DAEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents BBDays As System.Web.UI.WebControls.Button
    Protected WithEvents BADays As System.Web.UI.WebControls.Button
    Protected WithEvents BOverTime As System.Web.UI.WebControls.Button
    Protected WithEvents DJobAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTimeOffSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DInDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVDaysBlank As System.Web.UI.WebControls.TextBox

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

        Response.Cookies("PGM").Value = "HR_TimeOffSheet_01.aspx"
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
    End Sub
    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("TimeOffFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_TimeOffSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_TimeOffSheet")
        If DBDataSet1.Tables("F_TimeOffSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Date")                   '�ӽФ��
            DSalaryYM.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("SalaryYM")           '���ݦ~��

            SetFieldData("Name", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Name"))          '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DivisionCode")   '�����N�X

            SetFieldData("After", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("After"))        '�ƫe,�ƫ�
            DJobAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobAgent")           '¾�ȥN�z�H
            DTimeOffAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("TimeOffAgent")   '�N�а��H
            SetFieldData("Vacation", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VacationCode") + _
                                     ":" + _
                                     DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Vacation"))  '���O
            DEvidence.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Evidence")           '����
            DSalary.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Salary")               '�~��
            SetFieldData("DieType", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DieType"))    '�ల�O
            DVDays.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VDays").ToString        '�i�ФѼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1") <> "" Then                 '�[�ZNo1 
                LOTNo1.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
                LOTNo1.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
            Else
                LOTNo1.Visible = False
                DOTHours1.Visible = False
            End If
            DOTHours1.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours1"), 1)  '�[�ZNo1-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2") <> "" Then                 '�[�ZNo2 
                LOTNo2.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
                LOTNo2.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
            Else
                LOTNo2.Visible = False
                DOTHours2.Visible = False
            End If
            DOTHours2.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours2"), 1)  '�[�ZNo2-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3") <> "" Then                 '�[�ZNo3
                LOTNo3.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
                LOTNo3.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
            Else
                LOTNo3.Visible = False
                DOTHours3.Visible = False
            End If
            DOTHours3.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours3"), 1)  '�[�ZNo3-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4") <> "" Then                 '�[�ZNo4
                LOTNo4.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
                LOTNo4.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
            Else
                LOTNo4.Visible = False
                DOTHours4.Visible = False
            End If
            DOTHours4.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours4"), 1)  '�[�ZNo4-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5") <> "" Then                 '�[�ZNo5
                LOTNo5.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
                LOTNo5.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
            Else
                LOTNo5.Visible = False
                DOTHours5.Visible = False
            End If
            DOTHours5.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours5"), 1)  '�[�ZNo5-�ɼ�

            DOTHours.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours"), 1)    '�[�Z�`�ɼ�

            DBStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartDate")       '�w�w�}�l���
            SetFieldData("BStartH", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartH").ToString)   '�w�w�}�l��
            DBEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndDate")           '�w�w�������
            SetFieldData("BEndH", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndH").ToString)   '�w�w������
            DBDays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BDays"), 1)    '�w�w��

            DAStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartDate")       '��ڶ}�l���
            SetFieldData("AStartH", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartH").ToString)   '��ڶ}�l��
            DAEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndDate")           '��ڵ������
            SetFieldData("AEndH", DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndH").ToString)   '��ڵ�����
            DADays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("ADays"), 1)    '��ڤ�

            DFReason.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("FReason")             '�а��z��

            '���o�~��_���
            DBDataSet1.Clear()
            SQL = "Select StartDate As InDate From M_Emp "
            SQL = SQL & " Where Com_Code = '" & DDepoCode.Text & "'"
            SQL = SQL & "   And ID       = '" & DEmpID.Text & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_EMP")
            If DBDataSet1.Tables("M_EMP").Rows.Count > 0 Then
                DInDate.Text = DBDataSet1.Tables("M_EMP").Rows(0).Item("InDate")
            End If

            '������
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "T_WaitHandle")
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
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet9, "DecideHistory")
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
                DTimeOffSheet1.Visible = True   '���Sheet-1
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
                    Top = 712
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 600
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

                '�s�����---�ݦA�ק�
                LOTNo1.Visible = True          '�[�ZNo1
                LOTNo2.Visible = True          '�[�ZNo2
                LOTNo3.Visible = True          '�[�ZNo3
                LOTNo4.Visible = True          '�[�ZNo4
                LOTNo5.Visible = True          '�[�ZNo5
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
            Top = 520
            'Sheet����
            DDescSheet.Visible = False  '����Sheet
            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False    '����
            DReasonCode.Visible = False    '����z�ѥN�X
            DReason.Visible = False        '����z��
            DReasonDesc.Visible = False    '�����L����
            DHistoryLabel.Visible = False  '�֩w�i��
            '�s�����---�ݦA�ק�
            LOTNo1.Visible = False         '�[�ZNo1
            LOTNo2.Visible = False         '�[�ZNo2
            LOTNo3.Visible = False         '�[�ZNo3
            LOTNo4.Visible = False         '�[�ZNo4
            LOTNo5.Visible = False         '�[�ZNo5
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
        'BBStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  '�а��w�w�}�l���
        'BBEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBEndDate', '');"                     '�а��w�w�������
        'BAStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  '�а���ڶ}�l���
        'BAEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAEndDate', '');"                     '�а���ڵ������
        '
        BBStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  '�а��w�w�}�l���
        BBEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBEndDate', '');"                     '�а��w�w�������
        BAStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  '�а���ڶ}�l���
        BAEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAEndDate', '');"                     '�а���ڵ������
        'Modify-End

        BCardTime.Attributes("onclick") = "ShowCardTime();"    '��d�O��
        BVARecord.Attributes("onclick") = "ShowVacation();"    '�а��O��
        BOTRecord.Attributes("onclick") = "ShowOverTime();"    '�[�Z�ե�O��
        BOverTime.Attributes("onclick") = "ShowAOverTime();"   '�ե�O��
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
                    Top = 600
                Else
                    If DDelay.Visible = True Then
                        Top = 712
                    Else
                        Top = 600
                    End If
                End If
            End If
        Else
            Top = 520
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
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '��渹�X,�u�{���d���X,���W,����ݩ�
        SetFieldAttribute(pPost)                                            '���U����ݩʤ�����J�ˬd���]�w
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
        '���o�~��_���
        DBDataSet1.Clear()
        SQL = "Select StartDate As InDate From M_Emp "
        SQL = SQL & " Where Com_Code = '" & wDepoCode & "'"
        SQL = SQL & "   And ID       = '" & wEmpID & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_EMP")
        If DBDataSet1.Tables("M_EMP").Rows.Count > 0 Then
            DInDate.Text = DBDataSet1.Tables("M_EMP").Rows(0).Item("InDate")
        End If
        OleDbConnection1.Close()
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
        '�ƫe,�ƫ�
        Select Case FindFieldInf("After")
            Case 0  '���
                DAfter.BackColor = Color.LightGray
                DAfter.Visible = True
            Case 1  '�ק�+�ˬd
                DAfter.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAfterRqd", "DAfter", "���`�G�ݿ�J�ƫe")
                DAfter.Visible = True
            Case 2  '�ק�
                DAfter.BackColor = Color.Yellow
                DAfter.Visible = True
            Case Else   '����
                DAfter.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("After", "ZZZZZZ")
        '¾�ȥN�z�H
        Select Case FindFieldInf("JobAgent")
            Case 0  '���
                DJobAgent.BackColor = Color.LightGray
                DJobAgent.Visible = True
            Case 1  '�ק�+�ˬd
                DJobAgent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobAgentRqd", "DJobAgent", "���`�G�ݿ�J¾�ȥN�z�H")
                DJobAgent.Visible = True
            Case 2  '�ק�
                DJobAgent.BackColor = Color.Yellow
                DJobAgent.Visible = True
            Case Else   '����
                DJobAgent.Visible = False
        End Select
        If pPost = "New" Then DJobAgent.Text = ""
        '�N�а��H
        Select Case FindFieldInf("TimeOffAgent")
            Case 0  '���
                DTimeOffAgent.BackColor = Color.LightGray
                DTimeOffAgent.Visible = True
            Case 1  '�ק�+�ˬd
                DTimeOffAgent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTimeOffAgentRqd", "DTimeOffAgent", "���`�G�ݿ�J�N�а��H")
                DTimeOffAgent.Visible = True
            Case 2  '�ק�
                DTimeOffAgent.BackColor = Color.Yellow
                DTimeOffAgent.Visible = True
            Case Else   '����
                DTimeOffAgent.Visible = False
        End Select
        If pPost = "New" Then DTimeOffAgent.Text = ""
        '���O
        Select Case FindFieldInf("Vacation")
            Case 0  '���
                DVacation.BackColor = Color.LightGray
                DVacation.Visible = True
            Case 1  '�ק�+�ˬd
                DVacation.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVacationRqd", "DVacation", "���`�G�ݿ�J���O")
                DVacation.Visible = True
            Case 2  '�ק�
                DVacation.BackColor = Color.Yellow
                DVacation.Visible = True
            Case Else   '����
                DVacation.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Vacation", "ZZZZZZ")
        '����
        Select Case FindFieldInf("Evidence")
            Case 0  '���
                DEvidence.BackColor = Color.LightGray
                DEvidence.Visible = True
            Case 1  '�ק�+�ˬd
                DEvidence.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEvidenceRqd", "DEvidence", "���`�G�ݿ�J����")
                DEvidence.Visible = True
            Case 2  '�ק�
                DEvidence.BackColor = Color.Yellow
                DEvidence.Visible = True
            Case Else   '����
                DEvidence.Visible = False
        End Select
        If pPost = "New" Then DEvidence.Text = ""
        '�~��
        Select Case FindFieldInf("Salary")
            Case 0  '���
                DSalary.BackColor = Color.LightGray
                DSalary.Visible = True
            Case 1  '�ק�+�ˬd
                DSalary.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSalaryRqd", "DSalary", "���`�G�ݿ�J�~��")
                DSalary.Visible = True
            Case 2  '�ק�
                DSalary.BackColor = Color.Yellow
                DSalary.Visible = True
            Case Else   '����
                DSalary.Visible = False
        End Select
        If pPost = "New" Then DSalary.Text = ""
        '�ల�O
        Select Case FindFieldInf("DieType")
            Case 0  '���
                DDieType.BackColor = Color.LightGray
                DDieType.Visible = True
            Case 1  '�ק�+�ˬd
                DDieType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDieTypeRqd", "DDieType", "���`�G�ݿ�J�ల�O")
                DDieType.Visible = True
            Case 2  '�ק�
                DDieType.BackColor = Color.Yellow
                DDieType.Visible = True
            Case Else   '����
                DDieType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DieType", "ZZZZZZ")
        '�i�ФѼ�
        Select Case FindFieldInf("VDays")
            Case 0  '���
                DVDays.BackColor = Color.LightGray
                DVDays.Visible = True
            Case 1  '�ק�+�ˬd
                DVDays.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVDaysRqd", "DVDays", "���`�G�ݿ�J�i�ФѼ�")
                DVDays.Visible = True
            Case 2  '�ק�
                DVDays.BackColor = Color.Yellow
                DVDays.Visible = True
            Case Else   '����
                DVDays.Visible = False
        End Select
        If pPost = "New" Then DVDays.Text = "0"
        '�[�ZNo1
        Select Case FindFieldInf("OTNo1")
            Case 0  '���
                DOTNo1.BackColor = Color.LightGray
                DOTNo1.Visible = True
            Case 1  '�ק�+�ˬd
                DOTNo1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTNo1Rqd", "DOTNo1", "���`�G�ݿ�J�[�ZNo.")
                DOTNo1.Visible = True
            Case 2  '�ק�
                DOTNo1.BackColor = Color.Yellow
                DOTNo1.Visible = True
            Case Else   '����
                DOTNo1.Visible = False
        End Select
        If pPost = "New" Then DOTNo1.Text = ""
        '�[�ZNo1-�ɼ�
        Select Case FindFieldInf("OTHours1")
            Case 0  '���
                DOTHours1.BackColor = Color.LightGray
                DOTHours1.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHours1Rqd", "DOTHours1", "���`�G�ݿ�J�[�Z�ɼ�")
                DOTHours1.Visible = True
            Case 2  '�ק�
                DOTHours1.BackColor = Color.Yellow
                DOTHours1.Visible = True
            Case Else   '����
                DOTHours1.Visible = False
        End Select
        If pPost = "New" Then DOTHours1.Text = "0"
        '�[�ZNo2
        Select Case FindFieldInf("OTNo2")
            Case 0  '���
                DOTNo2.BackColor = Color.LightGray
                DOTNo2.Visible = True
            Case 1  '�ק�+�ˬd
                DOTNo2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTNo2Rqd", "DOTNo2", "���`�G�ݿ�J�[�ZNo.")
                DOTNo2.Visible = True
            Case 2  '�ק�
                DOTNo2.BackColor = Color.Yellow
                DOTNo2.Visible = True
            Case Else   '����
                DOTNo2.Visible = False
        End Select
        If pPost = "New" Then DOTNo2.Text = ""
        '�[�ZNo2-�ɼ�
        Select Case FindFieldInf("OTHours2")
            Case 0  '���
                DOTHours2.BackColor = Color.LightGray
                DOTHours2.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHours2Rqd", "DOTHours2", "���`�G�ݿ�J�[�Z�ɼ�")
                DOTHours2.Visible = True
            Case 2  '�ק�
                DOTHours2.BackColor = Color.Yellow
                DOTHours2.Visible = True
            Case Else   '����
                DOTHours2.Visible = False
        End Select
        If pPost = "New" Then DOTHours2.Text = "0"
        '�[�ZNo3
        Select Case FindFieldInf("OTNo3")
            Case 0  '���
                DOTNo3.BackColor = Color.LightGray
                DOTNo3.Visible = True
            Case 1  '�ק�+�ˬd
                DOTNo3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTNo3Rqd", "DOTNo3", "���`�G�ݿ�J�[�ZNo.")
                DOTNo3.Visible = True
            Case 2  '�ק�
                DOTNo3.BackColor = Color.Yellow
                DOTNo3.Visible = True
            Case Else   '����
                DOTNo3.Visible = False
        End Select
        If pPost = "New" Then DOTNo3.Text = ""
        '�[�ZNo3-�ɼ�
        Select Case FindFieldInf("OTHours3")
            Case 0  '���
                DOTHours3.BackColor = Color.LightGray
                DOTHours3.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHours3Rqd", "DOTHours3", "���`�G�ݿ�J�[�Z�ɼ�")
                DOTHours3.Visible = True
            Case 2  '�ק�
                DOTHours3.BackColor = Color.Yellow
                DOTHours3.Visible = True
            Case Else   '����
                DOTHours3.Visible = False
        End Select
        If pPost = "New" Then DOTHours3.Text = "0"
        '�[�ZNo4
        Select Case FindFieldInf("OTNo4")
            Case 0  '���
                DOTNo4.BackColor = Color.LightGray
                DOTNo4.Visible = True
            Case 1  '�ק�+�ˬd
                DOTNo4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTNo4Rqd", "DOTNo4", "���`�G�ݿ�J�[�ZNo.")
                DOTNo4.Visible = True
            Case 2  '�ק�
                DOTNo4.BackColor = Color.Yellow
                DOTNo4.Visible = True
            Case Else   '����
                DOTNo4.Visible = False
        End Select
        If pPost = "New" Then DOTNo4.Text = ""
        '�[�ZNo4-�ɼ�
        Select Case FindFieldInf("OTHours4")
            Case 0  '���
                DOTHours4.BackColor = Color.LightGray
                DOTHours4.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHours4Rqd", "DOTHours4", "���`�G�ݿ�J�[�Z�ɼ�")
                DOTHours4.Visible = True
            Case 2  '�ק�
                DOTHours4.BackColor = Color.Yellow
                DOTHours4.Visible = True
            Case Else   '����
                DOTHours4.Visible = False
        End Select
        If pPost = "New" Then DOTHours4.Text = "0"
        '�[�ZNo5
        Select Case FindFieldInf("OTNo5")
            Case 0  '���
                DOTNo5.BackColor = Color.LightGray
                DOTNo5.Visible = True
            Case 1  '�ק�+�ˬd
                DOTNo5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTNo5Rqd", "DOTNo5", "���`�G�ݿ�J�[�ZNo.")
                DOTNo5.Visible = True
            Case 2  '�ק�
                DOTNo5.BackColor = Color.Yellow
                DOTNo5.Visible = True
            Case Else   '����
                DOTNo5.Visible = False
        End Select
        If pPost = "New" Then DOTNo5.Text = ""
        '�[�ZNo5-�ɼ�
        Select Case FindFieldInf("OTHours5")
            Case 0  '���
                DOTHours5.BackColor = Color.LightGray
                DOTHours5.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHours5Rqd", "DOTHours5", "���`�G�ݿ�J�[�Z�ɼ�")
                DOTHours5.Visible = True
            Case 2  '�ק�
                DOTHours5.BackColor = Color.Yellow
                DOTHours5.Visible = True
            Case Else   '����
                DOTHours5.Visible = False
        End Select
        If pPost = "New" Then DOTHours5.Text = "0"
        '�[�Z�`�ɼ�
        Select Case FindFieldInf("OTHours")
            Case 0  '���
                DOTHours.BackColor = Color.LightGray
                DOTHours.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHours.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOTHoursRqd", "DOTHours", "���`�G�ݿ�J�[�Z�`�ɼ�")
                DOTHours.Visible = True
            Case 2  '�ק�
                DOTHours.BackColor = Color.Yellow
                DOTHours.Visible = True
            Case Else   '����
                DOTHours.Visible = False
        End Select
        If pPost = "New" Then DOTHours.Text = "0"
        '�w�w�}�l���
        Select Case FindFieldInf("BStartDate")
            Case 0  '���
                DBStartDate.BackColor = Color.LightGray
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
                BBStartDate.Visible = False
            Case 1  '�ק�+�ˬd
                DBStartDate.BackColor = Color.GreenYellow
                DBStartDate.ReadOnly = True
                ShowRequiredFieldValidator("DBStartDateRqd", "DBStartDate", "���`�G�ݿ�J�w�w�}�l���")
                DBStartDate.Visible = True
                BBStartDate.Visible = True
            Case 2  '�ק�
                DBStartDate.BackColor = Color.Yellow
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
                BBStartDate.Visible = True
            Case Else   '����
                DBStartDate.Visible = False
                BBStartDate.Visible = False
        End Select
        If pPost = "New" Then DBStartDate.Text = CStr(DateTime.Now.Today)
        '�w�w�}�l-��
        Select Case FindFieldInf("BStartH")
            Case 0  '���
                DBStartH.BackColor = Color.LightGray
                DBStartH.Visible = True
                BBDays.Visible = False
            Case 1  '�ק�+�ˬd
                DBStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartHRqd", "DBStartH", "���`�G�ݿ�J�w�w�}�l-��")
                DBStartH.Visible = True
                BBDays.Visible = True
            Case 2  '�ק�
                DBStartH.BackColor = Color.Yellow
                DBStartH.Visible = True
                BBDays.Visible = True
            Case Else   '����
                DBStartH.Visible = False
                BBDays.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartH", "ZZZZZZ")
        '�w�w�������
        Select Case FindFieldInf("BEndDate")
            Case 0  '���
                DBEndDate.BackColor = Color.LightGray
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
                BBEndDate.Visible = False
            Case 1  '�ק�+�ˬd
                DBEndDate.BackColor = Color.GreenYellow
                DBEndDate.ReadOnly = True
                ShowRequiredFieldValidator("DBEndDateRqd", "DBEndDate", "���`�G�ݿ�J�w�w�������")
                DBEndDate.Visible = True
                BBEndDate.Visible = True
            Case 2  '�ק�
                DBEndDate.BackColor = Color.Yellow
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
                BBEndDate.Visible = True
            Case Else   '����
                DBEndDate.Visible = False
                BBEndDate.Visible = False
        End Select
        If pPost = "New" Then DBEndDate.Text = CStr(DateTime.Now.Today)
        '�w�w����-��
        Select Case FindFieldInf("BEndH")
            Case 0  '���
                DBEndH.BackColor = Color.LightGray
                DBEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DBEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndHRqd", "DBEndH", "���`�G�ݿ�J�w�w����-��")
                DBEndH.Visible = True
            Case 2  '�ק�
                DBEndH.BackColor = Color.Yellow
                DBEndH.Visible = True
            Case Else   '����
                DBEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndH", "ZZZZZZ")
        '�w�w�Ѽ�
        Select Case FindFieldInf("BDays")
            Case 0  '���
                DBDays.BackColor = Color.LightGray
                DBDays.Visible = True
            Case 1  '�ק�+�ˬd
                DBDays.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBDaysRqd", "DBDays", "���`�G�ݿ�J�w�w�Ѽ�")
                DBDays.Visible = True
            Case 2  '�ק�
                DBDays.BackColor = Color.Yellow
                DBDays.Visible = True
            Case Else   '����
                DBDays.Visible = False
        End Select
        If pPost = "New" Then DBDays.Text = "0"
        '��ڶ}�l���
        Select Case FindFieldInf("AStartDate")
            Case 0  '���
                DAStartDate.BackColor = Color.LightGray
                DAStartDate.ReadOnly = True
                DAStartDate.Visible = True
                BAStartDate.Visible = False
            Case 1  '�ק�+�ˬd
                DAStartDate.BackColor = Color.GreenYellow
                DAStartDate.ReadOnly = True
                ShowRequiredFieldValidator("DAStartDateRqd", "DAStartDate", "���`�G�ݿ�J��ڶ}�l���")
                DAStartDate.Visible = True
                BAStartDate.Visible = True
            Case 2  '�ק�
                DAStartDate.BackColor = Color.Yellow
                DAStartDate.ReadOnly = True
                DAStartDate.Visible = True
                BAStartDate.Visible = True
            Case Else   '����
                DAStartDate.Visible = False
                BAStartDate.Visible = False
        End Select
        If pPost = "New" Then DAStartDate.Text = CStr(DateTime.Now.Today)
        '��ڶ}�l-��
        Select Case FindFieldInf("AStartH")
            Case 0  '���
                DAStartH.BackColor = Color.LightGray
                DAStartH.Visible = True
                BADays.Visible = False
            Case 1  '�ק�+�ˬd
                DAStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartHRqd", "DAStartH", "���`�G�ݿ�J��ڶ}�l-��")
                DAStartH.Visible = True
                BADays.Visible = True
            Case 2  '�ק�
                DAStartH.BackColor = Color.Yellow
                DAStartH.Visible = True
                BADays.Visible = True
            Case Else   '����
                DAStartH.Visible = False
                BADays.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartH", "ZZZZZZ")
        '��ڵ������
        Select Case FindFieldInf("AEndDate")
            Case 0  '���
                DAEndDate.BackColor = Color.LightGray
                DAEndDate.ReadOnly = True
                DAEndDate.Visible = True
                BAEndDate.Visible = False
            Case 1  '�ק�+�ˬd
                DAEndDate.BackColor = Color.GreenYellow
                DAEndDate.ReadOnly = True
                ShowRequiredFieldValidator("DAEndDateRqd", "DAEndDate", "���`�G�ݿ�J��ڵ������")
                DAEndDate.Visible = True
                BAEndDate.Visible = True
            Case 2  '�ק�
                DAEndDate.BackColor = Color.Yellow
                DAEndDate.ReadOnly = True
                DAEndDate.Visible = True
                BAEndDate.Visible = True
            Case Else   '����
                DAEndDate.Visible = False
                BAEndDate.Visible = False
        End Select
        If pPost = "New" Then DAEndDate.Text = CStr(DateTime.Now.Today)
        '��ڵ���-��
        Select Case FindFieldInf("AEndH")
            Case 0  '���
                DAEndH.BackColor = Color.LightGray
                DAEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DAEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndHRqd", "DAEndH", "���`�G�ݿ�J��ڵ���-��")
                DAEndH.Visible = True
            Case 2  '�ק�
                DAEndH.BackColor = Color.Yellow
                DAEndH.Visible = True
            Case Else   '����
                DAEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndH", "ZZZZZZ")
        '��ڤѼ�
        Select Case FindFieldInf("ADays")
            Case 0  '���
                DADays.BackColor = Color.LightGray
                DADays.Visible = True
            Case 1  '�ק�+�ˬd
                DADays.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DADaysRqd", "DADays", "���`�G�ݿ�J��ڤѼ�")
                DADays.Visible = True
            Case 2  '�ק�
                DADays.BackColor = Color.Yellow
                DADays.Visible = True
            Case Else   '����
                DADays.Visible = False
        End Select
        If pPost = "New" Then DADays.Text = "0"
        '�а��z��
        Select Case FindFieldInf("FReason")
            Case 0  '���
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '�ק�+�ˬd
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "���`�G�ݿ�J�а��z��")
                DFReason.Visible = True
            Case 2  '�ק�
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '����
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""
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
                    SQL = SQL + "  And FormNo = '001002' "
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
        '�ƫe,�ƫ�
        If pFieldName = "After" Then
            DAfter.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAfter.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='AFTER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAfter.Items.Add(ListItem1)
                Next
            End If
        End If
        '���O
        If pFieldName = "Vacation" Then
            DVacation.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DVacation.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='VACATION' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DVacation.Items.Add(ListItem1)
                Next
            End If
            '�ల�O
            If Left(DVacation.SelectedValue, 1) = "X" Then
                DDieType.Enabled = True
            Else
                DDieType.Enabled = False
            End If
        End If
        '�ల
        If pFieldName = "DieType" Then
            DDieType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDieType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='DIETYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDieType.Items.Add(ListItem1)
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
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
        '�w�w����-��
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
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
        '��ڶ}�l-��
        If pFieldName = "AStartH" Then
            DAStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
        '��ڵ���-��
        If pFieldName = "AEndH" Then
            DAEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1002' and DKey='HOUR' Order by Data "
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
        Dim BaseDate As Date
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        '--------------------------------------------------------------------------
        '--  �ˬd���U�����
        '--------------------------------------------------------------------------
        '�ե��`�ɼ�
        If ErrCode = 0 Then
            If Left(DVacation.SelectedValue, 1) = "9" Then
                ' 2012/1 ~ 2012/3 �Ȱ��ϥ�
                'If DOTHours.Text = "0" Then ErrCode = 9050
            End If
        End If
        '�а��Ѽ�
        If ErrCode = 0 Then
            If DBDays.Text = "0" Or DADays.Text = "0" Then ErrCode = 9051
        End If
        '¾�ȥN�z�H
        If ErrCode = 0 Then
            If DAfter.SelectedValue = "1.�ƫe" Then
                If DJobAgent.Text = "" Then ErrCode = 9052
            End If
        End If
        '�N�а��H
        If ErrCode = 0 Then
            If DAfter.SelectedValue = "2.�ƫ�" Then
                If DTimeOffAgent.Text = "" Then ErrCode = 9053
            End If
        End If
        '���O�ˬd
        If ErrCode = 0 Then
            If DVacation.SelectedValue = "" Then ErrCode = 9054
        End If
        '�_������ˬd
        If ErrCode = 0 Then
            '�w�w
            If CDate(DBEndDate.Text) < CDate(DBStartDate.Text) Then ErrCode = 8010
            If CDate(DBEndDate.Text) = CDate(DBStartDate.Text) Then
                If DBEndH.SelectedValue <= DBStartH.SelectedValue Then ErrCode = 8010
            End If
            '���
            If CDate(DAEndDate.Text) < CDate(DAStartDate.Text) Then ErrCode = 8010
            If CDate(DAEndDate.Text) = CDate(DAStartDate.Text) Then
                If DAEndH.SelectedValue <= DAStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '�_������P���ˬd
        If ErrCode = 0 Then
            '�w�w
            If CStr(Year(CDate(DBStartDate.Text))) + "/" + CStr(Month(CDate(DBStartDate.Text))) <> CStr(Year(CDate(DBEndDate.Text))) + "/" + CStr(Month(CDate(DBEndDate.Text))) Then
                ErrCode = 8020
            End If
            '���
            If CStr(Year(CDate(DAStartDate.Text))) + "/" + CStr(Month(CDate(DAStartDate.Text))) <> CStr(Year(CDate(DAEndDate.Text))) + "/" + CStr(Month(CDate(DAEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '�а��O�_���b����
        If ErrCode = 0 Then
            '�w�w
            If CDbl(DBDays.Text) - Fix(DBDays.Text) <> 0.5 And CDbl(DBDays.Text) - Fix(DBDays.Text) <> 0 Then
                ErrCode = 8030
            End If
            '���
            If CDbl(DADays.Text) - Fix(DADays.Text) <> 0.5 And CDbl(DADays.Text) - Fix(DADays.Text) <> 0 Then
                ErrCode = 8030
            End If
        End If
        '��Ƶ��⵲�G�ˬd
        If ErrCode = 0 Then
            '#�w�w
            Dim Hours As Double = 0
            Dim VHours As Double = 0
            Dim MHours As Double = 0
            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

            '2008/7�s�Ұ�k�����ץ�-�������t����
            If Left(DVacation.SelectedValue, 1) <> "O" And Left(DVacation.SelectedValue, 1) <> "P" And Left(DVacation.SelectedValue, 1) <> "Q" And Left(DVacation.SelectedValue, 1) <> "W" Then
                OleDbConnection1.Open()
                SQL = "Select Count(*) As Days From M_Vacation "
                SQL = SQL + "Where Active = '1' "
                'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                'SQL = SQL + "  and Depo = '" + wDepo + "' "
                '
                SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
                'Modify-End
                SQL = SQL + "  and YMD  >= '" + CDate(DBStartDate.Text) + "' "
                SQL = SQL + "  and YMD  <= '" + CDate(DBEndDate.Text) + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
                OleDbConnection1.Close()
            End If
            '�а��ɶ�
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '�а��ɼƮt�ɶ�
            If DBEndH.SelectedValue >= DBStartH.SelectedValue Then
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '�p��ɶ�
            If DBDays.Text <> FormatNumber((VHours + Hours - MHours) / 8, 1) Then
                ErrCode = 8040
            End If
            DBDataSet1.Clear()
            '----------------------------------------------------------------------------------------------
            '#���
            Hours = 0
            VHours = 0
            MHours = 0

            '2008/7�s�Ұ�k�����ץ�-�������t����
            If Left(DVacation.SelectedValue, 1) <> "O" And Left(DVacation.SelectedValue, 1) <> "P" And Left(DVacation.SelectedValue, 1) <> "Q" And Left(DVacation.SelectedValue, 1) <> "W" Then
                OleDbConnection1.Open()
                SQL = "Select Count(*) As Days From M_Vacation "
                SQL = SQL + "Where Active = '1' "
                'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                'SQL = SQL + "  and Depo = '" + wDepo + "' "
                '
                SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
                'Modify-End
                SQL = SQL + "  and YMD  >= '" + CDate(DAStartDate.Text) + "' "
                SQL = SQL + "  and YMD  <= '" + CDate(DAEndDate.Text) + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
                OleDbConnection1.Close()
            End If
            '�а��ɶ�
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '�а��ɼƮt�ɶ�
            If DAEndH.SelectedValue >= DAStartH.SelectedValue Then
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '�p��ɶ�
            If DADays.Text <> FormatNumber((VHours + Hours - MHours) / 8, 1) Then
                ErrCode = 8040
            End If
        End If
        '�~��_����ˬd
        If ErrCode = 0 Then
            If Left(DVacation.SelectedValue, 1) = "A" Then
                Dim SQL As String
                Dim DBDataSet1 As New DataSet
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                DBDataSet1.Clear()
                OleDbConnection1.Open()
                SQL = "Select StartDate As ArriveDate From M_Emp "
                SQL = SQL + "Where Com_Code = '" + DDepoCode.Text + "' "
                SQL = SQL + "  and ID  = '" + DEmpID.Text + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
                    BaseDate = CDate(Mid(DSalaryYM.Text, 1, 4) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Month) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Day))
                    If CDate(DAStartDate.Text) < BaseDate And CDate(DAEndDate.Text) < BaseDate Then
                        '�_���L��~��_���
                    Else
                        If CDate(DAStartDate.Text) >= BaseDate And CDate(DAEndDate.Text) >= BaseDate Then
                            '�_���L��~��_���
                        Else
                            '�_������~��_���
                            ErrCode = 8050
                        End If
                    End If
                Else
                    ErrCode = 8050
                End If
                OleDbConnection1.Close()
            End If
        End If
        '--------------------------------------------------------------------------
        'Check����z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If
        '�ˬd�e�U��No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("001002", wFormSno, wStep, DNo.Text)
                '��渹�X, ���y����, �u�{, �e�U��No
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
                Dim SQL As String

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '�P�_�O�_�_��
                        '���o���y����
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno)
                        '��渹�X, ���y����
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
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
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
            If ErrCode = 9050 Then Message = "�ե�ɶ��ݶ�g,�нT�{!"
            If ErrCode = 9051 Then Message = "�а��ɶ��ݶ�g,�нT�{!"
            If ErrCode = 9052 Then Message = "�ƫe�а���¾�ȥN�z�H�ݶ�g,�нT�{!"
            If ErrCode = 9053 Then Message = "�ƫ�а��ɥN�а��H�ݶ�g,�нT�{!"
            If ErrCode = 9054 Then Message = "���O�ݶ�g,�нT�{!"
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
            If ErrCode = 8030 Then Message = "�а����ݬO�b��,�нT�{!"
            If ErrCode = 8040 Then Message = "�а�����ɶ����ŦX,�нT�{!"
            If ErrCode = 8050 Then Message = "�а��_����������~�~��_���[" + BaseDate + "],�нT�{!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("TimeOffFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_TimeOffSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, SalaryYM, Name, EmpID, "                                '6~9
        SQl = SQl + "JobTitle, JobCode, DepoName, DepoCode, Division, "            '10~14
        SQl = SQl + "DivisionCode, After, JobAgent, TimeOffAgent, "                '15~18
        SQl = SQl + "VacationCode, Vacation, Evidence, Salary, DieType, VDays, "   '19~24
        SQl = SQl + "OTNo1, OTHours1, OTNo2, OTHours2, OTNo3, OTHours3, "          '25~30
        SQl = SQl + "OTNo4, OTHours4, OTNo5, OTHours5, OTHours, "                  '31~35
        SQl = SQl + "BStartDate, BStartH, BEndDate, BEndH, BDays, "                '36~40
        SQl = SQl + "AStartDate, AStartH, AEndDate, AEndH, ADays, "                '41~45
        SQl = SQl + "FReason, "                                                    '46
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '47~50
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '001002', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~9
        SQl = SQl + " '" + DDate.Text + "', "               '�ӽФ��
        SQl = SQl + " '" + DSalaryYM.Text + "', "           '���ݦ~��
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', " '�m�W
        SQl = SQl + " '" + DEmpID.Text + "', "              'EMP-ID
        '10~18
        SQl = SQl + " N'" + DJobTitle.Text + "', "          '¾��
        SQl = SQl + " '" + DJobCode.Text + "', "            '¾�٥N�X
        SQl = SQl + " N'" + DDepoName.Text + "', "          '���q�O
        SQl = SQl + " '" + DDepoCode.Text + "', "           '���q�OCode
        SQl = SQl + " N'" + DDivision.Text + "', "          '����
        SQl = SQl + " '" + DDivisionCode.Text + "', "       '�����N�X
        SQl = SQl + " N'" + DAfter.SelectedValue + "', "    '�ƫe,��
        SQl = SQl + " N'" + DJobAgent.Text + "', "          '¾�ȥN�z�H
        SQl = SQl + " N'" + DTimeOffAgent.Text + "', "      '�N�а���
        '19~24
        SQl = SQl + " '" + Left(DVacation.SelectedValue, 1) + "', "    '���O�N�X
        SQl = SQl + " N'" + Mid(DVacation.SelectedValue, 3) + "', " '���O
        SQl = SQl + " N'" + DEvidence.Text + "', "          '����
        SQl = SQl + " N'" + DSalary.Text + "', "            '�~��
        SQl = SQl + " N'" + DDieType.SelectedValue + "', "  '�ల
        SQl = SQl + " '" + FormatNumber(CDbl(DVDays.Text), 1) + "', "   '�i�ФѼ�
        '25~35
        SQl = SQl + " '" + DOTNo1.Text + "', "                            '�[�ZNo-1
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours1.Text), 1) + "', "  '�[�Z�ɼ�-1
        SQl = SQl + " '" + DOTNo2.Text + "', "                            '�[�ZNo-2
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours2.Text), 1) + "', "  '�[�Z�ɼ�-2
        SQl = SQl + " '" + DOTNo3.Text + "', "                            '�[�ZNo-3
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours3.Text), 1) + "', "  '�[�Z�ɼ�-3
        SQl = SQl + " '" + DOTNo4.Text + "', "                            '�[�ZNo-4
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours4.Text), 1) + "', "  '�[�Z�ɼ�-4
        SQl = SQl + " '" + DOTNo5.Text + "', "                            '�[�ZNo-5
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours5.Text), 1) + "', "  '�[�Z�ɼ�-5
        SQl = SQl + " '" + FormatNumber(CDbl(DOTHours.Text), 1) + "', "   '�[�Z�`�ɼ�
        '36~40
        SQl = SQl + " '" + DBStartDate.Text + "', "                       '�w�w-�}�l���
        SQl = SQl + " '" + CStr(CInt(DBStartH.SelectedValue)) + "', "     '�w�w-�}�l��
        SQl = SQl + " '" + DBEndDate.Text + "', "                         '�w�w-�������
        SQl = SQl + " '" + CStr(CInt(DBEndH.SelectedValue)) + "', "       '�w�w-������
        SQl = SQl + " '" + FormatNumber(CDbl(DBDays.Text), 1) + "', "     '�w�w-�ɼ�
        '41~45
        SQl = SQl + " '" + DAStartDate.Text + "', "                       '���-�}�l���
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "     '���-�}�l��
        SQl = SQl + " '" + DAEndDate.Text + "', "                         '���-�������
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "       '���-������
        SQl = SQl + " '" + FormatNumber(CDbl(DADays.Text), 1) + "', "     '���-�ɼ�
        '46
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "      '�а��z��
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "        '�@����
        SQl = SQl + " '" + NowDateTime + "', "                            '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                                     '�ק��
        SQl = SQl + " '" + NowDateTime + "' "                             '�ק�ɶ�
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("TimeOffFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_TimeOffSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " SalaryYM = '" & DSalaryYM.Text & "',"
        SQl = SQl + " Name = N'" & DName.SelectedItem.Text & "',"
        SQl = SQl + " EmpID = '" & DEmpID.Text & "',"

        SQl = SQl + " JobTitle = N'" & DJobTitle.Text & "',"
        SQl = SQl + " JobCode = '" & DJobCode.Text & "',"
        SQl = SQl + " DepoName = N'" & DDepoName.Text & "',"
        SQl = SQl + " DepoCode = '" & DDepoCode.Text & "',"
        SQl = SQl + " Division = N'" & DDivision.Text & "',"
        SQl = SQl + " DivisionCode = '" & DDivisionCode.Text & "',"

        SQl = SQl + " After = N'" & DAfter.SelectedValue & "',"
        SQl = SQl + " JobAgent = N'" & DJobAgent.Text & "',"
        SQl = SQl + " TimeOffAgent = N'" & DTimeOffAgent.Text & "',"

        SQl = SQl + " VacationCode = '" & Left(DVacation.SelectedValue, 1) & "',"
        SQl = SQl + " Vacation = N'" & Mid(DVacation.SelectedValue, 3) & "',"
        SQl = SQl + " Evidence = N'" & DEvidence.Text & "',"
        SQl = SQl + " Salary = N'" & DSalary.Text & "',"
        SQl = SQl + " DieType = N'" & DDieType.SelectedValue & "',"
        SQl = SQl + " VDays = '" & FormatNumber(CDbl(DVDays.Text), 1) & "',"

        If DOTNo1.Text <> "" Then SQl = SQl + " OTNo1 = '" & DOTNo1.Text & "',"
        If DOTNo2.Text <> "" Then SQl = SQl + " OTNo2 = '" & DOTNo2.Text & "',"
        If DOTNo3.Text <> "" Then SQl = SQl + " OTNo3 = '" & DOTNo3.Text & "',"
        If DOTNo4.Text <> "" Then SQl = SQl + " OTNo4 = '" & DOTNo4.Text & "',"
        If DOTNo5.Text <> "" Then SQl = SQl + " OTNo5 = '" & DOTNo5.Text & "',"

        SQl = SQl + " OTHours1 = '" & FormatNumber(CDbl(DOTHours1.Text), 1) & "',"
        SQl = SQl + " OTHours2 = '" & FormatNumber(CDbl(DOTHours2.Text), 1) & "',"
        SQl = SQl + " OTHours3 = '" & FormatNumber(CDbl(DOTHours3.Text), 1) & "',"
        SQl = SQl + " OTHours4 = '" & FormatNumber(CDbl(DOTHours4.Text), 1) & "',"
        SQl = SQl + " OTHours5 = '" & FormatNumber(CDbl(DOTHours5.Text), 1) & "',"
        SQl = SQl + " OTHours = '" & FormatNumber(CDbl(DOTHours.Text), 1) & "',"

        SQl = SQl + " BStartDate = '" & DBStartDate.Text & "',"
        SQl = SQl + " BStartH = '" & CStr(CInt(DBStartH.SelectedValue)) & "',"
        SQl = SQl + " BEndDate = '" & DBEndDate.Text & "',"
        SQl = SQl + " BEndH = '" & CStr(CInt(DBEndH.SelectedValue)) & "',"
        SQl = SQl + " BDays = '" & FormatNumber(CDbl(DBDays.Text), 1) & "',"

        SQl = SQl + " AStartDate = '" & DAStartDate.Text & "',"
        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AEndDate = '" & DAEndDate.Text & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " ADays = '" & FormatNumber(CDbl(DADays.Text), 1) & "',"

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
    '**     ��Button�p��w�w��� 
    '**
    '*****************************************************************
    Private Sub BBDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBDays.Click
        Dim Hours As Double = 0
        Dim VHours As Double = 0
        Dim MHours As Double = 0
        Dim ErrCode As Integer = 0
        Dim BaseDate As Date
        Dim SQL, Message As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        '�_������ˬd
        If ErrCode = 0 Then
            If CDate(DBEndDate.Text) < CDate(DBStartDate.Text) Then ErrCode = 8010
            If CDate(DBEndDate.Text) = CDate(DBStartDate.Text) Then
                If DBEndH.SelectedValue <= DBStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '�_������P���ˬd
        If ErrCode = 0 Then
            If CStr(Year(CDate(DBStartDate.Text))) + "/" + CStr(Month(CDate(DBStartDate.Text))) <> CStr(Year(CDate(DBEndDate.Text))) + "/" + CStr(Month(CDate(DBEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '����ɶ�
        If ErrCode = 0 Then
            '2008/7�s�Ұ�k�����ץ�-�������t����
            If Left(DVacation.SelectedValue, 1) <> "O" And Left(DVacation.SelectedValue, 1) <> "P" And Left(DVacation.SelectedValue, 1) <> "Q" And Left(DVacation.SelectedValue, 1) <> "W" Then
                OleDbConnection1.Open()
                SQL = "Select Count(*) As Days From M_Vacation "
                SQL = SQL + "Where Active = '1' "
                'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                'SQL = SQL + "  and Depo = '" + wDepo + "' "
                '
                SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
                'Modify-End
                SQL = SQL + "  and YMD  >= '" + CDate(DBStartDate.Text) + "' "
                SQL = SQL + "  and YMD  <= '" + CDate(DBEndDate.Text) + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
                OleDbConnection1.Close()
            End If
            '�а��ɶ�
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '�а��ɼƮt�ɶ�
            If DBEndH.SelectedValue >= DBStartH.SelectedValue Then
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DBStartH.SelectedValue <= "12" And DBEndH.SelectedValue <= "12") Then
                    Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                Else
                    If (DBStartH.SelectedValue >= "13" And DBEndH.SelectedValue >= "13") Then
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue)
                    Else
                        Hours = CInt(DBEndH.SelectedValue) - CInt(DBStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '�p��ɶ�
            DBDays.Text = FormatNumber((VHours + Hours - MHours) / 8, 1)
            '��ڦP�B
            DAStartDate.Text = DBStartDate.Text
            DAStartH.SelectedIndex = DAStartH.Items.IndexOf(DAStartH.Items.FindByValue(DBStartH.SelectedValue))
            DAEndDate.Text = DBEndDate.Text
            DAEndH.SelectedIndex = DAEndH.Items.IndexOf(DAEndH.Items.FindByValue(DBEndH.SelectedValue))
            DADays.Text = DBDays.Text
            '�а��O�_���b����
            If CDbl(DBDays.Text) - Fix(DBDays.Text) <> 0.5 And CDbl(DBDays.Text) - Fix(DBDays.Text) <> 0 Then
                ErrCode = 8030
            End If
        End If
        '�~��_����ˬd
        If ErrCode = 0 Then
            If Left(DVacation.SelectedValue, 1) = "A" Then
                DBDataSet1.Clear()
                OleDbConnection1.Open()
                SQL = "Select StartDate As ArriveDate From M_Emp "
                SQL = SQL + "Where Com_Code = '" + DDepoCode.Text + "' "
                SQL = SQL + "  and ID  = '" + DEmpID.Text + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
                    BaseDate = CDate(Mid(DSalaryYM.Text, 1, 4) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Month) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Day))
                    If CDate(DAStartDate.Text) < BaseDate And CDate(DAEndDate.Text) < BaseDate Then
                        '�_���L��~��_���
                    Else
                        If CDate(DAStartDate.Text) >= BaseDate And CDate(DAEndDate.Text) >= BaseDate Then
                            '�_���L��~��_���
                        Else
                            '�_������~��_���
                            ErrCode = 8050
                        End If
                    End If
                Else
                    ErrCode = 8050
                End If
                OleDbConnection1.Close()
            End If
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
            If ErrCode = 8030 Then Message = "�а����ݬO�b��,�нT�{!"
            If ErrCode = 8050 Then Message = "�а��_����������~�~��_���[" + BaseDate + "],�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��Button�p���ڤ�� 
    '**
    '*****************************************************************
    Private Sub BADays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BADays.Click
        Dim Hours As Double = 0
        Dim VHours As Double = 0
        Dim MHours As Double = 0
        Dim ErrCode As Integer = 0
        Dim BaseDate As Date
        Dim SQL, Message As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        '�_������ˬd
        If ErrCode = 0 Then
            If CDate(DAEndDate.Text) < CDate(DAStartDate.Text) Then ErrCode = 8010
            If CDate(DAEndDate.Text) = CDate(DAStartDate.Text) Then
                If DAEndH.SelectedValue <= DAStartH.SelectedValue Then ErrCode = 8010
            End If
        End If
        '�_������P���ˬd
        If ErrCode = 0 Then
            If CStr(Year(CDate(DAStartDate.Text))) + "/" + CStr(Month(CDate(DAStartDate.Text))) <> CStr(Year(CDate(DAEndDate.Text))) + "/" + CStr(Month(CDate(DAEndDate.Text))) Then
                ErrCode = 8020
            End If
        End If
        '����ɶ�
        If ErrCode = 0 Then
            '2008/7�s�Ұ�k�����ץ�-�������t����
            If Left(DVacation.SelectedValue, 1) <> "O" And Left(DVacation.SelectedValue, 1) <> "P" And Left(DVacation.SelectedValue, 1) <> "Q" And Left(DVacation.SelectedValue, 1) <> "W" Then
                OleDbConnection1.Open()
                SQL = "Select Count(*) As Days From M_Vacation "
                SQL = SQL + "Where Active = '1' "
                'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                'SQL = SQL + "  and Depo = '" + wDepo + "' "
                '
                SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
                'Modify-End
                SQL = SQL + "  and YMD  >= '" + CDate(DAStartDate.Text) + "' "
                SQL = SQL + "  and YMD  <= '" + CDate(DAEndDate.Text) + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
                OleDbConnection1.Close()
            End If
            '�а��ɶ�
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '�а��ɼƮt�ɶ�
            If DAEndH.SelectedValue >= DAStartH.SelectedValue Then
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) - 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            Else
                If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
                    Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                Else
                    If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue)
                    Else
                        Hours = CInt(DAEndH.SelectedValue) - CInt(DAStartH.SelectedValue) + 1
                    End If
                End If
                If Hours > 8 Then Hours = 8
            End If
            '�p��ɶ�
            DADays.Text = FormatNumber((VHours + Hours - MHours) / 8, 1)
            '�а��O�_���b����
            If CDbl(DADays.Text) - Fix(DADays.Text) <> 0.5 And CDbl(DADays.Text) - Fix(DADays.Text) <> 0 Then
                ErrCode = 8030
            End If
        End If
        '�~��_����ˬd
        If ErrCode = 0 Then
            If Left(DVacation.SelectedValue, 1) = "A" Then
                DBDataSet1.Clear()
                OleDbConnection1.Open()
                SQL = "Select StartDate As ArriveDate From M_Emp "
                SQL = SQL + "Where Com_Code = '" + DDepoCode.Text + "' "
                SQL = SQL + "  and ID  = '" + DEmpID.Text + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "Vacation")
                If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
                    BaseDate = CDate(Mid(DSalaryYM.Text, 1, 4) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Month) + "/" + CStr(CDate(DBDataSet1.Tables("Vacation").Rows(0).Item("ArriveDate")).Day))
                    If CDate(DAStartDate.Text) < BaseDate And CDate(DAEndDate.Text) < BaseDate Then
                        '�_���L��~��_���
                    Else
                        If CDate(DAStartDate.Text) >= BaseDate And CDate(DAEndDate.Text) >= BaseDate Then
                            '�_���L��~��_���
                        Else
                            '�_������~��_���
                            ErrCode = 8050
                        End If
                    End If
                Else
                    ErrCode = 8050
                End If
                OleDbConnection1.Close()
            End If
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
            If ErrCode = 8030 Then Message = "�а����ݬO�b��,�нT�{!"
            If ErrCode = 8050 Then Message = "�а��_����������~�~��_���[" + BaseDate + "],�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܧ� 
    '**
    '*****************************************************************
    Private Sub DVacation_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DVacation.SelectedIndexChanged
        Dim i As Integer
        Dim wDays As Double = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '�U�����(�ల����/�[�Z�O��)
        DEvidence.Text = ""
        DSalary.Text = ""
        DVDays.Text = "0"
        DDieType.Enabled = False
        SetFieldData("DieType", "ZZZZZZ")
        BOTRecord.Enabled = False

        OleDbConnection1.Open()
        '���o����
        SQL = "Select * From M_Referp "
        SQL = SQL + "Where Cat='1002' "
        SQL = SQL + "  and DKey = '" & "EVIDENCE-" + Left(DVacation.SelectedValue, 1) & "'"
        SQL = SQL + "Order by Data "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            DEvidence.Text = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        End If
        '���o�~��
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp "
        SQL = SQL + "Where Cat='1002' "
        SQL = SQL + "  and DKey = '" & "SALARY-" + Left(DVacation.SelectedValue, 1) & "'"
        SQL = SQL + "Order by Data "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            DSalary.Text = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        End If
        '9:�ե� --> Set�[�Z�O��
        If Left(DVacation.SelectedValue, 1) = "9" Then
            BOTRecord.Enabled = True
        End If
        '���o�w�ФѼ�(�D X:�ల)
        If Left(DVacation.SelectedValue, 1) <> "X" Then
            '�W�u�e���
            'DBDataSet1.Clear()
            'SQL = "Select IsNull(Sum(CodeValue), 0) As Days From HR_BeforeVacation "
            'SQL = SQL + "Where DepoCode = '" & DDepoCode.Text & "'"
            'SQL = SQL + "  and EmpID = '" & DEmpID.Text & "'"
            'SQL = SQL + "  and Code = '" & Left(DVacation.SelectedValue, 1) & "'"
            'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            'DBAdapter3.Fill(DBDataSet1, "Before")
            'wDays = DBDataSet1.Tables("Before").Rows(0).Item("Days")
            '�t�θ��
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(ADays), 0) As Days From F_TimeOffSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and DepoCode = '" & DDepoCode.Text & "'"
            SQL = SQL + "  and EmpID = '" & DEmpID.Text & "'"
            SQL = SQL + "  and AEndDate >= '" & Date.Now.Year.ToString + "/1/1" & "'"
            SQL = SQL + "  and AEndDate <= '" & Date.Now.Year.ToString + "/12/31" & "'"
            SQL = SQL + "  and VacationCode = '" & Left(DVacation.SelectedValue, 1) & "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet1, "TimeOffSheet")
            wDays = wDays + DBDataSet1.Tables("TimeOffSheet").Rows(0).Item("Days")

            DVDays.Text = FormatNumber(wDays, 1)
        End If
        '-----------------------------------------------------------------------------------------------
        '���o�i�ФѼ�(�DA:�~��, X:�ల)
        'If Left(DVacation.SelectedValue, 1) <> "A" And Left(DVacation.SelectedValue, 1) <> "X" Then
        'DBDataSet1.Clear()
        'SQL = "Select * From M_Referp "
        'SQL = SQL + "Where Cat='1002' "
        'SQL = SQL + "  and DKey = '" & "VACATIONDAY-" + Left(DVacation.SelectedValue, 1) & "'"
        'SQL = SQL + "Order by Data "
        'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter3.Fill(DBDataSet1, "M_Referp")
        'If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
        'DVDays.Text = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        'End If
        'End If
        '���o�i�ФѼ�(A:�~��)
        'If Left(DVacation.SelectedValue, 1) = "A" Then
        'DVDays.Text = CalVacation()
        'End If
        '-----------------------------------------------------------------------------------------------
        '���o�ల����
        If Left(DVacation.SelectedValue, 1) = "X" Then
            DDieType.Enabled = True
        End If
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ల�ܧ� 
    '**
    '*****************************************************************
    Private Sub DDieType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DDieType.SelectedIndexChanged
        Dim SQL As String
        Dim wDays As Double = 0
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        '���o�w�ФѼ�
        '�W�u�e���
        'DBDataSet1.Clear()
        'SQL = "Select IsNull(Sum(CodeValue), 0) As Days From HR_BeforeVacation "
        'SQL = SQL + "Where DepoCode = '" & DDepoCode.Text & "'"
        'SQL = SQL + "  and EmpID = '" & DEmpID.Text & "'"
        'SQL = SQL + "  and Code = '" & Left(DVacation.SelectedValue, 1) & "'"
        'SQL = SQL + "  and Code1 = '" & Left(DDieType.SelectedValue, 1) & "'"
        'Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter3.Fill(DBDataSet1, "Before")
        'wDays = DBDataSet1.Tables("Before").Rows(0).Item("Days")
        '�t�θ��
        DBDataSet1.Clear()
        SQL = "Select IsNull(Sum(ADays), 0) As Days From F_TimeOffSheet "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "  and DepoCode = '" & DDepoCode.Text & "'"
        SQL = SQL + "  and EmpID = '" & DEmpID.Text & "'"
        SQL = SQL + "  and VacationCode = '" & Left(DVacation.SelectedValue, 1) & "'"
        SQL = SQL + "  and DieType = '" & DDieType.SelectedValue & "'"
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "TimeOffSheet")
        wDays = wDays + DBDataSet1.Tables("TimeOffSheet").Rows(0).Item("Days")

        DVDays.Text = FormatNumber(wDays, 1)
        '-----------------------------------------------------------------------------------------------
        '���o�i�ФѼ�
        'SQL = "Select * From M_Referp "
        'SQL = SQL + "Where Cat='1002' "
        'SQL = SQL + "  and DKey = '" & "DIETYPEDAY-" + Left(DDieType.SelectedValue, 1) & "'"
        'SQL = SQL + "Order by Data "
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "M_Referp")
        'If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
        'DVDays.Text = DBDataSet1.Tables("M_Referp").Rows(0).Item("Data")
        'End If
        '-----------------------------------------------------------------------------------------------
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
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
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
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �p��i�Ц~�� 
    '**
    '*****************************************************************
    Function CalVacation() As String
        Dim wArriveDate, wC1, wC2, wC365 As DateTime
        Dim Years, Years1, Years2, wC1Days, wC2Days As Integer
        Dim Days, Days1, Days2, TDays As Integer

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        OleDbConnection1.Open()
        '���o�J����
        SQL = "Select ArriveDate From M_EMP "
        SQL = SQL & " Where Com_Code = '" & DDepoCode.Text & "'"
        SQL = SQL & "   And ID = '" & DEmpID.Text & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_EMP")
        If DBDataSet1.Tables("M_EMP").Rows.Count > 0 Then
            wArriveDate = DBDataSet1.Tables("M_EMP").Rows(0).Item("ArriveDate")
        End If
        OleDbConnection1.Close()
        '�~��
        Years = Fix(DateDiff("D", wArriveDate, Now) / 365)
        If CStr(Month(Now)) + CStr(Day(Now)) >= CStr(Month(wArriveDate)) + CStr(Day(wArriveDate)) Then
            Years1 = Years - 1
            Years2 = Years
        Else
            Years1 = Years
            Years2 = Years + 1
        End If
        If Years1 < 0 Then Years1 = 0
        '�Ѽ�
        wC1 = CDate(CStr(DateTime.Now.Year) + "/1/1")
        wC2 = CDate(CStr(DateTime.Now.Year) + "/" + CStr(Month(wArriveDate)) + "/" + CStr(Day(wArriveDate)))
        wC365 = CDate(CStr(DateTime.Now.Year) + "/12/31")
        wC1Days = DateDiff("D", wC1, wC2)
        wC2Days = DateDiff("D", wC2, wC365) + 1
        '���~��Base
        Days = 30
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp "
        SQL = SQL + "Where Cat='1002' "
        SQL = SQL + "  and DKey = '" & "SENIORITY-" + CStr(Years) & "'"
        SQL = SQL + "Order by Data "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            Days = CInt(DBDataSet1.Tables("M_Referp").Rows(0).Item("Data"))
        End If
        '�J����e�~��Base
        Days1 = 30
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp "
        SQL = SQL + "Where Cat='1002' "
        SQL = SQL + "  and DKey = '" & "SENIORITY-" + CStr(Years1) & "'"
        SQL = SQL + "Order by Data "
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            Days1 = CInt(DBDataSet1.Tables("M_Referp").Rows(0).Item("Data"))
        End If
        '�J�����~��Base
        Days2 = 30
        DBDataSet1.Clear()
        SQL = "Select * From M_Referp "
        SQL = SQL + "Where Cat='1002' "
        SQL = SQL + "  and DKey = '" & "SENIORITY-" + CStr(Years2) & "'"
        SQL = SQL + "Order by Data "
        Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter4.Fill(DBDataSet1, "M_Referp")
        If DBDataSet1.Tables("M_Referp").Rows.Count > 0 Then
            Days2 = CInt(DBDataSet1.Tables("M_Referp").Rows(0).Item("Data"))
        End If

        TDays = Fix((Days1 * wC1Days / 365) + (Days2 * wC2Days / 365) + 0.9999)
        If TDays > Days Then TDays = Days

        CalVacation = CStr(TDays)
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
    '**     �ˬd�W���ɮ�
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
        '���ݦ~��
        If InputCheck = 0 Then
            If FindFieldInf("SalaryYM") = 1 Then
                If DSalaryYM.Text = "" Then InputCheck = 1
            End If
        End If
        '�ƫe,�ƫ�
        If InputCheck = 0 Then
            If FindFieldInf("After") = 1 Then
                If DAfter.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '¾�ȥN�z�H
        If InputCheck = 0 Then
            If FindFieldInf("JobAgent") = 1 Then
                If DJobAgent.Text = "" Then InputCheck = 1
            End If
        End If
        '�N�а��H
        If InputCheck = 0 Then
            If FindFieldInf("TimeOffAgent") = 1 Then
                If DTimeOffAgent.Text = "" Then InputCheck = 1
            End If
        End If
        '���O
        If InputCheck = 0 Then
            If FindFieldInf("Vacation") = 1 Then
                If DVacation.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("Evidence") = 1 Then
                If DEvidence.Text = "" Then InputCheck = 1
            End If
        End If
        '�~��
        If InputCheck = 0 Then
            If FindFieldInf("Salary") = 1 Then
                If DSalary.Text = "" Then InputCheck = 1
            End If
        End If
        '�ల�O
        If InputCheck = 0 Then
            If FindFieldInf("DieType") = 1 Then
                If DDieType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�i�ФѼ�
        If InputCheck = 0 Then
            If FindFieldInf("VDays") = 1 Then
                If DVDays.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo1
        If InputCheck = 0 Then
            If FindFieldInf("OTNo1") = 1 Then
                If DOTNo1.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo1-�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours1") = 1 Then
                If DOTHours1.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo2
        If InputCheck = 0 Then
            If FindFieldInf("OTNo2") = 1 Then
                If DOTNo2.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo2-�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours2") = 1 Then
                If DOTHours2.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo3
        If InputCheck = 0 Then
            If FindFieldInf("OTNo3") = 1 Then
                If DOTNo3.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo3-�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours3") = 1 Then
                If DOTHours3.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo4
        If InputCheck = 0 Then
            If FindFieldInf("OTNo4") = 1 Then
                If DOTNo4.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo4-�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours4") = 1 Then
                If DOTHours4.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo5
        If InputCheck = 0 Then
            If FindFieldInf("OTNo5") = 1 Then
                If DOTNo5.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�ZNo5-�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours5") = 1 Then
                If DOTHours5.Text = "" Then InputCheck = 1
            End If
        End If
        '�[�Z�`�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("OTHours") = 1 Then
                If DOTHours.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w�}�l���
        If InputCheck = 0 Then
            If FindFieldInf("BStartDate") = 1 Then
                If DBStartDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w�}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("BStartH") = 1 Then
                If DBStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�������
        If InputCheck = 0 Then
            If FindFieldInf("BEndDate") = 1 Then
                If DBEndDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w����-��
        If InputCheck = 0 Then
            If FindFieldInf("BEndH") = 1 Then
                If DBEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w�Ѽ�
        If InputCheck = 0 Then
            If FindFieldInf("BDays") = 1 Then
                If DBDays.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڶ}�l���
        If InputCheck = 0 Then
            If FindFieldInf("AStartDate") = 1 Then
                If DAStartDate.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڶ}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("AStartH") = 1 Then
                If DAStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڵ������
        If InputCheck = 0 Then
            If FindFieldInf("AEndDate") = 1 Then
                If DAEndDate.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڵ���-��
        If InputCheck = 0 Then
            If FindFieldInf("AEndH") = 1 Then
                If DAEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��ڤѼ�
        If InputCheck = 0 Then
            If FindFieldInf("ADays") = 1 Then
                If DADays.Text = "" Then InputCheck = 1
            End If
        End If
        '�а��z��
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If
    End Function
End Class
