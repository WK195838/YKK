Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_AwaySheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BADays As System.Web.UI.WebControls.Button
    Protected WithEvents BBDays As System.Web.UI.WebControls.Button
    Protected WithEvents BAEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents BAStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DAEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartH As System.Web.UI.WebControls.DropDownList
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
    Protected WithEvents DPlace As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAHour As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOtherPlace As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBHour As System.Web.UI.WebControls.TextBox
    Protected WithEvents DADay As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBDay As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAwaySheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button

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

        Response.Cookies("PGM").Value = "HR_AwaySheet_01.aspx"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AwayFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_AwaySheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AwaySheet")
        If DBDataSet1.Tables("F_AwaySheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Date")                   '�ӽФ��
            DSalaryYM.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("SalaryYM")           '���ݦ~��

            SetFieldData("Name", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Name"))          '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("DivisionCode")   '�����N�X

            SetFieldData("Place", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("Place"))        '�ت��a
            DOtherPlace.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("OtherPlace")       '�ت��a(��L)
            DJobAgent.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("JobAgent")           '�N�z�H

            DBStartDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BStartDate")               '�w�w�}�l���
            SetFieldData("BStartH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BStartH").ToString)   '�w�w�}�l��
            DBEndDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BEndDate")                   '�w�w�������
            SetFieldData("BEndH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BEndH").ToString)       '�w�w������
            DBDay.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BDay"))     '�w�w��
            DBHour.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("BHour"))   '�w�w��

            DAStartDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AStartDate")               '��ڶ}�l���
            SetFieldData("AStartH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AStartH").ToString)   '��ڶ}�l��
            DAEndDate.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AEndDate")                   '��ڵ������
            SetFieldData("AEndH", DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AEndH").ToString)       '��ڵ�����
            DADay.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("ADay"))     '��ڤ�
            DAHour.Text = CStr(DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("AHour"))   '��ڮ�

            DFReason.Text = DBDataSet1.Tables("F_AwaySheet").Rows(0).Item("FReason")             '�~�X�z��
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
                DAwaySheet1.Visible = True   '���Sheet-1
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
                    Top = 576
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 464
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
            Top = 384
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
        'BBStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  
        'BBEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DBEndDate', '');"                     
        'BAStartDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  
        'BAEndDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAEndDate', '');"                     
        '
        BBStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBStartDate', 'Form1.DSalaryYM');"  '�~�X�w�w�}�l���
        BBEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DBEndDate', '');"                     '�~�X�w�w�������
        BAStartDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAStartDate', 'Form1.DSalaryYM');"  '�~�X��ڶ}�l���
        BAEndDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAEndDate', '');"                     '�~�X��ڵ������
        'Modify-End
        BCardTime.Attributes("onclick") = "ShowCardTime();"    '��d�O��
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
                    Top = 464
                Else
                    If DDelay.Visible = True Then
                        Top = 576
                    Else
                        Top = 464
                    End If
                End If
            End If
        Else
            Top = 384
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
        '�ت��a
        Select Case FindFieldInf("Place")
            Case 0  '���
                DPlace.BackColor = Color.LightGray
                DPlace.Visible = True
            Case 1  '�ק�+�ˬd
                DPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPlaceRqd", "DPlace", "���`�G�ݿ�J�ت��a")
                DPlace.Visible = True
            Case 2  '�ק�
                DPlace.BackColor = Color.Yellow
                DPlace.Visible = True
            Case Else   '����
                DPlace.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Place", "ZZZZZZ")
        '�ت��a(��L)
        Select Case FindFieldInf("OtherPlace")
            Case 0  '���
                DOtherPlace.BackColor = Color.LightGray
                DOtherPlace.Visible = True
            Case 1  '�ק�+�ˬd
                DOtherPlace.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOtherPlaceRqd", "DOtherPlace", "���`�G�ݿ�J�ت��a")
                DOtherPlace.Visible = True
            Case 2  '�ק�
                DOtherPlace.BackColor = Color.Yellow
                DOtherPlace.Visible = True
            Case Else   '����
                DOtherPlace.Visible = False
        End Select
        If pPost = "New" Then DOtherPlace.Text = ""
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
        Select Case FindFieldInf("BDay")
            Case 0  '���
                DBDay.BackColor = Color.LightGray
                DBDay.Visible = True
            Case 1  '�ק�+�ˬd
                DBDay.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBDayRqd", "DBDay", "���`�G�ݿ�J�w�w�Ѽ�")
                DBDay.Visible = True
            Case 2  '�ק�
                DBDay.BackColor = Color.Yellow
                DBDay.Visible = True
            Case Else   '����
                DBDay.Visible = False
        End Select
        If pPost = "New" Then DBDay.Text = "0"
        '�w�w�ɼ�
        Select Case FindFieldInf("BHour")
            Case 0  '���
                DBHour.BackColor = Color.LightGray
                DBHour.Visible = True
            Case 1  '�ק�+�ˬd
                DBHour.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBHourRqd", "DBHour", "���`�G�ݿ�J�w�w�ɼ�")
                DBHour.Visible = True
            Case 2  '�ק�
                DBHour.BackColor = Color.Yellow
                DBHour.Visible = True
            Case Else   '����
                DBHour.Visible = False
        End Select
        If pPost = "New" Then DBHour.Text = "0"
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
        Select Case FindFieldInf("ADay")
            Case 0  '���
                DADay.BackColor = Color.LightGray
                DADay.Visible = True
            Case 1  '�ק�+�ˬd
                DADay.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DADayRqd", "DADay", "���`�G�ݿ�J��ڤѼ�")
                DADay.Visible = True
            Case 2  '�ק�
                DADay.BackColor = Color.Yellow
                DADay.Visible = True
            Case Else   '����
                DADay.Visible = False
        End Select
        If pPost = "New" Then DADay.Text = "0"
        '��ڮɼ�
        Select Case FindFieldInf("AHour")
            Case 0  '���
                DAHour.BackColor = Color.LightGray
                DAHour.Visible = True
            Case 1  '�ק�+�ˬd
                DAHour.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAHourRqd", "DAHour", "���`�G�ݿ�J��ڮɼ�")
                DAHour.Visible = True
            Case 2  '�ק�
                DAHour.BackColor = Color.Yellow
                DAHour.Visible = True
            Case Else   '����
                DAHour.Visible = False
        End Select
        If pPost = "New" Then DAHour.Text = "0"
        '�~�X�z��
        Select Case FindFieldInf("FReason")
            Case 0  '���
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '�ק�+�ˬd
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "���`�G�ݿ�J�~�X�z��")
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
                    SQL = SQL + "  And FormNo = '001003' "
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
        '�ت��a
        If pFieldName = "Place" Then
            DPlace.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPlace.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='PLACE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPlace.Items.Add(ListItem1)
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
                SQL = "Select * From M_Referp Where Cat='1003' and DKey='HOUR' Order by Data "
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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        '--------------------------------------------------------------------------
        '--  �ˬd���U�����
        '--------------------------------------------------------------------------
        '�~�X�Ѽ�,�ɼ�
        If ErrCode = 0 Then
            If DBDay.Text = "0" And DBHour.Text = "0" Then ErrCode = 9051
            If DADay.Text = "0" And DAHour.Text = "0" Then ErrCode = 9051
        End If
        '�ت��a
        If ErrCode = 0 Then
            If Left(DPlace.SelectedValue, 2) <= "10" Or Left(DPlace.SelectedValue, 2) = "99" Then
                If DOtherPlace.Text = "" Then
                    ErrCode = 9052
                End If
            End If
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
        '��Ƶ��⵲�G�ˬd
        If ErrCode = 0 Then
            '#�w�w
            Dim Hours As Integer = 0
            Dim VHours As Integer = 0
            Dim MHours As Integer = 0
            Dim wTotal As Integer = 0

            Dim SQL As String
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DBStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DBEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '�~�X�ɶ�
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '�~�X�ɼƮt�ɶ�
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
            wTotal = VHours + Hours - MHours
            If DBDay.Text <> CStr(Fix(wTotal / 8)) Then
                ErrCode = 8040
            End If
            If DBHour.Text <> CStr(wTotal - (Fix(wTotal / 8) * 8)) Then
                ErrCode = 8040
            End If
            DBDataSet1.Clear()
            '----------------------------------------------------------------------------------------------
            '#���
            Hours = 0
            VHours = 0
            MHours = 0
            wTotal = 0
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DAStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DAEndDate.Text + "' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '�~�X�ɶ�
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '�~�X�ɼƮt�ɶ�
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
            wTotal = VHours + Hours - MHours
            If DADay.Text <> CStr(Fix(wTotal / 8)) Then
                ErrCode = 8040
            End If
            If DAHour.Text <> CStr(wTotal - (Fix(wTotal / 8) * 8)) Then
                ErrCode = 8040
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

                ErrCode = oCommon.CommissionNo("001003", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No

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
            If ErrCode = 9051 Then Message = "�~�X�ɶ��ݶ�g,�нT�{!"
            If ErrCode = 9052 Then Message = "�ت��a(�Ȥ�,�~�`��,��L)�ɻݶ�g���e,�нT�{!"
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
            If ErrCode = 8040 Then Message = "�~�X����ɶ����ŦX,�нT�{!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AwayFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_AwaySheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, SalaryYM, Name, EmpID, "                                '6~9
        SQl = SQl + "JobTitle, JobCode, DepoName, DepoCode, Division, "            '10~14
        SQl = SQl + "DivisionCode, Place, OtherPlace, JobAgent, "                  '15~18
        SQl = SQl + "BStartDate, BStartH, BEndDate, BEndH, BDay, BHour, "          '19~24
        SQl = SQl + "AStartDate, AStartH, AEndDate, AEndH, ADay, AHour, "          '25~30
        SQl = SQl + "FReason, "                                                    '31
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '32~35
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '001003', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~9
        SQl = SQl + " '" + DDate.Text + "', "               '�ӽФ��
        SQl = SQl + " '" + DSalaryYM.Text + "', "           '���ݦ~��
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', " '�m�W
        SQl = SQl + " '" + DEmpID.Text + "', "              'EMP-ID
        '10~15
        SQl = SQl + " N'" + DJobTitle.Text + "', "          '¾��
        SQl = SQl + " '" + DJobCode.Text + "', "            '¾�٥N�X
        SQl = SQl + " N'" + DDepoName.Text + "', "          '���q�O
        SQl = SQl + " '" + DDepoCode.Text + "', "           '���q�OCode
        SQl = SQl + " N'" + DDivision.Text + "', "          '����
        SQl = SQl + " '" + DDivisionCode.Text + "', "       '�����N�X
        '16~18
        SQl = SQl + " N'" + DPlace.SelectedValue + "', "    '�ت��a
        SQl = SQl + " N'" + DOtherPlace.Text + "', "        '�ت��a-��L
        SQl = SQl + " N'" + DJobAgent.Text + "', "          '¾�ȥN�z�H
        '19~24
        SQl = SQl + " '" + DBStartDate.Text + "', "                       '�w�w-�}�l���
        SQl = SQl + " '" + CStr(CInt(DBStartH.SelectedValue)) + "', "     '�w�w-�}�l��
        SQl = SQl + " '" + DBEndDate.Text + "', "                         '�w�w-�������
        SQl = SQl + " '" + CStr(CInt(DBEndH.SelectedValue)) + "', "       '�w�w-������
        SQl = SQl + " '" + CStr(CInt(DBDay.Text)) + "', "                 '�w�w-�Ѽ�
        SQl = SQl + " '" + CStr(CInt(DBHour.Text)) + "', "                '�w�w-�ɼ�
        '25~30
        SQl = SQl + " '" + DAStartDate.Text + "', "                       '���-�}�l���
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "     '���-�}�l��
        SQl = SQl + " '" + DAEndDate.Text + "', "                         '���-�������
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "       '���-������
        SQl = SQl + " '" + CStr(CInt(DADay.Text)) + "', "                 '���-�Ѽ�
        SQl = SQl + " '" + CStr(CInt(DAHour.Text)) + "', "                '���-�ɼ�
        '31
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "      '�~�X�z��
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AwayFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_AwaySheet Set "
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

        SQl = SQl + " Place = N'" & DPlace.SelectedValue & "',"
        SQl = SQl + " OtherPlace = N'" & DOtherPlace.Text & "',"
        SQl = SQl + " JobAgent = N'" & DJobAgent.Text & "',"

        SQl = SQl + " BStartDate = '" & DBStartDate.Text & "',"
        SQl = SQl + " BStartH = '" & CStr(CInt(DBStartH.SelectedValue)) & "',"
        SQl = SQl + " BEndDate = '" & DBEndDate.Text & "',"
        SQl = SQl + " BEndH = '" & CStr(CInt(DBEndH.SelectedValue)) & "',"
        SQl = SQl + " BDay = '" & CStr(CInt(DBDay.Text)) & "',"
        SQl = SQl + " BHour = '" & CStr(CInt(DBHour.Text)) & "',"

        SQl = SQl + " AStartDate = '" & DAStartDate.Text & "',"
        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AEndDate = '" & DAEndDate.Text & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " ADay = '" & CStr(CInt(DADay.Text)) & "',"
        SQl = SQl + " AHour = '" & CStr(CInt(DAHour.Text)) & "',"

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
        Dim Hours As Integer = 0
        Dim VHours As Integer = 0
        Dim MHours As Integer = 0
        Dim wTotal As Integer = 0
        Dim ErrCode As Integer = 0
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
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DBStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DBEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '�~�X�ɶ�
            If CDate(DBEndDate.Text) > CDate(DBStartDate.Text) Then
                VHours = DateDiff("d", CDate(DBStartDate.Text), CDate(DBEndDate.Text)) * 8
            End If
            '�~�X�ɼƮt�ɶ�
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
            wTotal = VHours + Hours - MHours
            DBDay.Text = CStr(Fix(wTotal / 8))
            DBHour.Text = CStr(wTotal - (Fix(wTotal / 8) * 8))
            '��ڦP�B
            DAStartDate.Text = DBStartDate.Text
            DAStartH.SelectedIndex = DAStartH.Items.IndexOf(DAStartH.Items.FindByValue(DBStartH.SelectedValue))
            DAEndDate.Text = DBEndDate.Text
            DAEndH.SelectedIndex = DAEndH.Items.IndexOf(DAEndH.Items.FindByValue(DBEndH.SelectedValue))
            DADay.Text = DBDay.Text
            DAHour.Text = DBHour.Text
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
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
        Dim Hours As Integer = 0
        Dim VHours As Integer = 0
        Dim MHours As Integer = 0
        Dim wTotal As Integer = 0
        Dim ErrCode As Integer = 0
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
            OleDbConnection1.Open()
            SQL = "Select Count(*) As Days From M_Vacation "
            SQL = SQL + "Where Active = '1' "
            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
            'SQL = SQL + "  and Depo = '" + wDepo + "' "
            '
            SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
            'Modify-End
            SQL = SQL + "  and YMD  >= '" + DAStartDate.Text + "' "
            SQL = SQL + "  and YMD  <= '" + DAEndDate.Text + "' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "Vacation")
            MHours = DBDataSet1.Tables("Vacation").Rows(0).Item("Days") * 8
            OleDbConnection1.Close()
            '�~�X�ɶ�
            If CDate(DAEndDate.Text) > CDate(DAStartDate.Text) Then
                VHours = DateDiff("d", CDate(DAStartDate.Text), CDate(DAEndDate.Text)) * 8
            End If
            '�~�X�ɼƮt�ɶ�
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
            wTotal = VHours + Hours - MHours
            DADay.Text = CStr(Fix(wTotal / 8))
            DAHour.Text = CStr(wTotal - (Fix(wTotal / 8) * 8))
        End If

        If ErrCode > 0 Then
            If ErrCode = 8010 Then Message = "����������i�p��}�l���,�нT�{!"
            If ErrCode = 8020 Then Message = "�}�l����P����������i���,�нT�{!"
            Response.Write(YKK.ShowMessage(Message))
        End If
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
        '�ت��a
        If InputCheck = 0 Then
            If FindFieldInf("Place") = 1 Then
                If DPlace.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ت��a(��L)
        If InputCheck = 0 Then
            If FindFieldInf("OtherPlace") = 1 Then
                If DOtherPlace.Text = "" Then InputCheck = 1
            End If
        End If
        '¾�ȥN�z�H
        If InputCheck = 0 Then
            If FindFieldInf("JobAgent") = 1 Then
                If DJobAgent.Text = "" Then InputCheck = 1
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
            If FindFieldInf("BDay") = 1 Then
                If DBDay.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w�ɼ�
        If InputCheck = 0 Then
            If FindFieldInf("BHour") = 1 Then
                If DBHour.Text = "" Then InputCheck = 1
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
            If FindFieldInf("ADay") = 1 Then
                If DADay.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڮɼ�
        If InputCheck = 0 Then
            If FindFieldInf("AHour") = 1 Then
                If DAHour.Text = "" Then InputCheck = 1
            End If
        End If
        '�~�X�z��
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
