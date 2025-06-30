Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_AddWorkSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BOverTime As System.Web.UI.WebControls.Button
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
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
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
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
    Protected WithEvents BWorkDate As System.Web.UI.WebControls.Button
    Protected WithEvents DWorkDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAWorkDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BAWorkDate As System.Web.UI.WebControls.Button
    Protected WithEvents DAddWorkSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents BACardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DAddWorkType As System.Web.UI.WebControls.DropDownList

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

        Response.Cookies("PGM").Value = "HR_AddWorkSheet_01.aspx"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AddWorkFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_AddWorkSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AddWorkSheet")
        If DBDataSet1.Tables("F_AddWorkSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Date")                   '�ӽФ��
            DSalaryYM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("SalaryYM")           '���ݦ~��

            SetFieldData("Name", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Name"))          '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DivisionCode")   '�����N�X

            DWorkDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WorkDate")                   '�ʶԤ��
            SetFieldData("WStartH", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WStartH").ToString)   '�ʶԶ}�l-��
            SetFieldData("WStartM", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WStartM").ToString)   '�ʶԶ}�l-��
            SetFieldData("WEndH", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WEndH").ToString)       '�ʶԲפ�-��
            SetFieldData("WEndM", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WEndM").ToString)       '�ʶԲפ�-��
            DWH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WH").ToString                      '�p�⵲�G-��
            DWM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WM").ToString                      '�p�⵲�G-��

            DAWorkDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AWorkDate")                 '�ɤu���
            SetFieldData("AStartH", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AStartH").ToString)   '�ɤu�}�l-��
            SetFieldData("AStartM", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AStartM").ToString)   '�ɤu�}�l-��
            SetFieldData("AEndH", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AEndH").ToString)       '�ɤu�פ�-��
            SetFieldData("AEndM", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AEndM").ToString)       '�ɤu�פ�-��
            DAH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AH").ToString                      '�p�⵲�G-��
            DAM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AM").ToString                      '�p�⵲�G-��

            DFReason.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("FReason")                     '�z��
            SetFieldData("AddWorkType", DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AddWorkType"))    '�ɤu���O
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
                DAddWorkSheet1.Visible = True    '���Sheet-1
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
                    Top = 608
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 496
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
            Top = 416
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
        'BWorkDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DWorkDate', 'Form1.DSalaryYM');"  '�ʶԤ��
        'BAWorkDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DAWorkDate', '');"               '�ɤu���
        '
        BWorkDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DWorkDate', 'Form1.DSalaryYM');"  '�ʶԤ��
        BAWorkDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DAWorkDate', '');"               '�ɤu���
        'Modify-End

        BCardTime.Attributes("onclick") = "ShowCardTime();"    '�ʶԨ�d�O��
        BACardTime.Attributes("onclick") = "ShowCardTimeA();"  '�ɤu��d�O��
        BOverTime.Attributes("onclick") = "ShowOverTime();"    '�[�Z�O��
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
                    Top = 496
                Else
                    If DDelay.Visible = True Then
                        Top = 608
                    Else
                        Top = 496
                    End If
                End If
            End If
        Else
            Top = 416
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
        '�ɤu���O
        Select Case FindFieldInf("AddWorkType")
            Case 0  '���
                DAddWorkType.BackColor = Color.LightGray
                DAddWorkType.Visible = True
            Case 1  '�ק�+�ˬd
                DAddWorkType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAddWorkTypeRqd", "DAddWorkType", "���`�G�ݿ�J�ɤu���O")
                DAddWorkType.Visible = True
            Case 2  '�ק�
                DAddWorkType.BackColor = Color.Yellow
                DAddWorkType.Visible = True
            Case Else   '����
                DAddWorkType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AddWorkType", "ZZZZZZ")
        '�ʶԤ��
        Select Case FindFieldInf("WorkDate")
            Case 0  '���
                DWorkDate.BackColor = Color.LightGray
                DWorkDate.ReadOnly = True
                DWorkDate.Visible = True
                BWorkDate.Visible = False
            Case 1  '�ק�+�ˬd
                DWorkDate.BackColor = Color.GreenYellow
                DWorkDate.ReadOnly = True
                ShowRequiredFieldValidator("DWorkDateRqd", "DWorkDate", "���`�G�ݿ�J�ʶԤ��")
                DWorkDate.Visible = True
                BWorkDate.Visible = True
            Case 2  '�ק�
                DWorkDate.BackColor = Color.Yellow
                DWorkDate.ReadOnly = True
                DWorkDate.Visible = True
                BWorkDate.Visible = True
            Case Else   '����
                DWorkDate.Visible = False
                BWorkDate.Visible = False
        End Select
        If pPost = "New" Then DWorkDate.Text = CStr(DateTime.Now.Today)
        '�ʶԶ}�l-��
        Select Case FindFieldInf("WStartH")
            Case 0  '���
                DWStartH.BackColor = Color.LightGray
                DWStartH.Visible = True
            Case 1  '�ק�+�ˬd
                DWStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWStartHRqd", "DWStartH", "���`�G�ݿ�J�ʶԶ}�l-��")
                DWStartH.Visible = True
            Case 2  '�ק�
                DWStartH.BackColor = Color.Yellow
                DWStartH.Visible = True
            Case Else   '����
                DWStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WStartH", "ZZZZZZ")
        '�ʶԶ}�l-��
        Select Case FindFieldInf("WStartM")
            Case 0  '���
                DWStartM.BackColor = Color.LightGray
                DWStartM.Visible = True
            Case 1  '�ק�+�ˬd
                DWStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWStartMRqd", "DWStartM", "���`�G�ݿ�J�ʶԶ}�l-��")
                DWStartM.Visible = True
            Case 2  '�ק�
                DWStartM.BackColor = Color.Yellow
                DWStartM.Visible = True
            Case Else   '����
                DWStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WStartM", "ZZZZZZ")
        '�ʶԲפ�-��
        Select Case FindFieldInf("WEndH")
            Case 0  '���
                DWEndH.BackColor = Color.LightGray
                DWEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DWEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWEndHRqd", "DWEndH", "���`�G�ݿ�J�ʶԲפ�-��")
                DWEndH.Visible = True
            Case 2  '�ק�
                DWEndH.BackColor = Color.Yellow
                DWEndH.Visible = True
            Case Else   '����
                DWEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WEndH", "ZZZZZZ")
        '�ʶԲפ�-��
        Select Case FindFieldInf("WEndM")
            Case 0  '���
                DWEndM.BackColor = Color.LightGray
                DWEndM.Visible = True
            Case 1  '�ק�+�ˬd
                DWEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWEndMRqd", "DWEndM", "���`�G�ݿ�J�ʶԲפ�-��")
                DWEndM.Visible = True
            Case 2  '�ק�
                DWEndM.BackColor = Color.Yellow
                DWEndM.Visible = True
            Case Else   '����
                DWEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WEndM", "ZZZZZZ")
        '�ʶԭp��-��
        Select Case FindFieldInf("WH")
            Case 0  '���
                DWH.BackColor = Color.LightGray
                DWH.ReadOnly = True
                DWH.Visible = True
            Case 1  '�ק�+�ˬd
                DWH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWHRqd", "DWH", "���`�G�ݿ�J�ʶԭp��-��")
                DWH.Visible = True
            Case 2  '�ק�
                DWH.BackColor = Color.Yellow
                DWH.Visible = True
            Case Else   '����
                DWH.Visible = False
        End Select
        If pPost = "New" Then DWH.Text = "0"
        '�ʶԭp��-��
        Select Case FindFieldInf("WM")
            Case 0  '���
                DWM.BackColor = Color.LightGray
                DWM.ReadOnly = True
                DWM.Visible = True
            Case 1  '�ק�+�ˬd
                DWM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWMRqd", "DWM", "���`�G�ݿ�J�ʶԭp��-��")
                DWM.Visible = True
            Case 2  '�ק�
                DWM.BackColor = Color.Yellow
                DWM.Visible = True
            Case Else   '����
                DWM.Visible = False
        End Select
        If pPost = "New" Then DWM.Text = "0"
        '�ɤu���
        Select Case FindFieldInf("AWorkDate")
            Case 0  '���
                DAWorkDate.BackColor = Color.LightGray
                DAWorkDate.ReadOnly = True
                DAWorkDate.Visible = True
                BAWorkDate.Visible = False
            Case 1  '�ק�+�ˬd
                DAWorkDate.BackColor = Color.GreenYellow
                DAWorkDate.ReadOnly = True
                ShowRequiredFieldValidator("DAWorkDateRqd", "DAWorkDate", "���`�G�ݿ�J�ɤu���")
                DAWorkDate.Visible = True
                BAWorkDate.Visible = True
            Case 2  '�ק�
                DAWorkDate.BackColor = Color.Yellow
                DAWorkDate.ReadOnly = True
                DAWorkDate.Visible = True
                BAWorkDate.Visible = True
            Case Else   '����
                DAWorkDate.Visible = False
                BAWorkDate.Visible = False
        End Select
        If pPost = "New" Then DAWorkDate.Text = CStr(DateTime.Now.Today)
        '�ɤu�}�l-��
        Select Case FindFieldInf("AStartH")
            Case 0  '���
                DAStartH.BackColor = Color.LightGray
                DAStartH.Visible = True
            Case 1  '�ק�+�ˬd
                DAStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartHRqd", "DAStartH", "���`�G�ݿ�J�ɤu�}�l-��")
                DAStartH.Visible = True
            Case 2  '�ק�
                DAStartH.BackColor = Color.Yellow
                DAStartH.Visible = True
            Case Else   '����
                DAStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartH", "ZZZZZZ")
        '�ɤu�}�l-��
        Select Case FindFieldInf("AStartM")
            Case 0  '���
                DAStartM.BackColor = Color.LightGray
                DAStartM.Visible = True
            Case 1  '�ק�+�ˬd
                DAStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartMRqd", "DAStartM", "���`�G�ݿ�J�ɤu�}�l-��")
                DAStartM.Visible = True
            Case 2  '�ק�
                DAStartM.BackColor = Color.Yellow
                DAStartM.Visible = True
            Case Else   '����
                DAStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartM", "ZZZZZZ")
        '�ɤu�פ�-��
        Select Case FindFieldInf("AEndH")
            Case 0  '���
                DAEndH.BackColor = Color.LightGray
                DAEndH.Visible = True
            Case 1  '�ק�+�ˬd
                DAEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndHRqd", "DAEndH", "���`�G�ݿ�J�ɤu�פ�-��")
                DAEndH.Visible = True
            Case 2  '�ק�
                DAEndH.BackColor = Color.Yellow
                DAEndH.Visible = True
            Case Else   '����
                DAEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndH", "ZZZZZZ")
        '�ɤu�פ�-��
        Select Case FindFieldInf("AEndM")
            Case 0  '���
                DAEndM.BackColor = Color.LightGray
                DAEndM.Visible = True
            Case 1  '�ק�+�ˬd
                DAEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndMRqd", "DAEndM", "���`�G�ݿ�J�ɤu�פ�-��")
                DAEndM.Visible = True
            Case 2  '�ק�
                DAEndM.BackColor = Color.Yellow
                DAEndM.Visible = True
            Case Else   '����
                DAEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndM", "ZZZZZZ")
        '�ɤu�p��-��
        Select Case FindFieldInf("AH")
            Case 0  '���
                DAH.BackColor = Color.LightGray
                DAH.ReadOnly = True
                DAH.Visible = True
            Case 1  '�ק�+�ˬd
                DAH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAHRqd", "DAH", "���`�G�ݿ�J�ɤu�p��-��")
                DAH.Visible = True
            Case 2  '�ק�
                DAH.BackColor = Color.Yellow
                DAH.Visible = True
            Case Else   '����
                DAH.Visible = False
        End Select
        If pPost = "New" Then DAH.Text = "0"
        '�ɤu�p��-��
        Select Case FindFieldInf("AM")
            Case 0  '���
                DAM.BackColor = Color.LightGray
                DAM.ReadOnly = True
                DAM.Visible = True
            Case 1  '�ק�+�ˬd
                DAM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAMRqd", "DAM", "���`�G�ݿ�J�ɤu�p��-��")
                DAM.Visible = True
            Case 2  '�ק�
                DAM.BackColor = Color.Yellow
                DAM.Visible = True
            Case Else   '����
                DAM.Visible = False
        End Select
        If pPost = "New" Then DAM.Text = "0"
        '�z��
        Select Case FindFieldInf("FReason")
            Case 0  '���
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '�ק�+�ˬd
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "���`�G�ݿ�J�ɤu�z��")
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
                    SQL = SQL + "  And FormNo = '001004' "
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
        '�ɤu���O
        If pFieldName = "AddWorkType" Then
            DAddWorkType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAddWorkType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='ADDWORKTYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAddWorkType.Items.Add(ListItem1)
                Next
            End If
        End If
        '�ʶԶ}�l-��
        If pFieldName = "WStartH" Then
            DWStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='WHOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '�ʶԶ}�l-��
        If pFieldName = "WStartM" Then
            DWStartM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWStartM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWStartM.Items.Add(ListItem1)
                Next
            End If
        End If
        '�ʶԲפ�-��
        If pFieldName = "WEndH" Then
            DWEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='WHOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '�ʶԲפ�-��
        If pFieldName = "WEndM" Then
            DWEndM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWEndM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWEndM.Items.Add(ListItem1)
                Next
            End If
        End If
        '�ɤu�}�l-��
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
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='AWHOUR' Order by Data "
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
        '�ɤu�}�l-��
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
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='MIN' Order by Data "
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
        '�ɤu�פ�-��
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
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='AWHOUR' Order by Data "
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
        '�ɤu�פ�-��
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
                SQL = "Select * From M_Referp Where Cat='1004' and DKey='MIN' Order by Data "
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

        '--------------------------------------------------------------------------
        '--  �ˬd���U�����
        '--------------------------------------------------------------------------
        '�ʶԭp�⵲�G
        If ErrCode = 0 Then
            If DWH.Text = "0" And DWM.Text = "0" Then ErrCode = 9050
        End If
        '�ɤu�p�⵲�G
        If ErrCode = 0 Then
            If DAH.Text = "0" And DAM.Text = "0" Then ErrCode = 9051
        End If
        '�ʶ�/�ɤu�ɼƤ��@�P
        If ErrCode = 0 Then
            If CInt(DWH.Text) * 60 + CInt(DWM.Text) > CInt(DAH.Text) * 60 + CInt(DAM.Text) Then ErrCode = 9053
        End If
        '�ʶ�/�ɤu�P���ˬd
        If ErrCode = 0 Then
            If CStr(Year(CDate(DWorkDate.Text))) + "/" + CStr(Month(CDate(DWorkDate.Text))) <> CStr(Year(CDate(DAWorkDate.Text))) + "/" + CStr(Month(CDate(DAWorkDate.Text))) Then
                ErrCode = 9052
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

                ErrCode = oCommon.CommissionNo("001004", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No

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
                '--�l��ǰe--------
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
            If ErrCode = 9050 Then Message = "�ʶԮɶ��ݶ�g,�нT�{!"
            If ErrCode = 9051 Then Message = "�ɤu�ɶ��ݶ�g,�нT�{!"
            If ErrCode = 9052 Then Message = "�ʶԤ���P�ɤu������i���,�нT�{!"
            If ErrCode = 9053 Then Message = "�ʶԻP�ɤu�ҵ��⤧�ɼƤ��@�P,�нT�{!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AddWorkFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_AddWorkSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, SalaryYM, Name, EmpID, JobTitle, "                      '6~10
        SQl = SQl + "JobCode, DepoName, DepoCode, Division, DivisionCode, "        '10~15
        SQl = SQl + "WorkDate, WStartH, WStartM, WEndH, WEndM, WH, WM, "           '16~22
        SQl = SQl + "AWorkDate, AStartH, AStartM, AEndH, AEndM, AH, AM, "          '23~29
        SQl = SQl + "FReason, AddWorkType, "                                       '30~31
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '32~34
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '001004', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~10
        SQl = SQl + " '" + DDate.Text + "', "               '�ӽФ��
        SQl = SQl + " '" + DSalaryYM.Text + "', "           '���ݦ~��
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', " '�m�W
        SQl = SQl + " '" + DEmpID.Text + "', "              'EMP-ID
        SQl = SQl + " N'" + DJobTitle.Text + "', "          '¾��
        '11~15
        SQl = SQl + " '" + DJobCode.Text + "', "            '¾�٥N�X
        SQl = SQl + " N'" + DDepoName.Text + "', "          '���q�O
        SQl = SQl + " '" + DDepoCode.Text + "', "           '���q�OCode
        SQl = SQl + " N'" + DDivision.Text + "', "          '����
        SQl = SQl + " '" + DDivisionCode.Text + "', "       '�����N�X
        '16~22
        SQl = SQl + " '" + DWorkDate.Text + "', "                         '�ʶԤ��
        SQl = SQl + " '" + CStr(CInt(DWStartH.SelectedValue)) + "', "     '�ʶԶ}�l-��
        SQl = SQl + " '" + CStr(CInt(DWStartM.SelectedValue)) + "', "     '�ʶԶ}�l-��
        SQl = SQl + " '" + CStr(CInt(DWEndH.SelectedValue)) + "', "       '�ʶԲפ�-��
        SQl = SQl + " '" + CStr(CInt(DWEndM.SelectedValue)) + "', "       '�ʶԲפ�-��
        SQl = SQl + " '" + CStr(CInt(DWH.Text)) + "', "                   '�ʶԭp��-��
        SQl = SQl + " '" + CStr(CInt(DWM.Text)) + "', "                   '�ʶԭp��-��
        '23~29
        SQl = SQl + " '" + DAWorkDate.Text + "', "                        '�ɤu���
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "     '�ɤu�}�l-��
        SQl = SQl + " '" + CStr(CInt(DAStartM.SelectedValue)) + "', "     '�ɤu�}�l-��
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "       '�ɤu�פ�-��
        SQl = SQl + " '" + CStr(CInt(DAEndM.SelectedValue)) + "', "       '�ɤu�פ�-��
        SQl = SQl + " '" + CStr(CInt(DAH.Text)) + "', "                   '�ɤu�p��-��
        SQl = SQl + " '" + CStr(CInt(DAM.Text)) + "', "                   '�ɤu�p��-��
        '31~32
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "      '�z��
        SQl = SQl + " N'" + DAddWorkType.SelectedItem.Text + "', "        '�ɤu���O
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("AddWorkFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_AddWorkSheet Set "
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

        SQl = SQl + " WorkDate = '" & DWorkDate.Text & "',"
        SQl = SQl + " WStartH = '" & CStr(CInt(DWStartH.SelectedValue)) & "',"
        SQl = SQl + " WStartM = '" & CStr(CInt(DWStartM.SelectedValue)) & "',"
        SQl = SQl + " WEndH = '" & CStr(CInt(DWEndH.SelectedValue)) & "',"
        SQl = SQl + " WEndM = '" & CStr(CInt(DWEndM.SelectedValue)) & "',"
        SQl = SQl + " WH = '" & CStr(CInt(DWH.Text)) & "',"
        SQl = SQl + " WM = '" & CStr(CInt(DWM.Text)) & "',"

        SQl = SQl + " AWorkDate = '" & DAWorkDate.Text & "',"
        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AStartM = '" & CStr(CInt(DAStartM.SelectedValue)) & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " AEndM = '" & CStr(CInt(DAEndM.SelectedValue)) & "',"
        SQl = SQl + " AH = '" & CStr(CInt(DAH.Text)) & "',"
        SQl = SQl + " AM = '" & CStr(CInt(DAM.Text)) & "',"

        SQl = SQl + " FReason = N'" & YKK.ReplaceString(DFReason.Text) & "',"
        SQl = SQl + " AddWorkType = N'" & DAddWorkType.SelectedItem.Text & "',"

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
    '**     ����ʶԲפ�-�� 
    '**
    '*****************************************************************
    Private Sub DWEndH_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DWEndH.SelectedIndexChanged
        Dim StartTime As String = DWStartH.SelectedValue + DWStartM.SelectedValue
        Dim EndTime As String = DWEndH.SelectedValue + DWEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DWStartH.SelectedValue, DWStartM.SelectedValue, DWEndH.SelectedValue, DWEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            '�������ȥ�1H
            If (DWStartH.SelectedValue <= "12" And DWEndH.SelectedValue <= "12") Then
            Else
                If (DWStartH.SelectedValue >= "13" And DWEndH.SelectedValue >= "13") Then
                Else
                    Hour = Hour - 1
                End If
            End If
            '
            DWH.Text = CStr(Hour)
            DWM.Text = CStr(Minute)
        Else
            DWH.Text = "0"
            DWM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����ʶԲפ�-�� 
    '**
    '*****************************************************************
    Private Sub DWEndM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DWEndM.SelectedIndexChanged
        Dim StartTime As String = DWStartH.SelectedValue + DWStartM.SelectedValue
        Dim EndTime As String = DWEndH.SelectedValue + DWEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DWStartH.SelectedValue, DWStartM.SelectedValue, DWEndH.SelectedValue, DWEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            '�������ȥ�1H
            If (DWStartH.SelectedValue <= "12" And DWEndH.SelectedValue <= "12") Then
            Else
                If (DWStartH.SelectedValue >= "13" And DWEndH.SelectedValue >= "13") Then
                Else
                    Hour = Hour - 1
                End If
            End If
            '
            DWH.Text = CStr(Hour)
            DWM.Text = CStr(Minute)
        Else
            DWH.Text = "0"
            DWM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����ɤu�פ�-�� 
    '**
    '*****************************************************************
    Private Sub DAEndH_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DAEndH.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            '�������ȥ�1H
            If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
            Else
                If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                Else
                    Hour = Hour - 1
                End If
            End If
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ����ɤu�פ�-�� 
    '**
    '*****************************************************************
    Private Sub DAEndM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DAEndM.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            '�������ȥ�1H
            If (DAStartH.SelectedValue <= "12" And DAEndH.SelectedValue <= "12") Then
            Else
                If (DAStartH.SelectedValue >= "13" And DAEndH.SelectedValue >= "13") Then
                Else
                    Hour = Hour - 1
                End If
            End If
            '
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
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
        '�ɤu���O
        If InputCheck = 0 Then
            If FindFieldInf("AddWorkType") = 1 Then
                If DAddWorkType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ʶԤ��
        If InputCheck = 0 Then
            If FindFieldInf("WorkDate") = 1 Then
                If DWorkDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�ʶԶ}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("WStartH") = 1 Then
                If DWStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ʶԶ}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("WStartM") = 1 Then
                If DWStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ʶԲפ�-��
        If InputCheck = 0 Then
            If FindFieldInf("WEndH") = 1 Then
                If DWEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ʶԲפ�-��
        If InputCheck = 0 Then
            If FindFieldInf("WEndM") = 1 Then
                If DWEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ʶԭp��-��
        If InputCheck = 0 Then
            If FindFieldInf("WH") = 1 Then
                If DWH.Text = "" Then InputCheck = 1
            End If
        End If
        '�ʶԭp��-��
        If InputCheck = 0 Then
            If FindFieldInf("WM") = 1 Then
                If DWM.Text = "" Then InputCheck = 1
            End If
        End If
        '�ɤu���
        If InputCheck = 0 Then
            If FindFieldInf("AWorkDate") = 1 Then
                If DAWorkDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("AStartH") = 1 Then
                If DAStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�}�l-��
        If InputCheck = 0 Then
            If FindFieldInf("AStartM") = 1 Then
                If DAStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�פ�-��
        If InputCheck = 0 Then
            If FindFieldInf("AEndH") = 1 Then
                If DAEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�פ�-��
        If InputCheck = 0 Then
            If FindFieldInf("AEndM") = 1 Then
                If DAEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�p��-��
        If InputCheck = 0 Then
            If FindFieldInf("AH") = 1 Then
                If DAH.Text = "" Then InputCheck = 1
            End If
        End If
        '�ɤu�p��-��
        If InputCheck = 0 Then
            If FindFieldInf("AM") = 1 Then
                If DAM.Text = "" Then InputCheck = 1
            End If
        End If
        '�z��
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
