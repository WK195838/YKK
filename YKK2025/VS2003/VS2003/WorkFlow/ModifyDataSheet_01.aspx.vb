Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ModifyDataSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DModifyDataSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DMapSheet4 As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSheet As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMReasonType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DStatus As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMContent1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMContent2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DIContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFDateTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DIPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DComNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSheet1 As System.Web.UI.WebControls.HyperLink

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
    'Dim wDepo As String = "CL"      '���c��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

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
            If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
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

        Response.Cookies("PGM").Value = "ModifyDataSheet_01.aspx"
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
        Dim Message As String = ""
        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
            Response.Write(YKK.ShowMessage(Message))
        End If
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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ModifyDataSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ModifyDataSheet")
        If DBDataSet1.Tables("F_ModifyDataSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("No")                         'No
            DDate.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Date")                     '���
            SetFieldData("Division", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Division"))    '�e�U����
            SetFieldData("Person", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Person"))        '�e�U���
            SetFieldData("MDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MDivision"))  '�קﳡ��
            SetFieldData("MPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MPerson"))      '�ק���
            SetFieldData("Sheet", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Sheet"))          '�e�U��
            DComNo.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("ComNo")                   '�e�UNo
            SetFieldData("Status", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("Status"))        '�e�U��
            SetFieldData("WDivision", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WDivision"))  '�ݳB�z�u�{����
            SetFieldData("WPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("WPerson"))      '�ݳB�z�u�{���

            SetFieldData("MReasonType", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReasonType"))   '�ק�z�����O
            DMReason.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MReason")               '�ק�z��
            DMContent1.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent1")           '�ק鷺�e-1
            DMContent2.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("MContent2")           '�ק鷺�e-2

            DIContent.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IContent")             '��ڭק鷺�e
            DFDateTime.Text = DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("FDateTime")           '�w�w�����ɶ�
            SetFieldData("IPerson", DBDataSet1.Tables("F_ModifyDataSheet").Rows(0).Item("IPerson"))      '��ڭק���

            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '����
                DBStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BStartTime")       '�w�w�}�l
                DBEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime")           '�w�w����
                DAStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AStartTime")       '��ڶ}�l
                If Not IsDBNull(DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")) Then
                    DAEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")           '��ڧ���
                Else
                    DAEndTime.Text = ""
                End If
                LBefOP.NavigateUrl = "BefOPList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & wStep

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
                DFormSno.Text = "�渹�G" & DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FormSno")       '�渹
            End If
        End If
        '���s��
        SQL = " select  formno,tablename1  from M_form"
        SQL &= " where formname = '" + Trim(DSheet.SelectedValue) + "'"
        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        Dim ds As New Data.DataSet
        DBAdapter3.Fill(ds)
        Dim SheetFormNo As String
        Dim SheetFormSno As String

        Dim dt As DataTable = ds.Tables(0)
        If dt.Rows.Count > 0 Then
            SheetFormNo = dt.Rows(0)("formno")
            SQL = " select formsno from f_" + dt.Rows(0)("tablename1")
            SQL &= " where no='" + Trim(DComNo.Text) + "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            Dim ds1 As New Data.DataSet
            DBAdapter4.Fill(ds1)
            Dim dt1 As DataTable = ds1.Tables(0)
            Dim sUrl As String
            Dim cun As Integer

            If dt1.Rows.Count > 0 Then
                SheetFormSno = dt1.Rows(0)("formsno")
                sUrl = dt.Rows(0)("tablename1") + "_02.aspx?pFormNo=" & SheetFormNo & "&pFormSno=" & SheetFormSno
                'Dim sUrlStr = Mid(sUrl, 11, sUrl.Length - 1)
                LSheet1.NavigateUrl = sUrl

            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
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
            Else
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

        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet���
                DMapSheet3.Visible = True   '�ϭ�Sheet
                DMapSheet4.Visible = True   '����Sheet
                DDelivery.Visible = True    '���Sheet

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
                    Top = 992
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 880
                End If

                '���
                DBStartTime.Visible = True      '�w�w�}�l
                DBEndTime.Visible = True        '�w�w����
                DAStartTime.Visible = True      '��ڶ}�l
                DAEndTime.Visible = True        '��ڧ���
                '�u�{�i���\Ū
                DOPReady.Visible = True
                DOPReadyDesc.Visible = True
                '�ݾ\Ū
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    '20190430 JESSICA ����
                    DOPReady.BackColor = Color.LightGray

                    'DOPReady.BackColor = Color.GreenYellow
                    'ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "���`�G�ݾ\Ū�u�{�i��")
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
                LBefOP.Visible = True          '�u�{�i��
                LSheet1.Visible = True          '�u�{�i��
                '���s��m
                BNG1.Style.Add("Top", Top)      'NG���s
                BNG2.Style.Add("Top", Top)     'NG���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
            Top = 696
            'Sheet����
            DMapSheet3.Visible = False  '�ϭ�Sheet
            DMapSheet4.Visible = False  '����Sheet
            DDelivery.Visible = False   '���Sheet

            DOPReady.Visible = False    '�u�{�i���\Ū
            DOPReadyDesc.Visible = False

            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False     '����
            DBStartTime.Visible = False     '�w�w�}�l
            DBEndTime.Visible = False       '�w�w����
            DAStartTime.Visible = False     '��ڶ}�l
            DAEndTime.Visible = False       '��ڧ���
            DReasonCode.Visible = False     '����z�ѥN�X
            DReason.Visible = False         '����z��
            DReasonDesc.Visible = False     '�����L����
            '�s������
            LBefOP.Visible = False          '�u�{�i��
            LSheet1.Visible = False          '�u�{�i��
            '���s��m
            BNG1.Style.Add("Top", Top)      'NG���s
            BNG2.Style.Add("Top", Top)     'NG���s
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      'OK���s
            DFormSno.Style.Add("Top", Top) '�渹
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
        '������
        BDate.Attributes("onclick") = "calendarPicker('Form1.DDate');"
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
    '**
    '**     �ק�̳����I���ƥ�
    '**
    '*****************************************************************
    Private Sub DMDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMDivision.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        DMPerson.Items.Clear()
        SQL = "Select UserName From M_Users "
        SQL = SQL & " Where DivName = '" & DMDivision.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")

        For i = 0 To DBTable1.Rows.Count - 1
            If i = 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = ""
                ListItem1.Value = ""
                ListItem1.Selected = True
                DMPerson.Items.Add(ListItem1)
            End If
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
            ListItem2.Value = DBTable1.Rows(i).Item("UserName")
            ListItem2.Selected = False
            DMPerson.Items.Add(ListItem2)
        Next
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ݳB�z�u�{�����I���ƥ�
    '**
    '*****************************************************************
    Private Sub DWDivision_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DWDivision.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        DWPerson.Items.Clear()
        SQL = "Select UserName From M_Users "
        SQL = SQL & " Where DivName = '" & DWDivision.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        DBTable1 = DBDataSet1.Tables("M_Users")

        For i = 0 To DBTable1.Rows.Count - 1
            If i = 0 Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = ""
                ListItem1.Value = ""
                ListItem1.Selected = True
                DWPerson.Items.Add(ListItem1)
            End If
            Dim ListItem2 As New ListItem
            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
            ListItem2.Value = DBTable1.Rows(i).Item("UserName")
            ListItem2.Selected = False
            DWPerson.Items.Add(ListItem2)
        Next
        'DB�s������
        OleDbConnection1.Close()
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
        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 880
                Else
                    If DDelay.Visible = True Then
                        Top = 992
                    Else
                        Top = 880
                    End If
                End If
            End If
        Else
            Top = 696
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
    '**(ShowSheetField)
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w
        'No
        Select Case FindFieldInf("No")
            Case 0  '���
                DNo.BackColor = Color.LightGray
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

        '���
        Select Case FindFieldInf("Date")
            Case 0  '���
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
                BDate.Visible = False
            Case 1  '�ק�+�ˬd
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J���")
                DDate.Visible = True
                BDate.Visible = True
            Case 2  '�ק�
                DDate.BackColor = Color.Yellow
                DDate.Visible = True
                BDate.Visible = True
            Case Else   '����
                DDate.Visible = False
                BDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)

        '�e�U����
        Select Case FindFieldInf("Division")
            Case 0  '���
                DDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
                DDivision.Visible = True
            Case 1  '�ק�+�ˬd
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "���`�G�ݿ�J�e�U����")
                DDivision.Visible = True
            Case 2  '�ק�
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '����
                DDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

        '�e�U���
        Select Case FindFieldInf("Person")
            Case 0  '���
                DPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
                DPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "���`�G�ݿ�J�e�U���")
                DPerson.Visible = True
            Case 2  '�ק�
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '����
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")

        '�קﳡ��
        Select Case FindFieldInf("MDivision")
            Case 0  '���
                DMDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
                DMDivision.Visible = True
            Case 1  '�ק�+�ˬd
                DMDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMDivisionRqd", "DMDivision", "���`�G�ݿ�J�קﳡ��")
                DMDivision.Visible = True
            Case 2  '�ק�
                DMDivision.BackColor = Color.Yellow
                DMDivision.Visible = True
            Case Else   '����
                DMDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MDivision", "ZZZZZZ")

        '�ק���
        Select Case FindFieldInf("MPerson")
            Case 0  '���
                DMPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
                DMPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DMPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMPersonRqd", "DMPerson", "���`�G�ݿ�J�ק���")
                DMPerson.Visible = True
            Case 2  '�ק�
                DMPerson.BackColor = Color.Yellow
                DMPerson.Visible = True
            Case Else   '����
                DMPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MPerson", "ZZZZZZ")

        '�e�U��
        Select Case FindFieldInf("Sheet")
            Case 0  '���
                DSheet.BackColor = Color.LightGray
                DSheet.Visible = True
            Case 1  '�ק�+�ˬd
                DSheet.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSheetRqd", "DSheet", "���`�G�ݿ�J�e�U��")
                DSheet.Visible = True
            Case 2  '�ק�
                DSheet.BackColor = Color.Yellow
                DSheet.Visible = True
            Case Else   '����
                DSheet.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sheet", "ZZZZZZ")

        '�e�UNo
        Select Case FindFieldInf("ComNo")
            Case 0  '���
                DComNo.BackColor = Color.LightGray
                DComNo.ReadOnly = True
                DComNo.Visible = True
            Case 1  '�ק�+�ˬd
                DComNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DComNoRqd", "DComNo", "���`�G�ݿ�J�ܢ�")
                DComNo.Visible = True
            Case 2  '�ק�
                DComNo.BackColor = Color.Yellow
                DComNo.Visible = True
            Case Else   '����
                DComNo.Visible = False
        End Select
        If pPost = "New" Then DComNo.Text = ""

        '���A
        Select Case FindFieldInf("Status")
            Case 0  '���
                DStatus.BackColor = Color.LightGray
                DStatus.Visible = True
            Case 1  '�ק�+�ˬd
                DStatus.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DStatusRqd", "DStatus", "���`�G�ݿ�J���A")
                DStatus.Visible = True
            Case 2  '�ק�
                DStatus.BackColor = Color.Yellow
                DStatus.Visible = True
            Case Else   '����
                DStatus.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Status", "ZZZZZZ")

        '�ݳB�z�u�{����
        Select Case FindFieldInf("WDivision")
            Case 0  '���
                DWDivision.BackColor = Color.LightGray
                DWDivision.Visible = True
            Case 1  '�ק�+�ˬd
                DWDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWDivisionRqd", "DWDivision", "���`�G�ݿ�J�ݳB�z�u�{����")
                DWDivision.Visible = True
            Case 2  '�ק�
                DWDivision.BackColor = Color.Yellow
                DWDivision.Visible = True
            Case Else   '����
                DWDivision.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WDivision", "ZZZZZZ")

        '�ݳB�z�u�{���
        Select Case FindFieldInf("WPerson")
            Case 0  '���
                DWPerson.BackColor = Color.LightGray
                DWPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DWPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWPersonRqd", "DWPerson", "���`�G�ݿ�J�ݳB�z�u�{���")
                DWPerson.Visible = True
            Case 2  '�ק�
                DWPerson.BackColor = Color.Yellow
                DWPerson.Visible = True
            Case Else   '����
                DWPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WPerson", "ZZZZZZ")

        '�ק�z�����O
        Select Case FindFieldInf("MReasonType")
            Case 0  '���
                DMReasonType.BackColor = Color.LightGray
                DMReasonType.Visible = True
            Case 1  '�ק�+�ˬd
                DMReasonType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMReasonTypeRqd", "DMReasonType", "���`�G�ݿ�J�ק�z�����O")
                DMReasonType.Visible = True
            Case 2  '�ק�
                DMReasonType.BackColor = Color.Yellow
                DMReasonType.Visible = True
            Case Else   '����
                DMReasonType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MReasonType", "ZZZZZZ")

        '�ק�z��
        Select Case FindFieldInf("MReason")
            Case 0  '���
                DMReason.BackColor = Color.LightGray
                DMReason.ReadOnly = True
                DMReason.Visible = True
            Case 1  '�ק�+�ˬd
                DMReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMReasonRqd", "DMReason", "���`�G�ݿ�J�ק�z��")
                DMReason.Visible = True
            Case 2  '�ק�
                DMReason.BackColor = Color.Yellow
                DMReason.Visible = True
            Case Else   '����
                DMReason.Visible = False
        End Select
        If pPost = "New" Then DMReason.Text = ""

        '�ק鷺�e-1
        Select Case FindFieldInf("MContent1")
            Case 0  '���
                DMContent1.BackColor = Color.LightGray
                DMContent1.ReadOnly = True
                DMContent1.Visible = True
            Case 1  '�ק�+�ˬd
                DMContent1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMContent1Rqd", "DMContent1", "���`�G�ݿ�J�ק鷺�e-1")
                DMContent1.Visible = True
            Case 2  '�ק�
                DMContent1.BackColor = Color.Yellow
                DMContent1.Visible = True
            Case Else   '����
                DMContent1.Visible = False
        End Select
        If pPost = "New" Then DMContent1.Text = ""

        '�ק鷺�e-2
        Select Case FindFieldInf("MContent2")
            Case 0  '���
                DMContent2.BackColor = Color.LightGray
                DMContent2.ReadOnly = True
                DMContent2.Visible = True
            Case 1  '�ק�+�ˬd
                DMContent2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMContent2Rqd", "DMContent2", "���`�G�ݿ�J�ק鷺�e-2")
                DMContent2.Visible = True
            Case 2  '�ק�
                DMContent2.BackColor = Color.Yellow
                DMContent2.Visible = True
            Case Else   '����
                DMContent2.Visible = False
        End Select
        If pPost = "New" Then DMContent2.Text = ""

        '��ڭק鷺�e
        Select Case FindFieldInf("IContent")
            Case 0  '���
                DIContent.BackColor = Color.LightGray
                DIContent.ReadOnly = True
                DIContent.Visible = True
            Case 1  '�ק�+�ˬd
                DIContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DIContentRqd", "DIContent", "���`�G�ݿ�J��ڭק鷺�e")
                DIContent.Visible = True
            Case 2  '�ק�
                DIContent.BackColor = Color.Yellow
                DIContent.Visible = True
            Case Else   '����
                DIContent.Visible = False
        End Select
        If pPost = "New" Then DIContent.Text = ""

        '�w�w�����ɶ�
        Select Case FindFieldInf("FDateTime")
            Case 0  '���
                DFDateTime.BackColor = Color.LightGray
                DFDateTime.ReadOnly = True
                DFDateTime.Visible = True
            Case 1  '�ק�+�ˬd
                DFDateTime.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFDateTimeRqd", "DFDateTime", "���`�G�ݿ�J�w�w�����ɶ�")
                DFDateTime.Visible = True
            Case 2  '�ק�
                DFDateTime.BackColor = Color.Yellow
                DFDateTime.Visible = True
            Case Else   '����
                DFDateTime.Visible = False
        End Select
        If pPost = "New" Then DFDateTime.Text = ""

        '��ƭק���
        Select Case FindFieldInf("IPerson")
            Case 0  '���
                DIPerson.BackColor = Color.LightGray
                DIPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DIPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DIPersonRqd", "DIPerson", "���`�G�ݿ�J��ƭק���")
                DIPerson.Visible = True
            Case 2  '�ק�
                DIPerson.BackColor = Color.Yellow
                DIPerson.Visible = True
            Case Else   '����
                DIPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("IPerson", "ZZZZZZ")

        '�渹
        If pPost = "New" Then DFormSno.Text = ""
    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     �إߪ��ݿ�J���
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top + 25))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
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

        '�e�U����
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '�e�U���
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '�קﳡ��
        If pFieldName = "MDivision" Then
            DMDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = DDivision.SelectedValue Then ListItem1.Selected = True
                    DMDivision.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ק���
        If pFieldName = "MPerson" Then
            DMPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where DivName = '" & DMDivision.SelectedValue & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = DPerson.SelectedValue Then ListItem1.Selected = True
                    DMPerson.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ݳB�z�u�{����
        If pFieldName = "WDivision" Then
            DWDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where Active = '1' "
                SQL = SQL & " Group By DivName "
                SQL = SQL & " Order By DivName "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWDivision.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWDivision.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Value = DBTable1.Rows(i).Item("DivName")
                        ListItem2.Selected = False
                        DWDivision.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '�ݳB�z�u�{���
        If pFieldName = "WPerson" Then
            DWPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")

                For i = 0 To DBTable1.Rows.Count - 1
                    If pName <> "ZZZZZZ" Then
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        If ListItem2.Value = pName Then ListItem2.Selected = True
                        DWPerson.Items.Add(ListItem2)
                    Else
                        If i = 0 Then
                            Dim ListItem1 As New ListItem
                            ListItem1.Text = ""
                            ListItem1.Value = ""
                            ListItem1.Selected = True
                            DWPerson.Items.Add(ListItem1)
                        End If
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Selected = False
                        DWPerson.Items.Add(ListItem2)
                    End If
                Next
            End If
        End If

        '�e�U��
        If pFieldName = "Sheet" Then
            DSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select FormName From M_Form "
                SQL = SQL & " Where FormNo < 1000 "
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Form")
                DBTable1 = DBDataSet1.Tables("M_Form")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("FormName")
                    ListItem1.Value = DBTable1.Rows(i).Item("FormName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSheet.Items.Add(ListItem1)
                Next
            End If
        End If

        '���A
        If pFieldName = "Status" Then
            DStatus.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DStatus.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='STATUS' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DStatus.Items.Add(ListItem1)
                Next
            End If
        End If

        '�ק�z�����O
        If pFieldName = "MReasonType" Then
            DMReasonType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMReasonType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYREASON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMReasonType.Items.Add(ListItem1)
                Next
            End If
        End If

        '��ڭק���
        If pFieldName = "IPerson" Then
            DIPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DIPerson.Items.Add(ListItem1)

                End If
            Else
                SQL = "Select * From M_Referp Where Cat='900' and DKey='MODIFYPERSON' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DIPerson.Items.Add(ListItem1)
                Next

                SQL = " Select  data  From M_Referp a,M_users b  Where Cat='900' and DKey='MODIFYPERSON'  "
                SQL &= " and userid ='" & Request.QueryString("pUserID") & "'  and  substring(a.data,4,6)=b.username "
                SQL &= " Order by DKey, Data"

                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                DIPerson.SelectedValue = DBTable1.Rows(i).Item("Data")
                DFDateTime.Text = Format(Now, "MM/dd hh:mm")


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
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DReason.Text = pName
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    DReason.Text = DBTable1.Rows(i).Item("Data")
                Next
            End If
        End If
        'DB�s������
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
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
                    If ErrCode = 9010 Then Message = "�W���ɮ׮榡���~,�нT�{!"
                    If ErrCode = 9020 Then Message = "�W���ɮ�Size�W�L1024KB,�нT�{!"
                    If ErrCode = 9030 Then Message = "�W���ɮרS���w,�нT�{!"
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
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
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
    '**     NG1���s�I���ƥ�
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
    '**     NG2���s�I���ƥ�
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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        'Check�ݳB�z�u�{
        If ErrCode = 0 Then
            'If DStatus.SelectedValue = "01-�}�o��" Then
            If DWDivision.SelectedValue = "" Or DWPerson.SelectedValue = "" Then
                ErrCode = 9070
            End If
            'End If
        End If

        'Check����z��
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--�ˬd�e�U��No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("800001", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '�O�_����
            Dim RepeatRun As Boolean = False    '�O�_���а���
            Dim MultiJob As Integer = 0         '�h�H�֩w
            Dim wLevel As String = ""           '�˫~������

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
                Dim i, RtnCode As Integer
                Dim NewFormSno As Integer = wFormSno    '���y����
                Dim pRunNextStep As Integer = 0         '�O�_����p��U�@��(�|ñ)
                Dim SQL As String

                '���o���y�����Χ�s������
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
                        '���o���y����
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '��渹�X, ���y����
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '�e�UNo�۰ʽs�C
                            DNo.Text = DComNo.Text & "-" & CStr(NewFormSno)

                            '�ӽЬy�{��ƫظm
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                            'Modify-End
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
                    Dim DBDataSet1 As New DataSet
                    Dim wAllocateID As String = ""
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                    OleDbConnection1.Open()

                    'jessica 20150209 �ק�W�[
                    SQL = "Select UserID From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    If wStep = 20 Or wStep = 22 Or wStep = 27 Or wStep = 46 Then
                        SQL = SQL & "   And UserName =  '" & DWPerson.SelectedValue & "'"
                    Else
                        SQL = SQL & "   And UserName =  '" & DMPerson.SelectedValue & "'"
                    End If
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "M_Users")
                    If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                    OleDbConnection1.Close()
                    '
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
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If  'pSeqno <> 0
                    Else    '�P�_�O�_�_��
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
            If ErrCode = 9050 Then Message = "���謰��L�ɻݶ�g����,�нT�{!"
            If ErrCode = 9060 Then Message = "�e�U��No.����,�нT�{�e�U��No.!"
            If ErrCode = 9070 Then Message = "�ݶ�J�u�{�ݳB�z�����Τu�{�ݳB�z��.!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_ModifyDataSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, No, "                  '1~5
        SQl = SQl + "Date, Division, Person, MDivision, MPerson, "               '6~10
        SQl = SQl + "Sheet, ComNo, Status, WDivision, WPerson, "                 '11~15
        SQl = SQl + "MReasonType, MReason, MContent1, MContent2, IContent, FDateTime, "   '16~21
        SQl = SQl + "IPerson, "                                                  '22
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "            '23~26
        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '800001', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO

        SQl = SQl + " '" + DDate.Text + "', "                '���
        SQl = SQl + " '" + DDivision.SelectedValue + "', "   '�e�U����
        SQl = SQl + " '" + DPerson.SelectedValue + "', "     '�e�U���
        SQl = SQl + " '" + DMDivision.SelectedValue + "', "  '�קﳡ��
        SQl = SQl + " '" + DMPerson.SelectedValue + "', "    '�ק���

        SQl = SQl + " '" + DSheet.SelectedValue + "', "      '�e�U��
        SQl = SQl + " N'" + YKK.ReplaceString(DComNo.Text) + "', "   '�e�UNO
        SQl = SQl + " '" + DStatus.SelectedValue + "', "     '���A
        SQl = SQl + " '" + DWDivision.SelectedValue + "', "  '�ݳB�z�u�{����
        SQl = SQl + " '" + DWPerson.SelectedValue + "', "    '�ݳB�z�u�{���

        SQl = SQl + " '" + DMReasonType.SelectedValue + "', "   '�ק�z�����O
        SQl = SQl + " N'" + YKK.ReplaceString(DMReason.Text) + "', "     '�ק�z��
        SQl = SQl + " N'" + YKK.ReplaceString(DMContent1.Text) + "', "   '�ק鷺�e-1
        SQl = SQl + " N'" + YKK.ReplaceString(DMContent2.Text) + "', "   '�ק鷺�e-2
        SQl = SQl + " N'" + YKK.ReplaceString(DIContent.Text) + "', "    '��ڭק鷺�e
        SQl = SQl + " N'" + YKK.ReplaceString(DFDateTime.Text) + "', "   '��ڭק�ɶ�

        SQl = SQl + " '" + DIPerson.SelectedValue + "', "    '��ڭק���

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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_ModifyDataSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"

        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " MDivision = '" & DMDivision.SelectedValue & "',"
        SQl = SQl + " MPerson = '" & DMPerson.SelectedValue & "',"

        SQl = SQl + " Sheet = '" & DSheet.SelectedValue & "',"
        SQl = SQl + " ComNo = N'" & YKK.ReplaceString(DComNo.Text) & "',"
        SQl = SQl + " Status = '" & DStatus.SelectedValue & "',"
        SQl = SQl + " WDivision = '" & DWDivision.SelectedValue & "',"
        SQl = SQl + " WPerson = '" & DWPerson.SelectedValue & "',"

        SQl = SQl + " MReasonType = '" & DMReasonType.SelectedValue & "',"
        SQl = SQl + " MReason = N'" & YKK.ReplaceString(DMReason.Text) & "',"
        SQl = SQl + " MContent1 = N'" & YKK.ReplaceString(DMContent1.Text) & "',"
        SQl = SQl + " MContent2 = N'" & YKK.ReplaceString(DMContent2.Text) & "',"
        SQl = SQl + " IContent = N'" & YKK.ReplaceString(DIContent.Text) & "',"
        SQl = SQl + " FDateTime = N'" & YKK.ReplaceString(DFDateTime.Text) & "',"

        SQl = SQl + " IPerson = '" & DIPerson.SelectedValue & "',"

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
                SQl = SQl + "FormNo, FormSno, No, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
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
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
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
    '**     �C�L���
    '**
    '*****************************************************************
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
        Dim URL As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()
        SQL = "Select ViewURL From V_WaitHandle_01 "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "URL")
        If DBDataSet1.Tables("URL").Rows.Count > 0 Then
            URL = DBDataSet1.Tables("URL").Rows(0).Item("ViewURL")
        End If
        'DB�s������
        OleDbConnection1.Close()

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenPrintSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �y�{�������
    '**
    '*****************************************************************
    Private Sub BFlow_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFlow.Click
        Dim URL As String = ""
        Dim SQL As String = ""
        Dim wLevel As String = ""           '�˫~������
        Dim wPerson As String = ""         '�s�Ϫ�

        If wStep = 1 Or wStep = 46 Then
            If DWPerson.SelectedValue <> "" Then
                Dim DBDataSet1 As New DataSet
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                OleDbConnection1.Open()

                SQL = "Select UserID From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserName =  '" & DWPerson.SelectedValue & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wPerson = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                OleDbConnection1.Close()
            End If
        Else
            If DMPerson.SelectedValue <> "" Then
                Dim DBDataSet1 As New DataSet
                Dim OleDbConnection1 As New OleDbConnection
                OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                OleDbConnection1.Open()

                SQL = "Select UserID From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserName =  '" & DMPerson.SelectedValue & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wPerson = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

                OleDbConnection1.Close()
            End If
        End If


        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel & _
                              "&pAllocateID=" & wPerson

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenSimulationSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub

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
        '����
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '���
        If InputCheck = 0 Then
            If FindFieldInf("Person") = 1 Then
                If DPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�קﳡ��
        If InputCheck = 0 Then
            If FindFieldInf("MDivision") = 1 Then
                If DMDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ק���
        If InputCheck = 0 Then
            If FindFieldInf("MPerson") = 1 Then
                If DMPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�e�U��
        If InputCheck = 0 Then
            If FindFieldInf("Sheet") = 1 Then
                If DSheet.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�e�UNo
        If InputCheck = 0 Then
            If FindFieldInf("ComNo") = 1 Then
                If DComNo.Text = "" Then InputCheck = 1
            End If
        End If
        '���A
        If InputCheck = 0 Then
            If FindFieldInf("Status") = 1 Then
                If DStatus.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ݳB�z�u�{����
        If InputCheck = 0 Then
            If FindFieldInf("WDivision") = 1 Then
                If DWDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ݳB�z�u�{���
        If InputCheck = 0 Then
            If FindFieldInf("WPerson") = 1 Then
                If DWPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ק�z�����O
        If InputCheck = 0 Then
            If FindFieldInf("MReasonType") = 1 Then
                If DMReasonType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�ק�z��
        If InputCheck = 0 Then
            If FindFieldInf("MReason") = 1 Then
                If DMReason.Text = "" Then InputCheck = 1
            End If
        End If
        '�ק鷺�e-1
        If InputCheck = 0 Then
            If FindFieldInf("MContent1") = 1 Then
                If DMContent1.Text = "" Then InputCheck = 1
            End If
        End If
        '�ק鷺�e-2
        If InputCheck = 0 Then
            If FindFieldInf("MContent2") = 1 Then
                If DMContent2.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڭק鷺�e
        If InputCheck = 0 Then
            If FindFieldInf("IContent") = 1 Then
                If DIContent.Text = "" Then InputCheck = 1
            End If
        End If
        '�w�w�����ɶ�
        If InputCheck = 0 Then
            If FindFieldInf("FDateTime") = 1 Then
                If DFDateTime.Text = "" Then InputCheck = 1
            End If
        End If
        '��ƭק���
        If InputCheck = 0 Then
            If FindFieldInf("IPerson") = 1 Then
                If DIPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Add-Start 2011/12/2 ����:�ݥθ�ƭק�e�U��
        '                    �ﵦ:1�H��1�i(Active=1)
        If wFormSno = 0 And wStep < 3 Then      ' �_��
            If oCommon.CanUseSheet("800001", Request.QueryString("pUserID")) = 1 Then
                InputCheck = 1
                Response.Write(YKK.ShowMessage("�W�@�i��ƭק�e�U�ѥ����� ! "))
            End If
        End If

    End Function
End Class
