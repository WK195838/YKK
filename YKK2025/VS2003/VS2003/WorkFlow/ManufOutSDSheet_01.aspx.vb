Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManufOutSDSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderGRCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOut As System.Web.UI.WebControls.Button
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DManufOutSDSheet As System.Web.UI.WebControls.Image
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSliderFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BImport As System.Web.UI.WebControls.Button

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

        Response.Cookies("DevNo").Value = ""        '�}�oNo, DevNoPicker�ϥ�
        Response.Cookies("SliderGRCode").Value = "" 'Slider Group Code, SliderGRCodePicker�ϥ�

        Response.Cookies("PGM").Value = "ManufOutSDSheet_01.aspx"
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
        'Check���Y�ӥت�
        If DSliderFile.Visible Then
            If DSliderFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "���Y�ӥت�"
                Else
                    Message = Message & ", " & "���Y�ӥت�"
                End If
            End If
        End If
        'Check����
        If DAttachFile.Visible Then
            If DAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "����"
                Else
                    Message = Message & ", " & "����"
                End If
            End If
        End If

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufOutSDFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufOutSDSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufOutSDSheet")
        If DBDataSet1.Tables("F_ManufOutSDSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("Person"))      '���
            DSliderGRCode.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("SliderGRCode")       'Slider Code
            DOFormNo.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("OFormNo")             '�ϸ�
            DOFormSno.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("OFormSno")             '�ϸ�
            DContent.Text = DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("Content")
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000007" Then
                    LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                End If

            Else
                LOFormNo.Visible = False
            End If
            If DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("SliderFile") <> "" Then          '����1
                LSliderFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("SliderFile")
            Else
                LSliderFile.Visible = False
            End If
            If DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("AttachFile") <> "" Then          '����1
                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufOutSDSheet").Rows(0).Item("AttachFile")
            Else
                LAttachFile.Visible = False
            End If

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
                DManufOutSDSheet.Visible = True   '���Sheet-1
                DDescSheet.Visible = True       '����Sheet
                DDelivery.Visible = True        '���Sheet

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
                    Top = 616
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 504
                End If

                '���
                DBStartTime.Visible = True      '�w�w�}�l
                DBEndTime.Visible = True        '�w�w����
                DAStartTime.Visible = True      '��ڶ}�l
                DAEndTime.Visible = True        '��ڧ���
                DOPReady.Visible = True     '�u�{�i���\Ū
                DOPReadyDesc.Visible = True
                '�ݾ\Ū
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    DOPReady.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "���`�G�ݾ\Ū�u�{�i��")
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
                LOFormNo.Visible = True        '��e�U
                LSliderFile.Visible = True     '����
                LAttachFile.Visible = True     '����
                LBefOP.Visible = True          '�u�{�i��
                '���s��m
                BNG1.Style.Add("Top", Top)      'NG���s
                BNG2.Style.Add("Top", Top)      'NG���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
            Top = 320
            'Sheet����
            DDescSheet.Visible = False  '����Sheet
            DDelivery.Visible = False   '���Sheet
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

            DOPReady.Visible = False    '�u�{�i���\Ū
            DOPReadyDesc.Visible = False

            '�s������
            LOFormNo.Visible = False        '��e�U
            LSliderFile.Visible = False     '����
            LAttachFile.Visible = False     '����
            LBefOP.Visible = False          '�u�{�i��
            '���s��m
            BNG1.Style.Add("Top", Top)       'NG���s
            BNG2.Style.Add("Top", Top)       'NG���s
            BSAVE.Style.Add("Top", Top)     '�x�s���s
            BOK.Style.Add("Top", Top)       'OK���s
            DFormSno.Style.Add("Top", Top)  '�渹
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
        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"  '������
        BOut.Attributes("onclick") = "DevNoPicker('Out','000010');"         '���s
        BImport.Attributes("onclick") = "DevNoPicker('Import','000010');"     '�i�f/�Ȩѩ��Y�e�U��-No
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
                    Top = 504
                Else
                    If DDelay.Visible = True Then
                        Top = 616
                    Else
                        Top = 504
                    End If
                End If
            End If
        Else
            Top = 320
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
                BOut.Visible = False
                BImport.Visible = False
            Case 1  '�ק�+�ˬd
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "���`�G�ݿ�J�ܢ�")
                DNo.Visible = True
                BOut.Visible = False
                BImport.Visible = True
            Case 2  '�ק�
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
                BOut.Visible = False
                BImport.Visible = True
            Case Else   '����
                DNo.Visible = False
                BOut.Visible = False
                BImport.Visible = False

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
        '����
        Select Case FindFieldInf("Division")
            Case 0  '���
                DDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
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
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")
        '���
        Select Case FindFieldInf("Person")
            Case 0  '���
                DPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
                DPerson.Visible = True
            Case 1  '�ק�+�ˬd
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "���`�G�ݿ�J���")
                DPerson.Visible = True
            Case 2  '�ק�
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '����
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")
        'Slider Code
        Select Case FindFieldInf("SliderGRCode")
            Case 0  '���
                DSliderGRCode.BackColor = Color.LightGray
                DSliderGRCode.ReadOnly = True
                DSliderGRCode.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderGRCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderGRCodeRqd", "DSliderGRCode", "���`�G�ݿ�J Slider Code")
                DSliderGRCode.Visible = True
            Case 2  '�ק�
                DSliderGRCode.BackColor = Color.Yellow
                DSliderGRCode.Visible = True
            Case Else   '����
                DSliderGRCode.Visible = False
        End Select
        If pPost = "New" Then DSliderGRCode.Text = ""
        'OFormNo
        Select Case FindFieldInf("OFormNo")
            Case 0  '���
                DOFormNo.BackColor = Color.LightGray
                DOFormNo.ReadOnly = True
                DOFormNo.Visible = True
            Case 1  '�ק�+�ˬd
                DOFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormNoRqd", "DOFormNo", "���`�G�ݿ�J��渹�X")
                DOFormNo.Visible = True
            Case 2  '�ק�
                DOFormNo.BackColor = Color.Yellow
                DOFormNo.Visible = True
            Case Else   '����
                DOFormNo.Visible = False
        End Select
        If pPost = "New" Then DOFormNo.Text = ""
        'OFormSno
        Select Case FindFieldInf("OFormSno")
            Case 0  '���
                DOFormSno.BackColor = Color.LightGray
                DOFormSno.ReadOnly = True
                DOFormSno.Visible = True
            Case 1  '�ק�+�ˬd
                DOFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormSnoRqd", "DOFormSno", "���`�G�ݿ�J�渹")
                DOFormSno.Visible = True
            Case 2  '�ק�
                DOFormSno.BackColor = Color.Yellow
                DOFormSno.Visible = True
            Case Else   '����
                DOFormSno.Visible = False
        End Select
        If pPost = "New" Then DOFormSno.Text = ""
        'Content
        Select Case FindFieldInf("Content")
            Case 0  '���
                DContent.BackColor = Color.LightGray
                DContent.ReadOnly = True
                DContent.Visible = True
            Case 1  '�ק�+�ˬd
                DContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContentRqd", "DContent", "���`�G�ݿ�J�������e")
                DContent.Visible = True
            Case 2  '�ק�
                DContent.BackColor = Color.Yellow
                DContent.Visible = True
            Case Else   '����
                DContent.Visible = False
        End Select
        If pPost = "New" Then DContent.Text = ""
        '���Y�ӥ�
        Select Case FindFieldInf("SliderFile")
            Case 0  '���
                DSliderFile.Visible = False
                DSliderFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DSliderFileRqd", "DSliderFile", "���`�G�ݿ�J���Y�ӥت�")
                DSliderFile.Visible = True
                DSliderFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DSliderFile.Visible = True
                DSliderFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DSliderFile.Visible = False
        End Select
        '����
        Select Case FindFieldInf("AttachFile")
            Case 0  '���
                DAttachFile.Visible = False
                DAttachFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DAttachFileRqd", "DAttachFile", "���`�G�ݿ�J����")
                DAttachFile.Visible = True
                DAttachFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DAttachFile.Visible = True
                DAttachFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DAttachFile.Visible = False
        End Select
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

        '����
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
        '���
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
        '����z�ѥN�X
        If pFieldName = "ReasonCode" Then
            SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
            DReasonCode.Items.Clear()
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

                'Check���Y�ӥ�
                If ErrCode = 0 Then
                    If DSliderFile.Visible Then
                        If DSliderFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DSliderFile)
                        End If
                    End If
                End If
                'Check����Size�ή榡
                If ErrCode = 0 Then
                    If DAttachFile.Visible Then
                        If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DAttachFile)
                        End If
                    End If
                End If

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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        'Check���Y�ӥ�
        If ErrCode = 0 Then
            If DSliderFile.Visible Then
                If DSliderFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DSliderFile)
                End If
            End If
        End If
        'Check����Size�ή榡
        If ErrCode = 0 Then
            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DAttachFile)
                End If
            End If
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
                ErrCode = oCommon.CommissionNo("000010", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                    If wFormSno = 0 And wStep < 3 Then      '�P�_�O�_�_��
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
                                wNextGate = wNextGate & "," & pNextGate(i)
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
                            AppendData(pFun, NewFormSno)    '�s�W�����
                            AddCommissionNo(wFormNo, NewFormSno)
                            ModifyManufOutData(pFun, 1, 1)   '��s���檬�A
                        End If  'pSeqno <> 0
                    Else    '�P�_�O�_�_��
                        If pNextStep = 999 Then     '�u�{������?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '��s�����
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            ModifyManufOutData(pFun, 0, 1)   '��s���檬�A
                        Else
                            ModifyData(pFun, "0")         '��s����� Sts=0(����)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)

                    End If  'Insert / Update
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
            If ErrCode = 9060 Then Message = "�e�U��No.����,�нT�{�e�U��No.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '�W���ɮ�ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufOutSDFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_ManufOutSDSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, "          '1~4
        SQl = SQl + "No, Date, Division, Person, SliderGRCode, "       '5~9
        SQl = SQl + "OFormNo, OFormSno, Content, SliderFile, AttachFile, "    '10~14
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "  '15~18
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '000010', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "                 'NO
        SQl = SQl + " '" + DDate.Text + "', "               '���
        SQl = SQl + " '" + DDivision.SelectedValue + "', "  '����
        SQl = SQl + " '" + DPerson.SelectedValue + "', "    '���
        SQl = SQl + " '" + YKK.ReplaceString(DSliderGRCode.Text) + "', "         'Slider Code
        SQl = SQl + " '" + DOFormNo.Text + "', "            '���No
        If DOFormSno.Text = "" Then                         '�渹
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DOFormSno.Text + "', "
        End If
        SQl = SQl + " N'" + YKK.ReplaceString(DContent.Text) + "', "            '�������e

        FileName = ""
        If DSliderFile.Visible Then                         '����
            If DSliderFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "SliderDetail" & "-" & UploadDateTime & "-" & Right(DSliderFile.PostedFile.FileName, InStr(StrReverse(DSliderFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "SliderDetail" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSliderFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DSliderFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DAttachFile.Visible Then                         '����
            If DAttachFile.PostedFile.FileName <> "" Then     '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufOutSDFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand



        SQl = "Update F_ManufOutSDSheet Set "

        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " SliderGRCode = '" & YKK.ReplaceString(DSliderGRCode.Text) & "',"
        SQl = SQl + " OFormNo = '" & DOFormNo.Text & "',"
        SQl = SQl + " OFormSno = '" & DOFormSno.Text & "',"
        SQl = SQl + " Content = N'" & YKK.ReplaceString(DContent.Text) & "',"

        If DSliderFile.Visible Then
            If DSliderFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "SliderDetail" & "-" & UploadDateTime & "-" & Right(DSliderFile.PostedFile.FileName, InStr(StrReverse(DSliderFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "SliderDetail" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSliderFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DSliderFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " SliderFile = N'" & FileName & "',"
            End If
        End If

        If DAttachFile.Visible Then
            If DAttachFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " AttachFile = N'" & FileName & "',"
            End If
        End If
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
    '**(ModifyManufOutData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyManufOutData(ByVal pFun As String, ByVal pSts As Integer, ByVal pContact As Integer)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand


        If DOFormNo.Text = "000007" Then
            SQl = "Update F_ManufOutSheet Set "
        End If
        If DOFormNo.Text = "000011" Then
            SQl = "Update F_ImportSheet Set "
        End If


        SQl = SQl + " SDSts = '" & CStr(pSts) & "',"
        SQl = SQl + " SliderDetail = '" & CStr(pContact) & "',"
        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where SliderGRCode =  '" & DSliderGRCode.Text & "'"
        ' SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
        ' SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
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
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx", ".docx", ".pptx"}   '�w�q���\���ɮ׮榡
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
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 * 2 Then
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
        Dim wLevel As String = "" '������

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel

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
        'Slider Code
        If InputCheck = 0 Then
            If FindFieldInf("SliderGRCode") = 1 Then
                If DSliderGRCode.Text = "" Then InputCheck = 1
            End If
        End If
        'OFormNo
        If InputCheck = 0 Then
            If FindFieldInf("OFormNo") = 1 Then
                If DOFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        'OFormSno
        If InputCheck = 0 Then
            If FindFieldInf("OFormSno") = 1 Then
                If DOFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        'Content
        If InputCheck = 0 Then
            If FindFieldInf("Content") = 1 Then
                If DContent.Text = "" Then InputCheck = 1
            End If
        End If
        '���Y�ӥ�
        If InputCheck = 0 Then
            If FindFieldInf("SliderFile") = 1 Then
                If DSliderFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("AttachFile") = 1 Then
                If DAttachFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
