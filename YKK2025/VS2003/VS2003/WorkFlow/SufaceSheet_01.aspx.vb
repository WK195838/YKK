Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SufaceSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOrderTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCResult2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCRemark3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCResult1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEADesc1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck12 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck11 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck10 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck9 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck8 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LFinalSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DEnglishName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBFinalDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBFinalDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAllowSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DColor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOrderTime As System.Web.UI.WebControls.Button
    Protected WithEvents BReqDelDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderSample As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqQty As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReqDelDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSufaceSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LQCReqFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOPManualFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LEACheckFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFinalFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LContactFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents LForcastFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LManufFlowFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAttachSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DORNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSpec As System.Web.UI.WebControls.Button
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DSufaceSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents LCustSampleFile As System.Web.UI.WebControls.Image
    Protected WithEvents DFinalSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCReqFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DCustSampleFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DOPManualFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFinalFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DContactFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DForcastFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DManufFlowFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DEACheckFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BQCDate1 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate2 As System.Web.UI.WebControls.Button
    Protected WithEvents BQCDate3 As System.Web.UI.WebControls.Button
    Protected WithEvents DQCLT As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents BOut As System.Web.UI.WebControls.Button
    Protected WithEvents BImport As System.Web.UI.WebControls.Button
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCap As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSchedule As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEACheckFile1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LEACheckFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCCheck13 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DQCCheck14 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DLOSS As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCHECK15 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYearSeason As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCCHECK16 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LPFASFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPFASFile As System.Web.UI.HtmlControls.HtmlInputFile

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

        Response.Cookies("DevNo").Value = ""        '�}�oNo, DevNoPicker�ϥ�
        Response.Cookies("PGM").Value = "SufaceSheet_01.aspx"
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
        'Check�Ȥ�˫~��
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then
                Message = "�Ȥ�˫~��"
            End If
        End If
        'Check�̲׼˫~��
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�̲׼˫~��"
                Else
                    Message = Message & ", " & "�̲׼˫~��"
                End If
            End If
        End If
        'Check�~��̿��
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~��̿��"
                Else
                    Message = Message & ", " & "�~��̿��"
                End If
            End If
        End If
        'Check�q�ὤ�p
        ' If DQCCheck6File.Visible Then
        ' If DQCCheck6File.PostedFile.FileName <> "" Then
        'If Message = "" Then
        'Message = "�q�ὤ�p"
        'Else
        '   Message = Message & ", " & "�q�ὤ�p"
        'End If
        ' End If
        ' End If
        'Check Oeko-tex���`������i
        If DEACheckFile.Visible Then
            If DEACheckFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "Oeko-tex���`������i"
                Else
                    Message = Message & ", " & "Oeko-tex���`������i"
                End If
            End If
        End If
        'Check A01���`������i
        If DEACheckFile1.Visible Then
            If DEACheckFile1.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "A01���`������i"
                Else
                    Message = Message & ", " & "A01���`������i"
                End If
            End If
        End If
        'Check���ճ��i��
        If DQCFinalFile.Visible Then
            If DQCFinalFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "���ճ��i��"
                Else
                    Message = Message & ", " & "���ճ��i��"
                End If
            End If
        End If

        'Check���ճ��i��
        If DPFASFile.Visible Then
            If DPFASFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "PFAS���i"
                Else
                    Message = Message & ", " & "PFAS���i"
                End If
            End If
        End If


        'Check�s�y�y�{��
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�s�y�y�{��"
                Else
                    Message = Message & ", " & "�s�y�y�{��"
                End If
            End If
        End If
        'Check������
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "������"
                Else
                    Message = Message & ", " & "������"
                End If
            End If
        End If
        'Check�@�~�зǮ�
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�@�~�зǮ�"
                Else
                    Message = Message & ", " & "�@�~�зǮ�"
                End If
            End If
        End If
        'Check������
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "������"
                Else
                    Message = Message & ", " & "������"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SufaceFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_SufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
        If DBDataSet1.Tables("F_SufaceSheet").Rows.Count > 0 Then
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile") <> "" Then        '�Ȥ�˫~��
                LCustSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("CustSampleFile")
            Else
                LCustSampleFile.Visible = False
            End If
            DNo.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Person"))      '���
            DSpec.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Spec")                   '�W��
            SetFieldData("Buyer", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Buyer"))        'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SellVendor")       '�e�U�t��
            DReqDelDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqDelDate")       '�Ʊ���
            DReqQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ReqQty")               '�w���q
            DSliderSample.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SliderSample")   '�˫~���Y
            SetFieldData("AttachSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AttachSample"))    '����
            DORNO.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ORNO")                   'OR-NO
            DOrderTime.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OrderTime")         '�U��ɶ�
            DPrice.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Price")                 '����
            DDevReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("DevReason")         '�}�o�z��
            SetFieldData("YearSeason", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("YearSeason"))    '�~�u
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile") <> "" Then       '�̲׼˫~��
                LFinalSampleFile.ImageUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FinalSampleFile")
            Else
                LFinalSampleFile.Visible = False
            End If
            SetFieldData("ManufType", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufType"))  '���s/�~�`
            SetFieldData("Suppiler", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Suppiler"))    '�~�`��
            DColor.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Color")                   'Color
            DQty.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Qty")                       '�ƶq
            DCap.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Cap")
            DSchedule.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Schedule")
            DFReason.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FReason")


            SetFieldData("AllowSample", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("AllowSample"))    '���׼˫~
            DBFinalDate.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("BFinalDate")        '�w�w������
            DCode.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Code")                    'Code
            DEnglishName.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EnglishName")      '�^��W��
            DLOSS.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("LOSS")                    'LOSS
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile") <> "" Then              '�~��̿��
                LQCReqFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCReqFile")
            Else
                LQCReqFile.Visible = False
            End If

            SetFieldData("QCCheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck1"))   '�f�|�o�k
            SetFieldData("QCCheck2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck2"))   '�P�ʩ��
            SetFieldData("QCCheck3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck3"))   'LOCK�j��
            SetFieldData("QCCheck4", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck4"))   '90�ױj��
            SetFieldData("QCCheck5", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck5"))   '��O
            SetFieldData("QCCheck15", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck15"))   'N-ANTI
            SetFieldData("QCCheck16", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck16")) 'FAFS


            '  If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File") <> "" Then           '�q�ὤ�p
            '  LQCCheck6File.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck6File")
            'Else
            '   LQCCheck6File.Visible = False
            'End If

            SetFieldData("QCCheck7", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck7"))   '�˰w
            SetFieldData("QCCheck8", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck8"))   'AATCC
            SetFieldData("QCCheck9", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck9"))   '���~
            SetFieldData("QCCheck10", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck10")) '�Q���Q��
            SetFieldData("QCCheck11", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck11")) '�@���K��
            SetFieldData("QCCheck12", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck12")) '�G���K��
            SetFieldData("QCCheck13", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck13")) 'Oeko-tex
            SetFieldData("QCCheck14", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCCheck14")) 'A01
            SetFieldData("EACheck1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheck1"))   '���`����
            DEADesc1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EADesc1")              '���`����Ƶ�

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile") <> "" Then            'Oeko-tex���`������i
                LEACheckFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile")
            Else
                LEACheckFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1") <> "" Then            'A01���`������i
                LEACheckFile1.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("EACheckFile1")
            Else
                LEACheckFile1.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile") <> "" Then            '���ճ��i��
                LQCFinalFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCFinalFile")
            Else
                LQCFinalFile.Visible = False
            End If


            DQCDate1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate1")               '���-1
            SetFieldData("QCResult1", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult1"))  '�˴����G-1
            DQCRemark1.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark1")           '�Ƶ�-1
            DQCDate2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate2")               '���-2
            SetFieldData("QCResult2", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult2"))  '�˴����G-2
            DQCRemark2.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark2")           '�Ƶ�-2
            DQCDate3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCDate3")               '���-3
            SetFieldData("QCResult3", DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCResult3"))  '�˴����G-3
            DQCRemark3.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCRemark3")           '�Ƶ�-3

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile") <> "" Then           '�s�y�y�{��
                LManufFlowFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ManufFlowFile")
            Else
                LManufFlowFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile") <> "" Then             '������
                LForcastFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ForcastFile")
            Else
                LForcastFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile") <> "" Then            '�@�~�зǮ�
                LOPManualFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OPManualFile")
            Else
                LOPManualFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile") <> "" Then             '������
                LContactFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("ContactFile")
            Else
                LContactFile.Visible = False
            End If

            If DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile") <> "" Then             'PFAS���i
                LPFASFile.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("PFASFile")
            Else
                LPFASFile.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OFormNo")             '����
            DOFormSno.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("OFormSno")           '��渹

            DQCLT.Text = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("QCLT")    'QC-L/T

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
                DSufaceSheet1.Visible = True   '���Sheet-1
                DSufaceSheet2.Visible = True   '���Sheet-2
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
                    Top = 1496
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 1384
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
                LCustSampleFile.Visible = True   '�Ȥ�˫~��
                LFinalSampleFile.Visible = True  '�̲׼˫~��
                LQCReqFile.Visible = True        '�~��̿��
                '  LQCCheck6File.Visible = True     '�q�ὤ�p
                LEACheckFile.Visible = True      'Oeko-tex���`������i
                LEACheckFile1.Visible = True     'A01���`������i
                LQCFinalFile.Visible = True      '���ճ��i��
                LPFASFile.Visible = True      '���ճ��i��

                LManufFlowFile.Visible = True    '�s�y�y�{��
                LForcastFile.Visible = True      '������
                LOPManualFile.Visible = True     '�@�~�зǮ�
                LContactFile.Visible = True      '������
                LBefOP.Visible = True            '�u�{�i��
                '���s��m
                BNG1.Style.Add("Top", Top)     'NG1���s
                BNG2.Style.Add("Top", Top)     'NG2���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s
                DFormSno.Style.Add("Top", Top) '�渹
            End If
        Else
            Top = 1200
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
            LCustSampleFile.Visible = False  '�Ȥ�˫~��
            LFinalSampleFile.Visible = False '�̲׼˫~��
            LQCReqFile.Visible = False       '�~��̿��
            '  LQCCheck6File.Visible = False    '�q�ὤ�p
            LEACheckFile.Visible = False     'Oeko-tex���`������i
            LEACheckFile1.Visible = False    'A01���`������i
            LQCFinalFile.Visible = False     '���ճ��i��
            LPFASFile.Visible = False     'FAfS
            LManufFlowFile.Visible = False   '�s�y�y�{��
            LForcastFile.Visible = False     '������
            LOPManualFile.Visible = False    '�@�~�зǮ�
            LContactFile.Visible = False     '������
            LBefOP.Visible = False           '�u�{�i��
            '���s��m
            BNG1.Style.Add("Top", Top)     'NG1���s
            BNG2.Style.Add("Top", Top)     'NG2���s
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
        BIn.Attributes("onclick") = "DevNoPicker('In','000014');"       '���s�e�U��-No
        BOut.Attributes("onclick") = "DevNoPicker('Out','000014');"     '�~�`�e�U��-No
        BImport.Attributes("onclick") = "DevNoPicker('Import','000014');"     '�i�f/�Ȩѩ��Y�e�U��-No

        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"              '���
        BQCDate1.Attributes("onclick") = "CalendarPicker('Form1.DQCDate1');"        '���
        BQCDate2.Attributes("onclick") = "CalendarPicker('Form1.DQCDate2');"        '���
        BQCDate3.Attributes("onclick") = "CalendarPicker('Form1.DQCDate3');"        '���

        BSpec.Attributes("onclick") = "SpecPicker('Form1.DSpec', 'SUFACE');"        'Spec
        BReqDelDate.Attributes("onclick") = "CalendarPicker('Form1.DReqDelDate');"  '�Ʊ���
        BOrderTime.Attributes("onclick") = "CalendarPicker('Form1.DOrderTime');"    '�U��ɶ�
        BBFinalDate.Attributes("onclick") = "CalendarPicker('Form1.DBFinalDate');"  '�w�w������
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
                    Top = 1384
                Else
                    If DDelay.Visible = True Then
                        Top = 1496
                    Else
                        Top = 1384
                    End If
                End If
            End If
        Else
            Top = 1200
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
        '�Ȥ�˫~��
        Select Case FindFieldInf("CustSampleFile")
            Case 0  '���
                DCustSampleFile.Visible = False
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DCustSampleFileRqd", "DCustSampleFile", "���`�G�ݿ�J�Ȥ�˫~��")
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DCustSampleFile.Visible = True
                DCustSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DCustSampleFile.Visible = False
        End Select

        'No
        Select Case FindFieldInf("No")
            Case 0  '���
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
                BIn.Visible = False
                BOut.Visible = False
                BImport.Visible = False
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

        '����
        Select Case FindFieldInf("Division")
            Case 0  '���
                DDivision.BackColor = Color.LightGray
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

        'Spec
        Select Case FindFieldInf("Spec")
            Case 0  '���
                DSpec.BackColor = Color.LightGray
                DSpec.ReadOnly = True
                DSpec.Visible = True
                BSpec.Visible = False
            Case 1  '�ק�+�ˬd
                DSpec.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSpecRqd", "DSpec", "���`�G�ݿ�J�W��(Size, Chain Type, ����)")
                DSpec.Visible = True
                BSpec.Visible = True
            Case 2  '�ק�
                DSpec.BackColor = Color.Yellow
                DSpec.Visible = True
                BSpec.Visible = True
            Case Else   '����
                DSpec.Visible = False
                BSpec.Visible = False
        End Select
        If pPost = "New" Then DSpec.Text = ""

        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '���
                DBuyer.BackColor = Color.LightGray
                DBuyer.Visible = True
            Case 1  '�ק�+�ˬd
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "���`�G�ݿ�JBuyer")
                DBuyer.Visible = True
            Case 2  '�ק�
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '����
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")

        '�e�U�t��
        Select Case FindFieldInf("SellVendor")
            Case 0  '���
                DSellVendor.BackColor = Color.LightGray
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '�ק�+�ˬd
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "���`�G�ݿ�J�e�U�t��")
                DSellVendor.Visible = True
            Case 2  '�ק�
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '����
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""

        '�Ʊ���
        Select Case FindFieldInf("ReqDelDate")
            Case 0  '���
                DReqDelDate.BackColor = Color.LightGray
                DReqDelDate.ReadOnly = True
                DReqDelDate.Visible = True
                BReqDelDate.Visible = False
            Case 1  '�ק�+�ˬd
                DReqDelDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqDelDateRqd", "DReqDelDate", "���`�G�ݿ�J�Ʊ���")
                DReqDelDate.Visible = True
                BReqDelDate.Visible = True
            Case 2  '�ק�
                DReqDelDate.BackColor = Color.Yellow
                DReqDelDate.Visible = True
                BReqDelDate.Visible = True
            Case Else   '����
                DReqDelDate.Visible = False
                BReqDelDate.Visible = False
        End Select
        If pPost = "New" Then DReqDelDate.Text = CStr(DateTime.Now.Today)

        '�w���q
        Select Case FindFieldInf("ReqQty")
            Case 0  '���
                DReqQty.BackColor = Color.LightGray
                DReqQty.ReadOnly = True
                DReqQty.Visible = True
            Case 1  '�ק�+�ˬd
                DReqQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqQtyRqd", "DReqQty", "���`�G�ݿ�J�w���q")
                DReqQty.Visible = True
            Case 2  '�ק�
                DReqQty.BackColor = Color.Yellow
                DReqQty.Visible = True
            Case Else   '����
                DReqQty.Visible = False
        End Select
        If pPost = "New" Then DReqQty.Text = ""

        '�˫~���Y
        Select Case FindFieldInf("SliderSample")
            Case 0  '���
                DSliderSample.BackColor = Color.LightGray
                DSliderSample.ReadOnly = True
                DSliderSample.Visible = True
            Case 1  '�ק�+�ˬd
                DSliderSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderSampleRqd", "DSliderSample", "���`�G�ݿ�J�˫~���Y")
                DSliderSample.Visible = True
            Case 2  '�ק�
                DSliderSample.BackColor = Color.Yellow
                DSliderSample.Visible = True
            Case Else   '����
                DSliderSample.Visible = False
        End Select
        If pPost = "New" Then DSliderSample.Text = ""

        '����
        Select Case FindFieldInf("AttachSample")
            Case 0  '���
                DAttachSample.BackColor = Color.LightGray
                DAttachSample.Visible = True
            Case 1  '�ק�+�ˬd
                DAttachSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAttachSampleRqd", "DAttachSample", "���`�G�ݿ�J����")
                DAttachSample.Visible = True
            Case 2  '�ק�
                DAttachSample.BackColor = Color.Yellow
                DAttachSample.Visible = True
            Case Else   '����
                DAttachSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AttachSample", "ZZZZZZ")

        '�~�u
        Select Case FindFieldInf("YearSeason")
            Case 0  '���
                DYearSeason.BackColor = Color.LightGray
                DYearSeason.Visible = True
            Case 1  '�ק�+�ˬd
                DYearSeason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DYearSeasonRqd", "DYearSeason", "���`�G�ݿ�~�Ωu")
                DYearSeason.Visible = True
            Case 2  '�ק�
                DYearSeason.BackColor = Color.Yellow
                DYearSeason.Visible = True
            Case Else   '����
                DYearSeason.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("YearSeason", "ZZZZZZ")

        'OR-NO
        Select Case FindFieldInf("ORNO")
            Case 0  '���
                DORNO.BackColor = Color.LightGray
                DORNO.ReadOnly = True
                DORNO.Visible = True
            Case 1  '�ק�+�ˬd
                DORNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "���`�G�ݿ�JOR-NO")
                DORNO.Visible = True
            Case 2  '�ק�
                DORNO.BackColor = Color.Yellow
                DORNO.Visible = True
            Case Else   '����
                DORNO.Visible = False
        End Select
        If pPost = "New" Then DORNO.Text = ""

        '�U��ɶ�
        Select Case FindFieldInf("OrderTime")
            Case 0  '���
                DOrderTime.BackColor = Color.LightGray
                DOrderTime.ReadOnly = True
                DOrderTime.Visible = True
                BOrderTime.Visible = False
            Case 1  '�ק�+�ˬd
                DOrderTime.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOrderTimeRqd", "DOrderTime", "���`�G�ݿ�J�U��ɶ�")
                DOrderTime.Visible = True
            Case 2  '�ק�
                DOrderTime.BackColor = Color.Yellow
                DOrderTime.Visible = True
            Case Else   '����
                DOrderTime.Visible = False
        End Select
        If pPost = "New" Then DOrderTime.Text = ""

        '����
        Select Case FindFieldInf("Price")
            Case 0  '���
                DPrice.BackColor = Color.LightGray
                DPrice.ReadOnly = True
                DPrice.Visible = True
            Case 1  '�ק�+�ˬd
                DPrice.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPriceRqd", "DPrice", "���`�G�ݿ�J����")
                DPrice.Visible = True
            Case 2  '�ק�
                DPrice.BackColor = Color.Yellow
                DPrice.Visible = True
            Case Else   '����
                DPrice.Visible = False
        End Select
        If pPost = "New" Then DPrice.Text = ""

        '�}�o�z��
        Select Case FindFieldInf("DevReason")
            Case 0  '���
                DDevReason.BackColor = Color.LightGray
                DDevReason.ReadOnly = True
                DDevReason.Visible = True
            Case 1  '�ק�+�ˬd
                DDevReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevReasonRqd", "DDevReason", "���`�G�ݿ�J�}�o�z��")
                DDevReason.Visible = True
            Case 2  '�ק�
                DDevReason.BackColor = Color.Yellow
                DDevReason.Visible = True
            Case Else   '����
                DDevReason.Visible = False
        End Select
        If pPost = "New" Then DDevReason.Text = ""

        '�̲׼˫~��
        Select Case FindFieldInf("FinalSampleFile")
            Case 0  '���
                DFinalSampleFile.Visible = False
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DFinalSampleFileRqd", "DFinalSampleFile", "���`�G�ݿ�J�̲׼˫~��")
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DFinalSampleFile.Visible = True
                DFinalSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DFinalSampleFile.Visible = False
        End Select

        '���s/�~�`
        Select Case FindFieldInf("ManufType")
            Case 0  '���
                DManufType.BackColor = Color.LightGray
                DManufType.Visible = True
            Case 1  '�ק�+�ˬd
                DManufType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufTypeRqd", "DManufType", "���`�G�ݿ�J���s/�~�`")
                DManufType.Visible = True
            Case 2  '�ק�
                DManufType.BackColor = Color.Yellow
                DManufType.Visible = True
            Case Else   '����
                DManufType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufType", "ZZZZZZ")

        '�~�`��
        Select Case FindFieldInf("Suppiler")
            Case 0  '���
                DSuppiler.BackColor = Color.LightGray
                DSuppiler.Visible = True
            Case 1  '�ק�+�ˬd
                DSuppiler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSuppilerRqd", "DSuppiler", "���`�G�ݿ�J�~�`��")
                DSuppiler.Visible = True
            Case 2  '�ק�
                DSuppiler.BackColor = Color.Yellow
                DSuppiler.Visible = True
            Case Else   '����
                DSuppiler.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Suppiler", "ZZZZZZ")

        '�C��
        Select Case FindFieldInf("Color")
            Case 0  '���
                DColor.BackColor = Color.LightGray
                DColor.ReadOnly = True
                DColor.Visible = True
            Case 1  '�ק�+�ˬd
                DColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorRqd", "DColor", "���`�G�ݿ�J�C��")
                DColor.Visible = True
            Case 2  '�ק�
                DColor.BackColor = Color.Yellow
                DColor.Visible = True
            Case Else   '����
                DColor.Visible = False
        End Select
        If pPost = "New" Then DColor.Text = ""

        '�ƶq
        Select Case FindFieldInf("Qty")
            Case 0  '���
                DQty.BackColor = Color.LightGray
                DQty.ReadOnly = True
                DQty.Visible = True
            Case 1  '�ק�+�ˬd
                DQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQtyRqd", "DQty", "���`�G�ݿ�J�ƶq")
                DQty.Visible = True
            Case 2  '�ק�
                DQty.BackColor = Color.Yellow
                DQty.Visible = True
            Case Else   '����
                DQty.Visible = False
        End Select
        If pPost = "New" Then DQty.Text = ""

        '�鲣��
        Select Case FindFieldInf("Cap")
            Case 0  '���
                DCap.BackColor = Color.LightGray
                DCap.ReadOnly = True
                DCap.Visible = True
            Case 1  '�ק�+�ˬd
                DCap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCapRqd", "DCap", "���`�G�ݿ�J�ƶq")
                DCap.Visible = True
            Case 2  '�ק�
                DCap.BackColor = Color.Yellow
                DCap.Visible = True
            Case Else   '����
                DCap.Visible = False
        End Select
        If pPost = "New" Then DCap.Text = ""

        '�鲣��
        Select Case FindFieldInf("Schedule")
            Case 0  '���
                DSchedule.BackColor = Color.LightGray
                DSchedule.ReadOnly = True
                DSchedule.Visible = True
            Case 1  '�ק�+�ˬd
                DSchedule.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DScheduleRqd", "DSchedule", "���`�G�ݿ�J�ƶq")
                DSchedule.Visible = True
            Case 2  '�ק�
                DSchedule.BackColor = Color.Yellow
                DSchedule.Visible = True
            Case Else   '����
                DSchedule.Visible = False
        End Select
        If pPost = "New" Then DSchedule.Text = ""

        '�z��
        Select Case FindFieldInf("Schedule")
            Case 0  '���
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '�ק�+�ˬd
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "���`�G�ݿ�J�z��")
                DFReason.Visible = True
            Case 2  '�ק�
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '����
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""

        '���׼˫~
        Select Case FindFieldInf("AllowSample")
            Case 0  '���
                DAllowSample.BackColor = Color.LightGray
                DAllowSample.Visible = True
            Case 1  '�ק�+�ˬd
                DAllowSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAllowSampleRqd", "DAllowSample", "���`�G�ݿ�J���׼˫~")
                DAllowSample.Visible = True
            Case 2  '�ק�
                DAllowSample.BackColor = Color.Yellow
                DAllowSample.Visible = True
            Case Else   '����
                DAllowSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AllowSample", "ZZZZZZ")

        '�w�w������
        Select Case FindFieldInf("BFinalDate")
            Case 0  '���
                DBFinalDate.BackColor = Color.LightGray
                DBFinalDate.ReadOnly = True
                DBFinalDate.Visible = True
                BBFinalDate.Visible = False
            Case 1  '�ק�+�ˬd
                DBFinalDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBFinalDateRqd", "DBFinalDate", "���`�G�ݿ�J�w�w������")
                DBFinalDate.Visible = True
                BBFinalDate.Visible = True
            Case 2  '�ק�
                DBFinalDate.BackColor = Color.Yellow
                DBFinalDate.Visible = True
                BBFinalDate.Visible = True
            Case Else   '����
                DBFinalDate.Visible = False
                BBFinalDate.Visible = False
        End Select
        If pPost = "New" Then DBFinalDate.Text = ""

        'Code
        Select Case FindFieldInf("Code")
            Case 0  '���
                DCode.BackColor = Color.LightGray
                DCode.ReadOnly = True
                DCode.Visible = True
            Case 1  '�ק�+�ˬd
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "���`�G�ݿ�JCode")
                DCode.Visible = True
            Case 2  '�ק�
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '����
                DCode.Visible = False
        End Select
        If pPost = "New" Then DCode.Text = ""

        '�^��W��
        Select Case FindFieldInf("EnglishName")
            Case 0  '���
                DEnglishName.BackColor = Color.LightGray
                DEnglishName.ReadOnly = True
                DEnglishName.Visible = True
            Case 1  '�ק�+�ˬd
                DEnglishName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEnglishNameRqd", "DEnglishName", "���`�G�ݿ�J�^��W��")
                DEnglishName.Visible = True
            Case 2  '�ק�
                DEnglishName.BackColor = Color.Yellow
                DEnglishName.Visible = True
            Case Else   '����
                DEnglishName.Visible = False
        End Select
        If pPost = "New" Then DEnglishName.Text = ""

        'LOSS
        Select Case FindFieldInf("LOSS")
            Case 0  '���
                DLOSS.BackColor = Color.LightGray
                DLOSS.ReadOnly = True
                DLOSS.Visible = True
            Case 1  '�ק�+�ˬd
                DLOSS.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLOSSRqd", "DLOSS", "���`�G�ݿ�JLOSS")
                DLOSS.Visible = True
            Case 2  '�ק�
                DLOSS.BackColor = Color.Yellow
                DLOSS.Visible = True
            Case Else   '����
                DLOSS.Visible = False
        End Select
        If pPost = "New" Then DLOSS.Text = ""

        '�~��̿��
        Select Case FindFieldInf("QCReqFile")
            Case 0  '���
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "���`�G�ݿ�J�~��̿��")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCReqFile.Visible = False
        End Select

        '�f�|�o�k
        Select Case FindFieldInf("QCCheck1")
            Case 0  '���
                DQCCheck1.BackColor = Color.LightGray
                DQCCheck1.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck1Rqd", "DQCCheck1", "���`�G�ݿ�J�f�|�o�k")
                DQCCheck1.Visible = True
            Case 2  '�ק�
                DQCCheck1.BackColor = Color.Yellow
                DQCCheck1.Visible = True
            Case Else   '����
                DQCCheck1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck1", "ZZZZZZ")

        '�P�ʩ��
        Select Case FindFieldInf("QCCheck2")
            Case 0  '���
                DQCCheck2.BackColor = Color.LightGray
                DQCCheck2.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck2Rqd", "DQCCheck2", "���`�G�ݿ�J�P�ʩ��")
                DQCCheck2.Visible = True
            Case 2  '�ק�
                DQCCheck2.BackColor = Color.Yellow
                DQCCheck2.Visible = True
            Case Else   '����
                DQCCheck2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck2", "ZZZZZZ")

        'LOCK�j��
        Select Case FindFieldInf("QCCheck3")
            Case 0  '���
                DQCCheck3.BackColor = Color.LightGray
                DQCCheck3.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck3Rqd", "DQCCheck3", "���`�G�ݿ�JLOCK�j��")
                DQCCheck3.Visible = True
            Case 2  '�ק�
                DQCCheck3.BackColor = Color.Yellow
                DQCCheck3.Visible = True
            Case Else   '����
                DQCCheck3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck3", "ZZZZZZ")

        '90�ױj��
        Select Case FindFieldInf("QCCheck4")
            Case 0  '���
                DQCCheck4.BackColor = Color.LightGray
                DQCCheck4.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck4Rqd", "DQCCheck4", "���`�G�ݿ�J90�ױj��")
                DQCCheck4.Visible = True
            Case 2  '�ק�
                DQCCheck4.BackColor = Color.Yellow
                DQCCheck4.Visible = True
            Case Else   '����
                DQCCheck4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck4", "ZZZZZZ")

        '��O
        Select Case FindFieldInf("QCCheck5")
            Case 0  '���
                DQCCheck5.BackColor = Color.LightGray
                DQCCheck5.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck5Rqd", "DQCCheck5", "���`�G�ݿ�J��O")
                DQCCheck5.Visible = True
            Case 2  '�ק�
                DQCCheck5.BackColor = Color.Yellow
                DQCCheck5.Visible = True
            Case Else   '����
                DQCCheck5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck5", "ZZZZZZ")


        'N-ANTI
        Select Case FindFieldInf("QCCheck15")
            Case 0  '���
                DQCCHECK15.BackColor = Color.LightGray
                DQCCHECK15.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCHECK15.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("QCCheck15Rqd", "DQCCheck15", "���`�G�ݿ�J��O")
                DQCCHECK15.Visible = True
            Case 2  '�ק�
                DQCCHECK15.BackColor = Color.Yellow
                DQCCHECK15.Visible = True
            Case Else   '����
                DQCCHECK15.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck15", "ZZZZZZ")


        'PFAS
        Select Case FindFieldInf("QCCheck15")
            Case 0  '���
                DQCCHECK16.BackColor = Color.LightGray
                DQCCHECK16.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCHECK16.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("QCCheck16Rqd", "DQCCHECK16", "���`�G�ݿ�JPFAS")
                DQCCHECK16.Visible = True
            Case 2  '�ק�
                DQCCHECK16.BackColor = Color.Yellow
                DQCCHECK16.Visible = True
            Case Else   '����
                DQCCHECK16.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck16", "ZZZZZZ")

        '�q�ὤ�p
        ' Select Case FindFieldInf("QCCheck6File")
        '     Case 0  '���
        ' DQCCheck6File.Visible = False
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "LightGrey")
        '     Case 1  '�ק�+�ˬd
        ' ShowRequiredFieldValidator("DQCCheck6FileRqd", "DQCCheck6File", "���`�G�ݿ�J�q�ὤ�p����")
        ' DQCCheck6File.Visible = True
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "GreenYellow")
        '     Case 2  '�ק�
        ' DQCCheck6File.Visible = True
        ' DQCCheck6File.Style.Add("BACKGROUND-COLOR", "Yellow")
        '     Case Else   '����
        ' DQCCheck6File.Visible = False
        ' End Select

        '�˰w
        Select Case FindFieldInf("QCCheck7")
            Case 0  '���
                DQCCheck7.BackColor = Color.LightGray
                DQCCheck7.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck7Rqd", "DQCCheck7", "���`�G�ݿ�J�˰w")
                DQCCheck7.Visible = True
            Case 2  '�ק�
                DQCCheck7.BackColor = Color.Yellow
                DQCCheck7.Visible = True
            Case Else   '����
                DQCCheck7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck7", "ZZZZZZ")

        'AATCC
        Select Case FindFieldInf("QCCheck8")
            Case 0  '���
                DQCCheck8.BackColor = Color.LightGray
                DQCCheck8.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck8.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck8Rqd", "DQCCheck8", "���`�G�ݿ�JAATCC")
                DQCCheck8.Visible = True
            Case 2  '�ק�
                DQCCheck8.BackColor = Color.Yellow
                DQCCheck8.Visible = True
            Case Else   '����
                DQCCheck8.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck8", "ZZZZZZ")

        '���~
        Select Case FindFieldInf("QCCheck9")
            Case 0  '���
                DQCCheck9.BackColor = Color.LightGray
                DQCCheck9.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck9.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck9Rqd", "DQCCheck9", "���`�G�ݿ�J���~")
                DQCCheck9.Visible = True
            Case 2  '�ק�
                DQCCheck9.BackColor = Color.Yellow
                DQCCheck9.Visible = True
            Case Else   '����
                DQCCheck9.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck9", "ZZZZZZ")

        '�Q���Q��
        Select Case FindFieldInf("QCCheck10")
            Case 0  '���
                DQCCheck10.BackColor = Color.LightGray
                DQCCheck10.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck10.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck10Rqd", "DQCCheck10", "���`�G�ݿ�J�Q���Q��")
                DQCCheck10.Visible = True
            Case 2  '�ק�
                DQCCheck10.BackColor = Color.Yellow
                DQCCheck10.Visible = True
            Case Else   '����
                DQCCheck10.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck10", "ZZZZZZ")

        '�@���K��
        Select Case FindFieldInf("QCCheck11")
            Case 0  '���
                DQCCheck11.BackColor = Color.LightGray
                DQCCheck11.Visible = True
            Case 1  '�ק�+�ˬd
                DQCCheck11.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck11Rqd", "DQCCheck11", "���`�G�ݿ�J�@���K��")
                DQCCheck11.Visible = True
            Case 2  '�ק�
                DQCCheck11.BackColor = Color.Yellow
                DQCCheck11.Visible = True
            Case Else   '����
                DQCCheck11.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCCheck11", "ZZZZZZ")

        '�G���K��,Oeko-tex,A01
        Select Case FindFieldInf("QCCheck12")
            Case 0  '���
                '�G���K��
                DQCCheck12.BackColor = Color.LightGray
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.LightGray
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.LightGray
                DQCCheck14.Visible = True
            Case 1  '�ק�+�ˬd
                '�G���K��
                DQCCheck12.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck12Rqd", "DQCCheck12", "���`�G�ݿ�J�G���K��")
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck13Rqd", "DQCCheck13", "���`�G�ݿ�JOeko-Tex")
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCCheck14Rqd", "DQCCheck14", "���`�G�ݿ�JA01")
                DQCCheck14.Visible = True
            Case 2  '�ק�
                '�G���K��
                DQCCheck12.BackColor = Color.Yellow
                DQCCheck12.Visible = True
                'Oeko-tex
                DQCCheck13.BackColor = Color.Yellow
                DQCCheck13.Visible = True
                'A01
                DQCCheck14.BackColor = Color.Yellow
                DQCCheck14.Visible = True
            Case Else   '����
                '�G���K��
                DQCCheck12.Visible = False
                'Oeko-tex
                DQCCheck13.Visible = False
                'A01
                DQCCheck14.Visible = False
        End Select
        If pPost = "New" Then
            SetFieldData("QCCheck12", "ZZZZZZ")
            SetFieldData("QCCheck13", "ZZZZZZ")
            SetFieldData("QCCheck14", "ZZZZZZ")
        End If

        '���`����
        Select Case FindFieldInf("EACheck1")
            Case 0  '���
                DEACheck1.BackColor = Color.LightGray
                DEACheck1.Visible = True
            Case 1  '�ק�+�ˬd
                DEACheck1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEACheck1Rqd", "DEACheck1", "���`�G�ݿ�J���`����")
                DEACheck1.Visible = True
            Case 2  '�ק�
                DEACheck1.BackColor = Color.Yellow
                DEACheck1.Visible = True
            Case Else   '����
                DEACheck1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("EACheck1", "ZZZZZZ")

        '���`����Ƶ�
        Select Case FindFieldInf("EADesc1")
            Case 0  '���
                DEADesc1.BackColor = Color.LightGray
                DEADesc1.ReadOnly = True
                DEADesc1.Visible = True
            Case 1  '�ק�+�ˬd
                DEADesc1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEADesc1Rqd", "DEADesc1", "���`�G�ݿ�J���`����Ƶ�")
                DEADesc1.Visible = True
            Case 2  '�ק�
                DEADesc1.BackColor = Color.Yellow
                DEADesc1.Visible = True
            Case Else   '����
                DEADesc1.Visible = False
        End Select
        If pPost = "New" Then DEADesc1.Text = ""

        'Oeko-tex,A01���`������i
        Select Case FindFieldInf("EACheckFile")
            Case 0  '���
                'Oeko-tex���`������i
                DEACheckFile.Visible = False
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                'A01���`������i
                DEACheckFile1.Visible = False
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                'Oeko-tex���`������i
                ShowRequiredFieldValidator("DEACheckFileRqd", "DEACheckFile", "���`�G�ݿ�JOeko-tex���`������i")
                DEACheckFile.Visible = True
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                'A01���`������i
                ShowRequiredFieldValidator("DEACheckFile1Rqd", "DEACheckFile1", "���`�G�ݿ�JA01���`������i")
                DEACheckFile1.Visible = True
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                'Oeko-tex���`������i
                DEACheckFile.Visible = True
                DEACheckFile.Style.Add("BACKGROUND-COLOR", "Yellow")
                'A01���`������i
                DEACheckFile1.Visible = True
                DEACheckFile1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                'Oeko-tex���`������i
                DEACheckFile.Visible = False
                'A01���`������i
                DEACheckFile1.Visible = False
        End Select

        '���ճ��i��
        Select Case FindFieldInf("QCFinalFile")
            Case 0  '���
                DQCFinalFile.Visible = False
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFinalFileRqd", "DQCFinalFile", "���`�G�ݿ�J���ճ��i��")
                DQCFinalFile.Visible = True
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFinalFile.Visible = True
                DQCFinalFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFinalFile.Visible = False
        End Select

        'FAFS
        Select Case FindFieldInf("QCFinalFile")
            Case 0  '���
                DPFASFile.Visible = False
                DPFASFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DPFASFileRqd", "DPFASFile", "���`�G�ݿ�JPFAS���i")
                DPFASFile.Visible = True
                DPFASFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DPFASFile.Visible = True
                DPFASFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DPFASFile.Visible = False
        End Select



        '���-1
        Select Case FindFieldInf("QCDate1")
            Case 0  '���
                DQCDate1.BackColor = Color.LightGray
                DQCDate1.ReadOnly = True
                DQCDate1.Visible = True
                BQCDate1.Visible = False
            Case 1  '�ק�+�ˬd
                DQCDate1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate1Rqd", "DQCDate1", "���`�G�ݿ�J���-1")
                DQCDate1.Visible = True
                BQCDate1.Visible = True
            Case 2  '�ק�
                DQCDate1.BackColor = Color.Yellow
                DQCDate1.Visible = True
                BQCDate1.Visible = True
            Case Else   '����
                DQCDate1.Visible = False
                BQCDate1.Visible = False
        End Select
        If pPost = "New" Then DQCDate1.Text = ""

        '�˴����G-1
        Select Case FindFieldInf("QCResult1")
            Case 0  '���
                DQCResult1.BackColor = Color.LightGray
                DQCResult1.Visible = True
            Case 1  '�ק�+�ˬd
                DQCResult1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult1Rqd", "DQCResult1", "���`�G�ݿ�J�˴����G-1")
                DQCResult1.Visible = True
            Case 2  '�ק�
                DQCResult1.BackColor = Color.Yellow
                DQCResult1.Visible = True
            Case Else   '����
                DQCResult1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult1", "ZZZZZZ")

        '�Ƶ�-1
        Select Case FindFieldInf("QCRemark1")
            Case 0  '���
                DQCRemark1.BackColor = Color.LightGray
                DQCRemark1.ReadOnly = True
                DQCRemark1.Visible = True
            Case 1  '�ק�+�ˬd
                DQCRemark1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark1Rqd", "DQCRemark1", "���`�G�ݿ�J�Ƶ�-1")
                DQCRemark1.Visible = True
            Case 2  '�ק�
                DQCRemark1.BackColor = Color.Yellow
                DQCRemark1.Visible = True
            Case Else   '����
                DQCRemark1.Visible = False
        End Select
        If pPost = "New" Then DQCRemark1.Text = ""

        '���-2
        Select Case FindFieldInf("QCDate2")
            Case 0  '���
                DQCDate2.BackColor = Color.LightGray
                DQCDate2.ReadOnly = True
                DQCDate2.Visible = True
                BQCDate2.Visible = False

                DQCDate3.BackColor = Color.LightGray
                DQCDate3.ReadOnly = True
                DQCDate3.Visible = True
                BQCDate3.Visible = False
            Case 1  '�ק�+�ˬd
                DQCDate2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate2Rqd", "DQCDate2", "���`�G�ݿ�J���-2")
                DQCDate2.Visible = True
                BQCDate2.Visible = True

                DQCDate3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCDate3Rqd", "DQCDate3", "���`�G�ݿ�J���-3")
                DQCDate3.Visible = True
                BQCDate3.Visible = True
            Case 2  '�ק�
                DQCDate2.BackColor = Color.Yellow
                DQCDate2.Visible = True
                BQCDate2.Visible = True

                DQCDate3.BackColor = Color.Yellow
                DQCDate3.Visible = True
                BQCDate1.Visible = True
            Case Else   '����
                DQCDate2.Visible = False
                BQCDate2.Visible = False
                DQCDate3.Visible = False
                BQCDate3.Visible = False

        End Select
        If pPost = "New" Then DQCDate2.Text = ""

        '�˴����G-2
        Select Case FindFieldInf("QCResult2")
            Case 0  '���
                DQCResult2.BackColor = Color.LightGray
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.LightGray
                DQCResult3.Visible = True
            Case 1  '�ק�+�ˬd
                DQCResult2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult2Rqd", "DQCResult2", "���`�G�ݿ�J�˴����G-2")
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResult3Rqd", "DQCResult3", "���`�G�ݿ�J�˴����G-3")
                DQCResult3.Visible = True
            Case 2  '�ק�
                DQCResult2.BackColor = Color.Yellow
                DQCResult2.Visible = True

                DQCResult3.BackColor = Color.Yellow
                DQCResult3.Visible = True
            Case Else   '����
                DQCResult2.Visible = False
                DQCResult3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult2", "ZZZZZZ")

        '�Ƶ�-2 �Ƶ�-3 �g�b�@�_ 20102/4 jessica �ק�
        Select Case FindFieldInf("QCRemark2")
            Case 0  '���
                DQCRemark2.BackColor = Color.LightGray
                DQCRemark2.ReadOnly = True
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.LightGray
                DQCRemark3.ReadOnly = True
                DQCRemark3.Visible = True

            Case 1  '�ק�+�ˬd
                DQCRemark2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark2Rqd", "DQCRemark2", "���`�G�ݿ�J�Ƶ�-2")
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemark3Rqd", "DQCRemark3", "���`�G�ݿ�J�Ƶ�-3")
                DQCRemark3.Visible = True

            Case 2  '�ק�
                DQCRemark2.BackColor = Color.Yellow
                DQCRemark2.Visible = True

                DQCRemark3.BackColor = Color.Yellow
                DQCRemark3.Visible = True

            Case Else   '����
                DQCRemark2.Visible = False
                DQCRemark3.Visible = False

        End Select
        If pPost = "New" Then DQCRemark2.Text = ""
        If pPost = "New" Then DQCDate3.Text = ""
        If pPost = "New" Then SetFieldData("QCResult3", "ZZZZZZ")
        If pPost = "New" Then DQCRemark3.Text = ""

        '�s�y�y�{��
        Select Case FindFieldInf("ManufFlowFile")
            Case 0  '���
                DManufFlowFile.Visible = False
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DManufFlowFileRqd", "DManufFlowFile", "���`�G�ݿ�J�s�y�y�{��")
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DManufFlowFile.Visible = True
                DManufFlowFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DManufFlowFile.Visible = False
        End Select

        '������
        Select Case FindFieldInf("ForcastFile")
            Case 0  '���
                DForcastFile.Visible = False
                DForcastFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DForcastFileRqd", "DForcastFile", "���`�G�ݿ�J������")
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DForcastFile.Visible = True
                DForcastFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DForcastFile.Visible = False
        End Select

        '�@�~�зǮ�
        Select Case FindFieldInf("OPManualFile")
            Case 0  '���
                DOPManualFile.Visible = False
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DOPManualFileRqd", "DOPManualFile", "���`�G�ݿ�J�@�~�зǮ�")
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DOPManualFile.Visible = True
                DOPManualFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DOPManualFile.Visible = False
        End Select

        '������
        Select Case FindFieldInf("ContactFile")
            Case 0  '���
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DContactFileRqd", "DContactFile", "���`�G�ݿ�J������")
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DContactFile.Visible = True
                DContactFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DContactFile.Visible = False
        End Select

        'QC-L/T
        Select Case FindFieldInf("QCLT")
            Case 0  '���
                DQCLT.BackColor = Color.LightGray
                DQCLT.ReadOnly = True
                DQCLT.Visible = True
            Case 1  '�ק�+�ˬd
                DQCLT.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCLTRqd", "DQCLT", "���`�G�ݿ�JQC L/T")
                DQCLT.Visible = True
            Case 2  '�ק�
                DQCLT.BackColor = Color.Yellow
                DQCLT.Visible = True
            Case Else   '����
                DQCLT.Visible = False
        End Select
        If pPost = "New" Then DQCLT.Text = ""

        ''**********************************************************************************************
        '������
        'Select Case FindFieldInf("Level")
        '    Case 0  '���
        'DLevel.BackColor = Color.LightGray
        'DLevel.Visible = True
        '    Case 1  '�ק�+�ˬd
        'DLevel.BackColor = Color.GreenYellow
        'ShowRequiredFieldValidator("DLevelRqd", "DLevel", "���`�G�ݿ�J������")
        'DLevel.Visible = True
        '    Case 2  '�ק�
        'DLevel.BackColor = Color.Yellow
        'DLevel.Visible = True
        '    Case Else   '����
        'DLevel.Visible = False
        'End Select
        'If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")
        ''**********************************************************************************************

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

        'Buyer
        If pFieldName = "Buyer" Then
            DBuyer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBuyer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='BUYER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBuyer.Items.Add(ListItem1)
                Next
            End If
        End If

        '����
        If pFieldName = "AttachSample" Then
            DAttachSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAttachSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='ATTACHSAMPLE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAttachSample.Items.Add(ListItem1)
                Next
            End If
        End If


        '�~�u
        If pFieldName = "YearSeason" Then
            DYearSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYearSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='YearSeason' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                DYearSeason.Items.Add("")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYearSeason.Items.Add(ListItem1)
                Next
            End If
        End If

        '���s/�~�`
        If pFieldName = "ManufType" Then
            DManufType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='MANUFTYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DManufType.Items.Add(ListItem1)
                Next
            End If
        End If

        'Suppiler
        If pFieldName = "Suppiler" Then
            DSuppiler.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSuppiler.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='SUPPLIER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSuppiler.Items.Add(ListItem1)
                Next
            End If
        End If

        '���׼˫~
        If pFieldName = "AllowSample" Then
            DAllowSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAllowSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='ALLOWSAMPLE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAllowSample.Items.Add(ListItem1)
                Next
            End If
        End If

        '�f�|�o�k
        If pFieldName = "QCCheck1" Then
            DQCCheck1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck1.Items.Add(ListItem1)
                Next
            End If
        End If

        '�P�ʩ��
        If pFieldName = "QCCheck2" Then
            DQCCheck2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck2.Items.Add(ListItem1)
                Next
            End If
        End If

        'LOCK�j��
        If pFieldName = "QCCheck3" Then
            DQCCheck3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK3' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck3.Items.Add(ListItem1)
                Next
            End If
        End If

        '90�ױj��
        If pFieldName = "QCCheck4" Then
            DQCCheck4.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck4.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK4' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck4.Items.Add(ListItem1)
                Next
            End If
        End If

        '��O
        If pFieldName = "QCCheck5" Then
            DQCCheck5.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck5.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK5' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck5.Items.Add(ListItem1)
                Next
            End If
        End If


        'N-ANTI
        If pFieldName = "QCCheck15" Then
            DQCCHECK15.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCHECK15.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK15' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCHECK15.Items.Add(ListItem1)
                Next
            End If
        End If

        'PFAS
        If pFieldName = "QCCheck16" Then
            DQCCHECK16.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCHECK16.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK16' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCHECK16.Items.Add(ListItem1)
                Next
            End If
        End If




        '�˰w
        If pFieldName = "QCCheck7" Then
            DQCCheck7.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck7.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK7' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck7.Items.Add(ListItem1)
                Next
            End If
        End If

        'AATCC
        If pFieldName = "QCCheck8" Then
            DQCCheck8.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck8.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK8' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck8.Items.Add(ListItem1)
                Next
            End If
        End If

        '���~
        If pFieldName = "QCCheck9" Then
            DQCCheck9.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck9.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK9' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck9.Items.Add(ListItem1)
                Next
            End If
        End If

        '�Q���Q��
        If pFieldName = "QCCheck10" Then
            DQCCheck10.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck10.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK10' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck10.Items.Add(ListItem1)
                Next
            End If
        End If

        '�@���K��
        If pFieldName = "QCCheck11" Then
            DQCCheck11.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck11.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK11' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck11.Items.Add(ListItem1)
                Next
            End If
        End If

        '�G���K��,Oeko-tex,A01
        If pFieldName = "QCCheck12" Then
            '�G���K��
            DQCCheck12.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck12.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK12' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck12.Items.Add(ListItem1)
                Next
            End If
        End If

        'Oeko-tex
        If pFieldName = "QCCheck13" Then
            DQCCheck13.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck13.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK13' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck13.Items.Add(ListItem1)
                Next
            End If
        End If

        'A01
        If pFieldName = "QCCheck14" Then
            DQCCheck14.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCCheck14.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCCHECK14' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCCheck14.Items.Add(ListItem1)
                Next
            End If
        End If

        '���`����
        If pFieldName = "EACheck1" Then
            DEACheck1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEACheck1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='EACHECK1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEACheck1.Items.Add(ListItem1)
                Next
            End If
        End If

        '�˴����G-1
        If pFieldName = "QCResult1" Then
            DQCResult1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT1' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult1.Items.Add(ListItem1)
                Next
            End If
        End If

        '�˴����G-2
        If pFieldName = "QCResult2" Then
            DQCResult2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult2.Items.Add(ListItem1)
                Next
            End If
        End If

        '�˴����G-3
        If pFieldName = "QCResult3" Then
            DQCResult3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='320' and DKey='QCRESULT3' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult3.Items.Add(ListItem1)
                Next
            End If
        End If

        '������
        'If pFieldName = "Level" Then
        'DLevel.Items.Clear()
        'If idx = 0 Then
        'If pName <> "ZZZZZZ" Then
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = pName
        'ListItem1.Value = pName
        'DLevel.Items.Add(ListItem1)
        'End If
        'Else
        'SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey = '') Order by DKey, Data "
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "M_Referp")
        'DBTable1 = DBDataSet1.Tables("M_Referp")
        'For i = 0 To DBTable1.Rows.Count - 1
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = DBTable1.Rows(i).Item("Data")
        'ListItem1.Value = DBTable1.Rows(i).Item("Data")
        'If ListItem1.Value = pName Then ListItem1.Selected = True
        'DLevel.Items.Add(ListItem1)
        'Next
        'End If
        'End If

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

                'Check�Ȥ�˫~��-Size�ή榡
                If ErrCode = 0 Then
                    If DCustSampleFile.Visible Then
                        If DCustSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DCustSampleFile)
                        End If
                    End If
                End If

                'Check�̲׼˫~��-Size�ή榡
                If ErrCode = 0 Then
                    If DFinalSampleFile.Visible Then
                        If DFinalSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DFinalSampleFile)
                        End If
                    End If
                End If

                'Check�~��̿��-Size�ή榡
                If ErrCode = 0 Then
                    If DQCReqFile.Visible Then
                        If DQCReqFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DQCReqFile)
                        End If
                    End If
                End If

                'Check�q�ὤ�p-Size�ή榡
                ' If ErrCode = 0 Then
                ' If DQCCheck6File.Visible Then
                ' If DQCCheck6File.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                ' ErrCode = UPFileIsNormal(DQCCheck6File)
                'End If
                'End If
                ' End If

                'Check Oeko-tex���`������i-Size�ή榡
                If ErrCode = 0 Then
                    If DEACheckFile.Visible Then
                        If DEACheckFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DEACheckFile)
                        End If
                    End If
                End If

                'Check A01���`������i-Size�ή榡
                If ErrCode = 0 Then
                    If DEACheckFile1.Visible Then
                        If DEACheckFile1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DEACheckFile1)
                        End If
                    End If
                End If

                'Check���ճ��i��-Size�ή榡
                If ErrCode = 0 Then
                    If DQCFinalFile.Visible Then
                        If DQCFinalFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DQCFinalFile)
                        End If
                    End If
                End If



                'Check FASFS-Size�ή榡
                If ErrCode = 0 Then
                    If DPFASFile.Visible Then
                        If DPFASFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DPFASFile)
                        End If
                    End If
                End If


                'Check�s�y�y�{��-Size�ή榡
                If ErrCode = 0 Then
                    If DManufFlowFile.Visible Then
                        If DManufFlowFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DManufFlowFile)
                        End If
                    End If
                End If

                'Check������-Size�ή榡
                If ErrCode = 0 Then
                    If DForcastFile.Visible Then
                        If DForcastFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DForcastFile)
                        End If
                    End If
                End If

                'Check�@�~�зǮ�-Size�ή榡
                If ErrCode = 0 Then
                    If DOPManualFile.Visible Then
                        If DOPManualFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DOPManualFile)
                        End If
                    End If
                End If

                'Check������-Size�ή榡
                If ErrCode = 0 Then
                    If DContactFile.Visible Then
                        If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                            ErrCode = UPFileIsNormal(DContactFile)
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
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '���o�����d�U���s��Data Status

        'Check�Ȥ�˫~��-Size�ή榡
        If ErrCode = 0 Then
            If DCustSampleFile.Visible Then
                If DCustSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DCustSampleFile)
                End If
            End If
        End If

        'Check�̲׼˫~��-Size�ή榡
        If ErrCode = 0 Then
            If DFinalSampleFile.Visible Then
                If DFinalSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DFinalSampleFile)
                End If
            End If
        End If

        'Check�~��̿��-Size�ή榡
        If ErrCode = 0 Then
            If DQCReqFile.Visible Then
                If DQCReqFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DQCReqFile)
                End If
            End If
        End If

        'Check�q�ὤ�p-Size�ή榡
        '  If ErrCode = 0 Then
        ' If DQCCheck6File.Visible Then
        ' If DQCCheck6File.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
        ' ErrCode = UPFileIsNormal(DQCCheck6File)
        'End If
        'End If
        '    End If

        'Check Oeko-tex���`������i-Size�ή榡
        If ErrCode = 0 Then
            If DEACheckFile.Visible Then
                If DEACheckFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DEACheckFile)
                End If
            End If
        End If

        'Check A01���`������i-Size�ή榡
        If ErrCode = 0 Then
            If DEACheckFile1.Visible Then
                If DEACheckFile1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DEACheckFile1)
                End If
            End If
        End If

        'Check���ճ��i��-Size�ή榡
        If ErrCode = 0 Then
            If DQCFinalFile.Visible Then
                If DQCFinalFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DQCFinalFile)
                End If
            End If
        End If


        'Check FAFS Size�ή榡
        If ErrCode = 0 Then
            If DPFASFile.Visible Then
                If DPFASFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DPFASFile)
                End If
            End If
        End If


        'Check�s�y�y�{��-Size�ή榡
        If ErrCode = 0 Then
            If DManufFlowFile.Visible Then
                If DManufFlowFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DManufFlowFile)
                End If
            End If
        End If

        'Check������-Size�ή榡
        If ErrCode = 0 Then
            If DForcastFile.Visible Then
                If DForcastFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DForcastFile)
                End If
            End If
        End If

        'Check�@�~�зǮ�-Size�ή榡
        If ErrCode = 0 Then
            If DOPManualFile.Visible Then
                If DOPManualFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DOPManualFile)
                End If
            End If
        End If

        'Check������-Size�ή榡
        If ErrCode = 0 Then
            If DContactFile.Visible Then
                If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                    ErrCode = UPFileIsNormal(DContactFile)
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

        'Check-QC-L/T
        If ErrCode = 0 Then
            If DQCLT.Text <> "" Then
                If Not YKK.IntegerData(DQCLT.Text) Then
                    ErrCode = 9050
                Else
                    wQCLT = CInt(DQCLT.Text)
                End If
            End If
        End If

        '--�ˬd�e�U��No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("000014", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
            If DAttachSample.SelectedValue = "YES" Then wLevel = "Z"


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
                    'Step=60 Action�վ�, ���s=���ܧ�(0), �~�`=�ܧ�(1)
                    If wStep = 60 Then
                        If DManufType.SelectedValue = "�~�`" Then pAction = 1
                    End If

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
                    If wFormSno = 0 And wStep < 3 Then '�P�_�O�_�_��
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                            ModifyManufData(pFun, 1, 1)   '��s���檬�A
                        End If
                    Else
                        If pNextStep = 999 Then     '�u�{������?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '��s�����
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            ModifyManufData(pFun, 0, 1)   '��s���檬�A
                        Else
                            ModifyData(pFun, "0")         '��s����� Sts=0(����)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    If wStep = 20 Then      '�P�_�O�_���u�{20
                        ModifyManufData(pFun, 1, 1)   '��s���檬�A
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
            If ErrCode = 9050 Then Message = "QC L/T�ݬ����ļƦr,�нT�{!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SufaceFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_SufaceSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "                    '1~5
        SQl = SQl + "Date, Division, Person, CustSampleFile, Spec, "               '6~10
        SQl = SQl + "Buyer, SellVendor, ReqDelDate, ReqQty, SliderSample, "        '11~15
        SQl = SQl + "AttachSample, ORNO, OrderTime, DevReason, FinalSampleFile, "  '16~20
        SQl = SQl + "ManufType, Suppiler, Color, Qty, AllowSample, "               '21~25
        SQl = SQl + "BFinalDate, Code, EnglishName, LOSS, QCReqFile, QCCheck1, "         '26~30
        SQl = SQl + "QCCheck2, QCCheck3, QCCheck4, QCCheck5, QCCheck15, QCCheck16,"       '31~35
        SQl = SQl + "QCCheck7, QCCheck8, QCCheck9, QCCheck10, QCCheck11,  "        '36~40
        SQl = SQl + "QCCheck12, QCCheck13, QCCheck14, EACheck1, EADesc1, EACheckFile, EACheckFile1, QCFinalFile, PFASFile,"     '41~47
        SQl = SQl + "QCDate1, QCResult1, QCRemark1, QCDate2, QCResult2, "          '48~52
        SQl = SQl + "QCRemark2, QCDate3, QCResult3, QCRemark3, ManufFlowFile, "    '53~57
        SQl = SQl + "ForcastFile, OPManualFile, ContactFile, Level, QCLT, "        '58~62
        SQl = SQl + "UpdSts, Suface, OFormNo, OFormSno, Price, YearSeason, "                   '63~67
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "              '68~71
        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "              '���פ�
        SQl = SQl + " '000014', "                           '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '���y����
        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        '6~10
        SQl = SQl + " N'" + DDate.Text + "', "                    '���
        SQl = SQl + " N'" + DDivision.SelectedValue + "', "       '����
        SQl = SQl + " N'" + DPerson.SelectedValue + "', "         '���
        FileName = ""
        If DCustSampleFile.Visible Then                          '�Ȥ�˫~��
            If DCustSampleFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCustSampleFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " N'" + YKK.ReplaceString(DSpec.Text) + "', "          '�W��
        '11~15
        SQl = SQl + " N'" + DBuyer.SelectedValue + "', "                  'Buyer
        SQl = SQl + " N'" + YKK.ReplaceString(DSellVendor.Text) + "', "   '�e�U�t��
        SQl = SQl + " N'" + DReqDelDate.Text + "', "                       '�Ʊ���
        SQl = SQl + " N'" + DReqQty.Text + "', "                           '�w���q
        SQl = SQl + " N'" + DSliderSample.Text + "', "                    '�˫~���Y
        '16~20
        SQl = SQl + " N'" + DAttachSample.SelectedValue + "', "           '����
        SQl = SQl + " N'" + DORNO.Text + "', "                            'OR-NO
        SQl = SQl + " N'" + DOrderTime.Text + "', "                       '�U��ɶ�
        SQl = SQl + " N'" + DDevReason.Text + "', "                       '�}�o�z��
        FileName = ""
        If DFinalSampleFile.Visible Then                                  '�̲׼˫~��
            If DFinalSampleFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFinalSampleFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        '21~25
        SQl = SQl + " N'" + DManufType.SelectedValue + "', "              '���s/�~�`
        SQl = SQl + " N'" + DSuppiler.SelectedValue + "', "               '�~�`��
        SQl = SQl + " N'" + DColor.Text + "', "                           'Color
        SQl = SQl + " N'" + DQty.Text + "', "                             '�ƶq
        SQl = SQl + " N'" + DAllowSample.SelectedValue + "', "            '���׼˫~
        '26~30
        SQl = SQl + " N'" + DBFinalDate.Text + "', "                      '�w�w������
        SQl = SQl + " N'" + DCode.Text + "', "                            'Code
        SQl = SQl + " N'" + DEnglishName.Text + "', "                     '�^��W��
        SQl = SQl + " N'" + DLOSS.Text + "', "                            'LOSS
        FileName = ""
        If DQCReqFile.Visible Then                                        '�~��̿��
            If DQCReqFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCReqFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCReqFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " N'" + DQCCheck1.SelectedValue + "', "               '�f�|�o�k
        '31~35
        SQl = SQl + " N'" + DQCCheck2.SelectedValue + "', "               '�P�ʩ��
        SQl = SQl + " N'" + DQCCheck3.SelectedValue + "', "               'LOCK�j��
        SQl = SQl + " N'" + DQCCheck4.SelectedValue + "', "               '90�ױj��
        SQl = SQl + " N'" + DQCCheck5.SelectedValue + "', "               '��O
        SQl = SQl + " N'" + DQCCHECK15.SelectedValue + "', "              'N-ANTI
        SQl = SQl + " N'" + DQCCHECK16.SelectedValue + "', "              'PFAS

        'FileName = ""
        ' If DQCCheck6File.Visible Then                                     '�q�ὤ�p
        ' If DQCCheck6File.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
        ' '*** IE8����-Start 2011/1/4
        ' 'FileName = CStr(NewFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & Right(DQCCheck6File.PostedFile.FileName, InStr(StrReverse(DQCCheck6File.PostedFile.FileName), "\") - 1)
        ' FileName = CStr(NewFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCheck6File.PostedFile.FileName)
        ' '*** IE8����-End
        ' Try    '�W�ǹ���
        ' DQCCheck6File.PostedFile.SaveAs(Path + FileName)
        ' Catch ex As Exception
        ' End Try
        ' End If
        ' Else
        ' FileName = ""
        ' End If
        'SQl = SQl + " N'" + FileName + "', "
        '36~40
        SQl = SQl + " N'" + DQCCheck7.SelectedValue + "', "               '�˰w
        SQl = SQl + " N'" + DQCCheck8.SelectedValue + "', "               'AATCC
        SQl = SQl + " N'" + DQCCheck9.SelectedValue + "', "               '���~
        SQl = SQl + " N'" + DQCCheck10.SelectedValue + "', "              '�Q���Q��
        SQl = SQl + " N'" + DQCCheck11.SelectedValue + "', "              '�@���K��
        '41~47
        SQl = SQl + " N'" + DQCCheck12.SelectedValue + "', "              '�G���K��
        SQl = SQl + " N'" + DQCCheck13.SelectedValue + "', "              'Oeko-tex
        SQl = SQl + " N'" + DQCCheck14.SelectedValue + "', "              'A01
        SQl = SQl + " N'" + DEACheck1.SelectedValue + "', "               '���`����
        SQl = SQl + " N'" + DEADesc1.Text + "', "                         '���`����Ƶ�

        FileName = ""
        If DEACheckFile.Visible Then                                      'Oeko-tex���`������i
            If DEACheckFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & Right(DEACheckFile.PostedFile.FileName, InStr(StrReverse(DEACheckFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DEACheckFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DEACheckFile1.Visible Then                                      'A01���`������i
            If DEACheckFile1.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & Right(DEACheckFile1.PostedFile.FileName, InStr(StrReverse(DEACheckFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile1.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DEACheckFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        FileName = ""
        If DQCFinalFile.Visible Then                                      '���ճ��i��
            If DQCFinalFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DQCFinalFile.PostedFile.FileName, InStr(StrReverse(DQCFinalFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFinalFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFinalFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DPFASFile.Visible Then                                      '���ճ��i��
            If DPFASFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DPFASFile.PostedFile.FileName, InStr(StrReverse(DPFASFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "FAFS" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPFASFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DPFASFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        '48~52
        SQl = SQl + " N'" + DQCDate1.Text + "', "                         '���-1
        SQl = SQl + " N'" + DQCResult1.SelectedValue + "', "              '�˴����G-1
        SQl = SQl + " N'" + DQCRemark1.Text + "', "                       '�Ƶ�-1
        SQl = SQl + " N'" + DQCDate2.Text + "', "                         '���-2
        SQl = SQl + " N'" + DQCResult2.SelectedValue + "', "              '�˴����G-2
        '53~57
        SQl = SQl + " N'" + DQCRemark2.Text + "', "                       '�Ƶ�-2
        SQl = SQl + " N'" + DQCDate3.Text + "', "                         '���-3
        SQl = SQl + " N'" + DQCResult3.SelectedValue + "', "              '�˴����G-3
        SQl = SQl + " N'" + DQCRemark3.Text + "', "                       '�Ƶ�-3
        FileName = ""
        If DManufFlowFile.Visible Then                                      '�s�y�y�{��
            If DManufFlowFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DManufFlowFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        '58~62
        FileName = ""
        If DForcastFile.Visible Then                                      '������
            If DForcastFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DForcastFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DForcastFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DOPManualFile.Visible Then                                      '�@�~�зǮ�
            If DOPManualFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DOPManualFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DOPManualFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        FileName = ""
        If DContactFile.Visible Then                                      '������
            If DContactFile.PostedFile.FileName <> "" Then    '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "
        SQl = SQl + " '" + "" + "', "
        SQl = SQl + " '" + DQCLT.Text + "', "            'QC-L/T
        '63~67
        SQl = SQl + " '" + "0" + "', "
        SQl = SQl + " '" + "0" + "', "
        SQl = SQl + " '" + DOFormNo.Text + "', "            '���No
        If DOFormSno.Text = "" Then                         '�渹
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DOFormSno.Text + "', "
        End If

        SQl = SQl + " N'" + DPrice.Text + "', "             '����
        SQl = SQl + " N'" + DYearSeason.SelectedValue + "', "  '�~�u
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SufaceFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String
        Dim RtnCode As Integer

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_SufaceSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"     'NO
        SQl = SQl + " Date = N'" & DDate.Text & "',"                   '���
        '20140714
        'SQl = SQl + " Division = N'" & DDivision.SelectedValue & "',"  '����
        'SQl = SQl + " Person = N'" & DPerson.SelectedValue & "',"      '���
        If DCustSampleFile.Visible Then                                '�Ȥ�˫~��
            If DCustSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "CustSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DCustSampleFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " CustSampleFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " Spec = N'" & YKK.ReplaceString(DSpec.Text) & "',"              '�W��
        SQl = SQl + " Buyer = N'" & DBuyer.SelectedValue & "',"                      'Buyer
        SQl = SQl + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"  '�e�U�t��
        SQl = SQl + " ReqDelDate = N'" & DReqDelDate.Text & "',"                     '�Ʊ���
        SQl = SQl + " ReqQty = N'" & DReqQty.Text & "',"                             '�w���q
        SQl = SQl + " SliderSample = N'" & DSliderSample.Text & "',"                 '�˫~���Y
        SQl = SQl + " AttachSample = N'" & DAttachSample.SelectedValue & "',"        '����
        SQl = SQl + " ORNO = N'" & DORNO.Text & "',"                                 'OR-NO
        SQl = SQl + " OrderTime = N'" & DOrderTime.Text & "',"                       '�U��ɶ�
        SQl = SQl + " Price = N'" & DPrice.Text & "',"                               '����
        SQl = SQl + " DevReason = N'" & DDevReason.Text & "',"                       '�}�o�z��
        If DFinalSampleFile.Visible Then                                             '�̲׼˫~��
            If DFinalSampleFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "FinalSample" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DFinalSampleFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " FinalSampleFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " ManufType = N'" & DManufType.SelectedValue & "',"        '���s/�~�`
        SQl = SQl + " Suppiler = N'" & DSuppiler.SelectedValue & "',"          '�~�`��
        SQl = SQl + " Color = N'" & DColor.Text & "',"                         'Color
        SQl = SQl + " Qty = N'" & DQty.Text & "',"                             '�ƶq
        SQl = SQl + " Cap = N'" & DCap.Text & "',"                             '�鲣��
        SQl = SQl + " Schedule = N'" & DSchedule.Text & "',"                   '��¦��{
        SQl = SQl + " FReason = N'" & DFReason.Text & "',"                     '�z��
        SQl = SQl + " BFinalDate = N'" & DBFinalDate.Text & "',"               '�w�w������
        SQl = SQl + " Code = N'" & DCode.Text & "',"                           'Code
        SQl = SQl + " EnglishName = N'" & DEnglishName.Text & "',"             '�^��W��
        SQl = SQl + " LOSS = N'" & DLOSS.Text & "',"                           'LOSS
        If DQCReqFile.Visible Then                                             '�~��̿��
            If DQCReqFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCReq" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCReqFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCReqFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCReqFile = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " QCCheck1 = N'" & DQCCheck1.SelectedValue & "',"    '�f�|�o�k
        SQl = SQl + " QCCheck2 = N'" & DQCCheck2.SelectedValue & "',"    '�P�ʩ��
        SQl = SQl + " QCCheck3 = N'" & DQCCheck3.SelectedValue & "',"    'LOCK�j��
        SQl = SQl + " QCCheck4 = N'" & DQCCheck4.SelectedValue & "',"    '90�ױj��
        SQl = SQl + " QCCheck5 = N'" & DQCCheck5.SelectedValue & "',"    '��O
        SQl = SQl + " QCCheck15 = N'" & DQCCHECK15.SelectedValue & "',"    'N-ANTI
        SQl = SQl + " QCCheck16  = N'" & DQCCHECK16.SelectedValue & "',"    'PFAS

        ' If DQCCheck6File.Visible Then                                       '�q�ὤ�p
        ' If DQCCheck6File.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
        ' '*** IE8����-Start 2011/1/4
        ' 'FileName = CStr(wFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & Right(DQCCheck6File.PostedFile.FileName, InStr(StrReverse(DQCCheck6File.PostedFile.FileName), "\") - 1)
        ' FileName = CStr(wFormSno) & "-" & "QCCheck6" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCCheck6File.PostedFile.FileName)
        ' '*** IE8����-End
        ' Try    '�W�ǹ���
        ' DQCCheck6File.PostedFile.SaveAs(Path + FileName)
        ' Catch ex As Exception
        ' End Try
        ' SQl = SQl + " QCCheck6File = N'" & FileName & "',"
        ' End If
        ' End If

        SQl = SQl + " QCCheck7 = N'" & DQCCheck7.SelectedValue & "',"    '�˰w
        SQl = SQl + " QCCheck8 = N'" & DQCCheck8.SelectedValue & "',"    'AATCC
        SQl = SQl + " QCCheck9 = N'" & DQCCheck9.SelectedValue & "',"    '���~
        SQl = SQl + " QCCheck10 = N'" & DQCCheck10.SelectedValue & "',"  '�Q���Q��
        SQl = SQl + " QCCheck11 = N'" & DQCCheck11.SelectedValue & "',"  '�@���K��
        SQl = SQl + " QCCheck12 = N'" & DQCCheck12.SelectedValue & "',"  '�G���K��
        SQl = SQl + " QCCheck13 = N'" & DQCCheck13.SelectedValue & "',"  'Oeko-tex
        SQl = SQl + " QCCheck14 = N'" & DQCCheck14.SelectedValue & "',"  'A01
        SQl = SQl + " EACheck1 = N'" & DEACheck1.SelectedValue & "',"    '���`����
        SQl = SQl + " EADesc1 = N'" & DEADesc1.Text & "',"               '���`����Ƶ�

        If DEACheckFile.Visible Then                                     'Oeko-tex���`������i
            If DEACheckFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & Right(DEACheckFile.PostedFile.FileName, InStr(StrReverse(DEACheckFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "EACheck" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DEACheckFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " EACheckFile = N'" & FileName & "',"
            End If
        End If

        If DEACheckFile1.Visible Then                                     'A01���`������i
            If DEACheckFile1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & Right(DEACheckFile1.PostedFile.FileName, InStr(StrReverse(DEACheckFile1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "EACheck1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DEACheckFile1.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DEACheckFile1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " EACheckFile1 = N'" & FileName & "',"
            End If
        End If

        If DQCFinalFile.Visible Then                                     '���`������i
            If DQCFinalFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DQCFinalFile.PostedFile.FileName, InStr(StrReverse(DQCFinalFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFinalFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFinalFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFinalFile = N'" & FileName & "',"
            End If
        End If

        If DPFASFile.Visible Then                                     '���`������i
            If DPFASFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFinal" & "-" & UploadDateTime & "-" & Right(DPFASFile.PostedFile.FileName, InStr(StrReverse(DPFASFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "FAFS" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPFASFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DPFASFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " PFASFile = N'" & FileName & "',"
            End If
        End If


        SQl = SQl + " QCDate1 = N'" & DQCDate1.Text & "',"               '���-1
        SQl = SQl + " QCResult1 = N'" & DQCResult1.SelectedValue & "',"  '�˴����G-1
        SQl = SQl + " QCRemark1 = N'" & DQCRemark1.Text & "',"           '�Ƶ�-1
        SQl = SQl + " QCDate2 = N'" & DQCDate2.Text & "',"               '���-2
        SQl = SQl + " QCResult2 = N'" & DQCResult2.SelectedValue & "',"  '�˴����G-2
        SQl = SQl + " QCRemark2 = N'" & DQCRemark2.Text & "',"           '�Ƶ�-2
        SQl = SQl + " QCDate3 = N'" & DQCDate3.Text & "',"               '���-3
        SQl = SQl + " QCResult3 = N'" & DQCResult3.SelectedValue & "',"  '�˴����G-3
        SQl = SQl + " QCRemark3 = N'" & DQCRemark3.Text & "',"           '�Ƶ�-3
        If DManufFlowFile.Visible Then                                   '�s�y�y�{��
            If DManufFlowFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "ManufFlow" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DManufFlowFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ManufFlowFile = N'" & FileName & "',"
            End If
        End If
        If DForcastFile.Visible Then                                   '�s�y�y�{��
            If DForcastFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Forcast" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DForcastFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DForcastFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ForcastFile = N'" & FileName & "',"
            End If
        End If
        If DOPManualFile.Visible Then                                   '�s�y�y�{��
            If DOPManualFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "OPManual" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DOPManualFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DOPManualFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " OPManualFile = N'" & FileName & "',"
            End If
        End If
        If DContactFile.Visible Then                                   '�s�y�y�{��
            If DContactFile.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Contact" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DContactFile.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DContactFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " ContactFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " OFormNo = '" & DOFormNo.Text & "',"
        SQl = SQl + " OFormSno = '" & DOFormSno.Text & "',"
        SQl = SQl + " QCLT = '" & DQCLT.Text & "',"  'QC-L/T
        SQl = SQl + " YearSeason = '" & DYearSeason.SelectedValue & "',"  'QC-L/T

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
            If UPFile.PostedFile.ContentLength <= 2500 * 1024 Then
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
    Private Sub BPrint_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
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
        Dim wLevel As String = ""   '������

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pUserID=" & Request.QueryString("pUserID") & _
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
    '**(ModifyManufData)
    '**     ��s������
    '**
    '*****************************************************************
    Sub ModifyManufData(ByVal pFun As String, ByVal pSts As Integer, ByVal pSurface As Integer)
        If DOFormNo.Text = "000003" Or DOFormNo.Text = "000007" Or DOFormNo.Text = "000011" Then    ''�P�_�O�_�����NO
            Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                                 CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
            Dim SQl As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
            Dim OleDBCommand1 As New OleDbCommand

            If DOFormNo.Text = "000003" Then
                SQl = "Update F_ManufInSheet Set "
            End If
            If DOFormNo.Text = "000007" Then
                SQl = "Update F_ManufOutSheet Set "
            End If
            If DOFormNo.Text = "000011" Then
                SQl = "Update F_ImportSheet Set "
            End If

            SQl = SQl + " SFSts = '" & CStr(pSts) & "',"
            SQl = SQl + " Surface = '" & CStr(pSurface) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
            SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()
        End If
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
        'Spec
        If InputCheck = 0 Then
            If FindFieldInf("Spec") = 1 Then
                If DSpec.Text = "" Then InputCheck = 1
            End If
        End If
        'Buyer
        If InputCheck = 0 Then
            If FindFieldInf("Buyer") = 1 Then
                If DBuyer.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�e�U�t��
        If InputCheck = 0 Then
            If FindFieldInf("SellVendor") = 1 Then
                If DSellVendor.Text = "" Then InputCheck = 1
            End If
        End If
        '�Ʊ���
        If InputCheck = 0 Then
            If FindFieldInf("ReqDelDate") = 1 Then
                If DReqDelDate.Text = "" Then InputCheck = 1
            End If
        End If
        '�w���q
        If InputCheck = 0 Then
            If FindFieldInf("ReqQty") = 1 Then
                If DReqQty.Text = "" Then InputCheck = 1
            End If
        End If
        '�˫~���Y
        If InputCheck = 0 Then
            If FindFieldInf("SliderSample") = 1 Then
                If DSliderSample.Text = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("AttachSample") = 1 Then
                If DAttachSample.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'OR-NO
        If InputCheck = 0 Then
            If FindFieldInf("ORNO") = 1 Then
                If DORNO.Text = "" Then InputCheck = 1
            End If
        End If
        '�U��ɶ�
        If InputCheck = 0 Then
            If FindFieldInf("OrderTime") = 1 Then
                If DOrderTime.Text = "" Then InputCheck = 1
            End If
        End If
        '����
        If InputCheck = 0 Then
            If FindFieldInf("Price") = 1 Then
                If DPrice.Text = "" Then InputCheck = 1
            End If
        End If
        '�}�o�z��
        If InputCheck = 0 Then
            If FindFieldInf("DevReason") = 1 Then
                If DDevReason.Text = "" Then InputCheck = 1
            End If
        End If
        '�̲׼˫~��
        If InputCheck = 0 Then
            If FindFieldInf("FinalSampleFile") = 1 Then
                If DFinalSampleFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '���s/�~�`
        If InputCheck = 0 Then
            If FindFieldInf("ManufType") = 1 Then
                If DManufType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�~�`��
        If InputCheck = 0 Then
            If FindFieldInf("Suppiler") = 1 Then
                If DSuppiler.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�C��
        If InputCheck = 0 Then
            If FindFieldInf("Color") = 1 Then
                If DColor.Text = "" Then InputCheck = 1
            End If
        End If
        '�ƶq
        If InputCheck = 0 Then
            If FindFieldInf("Qty") = 1 Then
                If DQty.Text = "" Then InputCheck = 1
            End If
        End If
        'QC-L/T
        If InputCheck = 0 Then
            If FindFieldInf("QCLT") = 1 Then
                If DQCLT.Text = "" Then InputCheck = 1
            End If
        End If
        '�鲣��
        If InputCheck = 0 Then
            If FindFieldInf("Cap") = 1 Then
                If DCap.Text = "" Then InputCheck = 1
            End If
        End If
        '��{
        If InputCheck = 0 Then
            If FindFieldInf("Schedule") = 1 Then
                If DSchedule.Text = "" Then InputCheck = 1
            End If
        End If
        '��{
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If
        '���׼˫~
        If InputCheck = 0 Then
            If FindFieldInf("AllowSample") = 1 Then
                If DAllowSample.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�w�w������
        If InputCheck = 0 Then
            If FindFieldInf("BFinalDate") = 1 Then
                If DBFinalDate.Text = "" Then InputCheck = 1
            End If
        End If
        'Code
        If InputCheck = 0 Then
            If FindFieldInf("Code") = 1 Then
                If DCode.Text = "" Then InputCheck = 1
            End If
        End If
        '�^��W��
        If InputCheck = 0 Then
            If FindFieldInf("EnglishName") = 1 Then
                If DEnglishName.Text = "" Then InputCheck = 1
            End If
        End If
        'LOSS
        If InputCheck = 0 Then
            If FindFieldInf("LOSS") = 1 Then
                If DLOSS.Text = "" Then InputCheck = 1
            End If
        End If
        '�~��̿��
        If InputCheck = 0 Then
            If FindFieldInf("QCReqFile") = 1 Then
                If DQCReqFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�f�|�o�k
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck1") = 1 Then
                If DQCCheck1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�P�ʩ��
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck2") = 1 Then
                If DQCCheck2.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'LOCK�j��
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck3") = 1 Then
                If DQCCheck3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '90�ױj��
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck4") = 1 Then
                If DQCCheck4.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '��O
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck5") = 1 Then
                If DQCCheck5.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        'N-ANTI
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck15") = 1 Then
                If DQCCHECK15.SelectedValue = "" Then InputCheck = 1
            End If
        End If

        'N-ANTI
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck16") = 1 Then
                If DQCCHECK16.SelectedValue = "" Then InputCheck = 1
            End If
        End If



        '�q�ὤ�p
        ' If InputCheck = 0 Then
        'If FindFieldInf("QCCheck6File") = 1 Then
        'If DQCCheck6File.PostedFile.FileName = "" Then InputCheck = 1
        'End If
        'End If

        '�˰w
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck7") = 1 Then
                If DQCCheck7.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'AATCC
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck8") = 1 Then
                If DQCCheck8.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '���~
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck9") = 1 Then
                If DQCCheck9.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�Q���Q��
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck10") = 1 Then
                If DQCCheck10.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�@���K��
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck11") = 1 Then
                If DQCCheck11.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�G���K��,Oeko-tex,A01
        If InputCheck = 0 Then
            If FindFieldInf("QCCheck12") = 1 Then
                If DQCCheck12.SelectedValue = "" Then InputCheck = 1
                If DQCCheck13.SelectedValue = "" Then InputCheck = 1
                If DQCCheck14.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '���`����
        If InputCheck = 0 Then
            If FindFieldInf("EACheck1") = 1 Then
                If DEACheck1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '���`����Ƶ�
        If InputCheck = 0 Then
            If FindFieldInf("EADesc1") = 1 Then
                If DEADesc1.Text = "" Then InputCheck = 1
            End If
        End If
        'Oeko-tex,A01���`������i
        If InputCheck = 0 Then
            If FindFieldInf("EACheckFile") = 1 Then
                If DEACheckFile.PostedFile.FileName = "" Then InputCheck = 1
                If DEACheckFile1.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '���ճ��i��
        If InputCheck = 0 Then
            If FindFieldInf("QCFinalFile") = 1 Then
                If DQCFinalFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

        'FAFS
        If InputCheck = 0 Then
            If FindFieldInf("QCFinalFile") = 1 Then
                If DPFASFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If


        '���-1
        If InputCheck = 0 Then
            If FindFieldInf("QCDate1") = 1 Then
                If DQCDate1.Text = "" Then InputCheck = 1
            End If
        End If
        '�˴����G-1
        If InputCheck = 0 Then
            If FindFieldInf("QCResult1") = 1 Then
                If DQCResult1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�Ƶ�-1
        If InputCheck = 0 Then
            If FindFieldInf("QCRemark1") = 1 Then
                If DQCRemark1.Text = "" Then InputCheck = 1
            End If
        End If
        '���-2
        If InputCheck = 0 Then
            If FindFieldInf("QCDate2") = 1 Then
                If DQCDate2.Text = "" Then InputCheck = 1
                If DQCDate3.Text = "" Then InputCheck = 1
            End If
        End If
        '�˴����G-2
        If InputCheck = 0 Then
            If FindFieldInf("QCResult2") = 1 Then
                If DQCResult2.SelectedValue = "" Then InputCheck = 1
                If DQCResult3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�Ƶ�-2 �Ƶ�-3
        If InputCheck = 0 Then
            If FindFieldInf("QCRemark2") = 1 Then
                If DQCRemark2.Text = "" Then InputCheck = 1
                If DQCRemark3.Text = "" Then InputCheck = 1
            End If
        End If
        '�s�y�y�{��
        If InputCheck = 0 Then
            If FindFieldInf("ManufFlowFile") = 1 Then
                If DManufFlowFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '������
        If InputCheck = 0 Then
            If FindFieldInf("ForcastFile") = 1 Then
                If DForcastFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�@�~�зǮ�
        If InputCheck = 0 Then
            If FindFieldInf("OPManualFile") = 1 Then
                If DOPManualFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '������
        If InputCheck = 0 Then
            If FindFieldInf("ContactFile") = 1 Then
                If DContactFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�Ȥ�˫~��
        If InputCheck = 0 Then
            If FindFieldInf("CustSampleFile") = 1 Then
                If DCustSampleFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
    End Function

    Private Sub DAttachSample_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DAttachSample.SelectedIndexChanged

    End Sub
End Class
