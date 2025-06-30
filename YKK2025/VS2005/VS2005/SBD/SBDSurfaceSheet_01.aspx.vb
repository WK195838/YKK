Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDSurfaceSheet_01
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '�U���
    Dim Attribute(60) As Integer    '�U����ݩ�    
    Dim Top As String               '�w�]�������m
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
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wApplyCalendar As String = ""       '�ӽЪ�
    Dim wDecideCalendar As String = ""      '�֩w��
    Dim wNextGateCalendar As String = ""    '�U�@�֩w��
    'Modify-End

    Dim wUserName As String = ""    '�m�W�N�z��
    Dim HolidayList As New List(Of Integer) '�ΥH�O�����骺�����ޭ�
    Dim indexList As New List(Of Integer)   '�ΥH�O�����ݩ�������������ޭ�
    Dim DateList As New List(Of String)     '�ΥH�O���ҿ�����@�P���

    ''' <summary>
    ''' �H�U���@�Ψ禡���ŧi
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD�t�@�q�[��

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�
        '  TopPosition()           '���s��RequestedField��Top��m
        SetControlPosition()    ' �]�w�����m
        If Not Me.IsPostBack Then   '���OPostBack
            ShowSheetField("New")   '��������ܤ�����J�ˬd
            ShowSheetFunction()     '���\����s���

            If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
                ShowFormData()      '��ܪ����
                UpdateTranFile()    '��s������
                SetControlPosition()    ' �]�w�����m
            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�

        Else
            ShowSheetFunction()     '���\����s���
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text, Request.QueryString("pUserID")) ' �]�w�w�]��ñ�֪�
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
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '�{�b���
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

        Response.Cookies("PGM").Value = "SBDSurfaceSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X
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
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDSufaceFilePath")



        SQL = "Select * From F_SBDSufaceSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            SetFieldData("Buyer", DBAdapter1.Rows(0).Item("Buyer"))    'Buyer
            DSellVendor.Text = DBAdapter1.Rows(0).Item("SellVendor")              'Vendor
            DReqDelDate.Value = DBAdapter1.Rows(0).Item("ReqDelDate")              '�Ʊ���
            DReqQty.Text = DBAdapter1.Rows(0).Item("ReqQty")              '�Ʊ���
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              '�~�W
            SetFieldData("AttachSample", DBAdapter1.Rows(0).Item("AttachSample"))    '����
            DORNO.Text = DBAdapter1.Rows(0).Item("ORNO")              'ORNO

            If Mid(DBAdapter1.Rows(0).Item("OrderTime").ToString, 1, 4) = "1900" Then
                DOrderTime.Value = ""
            Else
                DOrderTime.Value = DBAdapter1.Rows(0).Item("OrderTime")               '�U��ɶ�
            End If


            DDevReason.Text = DBAdapter1.Rows(0).Item("DevReason")              '�}�o�z��
            DColor.Text = DBAdapter1.Rows(0).Item("Color")              'color code
            SetFieldData("SampleQty", DBAdapter1.Rows(0).Item("SampleQty"))    '����
            DEnglishName.Text = DBAdapter1.Rows(0).Item("EnglishName")              '�^��W��
            SetFieldData("Supplier", DBAdapter1.Rows(0).Item("Supplier"))    '�~�`�t��
            SetFieldData("AllowSample", DBAdapter1.Rows(0).Item("AllowSample"))    '���׼˫~
            DCap.Text = DBAdapter1.Rows(0).Item("Cap")              '�鲣��

            If Mid(DBAdapter1.Rows(0).Item("BFinalDate").ToString, 1, 4) = "1900" Then
                DBFinalDate.Value = ""
            Else
                DBFinalDate.Value = DBAdapter1.Rows(0).Item("BFinalDate")                '�w�w������
            End If




            DSchedule.Text = DBAdapter1.Rows(0).Item("Schedule")              '��¦��{

            If Trim(DBAdapter1.Rows(0).Item("QCReqFile")) <> "" Then
                LQCReqFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QCReqFile")  '�~��̿��
                LQCReqFile.Visible = True
            Else
                LQCReqFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("QAFile")) <> "" Then
                LQAFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QAFile")  '�~�����i��
                LQAFile.Visible = True
            Else
                LQAFile.Visible = False
            End If


            If Mid(DBAdapter1.Rows(0).Item("QCDate").ToString, 1, 4) = "1900" Then
                DQCDate.Value = ""
            Else
                DQCDate.Value = DBAdapter1.Rows(0).Item("QCDate")          '�~��P�w���
            End If




            SetFieldData("QCResult", DBAdapter1.Rows(0).Item("QCResult"))    '�˴����G
            DQCRemark.Text = DBAdapter1.Rows(0).Item("QCRemark")              '�~��Ƶ�


            If Trim(DBAdapter1.Rows(0).Item("ManufFlowFile")) <> "" Then
                LManufFlowFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ManufFlowFile")  '�s�y�y�{
                LManufFlowFile.Visible = True
            Else
                LManufFlowFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("OPManualFile")) <> "" Then
                LOPManualFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("OPManualFile")  '�@�~�y�{
                LOPManualFile.Visible = True
            Else
                LOPManualFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ForcastFile")) <> "" Then
                LForcastFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ForcastFile")  ' ������
                LForcastFile.Visible = True
            Else
                LForcastFile.Visible = False
            End If


            If Trim(DBAdapter1.Rows(0).Item("ContactFile")) <> "" Then
                LContactFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContactFile")  '������
                LContactFile.Visible = True
            Else
                LContactFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("CustSampleFile")) <> "" Then
                LCustSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("CustSampleFile")  '������
                LCustSampleFile.Visible = True
            Else
                LCustSampleFile.Visible = False
            End If

            If Trim(DBAdapter1.Rows(0).Item("FinalSampleFile")) <> "" Then
                LFinalSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("FinalSampleFile")  '������
                LFinalSampleFile.Visible = True
            Else
                LFinalSampleFile.Visible = False
            End If




            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"

            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then
                DDecideDesc.Text = DBAdapter2.Rows(0).Item("DecideDesc")       '����


                If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBAdapter2.Rows(0).Item("ReasonCode"))    '����z�ѥN�X
                        If DBAdapter2.Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '����z��
                            DReasonDesc.Text = ""   '�����L����
                        Else
                            DReason.Text = DBAdapter2.Rows(0).Item("Reason")  '����z��
                            DReasonDesc.Text = DBAdapter2.Rows(0).Item("ReasonDesc")     '�����L����
                        End If
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
        SQL = SQL + "Order by Unique_ID Desc "
        Dim dtWaitHandle1 As DataTable = uDataBase.GetDataTable(SQL)

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet9, "DecideHistory")
        GridView2.DataSource = dtWaitHandle1
        GridView2.DataBind()

        'DB�s������
        'OleDbConnection1.Close()

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

        'DB�s���]�w

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)


        If DBAdapter1.Rows.Count > 0 Then
            '�q�lñ�����ϥ�
            If DBAdapter1.Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '���[�ɮץ��ϥ�(��FormField���]�w)
            If DBAdapter1.Rows(0).Item("Attach") = 1 Then
            Else
            End If
            '�x�s���s
            If DBAdapter1.Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBAdapter1.Rows(0).Item("SaveDesc")
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1���s
            If DBAdapter1.Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBAdapter1.Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBAdapter1.Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2���s
            If DBAdapter1.Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBAdapter1.Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBAdapter1.Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK���s
            If DBAdapter1.Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBAdapter1.Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                wOKSts = DBAdapter1.Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '��Ǻ޲z
            If DBAdapter1.Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter2.Rows.Count > 0 Then

                '��Ǻ޲z
                If wDelay = 1 Then
                    If DBAdapter2.Rows(0).Item("BEndTime") < NowDateTime Then
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
                    Top = 1100
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 1000
                End If



                '������
                DDecideDesc.Visible = True      '����
                '�����ݿ�J
                If DBAdapter2.Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "���`�G�ݿ�J����")
                End If

                '���s��m
                BNG1.Style.Add("Top", Top)      'NG���s
                BNG2.Style.Add("Top", Top)     'NG���s
                BSAVE.Style.Add("Top", Top)    '�x�s���s
                BOK.Style.Add("Top", Top)      'OK���s

            End If
        Else
            Top = 800
            'Sheet����

            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False     '����
            DDescSheet.Visible = False
            DReasonCode.Visible = False     '����z�ѥN�X
            DReason.Visible = False         '����z��
            DReasonDesc.Visible = False     '�����L����
            DHistoryLabel.Visible = False  '�֩w�i��
            '���s��m
            BNG1.Style.Add("Top", Top)      'NG���s
            BNG2.Style.Add("Top", Top)     'NG���s
            BSAVE.Style.Add("Top", Top)    '�x�s���s
            BOK.Style.Add("Top", Top)      'OK���s

        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     �]�w�u�X�����ƥ�
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        '������
        BDate1.Attributes.Add("onClick", "calendarPicker('DOrderTime')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DBFinalDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DQCDate')")
        BDate4.Attributes.Add("onClick", "calendarPicker('DReqDelDate')")

        'BDate1.Attributes("onclick") = "calendarPicker('Form1.DMapDate');"
        'BDate2.Attributes("onclick") = "calendarPicker('Form1.DSampleDate');"
        'BDate3.Attributes("onclick") = "calendarPicker('Form1.DHalfFinishDdate');"

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
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String


        '���s��RequestedField��Top��m
        If wFormSno > 0 And wStep > 2 Then      '�P�_�O�_[ñ��]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

            If DBAdapter1.Rows.Count > 0 Then
                If DBAdapter1.Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 1100
                Else
                    If DDelay.Visible = True Then
                        Top = 1600
                    Else
                        Top = 1100
                    End If
                End If
            End If
        Else
            Top = 696
        End If
        '----


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
        LCustSampleFile.Visible = False

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
        LFinalSampleFile.Visible = False

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
        Select Case FindFieldInf("AppDate")
            Case 0  '���
                DAppDate.BackColor = Color.LightGray
                DAppDate.ReadOnly = True
                DAppDate.Visible = True

            Case 1  '�ק�+�ˬd
                DAppDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppDateRqd", "DDate", "���`�G�ݿ�J���")
                DAppDate.Visible = True

            Case 2  '�ק�
                DAppDate.BackColor = Color.Yellow
                DAppDate.Visible = True

            Case Else   '����
                DAppDate.Visible = False

        End Select
        If pPost = "New" Then DAppDate.Text = Now.ToString("yyyy/MM/dd") '�{�b���

        '�קﳡ��
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
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")

        '���
        Select Case FindFieldInf("Appper")
            Case 0  '���
                DAppper.BackColor = Color.LightGray
                DAppper.Visible = True
                DAppper.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DAppper.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppperRqd", "DAppper", "���`�G�ݿ�J���")
                DAppper.Visible = True
            Case 2  '�ק�
                DAppper.BackColor = Color.Yellow
                DAppper.Visible = True
            Case Else   '����
                DAppper.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Appper", "ZZZZZZ")



        'buyer
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
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "lightgrey")
                DReqDelDate.Attributes.Add("readonly", "true")
                BDate4.Disabled = True

            Case 1  '�ק�+�ˬd
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "greenyellow")
                DReqDelDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DReqDelDateRqd", "DReqDelDate", "���`�G�ݿ�J�Ʊ���")
                BDate4.Disabled = False

            Case 2  '�ק�
                DReqDelDate.Visible = True
                DReqDelDate.Style.Add("background-color", "yellow")
                DReqDelDate.Attributes.Add("readonly", "true")
                BDate4.Disabled = False

            Case Else   '����
                DReqDelDate.Visible = False
                BDate4.Disabled = True

        End Select
        If pPost = "New" Then DReqDelDate.Value = ""


        '�j�f�w���q
        Select Case FindFieldInf("ReqQty")
            Case 0  '���
                DReqQty.BackColor = Color.LightGray
                DReqQty.Visible = True
                DReqQty.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DReqQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReqQtyRqd", "DReqQty", "���`�G�ݿ�J�j�f�w���q")
                DReqQty.Visible = True
            Case 2  '�ק�
                DReqQty.BackColor = Color.Yellow
                DReqQty.Visible = True
            Case Else   '����
                DReqQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ReqQty", "ZZZZZZ")



        '�~�W
        Select Case FindFieldInf("Code")
            Case 0  '���
                DCode.BackColor = Color.LightGray
                DCode.Visible = True
                DCode.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "���`�G�ݿ�J�~�W")
                DCode.Visible = True
            Case 2  '�ק�
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '����
                DCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Code", "ZZZZZZ")



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



        'ORNO
        Select Case FindFieldInf("ORNO")
            Case 0  '���
                DORNO.BackColor = Color.LightGray
                DORNO.Visible = True
                DORNO.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DORNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DORNORqd", "DORNO", "���`�G�ݿ�JORNO")
                DORNO.Visible = True
            Case 2  '�ק�
                DORNO.BackColor = Color.Yellow
                DORNO.Visible = True
            Case Else   '����
                DORNO.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ORNO", "ZZZZZZ")


        '�U��ɶ�
        Select Case FindFieldInf("OrderTime")
            Case 0  '���
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "lightgrey")
                DOrderTime.Attributes.Add("readonly", "true")
                BDate1.Disabled = True

            Case 1  '�ק�+�ˬd
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "greenyellow")
                DOrderTime.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DOrderTimeRqd", "DOrderTime", "���`�G�ݿ�J�U��ɶ�")
                BDate1.Disabled = False

            Case 2  '�ק�
                DOrderTime.Visible = True
                DOrderTime.Style.Add("background-color", "yellow")
                DOrderTime.Attributes.Add("readonly", "true")
                BDate1.Disabled = False

            Case Else   '����
                DOrderTime.Visible = False
                BDate1.Disabled = True

        End Select
        If pPost = "New" Then DOrderTime.Value = ""



        '�}�o�z��
        Select Case FindFieldInf("DevReason")
            Case 0  '���
                DDevReason.BackColor = Color.LightGray
                DDevReason.Visible = True
                DDevReason.ReadOnly = True
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




        'Color Code
        Select Case FindFieldInf("Color")
            Case 0  '���
                DColor.BackColor = Color.LightGray
                DColor.Visible = True
                DColor.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorRqd", "DColor", "���`�G�ݿ�JColor Code")
                DColor.Visible = True
            Case 2  '�ק�
                DColor.BackColor = Color.Yellow
                DColor.Visible = True
            Case Else   '����
                DColor.Visible = False
        End Select
        If pPost = "New" Then DColor.Text = ""


        '�˫~�ݨD�q
        Select Case FindFieldInf("SampleQty")
            Case 0  '���
                DSampleQty.BackColor = Color.LightGray
                DSampleQty.Visible = True

            Case 1  '�ק�+�ˬd
                DSampleQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleQtyRqd", "DSampleQty", "���`�G�ݿ�J�˫~�ݨD�q")
                DSampleQty.Visible = True
            Case 2  '�ק�
                DSampleQty.BackColor = Color.Yellow
                DSampleQty.Visible = True
            Case Else   '����
                DSampleQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SampleQty", "ZZZZZZ")



        '�^��W��
        Select Case FindFieldInf("EnglishName")
            Case 0  '���
                DEnglishName.BackColor = Color.LightGray
                DEnglishName.Visible = True
                DEnglishName.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEnglishName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEnglishNameRqd", "DEnglishName", "���`�G�ݿ�JColor Code")
                DEnglishName.Visible = True
            Case 2  '�ק�
                DEnglishName.BackColor = Color.Yellow
                DEnglishName.Visible = True
            Case Else   '����
                DEnglishName.Visible = False
        End Select
        If pPost = "New" Then DEnglishName.Text = ""


        '�~�`�t��
        Select Case FindFieldInf("Supplier")
            Case 0  '���
                DSupplier.BackColor = Color.LightGray
                DSupplier.Visible = True

            Case 1  '�ק�+�ˬd
                DSupplier.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSupplierRqd", "DSupplier", "���`�G�ݿ�J�~�`�t��")
                DSupplier.Visible = True
            Case 2  '�ק�
                DSupplier.BackColor = Color.Yellow
                DSupplier.Visible = True
            Case Else   '����
                DSupplier.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Supplier", "ZZZZZZ")



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



        '�鲣��
        Select Case FindFieldInf("Cap")
            Case 0  '���
                DCap.BackColor = Color.LightGray
                DCap.Visible = True
                DCap.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DCap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCapRqd", "DCap", "���`�G�ݿ�J�鲣��")
                DCap.Visible = True
            Case 2  '�ק�
                DCap.BackColor = Color.Yellow
                DCap.Visible = True
            Case Else   '����
                DCap.Visible = False
        End Select
        If pPost = "New" Then DCap.Text = ""


        '�w�w�����鶡
        Select Case FindFieldInf("BFinalDate")
            Case 0  '���
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "lightgrey")
                DBFinalDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = True

            Case 1  '�ק�+�ˬd
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "greenyellow")
                DBFinalDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DBFinalDateRqd", "DBFinalDate", "���`�G�ݿ�J�w�w�����鶡")
                BDate2.Disabled = False

            Case 2  '�ק�
                DBFinalDate.Visible = True
                DBFinalDate.Style.Add("background-color", "yellow")
                DBFinalDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = False

            Case Else   '����
                DBFinalDate.Visible = False
                BDate2.Disabled = True

        End Select
        If pPost = "New" Then DBFinalDate.Value = ""


        '��Ǥ�{
        Select Case FindFieldInf("Schedule")
            Case 0  '���
                DSchedule.BackColor = Color.LightGray
                DSchedule.Visible = True
                DSchedule.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DSchedule.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DScheduleRqd", "DSchedule", "���`�G�ݿ�J��Ǥ�{")
                DSchedule.Visible = True
            Case 2  '�ק�
                DSchedule.BackColor = Color.Yellow
                DSchedule.Visible = True
            Case Else   '����
                DSchedule.Visible = False
        End Select
        If pPost = "New" Then DSchedule.Text = ""



        '�~��̿��
        Select Case FindFieldInf("QCReqFile")

            Case 0  '���
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "���`�G�ݿ�J���")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCReqFile.Visible = False
        End Select
        LQCReqFile.Visible = False


        '�~�����i��
        Select Case FindFieldInf("QAFile")

            Case 0  '���
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQAFileRqd", "DQAFile", "���`�G�ݿ�J�~�����i��")
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQAFile.Visible = True
                DQAFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQAFile.Visible = False
        End Select
        LQAFile.Visible = False



        '�~��P�w�ɶ�
        Select Case FindFieldInf("QCDate")
            Case 0  '���
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "lightgrey")
                DQCDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = True

            Case 1  '�ק�+�ˬd
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "greenyellow")
                DQCDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DQCDateRqd", "DQCDate", "���`�G�ݿ�J�~��P�w�ɶ�")
                BDate3.Disabled = False

            Case 2  '�ק�
                DQCDate.Visible = True
                DQCDate.Style.Add("background-color", "yellow")
                DQCDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = False

            Case Else   '����
                DQCDate.Visible = False
                BDate3.Disabled = True

        End Select
        If pPost = "New" Then DQCDate.Value = ""



        '�˴����G
        Select Case FindFieldInf("QCResult")
            Case 0  '���
                DQCResult.BackColor = Color.LightGray
                DQCResult.Visible = True
            Case 1  '�ק�+�ˬd
                DQCResult.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCResultRqd", "DQCResult", "���`�G�ݿ�J�˴����G")
                DQCResult.Visible = True
            Case 2  '�ק�
                DQCResult.BackColor = Color.Yellow
                DQCResult.Visible = True
            Case Else   '����
                DQCResult.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("QCResult", "ZZZZZZ")


        '�~���Ƶ� 
        Select Case FindFieldInf("QCRemark")
            Case 0  '���
                DQCRemark.BackColor = Color.LightGray
                DQCRemark.Visible = True
                DQCRemark.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DQCRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQCRemarkRqd", "DQCRemark", "���`�G�ݿ�J�~���Ƶ�")
                DQCRemark.Visible = True
            Case 2  '�ק�
                DQCRemark.BackColor = Color.Yellow
                DQCRemark.Visible = True
            Case Else   '����
                DQCRemark.Visible = False
        End Select
        If pPost = "New" Then DQCRemark.Text = ""


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
        LManufFlowFile.Visible = False


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
        LOPManualFile.Visible = False






        '������
        Select Case FindFieldInf("ForcastFile")
            Case 0  '���
                DForcastFile.Visible = False
                DForcastFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DForcastFile.Attributes.Add("readonly", "true")
                LForcastFile.Visible = True
                LForcastFile.BackColor = Color.LightGray
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
        LForcastFile.Visible = False


        '������
        Select Case FindFieldInf("ContactFile")
            Case 0  '���
                DContactFile.Visible = False
                DContactFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContactFile.Attributes.Add("readonly", "true")
                LContactFile.Visible = True
                LContactFile.BackColor = Color.LightGray

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
        LContactFile.Visible = False




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
        rqdVal.Display = ValidatorDisplay.None              ' �]�b�����W�[�JValidationSummary , �G���ұ���Τ@���
        rqdVal.Style.Add("Top", Top + 300 & "px")
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet

        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer


        idx = FindFieldInf(pFieldName)

        '���̤γ��� 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAppper.Text = DBUser.Rows(0).Item("Username")
        DDivision.Text = DBUser.Rows(0).Item("Divname")

        'buyer
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
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Buyer' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DBuyer.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
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
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'AttachSample' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DAttachSample.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAttachSample.Items.Add(ListItem1)
                Next
            End If
        End If

        '�˫~�ݨD�q
        If pFieldName = "SampleQty" Then
            DSampleQty.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSampleQty.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'SampleQty' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSampleQty.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSampleQty.Items.Add(ListItem1)
                Next
            End If
        End If




        '�~�`�t��
        If pFieldName = "Supplier" Then
            DSupplier.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSupplier.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Supplier'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSupplier.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSupplier.Items.Add(ListItem1)
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
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'AllowSample'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DAllowSample.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAllowSample.Items.Add(ListItem1)
                Next
            End If
        End If


        'QC �˴����G
        If pFieldName = "QCResult" Then
            DQCResult.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DQCResult.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'QCResult'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DQCResult.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DQCResult.Items.Add(ListItem1)
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
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("DKey")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("DKey")
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
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

                For i = 0 To DBAdapter1.Rows.Count - 1
                    DReason.Text = DBAdapter1.Rows(i).Item("Data")
                Next
            End If
        End If



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
    '**     �x�s���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            ' If InputCheck() = 0 Then
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

                Dim URL As String = uCommon.GetAppSetting("RedirectURL") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
            'End If
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
            DisabledButton()   '����Button�B�@
            FlowControl("OK", 0, "1")
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
            DisabledButton()   '����Button�B�@
            FlowControl("NG1", 1, "2")
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
            DisabledButton()   '����Button�B�@
            FlowControl("NG2", 2, "3")
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
                ErrCode = oCommon.CommissionNo("003002", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
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
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            'NG-1���s
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2���s
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK���s
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                   System.Configuration.ConfigurationManager.AppSettings("SBDSufaceFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String


        SQl = "Insert into F_SBDSufaceSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate, Division,APPPER,Buyer,"                  '1~5
        SQl = SQl + "SellVendor,ReqDelDate,ReqQty,Code,AttachSample, "               '6~10
        SQl = SQl + "ORNO,OrderTime,DevReason,Color,SampleQty, "                 '11~15
        SQl = SQl + "EnglishName,Supplier,AllowSample,Cap,BFinalDate, "   '16~20
        SQl = SQl + "Schedule,QCReqFile,QAFile,QCDate,QCResult, "                                                  '21~25
        SQl = SQl + "QCRemark,ManufFlowFile,OPManualFile,ForcastFile,ContactFile, "            '26~30
        SQl = SQl + "CustSampleFile,FinalSampleFile, "
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "


        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '003002', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Text + "', "                '���
        SQl = SQl + " '" + DDivision.Text + "', "   '����
        SQl = SQl + " '" + DAppper.Text + "', "     '���
        SQl = SQl + " '" + DBuyer.SelectedValue + "', "  'buyer

        SQl = SQl + " '" + DSellVendor.Text + "', "    '�e�U�t��
        SQl = SQl + " '" + DReqDelDate.Value + "', "      '�Ʊ���
        SQl = SQl + " N'" + YKK.ReplaceString(DReqQty.Text) + "', "   '�j�f�w���q
        SQl = SQl + " '" + DCode.Text + "', "     '�~�W
        SQl = SQl + " '" + DAttachSample.SelectedValue + "', "  '����


        SQl = SQl + " '" + DORNO.Text + "', "    'ORNO
        SQl = SQl + " '" + DOrderTime.Value + "',"   '�U��ɶ�
        SQl = SQl + " '" + DDevReason.Text + "', "   '�}�o�z��
        SQl = SQl + " '" + DColor.Text + "', "   'color code
        SQl = SQl + " '" + DSampleQty.SelectedValue + "', "   '�˫~�ݨD�q


        SQl = SQl + " '" + DEnglishName.Text + "', "    '�^��W��
        SQl = SQl + " '" + DSupplier.SelectedValue + "',"   '�~�`�t��
        SQl = SQl + " '" + DAllowSample.SelectedValue + "', "   '���׼˫~
        SQl = SQl + " '" + DCap.Text + "', "   '�鲣��
        SQl = SQl + " '" + DBFinalDate.Value + "', "   '�w�w������



        SQl = SQl + " '" + DSchedule.Text + "', "    '��Ǥ�{
 
        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)

                DQCReqFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '�~��̿�� 


        FileName = ""
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)

                DQAFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '�~�����i�� 

        SQl = SQl + " '" + DQCDate.Value + "', "    '���դ��
        SQl = SQl + " '" + DQCResult.SelectedValue + "', "    '�˴����G 



        SQl = SQl + " '" + DQCRemark.Text + "', "    '�˴��Ƶ� 
        FileName = ""
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ManufFlowFile" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                DManufFlowFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '�s�y�y�{��



        FileName = ""
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "OPManualFile" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                DOPManualFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '�@�~�зǮ�


        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "',"     '������

     
        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '������

        FileName = ""
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "CustSampleFile" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                DCustSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�U�ȼ˫~��


        FileName = ""
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "FinalSampleFile" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�̲׼˫~��



        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '�@����
        SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                       '�ק��
        SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     ��s�����
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
               System.Configuration.ConfigurationManager.AppSettings("SBDSufaceFilePath")
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String


        SQl = "Update F_SBDSufaceSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " AppDate = '" & DAppDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.Text & "',"
        SQl = SQl + " APPPER = '" & DAppper.Text & "',"
        SQl = SQl + " Buyer = '" & DBuyer.SelectedValue & "',"
        SQl = SQl + " SellVendor = '" & DSellVendor.Text & "',"

        SQl = SQl + " ReqDelDate = '" & DReqDelDate.Value & "',"
        SQl = SQl + " ReqQty = '" & DReqQty.Text & "',"

        SQl = SQl + " Code = '" & DCode.Text & "',"
        SQl = SQl + " AttachSample = '" & DAttachSample.SelectedValue & "',"
        SQl = SQl + " ORNO = '" & DORNO.Text & "',"
        SQl = SQl + " DevReason = '" & DDevReason.Text & "',"
        SQl = SQl + " Color = '" & DColor.Text & "',"
        SQl = SQl + " SampleQty = '" & DSampleQty.Text & "',"
        SQl = SQl + " EnglishName= '" & DEnglishName.Text & "',"

        SQl = SQl + " Supplier= '" & DSupplier.SelectedValue & "',"
        SQl = SQl + " AllowSample= '" & DAllowSample.Text & "',"
        SQl = SQl + " Cap = '" & DCap.Text & "',"

        SQl = SQl + " BFinalDate= '" & DBFinalDate.Value & "',"
        SQl = SQl + " Schedule= '" & DSchedule.Text & "',"


        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QCReqFile= N'" + YKK.ReplaceString(FileName) + "',"           '�~��̿��
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QAFile= N'" + YKK.ReplaceString(FileName) + "',"           '�~�����i��
            Else
                FileName = ""
            End If
        End If




        SQl = SQl + " QCDate= '" & DQCDate.Value & "',"                      '�˴����
        SQl = SQl + " QCResult= '" & DQCResult.SelectedValue & "',"                               '�˴����G
        SQl = SQl + " QCRemark= '" & DQCRemark.Text & "',"                                 '�˴��Ƶ�

        FileName = ""
        If DManufFlowFile.Visible Then
            If DManufFlowFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ManufFlowFile" & "-" & UploadDateTime & "-" & Right(DManufFlowFile.PostedFile.FileName, InStr(StrReverse(DManufFlowFile.PostedFile.FileName), "\") - 1)
                DManufFlowFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ManufFlowFile= N'" + YKK.ReplaceString(FileName) + "',"           '�s�y�y�{��
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DOPManualFile.Visible Then
            If DOPManualFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "OPManualFile" & "-" & UploadDateTime & "-" & Right(DOPManualFile.PostedFile.FileName, InStr(StrReverse(DOPManualFile.PostedFile.FileName), "\") - 1)
                DOPManualFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " OPManualFile= N'" + YKK.ReplaceString(FileName) + "',"           '�@�~�зǮ�
            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ForcastFile= N'" + YKK.ReplaceString(FileName) + "',"                 '������
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContactFile.Visible Then
            If DContactFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContactFile" & "-" & UploadDateTime & "-" & Right(DContactFile.PostedFile.FileName, InStr(StrReverse(DContactFile.PostedFile.FileName), "\") - 1)
                DContactFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContactFile= N'" + YKK.ReplaceString(FileName) + "',"                     '������
            Else
                FileName = ""
            End If
        End If


        FileName = ""
        If DCustSampleFile.Visible Then
            If DCustSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "CustSampleFile" & "-" & UploadDateTime & "-" & Right(DCustSampleFile.PostedFile.FileName, InStr(StrReverse(DCustSampleFile.PostedFile.FileName), "\") - 1)
                DCustSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " CustSampleFile= N'" + YKK.ReplaceString(FileName) + "',"
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DFinalSampleFile.Visible Then
            If DFinalSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "FinalSampleFile" & "-" & UploadDateTime & "-" & Right(DFinalSampleFile.PostedFile.FileName, InStr(StrReverse(DFinalSampleFile.PostedFile.FileName), "\") - 1)
                DFinalSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " FinalSampleFile= N'" + YKK.ReplaceString(FileName) + "',"
            Else
                FileName = ""
            End If
        End If





        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQl)
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
        uDataBase.ExecuteNonQuery(SQl)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     �l�[�����ƩM�e�U���Ӫ�
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String


        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQl)

        If DBAdapter1.Rows.Count <= 0 Then
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
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    uDataBase.ExecuteNonQuery(SQl)
                End If
            End If
        End If

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
    '**(SetControlPosition)
    '**    �]�w���s,����,����,�֩w�i������m
    '**
    '*****************************************************************
    Sub SetControlPosition()
        TopPosition()
        If DDescSheet.Visible Then                                      ' ����
            DDescSheet.Style("top") = Top - 250 & "px"
            DDecideDesc.Style("top") = Top - 250 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' ���𻡩�
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If

        BSAVE.Style("top") = Top + 10 & "px"
        BNG1.Style("top") = Top + 10 & "px"
        BNG2.Style("top") = Top + 10 & "px"
        BOK.Style("top") = Top + 10 & "px"

        Top += 48

        If GridView2.Rows.Count > 0 Then                                ' �֩w�i��
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
    End Sub



    Protected Sub DNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.TextChanged

    End Sub

    Protected Sub DBFinalDate_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBFinalDate.ServerChange

    End Sub
End Class

