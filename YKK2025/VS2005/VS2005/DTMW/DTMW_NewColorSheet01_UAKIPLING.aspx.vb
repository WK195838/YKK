Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet01_UAKIPLING
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
    Dim wbFormSno As Integer        '�s��_��-���y����
    Dim wbStep As Integer           '�s��_��-�u�{�N�X
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim wUserID As String          '�ϥΪ�ID
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
    Dim oWaves As New WAVES.CommonService
    Dim aAgentID As String
    Dim Makemap1, Makemap2, Makemap3, Makemap4, Makemap5, Makemap6 As Integer
    Dim AQty, DevNo, Manufout As String
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim LastStep As String






    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�
        TopPosition()           '���s��RequestedField��Top��m
        SetControlPosition()    ' �]�w�����m

        If Not Me.IsPostBack Then   '���OPostBack
            ShowSheetField("New")   '��������ܤ�����J�ˬd
            ShowSheetFunction()     '���\����s���

            If wFormSno > 0 And wStep > 2 Then    '�P�_�O�_[ñ��]
                ShowFormData()      '��ܪ����
                UpdateTranFile()    '��s������
                TopPosition()
                SetControlPosition()    ' �]�w�����m

            End If
            SetPopupFunction()      '�]�w�u�X�����ƥ�



        Else
            ShowSheetFunction()     '���\����s���
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,wUserID) ' �]�w�w�]��ñ�֪�
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
        wbFormSno = Request.QueryString("pFormSno")  '�s��_��άy����
        wbStep = Request.QueryString("pStep")        '�s��_��Τu�{�N�X

        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID

        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        If wStep = 10 Or wStep = 510 Or wStep = 20 Or wStep = 520 Or wStep = 520 Or wStep = 530 Then '���w�}�z��
            Dim SQL As String

            SQL = "Select * From T_waitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and step ='" + CStr(wStep) + "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                wUserID = DBAdapter1.Rows(0).Item("Workid")
            End If

        Else
            wUserID = Request.QueryString("pUserID")
        End If

        'wUserID = Request.QueryString("UserID")
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "DTMW_NewColorSheet01_03CFP12.aspx"
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
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo,wUserID)
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, wUserID)
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

        SQL = "Select * From F_NewColorUAKIPLING "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then '�Ȥ�ۮ�
                DCustomerCheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = True   ' ����20
                    BNG2.Visible = True
                End If
            End If

            If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then '�u�t�ۮ�
                DFactoryCheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = False
                    BNG2.Visible = True   ' ����40
                End If
            End If

            If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then '�u�t�ۮ�
                DVCACheck.Checked = True
                If (wStep = 10) Or (wStep = 510) Then
                    BOK.Visible = False
                    BNG2.Visible = True   ' ����40
                End If
            End If


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '�~
            SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '�u

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DReceiveDate.Text = ""
            Else
                DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '�U��ɶ�
            End If

            SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '���O1
            SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '���O2
            SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '���O2

            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '���O2

            If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                DDeliveryDate.Value = ""
            Else
                DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '�U��ɶ�
            End If

            If wStep = 500 Then  '(NG�A�e�X�Ʊ����ݭ��s���)
                DDeliveryDate.Value = ""
            End If


            DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
            If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                DNOCCS.Checked = True
            End If
            DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
            DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
            SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '���O2

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DCheckNo.Text = DBAdapter1.Rows(0).Item("CheckNo")
            DPFBWire.Value = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Value = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '���O2

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("YKKColorCodeVF")
            DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("YKKColorCodeSLD")

            If wStep = 70 Then
                If DBAdapter1.Rows(0).Item("DTNO") = 0 Then

                    DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("VFColor")
                    DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("SLDColor")
                End If

                If DYKKColorCodeVF.Value = "" Then
             
                    DYKKColorCodeVF.Value = DBAdapter1.Rows(0).Item("VFColor")
                End If

                If DYKKColorCodeSLD.Value = "" Then
                
                    DYKKColorCodeSLD.Value = DBAdapter1.Rows(0).Item("SLDColor")
                End If




            End If



            Dim Version As Integer
            Dim Formname As String

            If wStep = 20 Or wStep = 520 Then
                Formname = "3CF DTM"
            Else
                Formname = "5VS  DTM"
            End If

            '�p��O�ĴX��         
            SQL = " select  count(*)cun  from f_newcolorcomplete "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And FormName =  '" & Formname & "'"

            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter3.Rows.Count > 0 Then
                Version = DBAdapter3.Rows(0).Item("cun")
            Else
                Version = 1
            End If

            If wStep = 520 Or wStep = 530 Then
                Version = Version + 1
            End If



            DVersion.Text = Version
            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            SetFieldData("ColorType", DBAdapter1.Rows(0).Item("again"))    '�@�H��
            If DBAdapter1.Rows(0).Item("again") = 1 Then
                DAgain.SelectedValue = "�H��"
            ElseIf DBAdapter1.Rows(0).Item("again") = 2 Then
                DAgain.SelectedValue = "�@��"
            End If

            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
            DVFColor1.Text = DBAdapter1.Rows(0).Item("VFColor1")


            If DBAdapter1.Rows(0).Item("Complete") = 1 Then       '���s��̿৹����
                LComplete.NavigateUrl = "NewColoCompleteList.aspx?pNo=" & DNo.Text
                LComplete.Visible = True
            Else
                LComplete.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("DTNO") = 1 Then       '���l�[�֥i��
                LDTSheet.NavigateUrl = "NewColorDTSheetList.aspx?pNo=" & DNo.Text
                LDTSheet.Visible = True
            Else
                LDTSheet.Visible = False
            End If

        End If



        If wStep = 80 Then  '�j���ܦ��ťճ̫��J���
            DSLDColor.Text = ""
            DAgain.SelectedValue = ""
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




        '�֩w�i�����
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        ' SQL = SQL + "Order by Unique_ID Desc "
        SQL = SQL + " Order by CreateTime Desc, Step Desc, SeqNo Desc "
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
                    Top = 1200
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    If DBAdapter2.Rows(0).Item("Flowtype") = "1" Then
                        Top = 1120
                    Else
                        Top = 1200
                    End If

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
            Top = 1100
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
        BDate.Attributes.Add("onClick", "calendarPicker('DDeliveryDate')")
        BCustomer.Attributes.Add("onClick", "GetCustomer()") '��Ȥ�
        BBuyer.Attributes.Add("onClick", "GetBuyer()") '��buyer
        BCopySheet.Attributes.Add("onClick", "CopyNewColor('" + wFormNo + "')") '��buyer

        BYKKColor.Attributes.Add("onClick", "GetYKKColor()") '��buyer
        BYKKColorCode.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCode')") '��buyer
        BYKKColorCodeVF.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCodeVF')") '��buyer
        BYKKColorCodeSLD.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCodeSLD')") '��buyer

        Me.DColorSystem.Attributes.CssStyle.Add("text-transform", "uppercase")

        BSLDColor.Attributes.Add("onClick", "CheckColor('" + DYKKColorCode.Value + "')") '��buyer


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
                    Top = 1200
                Else
                    If DDelay.Visible = True Then
                        Top = 1200
                    Else
                        Top = 1200
                    End If
                End If
            End If
        Else
            Top = 1600

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



        If wFormNo = "005011" Then
            LSheetName.Text = "����ZIPPER CHAIN (UA��KIPLING)" '���W��
        End If



        Dim SQL As String
        Dim InputData1 As String
        Dim InputData2 As String


        InputData1 = D1.Text
        InputData2 = D2.Text


        '�a�JCUSTOMER
        If InputData1 <> "" Then

            SQL = " Select  * from  MST_Custmer where 1=1 "
            If InputData1 <> "" Then
                SQL = SQL + " and ( custmer like '%" + InputData1 + "%' or name_c like '%" + InputData1 + "%')"
            End If
            SQL = SQL + " order by custmer,name_c "

            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DCustomer.Text = DBData.Rows(0).Item("Name_C")
            DCustomerCode.Text = DBData.Rows(0).Item("Custmer")
        End If

        '�a�JBUYER
        If InputData2 <> "" Then


            SQL = " Select  * from  MST_Buyer where 1=1 "
            If InputData2 <> "" Then
                SQL = SQL + " and ( buyer like '%" + InputData2 + "%' or buyer_name like '%" + InputData2 + "%')"
            End If

            SQL = SQL + " order by buyer_name,buyer "
            Dim DBData As DataTable = uDataBase.GetDataTable(SQL)

            DBuyer.Text = DBData.Rows(0).Item("Buyer_Name")
            DBuyerCode.Text = DBData.Rows(0).Item("Buyer")

        End If

        'jessica 20221202 DTMW �⸹�ˬd
        '�a�JSLDCOLOR
        If D3.Text <> "" Then
            DSLDColor.Text = D3.Text
            DYKKColorCodeSLD.Value = D3.Text
        End If


        'Modify 20150417
        '�a�J�¦�
        If wStep = 70 Then
            If DNewOldColor.SelectedValue = "�¦�" Then

                SQL = " Select  ltrim(CPS1XX)CPS1XX,ltrim(CPT1XX)CPT1XX from   MST_color_structure where  paccxx ='" + DYKKColorCode.Value + "'"
                SQL = SQL + " and ctbnxx ='1'"


                Dim DBData As DataTable = uDataBase.GetDataTable(SQL)
                If DBData.Rows.Count > 0 Then
                    DSLDColor.Text = DBData.Rows(0).Item("CPS1XX")
                    DVFColor.Text = DBData.Rows(0).Item("CPT1XX")
                End If
            Else
                DSLDColor.Text = ""
                DVFColor.Text = ""
            End If
        End If




        '�Ȥ�ۮ�
        Select Case FindFieldInf("CustomerCheck")
            Case 0  '���
                ' DCustomerCheck.BackColor = Color.LightGray
                DCustomerCheck.Enabled = False

            Case 1  '�ק�+�ˬd
                DCustomerCheck.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DCustomerCheckRqd", "DCustomerCheck", "���`�G�ݿ�J�Ȥ�ۮ�")

            Case 2  '�ק�
                DCustomerCheck.BackColor = Color.Yellow

            Case Else   '����
                DCustomerCheck.Visible = False
        End Select
        If pPost = "New" Then DCustomerCheck.Checked = True

        '�u�t�ۮ�
        Select Case FindFieldInf("FactoryCheck")
            Case 0  '���
                '    DCustomerCheck.BackColor = Color.LightGray
                DFactoryCheck.Enabled = False

            Case 1  '�ק�+�ˬd
                DFactoryCheck.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFactoryCheckRqd", "DFactoryCheck", "���`�G�ݿ�J�u�t�ۮ�")

            Case 2  '�ק�
                DFactoryCheck.BackColor = Color.Yellow

            Case Else   '����
                DFactoryCheck.Visible = False
        End Select
        If pPost = "New" Then DFactoryCheck.Checked = False


        'VCA��
        Select Case FindFieldInf("VCACheck")
            Case 0  '���
                ' DCustomerCheck.BackColor = Color.LightGray
                DVCACheck.Enabled = False

            Case 1  '�ק�+�ˬd
                DVCACheck.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVCACheckRqd", "VCACheck", "���`�G�ݿ�JVCA")

            Case 2  '�ק�
                DVCACheck.BackColor = Color.Yellow

            Case Else   '����
                DVCACheck.Visible = False
        End Select
        If pPost = "New" Then DVCACheck.Checked = False



        'VCA��
        Select Case FindFieldInf("NOCCS")
            Case 0  '���
                ' DCustomerCheck.BackColor = Color.LightGray
                DNOCCS.Enabled = False

            Case 1  '�ק�+�ˬd
                DNOCCS.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DNOCCSRqd", "NOCCS", "���`�G�ݿ�J�L�kCCS")

            Case 2  '�ק�
                DNOCCS.BackColor = Color.Yellow

            Case Else   '����
                DNOCCS.Visible = False
        End Select
        If pPost = "New" Then DNOCCS.Checked = False


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

            Case 1  '�ק�+�ˬd
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J���")
                DDate.Visible = True

            Case 2  '�ק�
                DDate.BackColor = Color.Yellow
                DDate.Visible = True

            Case Else   '����
                DDate.Visible = False

        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '�{�b���

        'DepoName
        Select Case FindFieldInf("DepoName")
            Case 0  '���
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '�ק�+�ˬd
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "���`�G�ݿ�J����")
                DDepoName.Visible = True
            Case 2  '�ק�
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '����
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DepoName", "ZZZZZZ")


        'Name
        Select Case FindFieldInf("Name")
            Case 0  '���
                DName.BackColor = Color.LightGray
                DName.ReadOnly = True
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


        'Customer
        Select Case FindFieldInf("Customer")
            Case 0  '���
                DCustomer.BackColor = Color.LightGray
                DCustomer.ReadOnly = True
                DCustomer.Visible = True
                BCustomer.Visible = False
                BCopySheet.Visible = False
            Case 1  '�ק�+�ˬd
                DCustomer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerRqd", "DCustomer", "���`�G�ݿ�J�Ȥ�W��")
                DCustomer.Visible = True
                DCustomer.ReadOnly = True
                BCustomer.Visible = True
                BCopySheet.Visible = True
            Case 2  '�ק�
                DCustomer.BackColor = Color.Yellow
                DCustomer.Visible = True
                BCustomer.Visible = True
                BCopySheet.Visible = True
            Case Else   '����
                DCustomer.Visible = False
                BCustomer.Visible = False
                BCopySheet.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Customer", "ZZZZZZ")


        'CustomerCode
        Select Case FindFieldInf("CustomerCode")
            Case 0  '���
                DCustomerCode.BackColor = Color.LightGray
                DCustomerCode.ReadOnly = True
                DCustomerCode.Visible = True
            Case 1  '�ק�+�ˬd
                DCustomerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerCodeRqd", "DCustomerCode", "���`�G�ݿ�J�Ȥ�W��")
                DCustomerCode.Visible = True
                DCustomer.ReadOnly = True
            Case 2  '�ק�
                DCustomerCode.BackColor = Color.Yellow
                DCustomerCode.Visible = True
            Case Else   '����
                DCustomerCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerCode", "ZZZZZZ")



        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '���
                DBuyer.BackColor = Color.LightGray
                DBuyer.ReadOnly = True
                DBuyer.Visible = True
                BBuyer.Visible = False
            Case 1  '�ק�+�ˬd
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "���`�G�ݿ�JBuyer")
                DBuyer.Visible = True
                DBuyer.ReadOnly = True
                BBuyer.Visible = True
            Case 2  '�ק�
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
                BBuyer.Visible = True
            Case Else   '����
                DBuyer.Visible = False
                BBuyer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Buyer", "ZZZZZZ")


        'BuyerCode
        Select Case FindFieldInf("BuyerCode")
            Case 0  '���
                DBuyerCode.BackColor = Color.LightGray
                DBuyerCode.ReadOnly = True
                DBuyerCode.Visible = True
            Case 1  '�ק�+�ˬd
                DBuyerCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerCodeRqd", "DBuyerCode", "���`�G�ݿ�JBuyer")
                DBuyerCode.Visible = True
                DBuyerCode.ReadOnly = True
            Case 2  '�ק�
                DBuyerCode.BackColor = Color.Yellow
                DBuyerCode.Visible = True
            Case Else   '����
                DBuyerCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BuyerCode", "ZZZZZZ")




        'CustomerColor
        Select Case FindFieldInf("CustomerColor")
            Case 0  '���
                DCustomerColor.BackColor = Color.LightGray
                DCustomerColor.ReadOnly = True
                DCustomerColor.Visible = True
            Case 1  '�ק�+�ˬd
                DCustomerColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerColorRqd", "DCustomerColor", "���`�G�ݿ�J�Ȥ��W")
                DCustomerColor.Visible = True
                DCustomerColor.ReadOnly = True
            Case 2  '�ק�
                DCustomerColor.BackColor = Color.Yellow
                DCustomerColor.Visible = True
            Case Else   '����
                DCustomerColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerColor", "ZZZZZZ")

        'CustomerColorCode
        Select Case FindFieldInf("CustomerColorCode")
            Case 0  '���
                DCustomerColorCode.BackColor = Color.LightGray
                DCustomerColorCode.ReadOnly = True
                DCustomerColorCode.Visible = True
            Case 1  '�ק�+�ˬd
                DCustomerColorCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerColorCodeRqd", "DCustomerColorCode", "���`�G�ݿ�J�Ȥ�⸹")
                DCustomerColorCode.Visible = True
                DCustomerColorCode.ReadOnly = True
            Case 2  '�ק�
                DCustomerColorCode.BackColor = Color.Yellow
                DCustomerColorCode.Visible = True
            Case Else   '����
                DCustomerColorCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerColorCode", "ZZZZZZ")


        'OverseaYKKCode
        Select Case FindFieldInf("OverseaYKKCode")
            Case 0  '���
                DOverseaYKKCode.BackColor = Color.LightGray
                DOverseaYKKCode.ReadOnly = True
                DOverseaYKKCode.Visible = True
            Case 1  '�ק�+�ˬd
                DOverseaYKKCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOverseaYKKCodeRqd", "DOverseaYKKCode", "���`�G�ݿ�J���~YKK�⸹")
                DOverseaYKKCode.Visible = True
                DOverseaYKKCode.ReadOnly = True
            Case 2  '�ק�
                DOverseaYKKCode.BackColor = Color.Yellow
                DOverseaYKKCode.Visible = True
            Case Else   '����
                DOverseaYKKCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("OverseaYKKCode", "ZZZZZZ")


        'PANTONECode
        Select Case FindFieldInf("PANTONECode")
            Case 0  '���
                DPANTONECode.BackColor = Color.LightGray
                DPANTONECode.ReadOnly = True
                DPANTONECode.Visible = True
            Case 1  '�ק�+�ˬd
                DPANTONECode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPANTONECodeRqd", "DPANTONECode", "���`�G�ݿ�JPANTONE�⸹")
                DPANTONECode.Visible = True
                DPANTONECode.ReadOnly = True
            Case 2  '�ק�
                DPANTONECode.BackColor = Color.Yellow
                DPANTONECode.Visible = True
            Case Else   '����
                DPANTONECode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("PANTONECode", "ZZZZZZ")




        '�~
        Select Case FindFieldInf("DevYear")
            Case 0  '���
                DDevYear.BackColor = Color.LightGray
                DDevYear.Visible = True

            Case 1  '�ק�+�ˬd
                DDevYear.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevYearRqd", "DDevYear", "���`�G�ݿ�J�~")
                DDevYear.Visible = True
            Case 2  '�ק�
                DDevYear.BackColor = Color.Yellow
                DDevYear.Visible = True
            Case Else   '����
                DDevYear.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DevYear", "ZZZZZZ")


        '�~
        Select Case FindFieldInf("DevSeason")
            Case 0  '���
                DDevSeason.BackColor = Color.LightGray
                DDevSeason.Visible = True

            Case 1  '�ק�+�ˬd
                DDevSeason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevSeasonRqd", "DevSeason", "���`�G�ݿ�J�u")
                DDevSeason.Visible = True
            Case 2  '�ק�
                DDevSeason.BackColor = Color.Yellow
                DDevSeason.Visible = True
            Case Else   '����
                DDevSeason.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("DevSeason", "ZZZZZZ")

        'DuplicateNo
        Select Case FindFieldInf("DuplicateNo")
            Case 0  '���
                DDuplicateNo.BackColor = Color.LightGray
                DDuplicateNo.ReadOnly = True
                DDuplicateNo.Visible = True
            Case 1  '�ק�+�ˬd
                DDuplicateNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDuplicateNoRqd", "DDuplicateNo", "���`�G�ݿ�J���Ш̿઺�s��")
                DDuplicateNo.Visible = True
                DDuplicateNo.ReadOnly = True
            Case 2  '�ק�
                DDuplicateNo.BackColor = Color.Yellow
                DDuplicateNo.Visible = True
            Case Else   '����
                DDuplicateNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DuplicateNo", "ZZZZZZ")




        'NOCCSReason
        Select Case FindFieldInf("NOCCSReason")
            Case 0  '���
                DNOCCSReason.BackColor = Color.LightGray
                DNOCCSReason.ReadOnly = True
                DNOCCSReason.Visible = True
            Case 1  '�ק�+�ˬd
                DNOCCSReason.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DNOCCSReasonRqd", "DNOCCSReason", "���`�G�ݿ�J�L�kCCS��]")
                DNOCCSReason.Visible = True
            Case 2  '�ק�
                DNOCCSReason.BackColor = Color.Yellow
                DNOCCSReason.Visible = True
            Case Else   '����
                DNOCCSReason.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NOCCSReason", "ZZZZZZ")

        'CustomerNGColor
        Select Case FindFieldInf("CustomerNGColor")
            Case 0  '���
                DCustomerNGColor.BackColor = Color.LightGray
                DCustomerNGColor.ReadOnly = True
                DCustomerNGColor.Visible = True
            Case 1  '�ק�+�ˬd
                DCustomerNGColor.BackColor = Color.GreenYellow
                '  ShowRequiredFieldValidator("DCustomerNGColorRqd", "DCustomerNGColor", "���`�G�ݿ�J�Ȥ�_�M��YKK�⸹")
                DCustomerNGColor.Visible = True

            Case 2  '�ק�       

                DCustomerNGColor.BackColor = Color.Yellow
                DCustomerNGColor.Visible = True

            Case Else   '����
                DCustomerNGColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerNGColor", "ZZZZZZ")



        'Remark
        Select Case FindFieldInf("Remark")
            Case 0  '���
                DRemark.BackColor = Color.LightGray
                DRemark.ReadOnly = True
                DRemark.Visible = True
            Case 1  '�ק�+�ˬd
                DRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "���`�G�ݿ�J�Ƶ�")
                DRemark.Visible = True
            Case 2  '�ק�
                DRemark.BackColor = Color.Yellow
                DRemark.Visible = True
            Case Else   '����
                DRemark.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Remark", "ZZZZZZ")


        '�Ȥ���
        Select Case FindFieldInf("CustomerSample")
            Case 0  '���
                DCustomerSample.BackColor = Color.LightGray
                DCustomerSample.Visible = True

            Case 1  '�ק�+�ˬd
                DCustomerSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCustomerSampleRqd", "DCustomerSample", "���`�G�ݿ�Ȥ���")
                DCustomerSample.Visible = True
            Case 2  '�ק�
                DCustomerSample.BackColor = Color.Yellow
                DCustomerSample.Visible = True
            Case Else   '����
                DCustomerSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CustomerSample", "ZZZZZZ")





        'Version
        Select Case FindFieldInf("Version")
            Case 0  '���
                DVersion.BackColor = Color.LightGray
                DVersion.ReadOnly = True
                DVersion.Visible = True
            Case 1  '�ק�+�ˬd
                DVersion.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVersionRqd", "DVersion", "���`�G�ݿ�J����")
                DVersion.Visible = True
                DVersion.ReadOnly = True
            Case 2  '�ק�
                DVersion.BackColor = Color.Yellow
                DVersion.Visible = True
            Case Else   '����
                DVersion.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Version", "ZZZZZZ")


        'CheckNo
        Select Case FindFieldInf("CheckNo")
            Case 0  '���
                DCheckNo.BackColor = Color.LightGray
                DCheckNo.ReadOnly = True
                DCheckNo.Visible = True
            Case 1  '�ק�+�ˬd
                DCheckNo.BackColor = Color.GreenYellow
                '    ShowRequiredFieldValidator("DCheckNoRqd", "DCheckNo", "���`�G�ݿ�J�T�{�֥i�N��")
                DCheckNo.Visible = True

            Case 2  '�ק�
                DCheckNo.BackColor = Color.Yellow
                DCheckNo.Visible = True
            Case Else   '����
                DCheckNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CheckNo", "ZZZZZZ")



        'ReceiveDate
        Select Case FindFieldInf("ReceiveDate")
            Case 0  '���
                DReceiveDate.BackColor = Color.LightGray
                DReceiveDate.ReadOnly = True
                DReceiveDate.Visible = True
            Case 1  '�ק�+�ˬd
                DReceiveDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReceiveDateRqd", "DReceiveDate", "���`�G�ݿ�J����")
                DReceiveDate.Visible = True
                DReceiveDate.ReadOnly = True
            Case 2  '�ק�
                DReceiveDate.BackColor = Color.Yellow
                DReceiveDate.Visible = True
            Case Else   '����
                DReceiveDate.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ReceiveDate", "ZZZZZZ")



        'SLDColor
        Select Case FindFieldInf("SLDColor")
            Case 0  '���
                DSLDColor.BackColor = Color.LightGray
                DSLDColor.ReadOnly = True
                DSLDColor.Visible = True
                BYKKColor.Visible = False
            Case 1  '�ק�+�ˬd
                DSLDColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDColorRqd", "DSLDColor", "���`�G�ݿ�J�T�{���Y�ݥΦ�")
                DSLDColor.Visible = True

            Case 2  '�ק�
                DSLDColor.BackColor = Color.Yellow
                DSLDColor.Visible = True
                DSLDColor.ReadOnly = False
            Case Else   '����
                DSLDColor.Visible = False
                BYKKColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("SLDColor", "ZZZZZZ")

        'VFColor1 
        Select Case FindFieldInf("VFColor")
            Case 0  '���
                DVFColor.BackColor = Color.LightGray
                DVFColor.ReadOnly = True
                DVFColor.Visible = True
            Case 1  '�ק�+�ˬd
                DVFColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColorRqd", "DVFColor", "���`�G�ݿ�J�T�{VF����")
                DVFColor.Visible = True

            Case 2  '�ק�
                DVFColor.BackColor = Color.Yellow
                DVFColor.Visible = True
            Case Else   '����
                DVFColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor", "ZZZZZZ")

        'VFColor11 
        Select Case FindFieldInf("VFColor1")
            Case 0  '���
                DVFColor1.BackColor = Color.LightGray
                DVFColor1.ReadOnly = True
                DVFColor1.Visible = True
            Case 1  '�ק�+�ˬd
                DVFColor1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColor1Rqd", "DVFColor1", "���`�G�ݿ�J�T�{VF�ݥΦ�")
                DVFColor1.Visible = True

            Case 2  '�ק�
                DVFColor1.BackColor = Color.Yellow
                DVFColor1.Visible = True
            Case Else   '����
                DVFColor1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor1", "ZZZZZZ")


        '������
        Select Case FindFieldInf("ColorLight1")
            Case 0  '���
                DColorLight1.BackColor = Color.LightGray
                DColorLight1.Visible = True

            Case 1  '�ק�+�ˬd
                DColorLight1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorLight1Rqd", "DColorLight1", "���`�G�ݿ�Ĥ@������")
                DColorLight1.Visible = True
            Case 2  '�ק�
                DColorLight1.BackColor = Color.Yellow
                DColorLight1.Visible = True
            Case Else   '����
                DColorLight1.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorLight1", "ZZZZZZ")
        '������
        Select Case FindFieldInf("ColorLight2")
            Case 0  '���
                DColorLight2.BackColor = Color.LightGray
                DColorLight2.Visible = True

            Case 1  '�ק�+�ˬd
                DColorLight2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DColorLight2Rqd", "DColorLight2", "���`�G�ݿ�ĤG������")
                DColorLight2.Visible = True
            Case 2  '�ק�
                DColorLight2.BackColor = Color.Yellow
                DColorLight2.Visible = True
            Case Else   '����
                DColorLight1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ColorLight2", "ZZZZZZ")


        '������
        Select Case FindFieldInf("ColorLight3")
            Case 0  '���
                DColorLight3.BackColor = Color.LightGray
                DColorLight3.Visible = True

            Case 1  '�ק�+�ˬd
                DColorLight3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("Dcolorlight3Rqd", "Dcolorlight3", "���`�G�ݿ�ĤT������")
                DColorLight3.Visible = True
            Case 2  '�ק�
                DColorLight3.BackColor = Color.Yellow
                DColorLight3.Visible = True
            Case Else   '����
                DColorLight3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ColorLight3", "ZZZZZZ")



        'DeliveryDate
        Select Case FindFieldInf("DeliveryDate")
            Case 0  '���
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "lightgrey")
                DDeliveryDate.Attributes.Add("readonly", "true")
                BDate.Visible = False

            Case 1  '�ק�+�ˬd
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "greenyellow")
                DDeliveryDate.Attributes.Add("readonly", "true")
                ' ShowRequiredFieldValidator("DDeliveryDateRqd", "DDeliveryDate", "���`�G�ݿ�J�Ʊ���")
                BDate.Visible = True

            Case 2  '�ק�
                DDeliveryDate.Visible = True
                DDeliveryDate.Style.Add("background-color", "yellow")
                DDeliveryDate.Attributes.Add("readonly", "true")
                BDate.Visible = True
            Case Else   '����
                DDeliveryDate.Visible = False
                BDate.Visible = False

        End Select
        If pPost = "New" Then DDeliveryDate.Value = ""





        'YKK��O
        Select Case FindFieldInf("YKKColorType")
            Case 0  '���
                DYKKColorType.BackColor = Color.LightGray
                DYKKColorType.Visible = True

            Case 1  '�ק�+�ˬd
                DYKKColorType.BackColor = Color.GreenYellow
                '      ShowRequiredFieldValidator("DYKKColorTypeRqd", "DYKKColorType", "���`�G�ݿ�YKK��O")
                DYKKColorType.Visible = True
            Case 2  '�ק�
                DYKKColorType.BackColor = Color.Yellow
                DYKKColorType.Visible = True
            Case Else   '����
                DYKKColorType.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("YKKColorType", "ZZZZZZ")



        '�@�H��
        Select Case FindFieldInf("SLDColor")
            Case 0  '���
                DAgain.BackColor = Color.LightGray
                DAgain.Visible = True

            Case 1  '�ק�+�ˬd
                DAgain.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAgainRqd", "DAgain", "���`�G�ݿ�J�@�H��")
                DAgain.Visible = True
            Case 2  '�ק�
                DAgain.BackColor = Color.Yellow
                DAgain.Visible = True
            Case Else   '����
                DAgain.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorType", "ZZZZZZ")





        'YKKColorCode
        Select Case FindFieldInf("YKKColorCode")
            Case 0  '���
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "lightgrey")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = False

            Case 1  '�ק�+�ˬd
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "greenyellow")
                '  DYKKColorCode.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorCodeRqd", "DYKKColorCode", "���`�G�ݿ�JYKK�⸹")
                BYKKColorCode.Visible = True

            Case 2  '�ק�
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "yellow")
                DYKKColorCode.Attributes.Add("readonly", "true")
                BYKKColorCode.Visible = True
            Case Else   '����
                DYKKColorCode.Visible = False
                BYKKColorCode.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCode.Value = ""


        'YKKColorcodeVF
        Select Case FindFieldInf("YKKColorCodeVF")
            Case 0  '���
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "lightgrey")
                DYKKColorCodeVF.Attributes.Add("readonly", "true")
                BYKKColorCodeVF.Visible = False

            Case 1  '�ק�+�ˬd
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "greenyellow")
                '  DYKKColorcodeVF.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorcodeVFRqd", "DYKKColorcodeVF", "���`�G�ݿ�JYKK�⸹")
                BYKKColorCodeVF.Visible = True

            Case 2  '�ק�
                DYKKColorCodeVF.Visible = True
                DYKKColorCodeVF.Style.Add("background-color", "yellow")
                'DYKKColorCodeVF.Attributes.Add("readonly", "true")
                BYKKColorCodeVF.Visible = True
            Case Else   '����
                DYKKColorCodeVF.Visible = False
                BYKKColorCodeVF.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCodeVF.Value = ""


        'YKKColorcodeSLD
        Select Case FindFieldInf("YKKColorCodeSLD")
            Case 0  '���
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "lightgrey")
                DYKKColorCodeSLD.Attributes.Add("readonly", "true")
                BYKKColorCodeSLD.Visible = False

            Case 1  '�ק�+�ˬd
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "greenyellow")
                '  DYKKColorcodeSLD.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorcodeSLDRqd", "DYKKColorcodeSLD", "���`�G�ݿ�JYKK�⸹")
                BYKKColorCodeSLD.Visible = True

            Case 2  '�ק�
                DYKKColorCodeSLD.Visible = True
                DYKKColorCodeSLD.Style.Add("background-color", "yellow")
                'DYKKColorCodeSLD.Attributes.Add("readonly", "true")
                BYKKColorCodeSLD.Visible = True
            Case Else   '����
                DYKKColorCodeSLD.Visible = False
                BYKKColorCodeSLD.Visible = False

        End Select
        If pPost = "New" Then DYKKColorCodeSLD.Value = ""




        'PFBWire
        Select Case FindFieldInf("PFBWire")
            Case 0  '���
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "lightgrey")
                DPFBWire.Attributes.Add("readonly", "true")
                BYKKColor.Visible = False

            Case 1  '�ק�+�ˬd
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "greenyellow")
                DPFBWire.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DPFBWireRqd", "DPFBWire", "���`�G�ݿ�JPF�Q��U��ݥΦ�")
                BYKKColor.Visible = True

            Case 2  '�ק�
                DPFBWire.Visible = True
                DPFBWire.Style.Add("background-color", "yellow")
                DPFBWire.Attributes.Add("readonly", "true")
                BYKKColor.Visible = True
            Case Else   '����
                DPFBWire.Visible = False
                BYKKColor.Visible = False

        End Select
        If pPost = "New" Then DPFBWire.Value = ""


        'DPFOpenParts
        Select Case FindFieldInf("PFOpenParts")
            Case 0  '���
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "lightgrey")
                DPFOpenParts.Attributes.Add("readonly", "true")


            Case 1  '�ק�+�ˬd
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "greenyellow")
                DPFOpenParts.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DPFOpenPartsRqd", "DPFOpenParts", "���`�G�ݿ�JPF�}���")


            Case 2  '�ק�
                DPFOpenParts.Visible = True
                DPFOpenParts.Style.Add("background-color", "yellow")
                DPFOpenParts.Attributes.Add("readonly", "true")

            Case Else   '����
                DPFOpenParts.Visible = False


        End Select
        If pPost = "New" Then DPFOpenParts.Value = ""




        'YKK��O
        Select Case FindFieldInf("ColorSystem")
            Case 0  '���
                DColorSystem.BackColor = Color.LightGray
                DColorSystem.Visible = True

            Case 1  '�ק�+�ˬd
                DColorSystem.BackColor = Color.GreenYellow
                '     ShowRequiredFieldValidator("DColorSystemRqd", "DColorSystem", "���`�G�ݿ��t")
                DColorSystem.Visible = True
            Case 2  '�ק�
                DColorSystem.BackColor = Color.Yellow
                DColorSystem.Visible = True
            Case Else   '����
                DColorSystem.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("ColorSystem", "ZZZZZZ")




        '�s�¦�
        Select Case FindFieldInf("NewOldColor")
            Case 0  '���
                DNewOldColor.BackColor = Color.LightGray
                DNewOldColor.Visible = True

            Case 1  '�ק�+�ˬd
                DNewOldColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNewOldColorRqd", "DNewOldColor", "���`�G�ݿ�s�¦�")
                DNewOldColor.Visible = True
            Case 2  '�ק�
                DNewOldColor.BackColor = Color.Yellow
                DNewOldColor.Visible = True
            Case Else   '����
                DNewOldColor.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("NewOldColor", "ZZZZZZ")


        'ReRegister
        Select Case FindFieldInf("ReRegister")
            Case 0  '���
                DReRegister.BackColor = Color.LightGray
                DReRegister.Visible = True
            Case 1  '�ק�+�ˬd
                DReRegister.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReRegisterRqd", "DReRegister", "���`�G�ݿ�J�~��ӽ�")
                DReRegister.Visible = True
            Case 2  '�ק�
                DReRegister.BackColor = Color.Yellow
                DReRegister.Visible = True
            Case Else   '����
                DReRegister.Visible = False
                DReRegister.Checked = False
        End Select


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
        DDepoName.Text = DBUser.Rows(0).Item("Divname")
        DName.Text = DBUser.Rows(0).Item("Username")

        '����
        If pFieldName = "ColorLight1" Then
            DColorLight1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight1.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight1.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight1.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '����
        If pFieldName = "ColorLight2" Then
            DColorLight2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight2.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorLight2' "
                SQL = SQL & " order by createtime desc "



                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight2.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight2.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '����
        If pFieldName = "ColorLight3" Then
            DColorLight3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DColorLight3.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  data from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'colorlight2' "
                SQL = SQL & " order by createtime desc "



                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DColorLight3.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DColorLight3.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK��O
        If pFieldName = "YKKColorType" Then
            DYKKColorType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DYKKColorType.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'YKKColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DYKKColorType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DYKKColorType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK��O
        If pFieldName = "ColorType" Then
            DAgain.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAgain.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'ColorType'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DAgain.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAgain.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        'YKK��O
        If pFieldName = "CustomerSample" Then

            DCustomerSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCustomerSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'CustomerSample'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DCustomerSample.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCustomerSample.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '�}�o�~
        If pFieldName = "DevYear" Then

            DDevYear.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevYear.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevYear' order by data"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevYear.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevYear.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If


        '�}�o�u
        If pFieldName = "DevSeason" Then

            DDevSeason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDevSeason.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DevSeason'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDevSeason.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDevSeason.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
            End If
        End If

        '�s���¦�
        If pFieldName = "NewOldColor" Then
            DNewOldColor.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DNewOldColor.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'NewOldColor'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DNewOldColor.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DNewOldColor.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
        If Request.Cookies("RunBSAVE").Value = True And OK() = True Then

            DisabledButton()   '����Button�B�@
            FlowControl("OK", 3, "1")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True And OK() = True Then

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

            If wStep = 10 Then
                If OK() = True Then
                    DisabledButton()   '����Button�B�@
                    FlowControl("NG2", 2, "3")

                End If
            Else
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
                ErrCode = oCommon.CommissionNo("005011", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo,wUserID, wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, wUserID, wApplyID)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                            'Modify-End
                            '�]�w�e�UNo
                            DNo.Text = SetNo(NewFormSno)

                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '���O�q�������а���
                            '��s������
                            'jessica 2021/7/21 step 65 �hDYE
                            If pAction = 3 Then
                                ModifyTranData(pFun, "4")
                            Else
                                ModifyTranData(pFun, pSts)
                            End If



                            '�y�{��Ƶ���
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,wUserID, pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, wUserID, pRunNextStep)
                            '��渹�X,���y����,�u�{���d���X,��ƾ�,ñ�֪�, �y�{�����_(�|ñ)
                            'Modify-End

                            'If wStep = 541 Or wStep = 546 Then
                            'pRunNextStep = 1
                            'End If

                            '�ˬd�O�_�������~�e��60
                            Dim SQL1 As String
                            If wStep = 51 Then
                                SQL1 = " Select step From T_WaitHandle Where "
                                SQL1 = SQL1 + "    FormNo  = '005011'   And FormSno = '" + CStr(wFormSno) + "'   And Step = '52'   and  modifyuser <> 'Operation_02'  "
                                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                                If dtFlow.Rows.Count > 0 Then
                                    pRunNextStep = 1
                                Else
                                    pRunNextStep = 0
                                End If
                            End If

                            If wStep = 52 Then
                                SQL1 = " Select step From T_WaitHandle Where  "
                                SQL1 = SQL1 + "    FormNo  = '005011'   And FormSno = '" + CStr(wFormSno) + "'   And Step = '51'  and  modifyuser <> 'Operation_02'  "
                                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                                If dtFlow.Rows.Count > 0 Then
                                    pRunNextStep = 1
                                Else
                                    pRunNextStep = 0
                                End If
                            End If

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '�O�q�������а���
                    End If
                    End If
                End If

                '���o�U�@��
                If ErrCode = 0 And pRunNextStep = 1 Then

                    Dim wAllocateID As String = ""
                    '�u���å����~CC���ɻT�A�Y���O�N�������_���
                    'If wStep = 40 And pAction = 0 And DYKKColorType.SelectedValue <> "�å�" Then
                    ' pAction = 3
                    'End If
                    ' 20150408 �ק�40�u�{�p�G���O�¦�N����70�u�{

                    If (wStep = 30 Or wStep = 530) And pFun = "OK" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 2
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 0
                        End If
                    End If


                    If wStep = 60 And pFun = "NG1" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 1
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 3
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 1
                        End If
                    End If


                    If (wStep = 48 Or wStep = 548) And pFun = "OK" Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 2
                        ElseIf DBuyerCode.Text = "000021" Then 'KIPLING
                            pAction = 0
                        End If
                    End If


                    If wStep = 70 Then
                        Dim SQL1 As String
                        Dim DTNO As Integer
                        SQL1 = " select * from F_NewcolorUAKIPLINGDT"
                        SQL1 = SQL1 + " where dtsheet <> '���Y' and OFORMSNO = '" + CStr(wFormSno) + "'  AND STS <>'2' "
                        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL1)
                        If dtFlow.Rows.Count = 0 Then
                            DTNO = 0  '�P�_�O�_���֥i��
                        Else
                            DTNO = 1
                        End If

                        If DNewOldColor.SelectedValue = "�s��" And DTNO = 0 Then
                            pAction = 0
                        Else
                            If DYKKColorType.SelectedValue = "�å�" Then
                                pAction = 1  '151
                            Else
                                pAction = 2
                            End If
                        End If

                    End If

                    If wStep = 75 And pFun = "OK" Then
                        If DYKKColorType.SelectedValue = "�å�" Then
                            pAction = 0
                        Else
                            pAction = 1
                        End If
                    End If


                    If wStep = 155 Then
                        If DBuyerCode.Text = "TW0371" Then 'UA
                            pAction = 0
                        ElseIf DBuyerCode.Text = "TW0011" Then 'KIPLING
                            pAction = 1
                        ElseIf DBuyerCode.Text = "000021" Then 'TNF
                            pAction = 1
                        End If
                    End If

                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, wUserID, wApplyID, wAgentID, wAllocateID, MultiJob, _
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
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo,wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
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
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, wUserID, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '��渹�X,���y����,�u�{���d���X,�Ǹ�,�ӽЪ�ID,��ƾ�,�ظm��, ñ�֪�, �Q�N�z��, �ӽЪ�, �w�w�}�l��, �w�w������, ���n��
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo,wUserID, wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, wUserID, wApplyID)
                        '��渹�X,���y����,�u�{���d���X,�Ǹ�,��ƾ�,ñ�֪�, �ӽЪ�
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '��u�{��{�վ�
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010��ƾ����)
                        'RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(wUserID, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
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
                            oCommon.Send(wUserID, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '�e���, �����, �ӽЪ�, ��渹�X, ���y����, �u�{���d���X,�T�����O
                        Next i
                    Else
                        oCommon.Send(wUserID, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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


                    LastStep = wStep  '�O���W�@�Ӥu�{
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

                Dim URL As String

                ' MODIFY-START BY JOY 2015/4/7 (�}�z��)
                ' If (LastStep = 20 Or LastStep = 520 Or LastStep = 25) And pFun = "OK" Then '�s��T�{�������OK�~�i��
                '�W�[�s��_�� Modify 20150419
                If wbFormSno = 0 And wbStep < 3 And DReRegister.Checked = True Then    '�P�_�O�_�_��

                    uJavaScript.PopMsg(Me, "�w���\�e�X�z�Ҷ�g���ӽг�A���~��ӽСI")
                    EnabledButton()                             '�_��Button�B�@


                    DNo.Text = ""
                    wFormSno = wbFormSno
                    wStep = wbStep


                Else


                    If ((wbStep = 20 Or wbStep = 520) And pFun = "OK") Or ((wbStep = 30 Or wbStep = 530) And pFun = "OK") Then '�s��T�{�������OK�~�i��
                        Dim FormName As String
                        If LastStep = 20 Or LastStep = 520 Then
                            FormName = "3CF DTM"
                        Else
                            FormName = "5VS  DTM"
                        End If

                        'MODIFY-END
                        Dim SQL As String  '�ˬd�O�_���ۦP�������s��̿৹����
                        SQL = " select  *  from F_NewColorUAKIPLING a, F_NewColorComplete b"
                        SQL = SQL & " where a.formno = b.formno And a.formsno = b.formsno"
                        SQL = SQL & " and a.version = b.version "
                        SQL = SQL & " and formname ='" & FormName & "'"
                        SQL = SQL & " and a.No = '" & DNo.Text + "'"
                        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
                        If dtFlow.Rows.Count = 0 Then '�p�G�S���~�i�H�s�W
                            URL = uCommon.GetAppSetting("NewColorCompleteUrlUA") & "?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & LastStep & "&pNextGate=" & wNextGate & _
                                                                                                       "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                        Else
                            URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                  "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                        End If


                    Else

                        URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                   "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
                    End If
                    Response.Redirect(URL)
                End If '

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

        Dim SQl As String
        SQl = "Insert into F_NewColorUAKIPLING "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, DepoName,Date,Name,CustomerCheck,"  '1~5                
        SQl = SQl + "FactoryCheck,VCACheck,Customer,CustomerCode,Buyer,"   '6~10              
        SQl = SQl + "BuyerCode,CustomerColor,CustomerColorCode,OverseaYKKCode,PANTONECode,"  '11~15
        SQl = SQl + "DevYear,DevSeason,ReceiveDate,ColorLight1,ColorLight2,ColorLight3,NewOldColor," ' 16~20
        SQl = SQl + "DeliveryDate,DuplicateNo,NOCCS,NOCCSReason,CustomerNGColor," '21~25
        SQl = SQl + "Remark,CustomerSample,ColorSystem,PFBWire,PFOpenParts," '26~30
        SQl = SQl + "YKKColorType,YKKColorCode,YKKColorCodeVF,YKKColorCodeSLD,Version,CheckNo,SLDColor,Again," '31~35
        SQl = SQl + "VFCOLOR,VFCOLOR1," '36
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '005011', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        SQl = SQl + " N'" + DDepoName.Text + "', "                '����2
        SQl = SQl + " N'" + DDate.Text + "', "                '���3
        SQl = SQl + " N'" + DName.Text + "', "                '�m�W4
        If DCustomerCheck.Checked = True Then
            SQl = SQl + " '1', "  '5
        Else
            SQl = SQl + " '0', "  '5:
        End If

        If DFactoryCheck.Checked = True Then
            SQl = SQl + " '1', "  '6
        Else
            SQl = SQl + " '0', "  '6
        End If


        If DVCACheck.Checked = True Then
            SQl = SQl + " '1', "  '7
        Else
            SQl = SQl + " '0', "  '7
        End If


        SQl = SQl + " N'" + DCustomer.Text.ToUpper + "', "                '�Ȥ�8
        SQl = SQl + " N'" + DCustomerCode.Text.ToUpper + "', "                '�Ȥ�9
        SQl = SQl + " N'" + DBuyer.Text.ToUpper + "', "                'BUYER10
        SQl = SQl + " N'" + DBuyerCode.Text.ToUpper + "', "                'BUYER11
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerColor.Text.ToUpper) + "', "                ' �Ȥ��W12
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerColorCode.Text.ToUpper) + "', "                ' �Ȥ�⸹13
        SQl = SQl + " N'" + YKK.ReplaceString(DOverseaYKKCode.Text.ToUpper) + "', "                ' ���~YKK�⸹14
        SQl = SQl + " N'" + YKK.ReplaceString(DPANTONECode.Text.ToUpper) + "', "                ' PANTONECode�⸹15
        SQl = SQl + " N'" + DDevYear.SelectedValue + "', "                ' �}�o�~16
        SQl = SQl + " N'" + DDevSeason.SelectedValue + "', "                ' �}�o�u17
        SQl = SQl + " N'" + DReceiveDate.Text + "', "                ' �������18
        SQl = SQl + " N'" + DColorLight1.SelectedValue.ToUpper + "', "                ' �Ĥ@������19
        SQl = SQl + " N'" + DColorLight2.SelectedValue.ToUpper + "', "                ' �ĤG������20
        SQl = SQl + " N'" + DColorLight3.SelectedValue.ToUpper + "', "                ' �ĤG������20
        SQl = SQl + " N'" + DNewOldColor.SelectedValue.ToUpper + "', "                ' �ĤG������20
        SQl = SQl + " N'" + DDeliveryDate.Value + "', "                ' ���Ш̿઺�s��22
        SQl = SQl + " N'" + DDuplicateNo.Text.ToUpper + "', "                ' �������21
        If DNOCCS.Checked = True Then
            SQl = SQl + " '1', " '23
        Else
            SQl = SQl + " '0', " '23
        End If
        SQl = SQl + " N'" + YKK.ReplaceString(DNOCCSReason.Text) + "', "                ' �L�kCCS��] 24
        SQl = SQl + " N'" + YKK.ReplaceString(DCustomerNGColor.Text.ToUpper) + "', "                ' �Ȥ�_�M���⸹25
        SQl = SQl + " N'" + YKK.ReplaceString(DRemark.Text) + "', "                ' �Ƶ�26
        SQl = SQl + " N'" + DCustomerSample.SelectedValue.ToUpper + "', "                ' �Ȥ�˫~27
        SQl = SQl + " N'" + DColorSystem.Text.ToUpper + "', "                ' ��t29
        SQl = SQl + " N'" + DPFBWire.Value.ToUpper + "', "                ' BWire29
        SQl = SQl + " N'" + DPFOpenParts.Value.ToUpper + "', "                ' �}���30
        SQl = SQl + " N'" + DYKKColorType.SelectedValue.ToUpper + "', "                ' YKK��O31
        SQl = SQl + " N'" + DYKKColorCode.Value.ToUpper + "', "                ' YKK�⸹32
        SQl = SQl + " N'" + DYKKColorCodeVF.Value.ToUpper + "', "                ' YKK�⸹32
        SQl = SQl + " N'" + DYKKColorCodeSLD.Value.ToUpper + "', "                ' YKK�⸹32

        If DVersion.Text = "" Then
            SQl = SQl + " 0, "                ' ����33
        Else
            SQl = SQl + " N'" + DVersion.Text + "', "                ' ����33
        End If

        SQl = SQl + " N'" + DCheckNo.Text.ToUpper + "', "                ' �T�{�֥i�N��34
        SQl = SQl + " N'" + DSLDColor.Text.ToUpper + "', "                ' ���Y���⸹35

        If DAgain.SelectedValue = "�H��" Then
            SQl = SQl + "1, "                ' �H��
        ElseIf DAgain.SelectedValue = "�@��" Then
            SQl = SQl + "2, "                ' �@��
        Else
            SQl = SQl + "0, "
        End If

        SQl = SQl + " N'" + DVFColor.Text.ToUpper + "', "                ' VF�⸹36
        SQl = SQl + " N'" + DVFColor1.Text.ToUpper + "', "                ' VF�⸹36
        SQl = SQl + " '" + wUserID + "', "     '�@����
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
        Dim SQL As String = ""

        SQL = " Update F_NewColorUAKIPLING"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = N'" & DDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " Name = N'" & DName.Text & "',"
        If DCustomerCheck.Checked = True Then
            SQL = SQL + " CustomerCheck =1,"
        Else
            SQL = SQL + " CustomerCheck =0,"
        End If
        If DFactoryCheck.Checked = True Then
            SQL = SQL + " FactoryCheck =1,"
        Else
            SQL = SQL + " FactoryCheck =0,"
        End If
        If DVCACheck.Checked = True Then
            SQL = SQL + " VCACheck =1,"
        Else
            SQL = SQL + " VCACheck =0,"
        End If
        SQL = SQL + " Customer = N'" & DCustomer.Text.ToUpper & "',"
        SQL = SQL + " CustomerCode = N'" & DCustomerCode.Text.ToUpper & "',"
        SQL = SQL + " Buyer = N'" & DBuyer.Text & "',"
        SQL = SQL + " BuyerCode = N'" & DBuyerCode.Text & "',"
        SQL = SQL + " CustomerColor = N'" & YKK.ReplaceString(DCustomerColor.Text.ToUpper) & "',"
        SQL = SQL + " CustomerColorCode = N'" & YKK.ReplaceString(DCustomerColorCode.Text.ToUpper) & "',"
        SQL = SQL + " OverseaYKKCode = N'" & YKK.ReplaceString(DOverseaYKKCode.Text.ToUpper) & "',"
        SQL = SQL + " PANTONECode = N'" & YKK.ReplaceString(DPANTONECode.Text.ToUpper) & "',"
        SQL = SQL + " DevYear = N'" & DDevYear.SelectedValue & "',"
        SQL = SQL + " DevSeason = N'" & DDevSeason.SelectedValue & "',"
        If wStep = 10 Then
            SQL = SQL + " ReceiveDate = N'" & CDate(NowDateTime).ToString("yyyy/MM/dd") & "',"
        Else
            SQL = SQL + " ReceiveDate = N'" & DReceiveDate.Text & "',"
        End If

        SQL = SQL + " ColorLight1 = N'" & DColorLight1.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorLight2 = N'" & DColorLight2.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorLight3 = N'" & DColorLight3.SelectedValue.ToUpper & "',"
        SQL = SQL + " NewOldColor = N'" & DNewOldColor.SelectedValue.ToUpper & "',"
        SQL = SQL + " DeliveryDate = N'" & DDeliveryDate.Value & "',"
        SQL = SQL + " DuplicateNo = N'" & DDuplicateNo.Text.ToUpper & "',"

        If DNOCCS.Checked = True Then
            SQL = SQL + " NOCCS =1,"
        Else
            SQL = SQL + " NOCCS =0,"
        End If


        SQL = SQL + " NOCCSReason = N'" & YKK.ReplaceString(DNOCCSReason.Text) & "',"
        SQL = SQL + " CustomerNGColor = N'" & YKK.ReplaceString(DCustomerNGColor.Text.ToUpper) & "',"
        SQL = SQL + " Remark = N'" & YKK.ReplaceString(DRemark.Text) & "',"
        SQL = SQL + " CustomerSample = N'" & DCustomerSample.SelectedValue.ToUpper & "',"
        SQL = SQL + " ColorSystem = N'" & DColorSystem.Text.ToUpper & "',"
        SQL = SQL + " PFBWire = N'" & DPFBWire.Value.ToUpper & "',"
        SQL = SQL + " PFOpenParts = N'" & DPFOpenParts.Value.ToUpper & "',"
        SQL = SQL + " YKKColorType = N'" & DYKKColorType.SelectedValue.ToUpper & "',"
        SQL = SQL + " YKKColorCode = N'" & DYKKColorCode.Value.ToUpper & "',"
        SQL = SQL + " YKKColorCodeVF = N'" & DYKKColorCodeVF.Value.ToUpper & "',"
        SQL = SQL + " YKKColorCodeSLD = N'" & DYKKColorCodeSLD.Value.ToUpper & "',"

        ' If wStep = 20 Then
        ' SQL = SQL + " Version =1,"
        ' Else
        SQL = SQL + " Version ='" & DVersion.Text.ToUpper & "',"
        ' End If
        SQL = SQL + " CheckNo = N'" & DCheckNo.Text.ToUpper & "',"

        If DSLDColor.Text.ToUpper <> "" Then
            SQL = SQL + " SLDColor = N'" & DSLDColor.Text.ToUpper & "',"
        End If


        If DVFColor.Text.ToUpper <> "" Then
            SQL = SQL + " VFColor = N'" & DVFColor.Text.ToUpper & "',"
        End If


        If DAgain.SelectedValue <> "" Then
            If DAgain.SelectedValue = "�H��" Then
                SQL = SQL + " again =1,"
            ElseIf DAgain.SelectedValue = "�@��" Then
                SQL = SQL + " again =2,"
            End If

        End If


      

        If DVFColor1.Text.ToUpper <> "" Then
            SQL = SQL + " VFColor1 = N'" & DVFColor1.Text.ToUpper & "',"
        End If


        SQL = SQL + " ModifyUser = '" & wUserID & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL)
        Dim Code As String = ""

        If wStep = 90 And DNewOldColor.SelectedValue = "�s��" Then 'Master ���ɫ����ƨ�WINS COLOR STRUCTURE

            If DSLDColor.Text <> "" And DAgain.SelectedValue <> "" Then
                ' INSERT �@�H��
                Dim YB1CP2 As String = ""
                If DAgain.SelectedValue = "�H��" Then
                    YB1CP2 = "0"
                ElseIf DAgain.SelectedValue = "�@��" Then
                    YB1CP2 = "1"
                End If

                '���T�wcolorcode �O�_�s�b, ���s�b�~INSERT 
                Dim SLDColor As String

                If Len(DSLDColor.Text) = 3 Then  '�T�X�[2��ť�
                    SLDColor = "  " + DSLDColor.Text
                Else
                    SLDColor = DSLDColor.Text
                End If

                SQL = "Select * From R_NCTP200WK  "
                SQL = SQL & " Where CLRCP2 = '" + SLDColor + "'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter1.Rows.Count = 0 Then
                    SQL = " INSERT INTO R_NCTP200WK (NO,STS,CLRCP2,YB1CP2,UIDCP2,PRGCP2,DEVCP2,RADUP2,RADTP2,RUPUP2,RUPTP2,RCMCP2)"
                    SQL = SQL + "VALUES('" + DNo.Text + "','0','" + SLDColor + "','" + YB1CP2 + "','RPA','UPLOAD','WINGS',convert(char(10),getdate(),112) ,"
                    SQL = SQL + " substring(convert(char(10),getdate(),108),1,2)+substring(convert(char(10),getdate(),108),4,2)+substring(convert(char(10),getdate(),108),7,2) ,"
                    SQL = SQL + " '0','0','000034')"
                    uDataBase.ExecuteNonQuery(SQL)

                    Code = oWaves.NewColorType2Wings(DNo.Text)
                    If Code <> "0" Then
                        uJavaScript.PopMsg(Me, "�פJWINS���`�A�Ь��q����")

                    End If
                End If


            End If



            SQL = "insert into R_NCFA300WK (no,sts,PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,CCLCW3,UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3)"
            SQL = SQL + " select no,STS,PDPCW3,CTBNW3,PACCW3,ST1CW3,ST6CW3,SCSVW3,"
            SQL = SQL + "  CASE WHEN CCLCW3 ='V7' THEN '  V7+' ELSE RIGHT(REPLICATE(' ', 5) + CAST(CCLCW3 as NVARCHAR), 5) END AS CCLCW3,UIDCW3,PRGCW3,DEVCW3,RADUW3,RADTW3,RUPUW3,RUPTW3,RCMCW3, PRBFW3, CHKFW3,WKDTW3 from ("
            SQL = SQL + " select NO,'0' AS STS,'01' as PDPCW3 ,CTBNW3,ykkcolorcode as PACCW3,ST1CW3,ST6CW3,SCSVW3,"
            SQL = SQL + " Case when CCLCW3 ='YKKColorCode' then YKKColorCode "
            SQL = SQL + " when CCLCW3 ='VFColor' then VFColor"
            SQL = SQL + " when CCLCW3 ='PFBWire' then PFBWire"
            SQL = SQL + " when CCLCW3 ='PFOpenParts' then PFOpenParts"
            SQL = SQL + " when CCLCW3 ='SLDColor' then SLDColor "
            SQL = SQL + " when CCLCW3 ='YKKColorCodeSLD' then YKKColorCodeSLD"
            SQL = SQL + " when CCLCW3 ='YKKColorCodeVF' then YKKColorCodeVF"
            SQL = SQL + " when CCLCW3 ='TSCC' then CCLCW3"
            SQL = SQL + " when CCLCW3 ='C5+' then CCLCW3"
            SQL = SQL + " end as CCLCW3 ,'' as UIDCW3,'' as PRGCW3,'' as DEVCW3,convert(char(10),getdate(),112)as  RADUW3,"
            SQL = SQL + " substring(convert(char(10),getdate(),108),1,2)+substring(convert(char(10),getdate(),108),4,2)+substring(convert(char(10),getdate(),108),7,2) as RADTW3,"
            SQL = SQL + " '' as RUPUW3,'' as RUPTW3,'' as RCMCW3,'' as PRBFW3,'' as CHKFW3,'' as WKDTW3"
            SQL = SQL + " from v_newcolor a,R_NewColorST_01 b"
            SQL = SQL + " where sts in(1,0) and no = '" + DNo.Text + "' AND TYPE ='N2'"
            SQL = SQL + " )a"
            uDataBase.ExecuteNonQuery(SQL)


            Code = oWaves.NewColor2Wings(DNo.Text)

            If Code <> "0" Then
                uJavaScript.PopMsg(Me, "�פJWINS���`�A�Ь��q����")
            End If
        End If




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
            'jessica 2021/7/21 step 65 �hDYE
            If pSts = "4" Then SQl = SQl + " StsDes = '" & BSAVE.Value & "',"
            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
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
            SQl = SQl + " ModifyUser = '" & wUserID & "',"
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
                SQl = SQl + " '" + wUserID + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " CreateUser = '" & wUserID & "',"
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



        If DDescSheet.Visible Then                                      ' ����
            DDescSheet.Style("top") = Top - 130 & "px"
            DDecideDesc.Style("top") = Top - 130 + 6 & "px"
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' ���𻡩�
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top + 20 & "px"
        BNG1.Style("top") = Top + 20 & "px"
        BNG2.Style("top") = Top + 20 & "px"
        BOK.Style("top") = Top + 20 & "px"


        Top = Top + 50
        If GridView2.Rows.Count > 0 Then                                ' �֩w�i��
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
        End If
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
        Dim SQL As String = ""

        'Set�����
        Str2 = CStr(DateTime.Now.Month)  '��
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '��
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = "J" + CStr(DateTime.Now.Year) + Str2
        '�~
        'Set�渹
        '�������Ʀ��X��  20150414 Modify by Jessica
        SQL = " select   isnull(max(convert(int,substring(no,8,4))),0) cun from  F_NewColorUAKIPLING"
        SQL = SQL + " where  left(convert(char(10),date ,111),7) = left(convert(char(10), getdate(),111),7)"
        Dim dt1 As DataTable = uDataBase.GetDataTable(SQL)
        If dt1.Rows.Count > 0 Then
            Str1 = CStr(CInt(dt1.Rows(0).Item("cun")) + 1)
        End If

        ' Str1 = CStr(Seq)

        For i = Str1.Length To 4 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function OK() As Boolean
        Top = 1200
        SetControlPosition()
        '  ShowFormData()
        Dim SQL As String
        Dim DOUBLENo As Integer = 0

        Dim Errcode As Integer = 0


        If wStep = 1 Then
            If DCustomerCheck.Checked = False And DFactoryCheck.Checked = False And DVCACheck.Checked = False Then
                isOK = False
                Message = "���`�G�ФĿ�Ȥ�ۮ֩Τu�t�ۮ֩�VCA!"

            End If

            If DVCACheck.Checked = False Then
                If DCustomerNGColor.Text = "" Then
                    DCustomerNGColor.BackColor = Color.GreenYellow
                    DCustomerNGColor.Visible = True
                    isOK = False
                    Message = "���`�G�ݿ�J�Ȥ�_�M��YKK�⸹!"

                End If
            End If
        End If



        If wStep = 20 Or wStep = 30 Or wStep = 530 Or wStep = 520 Then
            If DColorSystem.Text = "" Then
                isOK = False
                Message = "���`�G�п�J��t!"
            End If

            If DPFBWire.Value = "" Then
                isOK = False
                Message = Message + "\n" + "���`�G�п�JPF�U��Q��ݥΦ�!"
            End If
            If DPFOpenParts.Value = "" Then
                isOK = False
                Message = Message + "\n" + "���`�G�п�JPF�}���!"
            End If

        End If

        '  If wStep = 60 Then
        ' If DCheckNo.Text = "" Then
        'isOK = False
        'Message = "���`�G�п�J�T�{�֥i�N��!"
        'End If

        'End If

        If wStep = 500 Or wStep = 1 Then  '(NG�A�e�X�Ʊ����ݭ��s���)
            If DDeliveryDate.Value = "" Then
                isOK = False
                Message = "���`�G�п�J�Ʊ���!"
            End If
        End If

        If DBuyerCode.Text <> "TW0371" And DBuyerCode.Text <> "TW0011" And DBuyerCode.Text <> "000021" Then
            isOK = False
            Message = "���`�GBUYER �п�JUA �� KIPPLING!"
        End If



        If wStep = 1 Then

            '�ˬd�O�_�����Ш̿઺�s��
            ' Dim SQL As String
            SQL = "select no,customer,customercode,buyer,buyercode,customercolor,customercolorcode,overseaykkcode,pantonecode devyear,devseason "
            SQL = SQL + " from  v_newcolor"
            SQL = SQL + " where sts <>2 and customercode='" + DCustomerCode.Text + "'"
            SQL = SQL + " and  buyercode='" + DBuyerCode.Text + "'"
            SQL = SQL + " and customercolor='" + YKK.ReplaceString(DCustomerColor.Text) + "'"
            SQL = SQL + " and  customercolorcode='" + YKK.ReplaceString(DCustomerColorCode.Text) + "'"
            SQL = SQL + " and overseaykkcode='" + YKK.ReplaceString(DOverseaYKKCode.Text) + "'"
            SQL = SQL + " and pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
            SQL = SQL + " and  devyear='" + DDevYear.SelectedValue + "'"
            SQL = SQL + " and devseason='" + DDevSeason.SelectedValue + "'"
            SQL = SQL + " and formno = '" + wFormNo + "'"


            Dim NoStr As String = ""
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                For Each dtr As Data.DataRow In DBAdapter1.Rows
                    If NoStr = "" Then
                        NoStr = dtr("No")
                    Else
                        NoStr = NoStr + "," + dtr("No")
                    End If

                Next
                If NoStr = DDuplicateNo.Text Then  '�p�G���Ъ��s���@�˴N���ΦA����@��
                    DOUBLENo = DOUBLENo + 1
                End If
            Else
                DDuplicateNo.Text = "" '�p�G�S�����дN�M��

            End If


            '20180612 ���Ш���
            DDuplicateNo.Text = NoStr
            If DDuplicateNo.Text <> "" Then

                Message = Message + "\n" + "�����Ш̿઺�s���нT�{��A�~�����?!"
                isOK = False
            End If

            'If NoStr <> "" And DOUBLENo = 0 Then

            '    DDuplicateNo.Text = NoStr

            '    Message = Message + "\n" + "�����Ш̿઺�s���нT�{��A�~�����?!"

            '    isOK = False


            'End If

        End If

        '�ˬd�O�_�����Ш�PANTONE�⸹ 20150521
        ' Dim SQL As String
        If wStep = 1 Or wStep = 500 Then
            If YKK.ReplaceString(DPANTONECode.Text) <> "" Then

                SQL = "select no "
                SQL = SQL + " from V_newColor"
                SQL = SQL + " where sts =0 and  pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
                If wStep = 500 Then
                    SQL = SQL + " and no <> '" + DNo.Text + "'"
                End If





                Dim NoStr As String = ""
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter1.Rows.Count > 0 Then
                    For Each dtr As Data.DataRow In DBAdapter1.Rows
                        If NoStr = "" Then
                            NoStr = dtr("No")
                        Else
                            NoStr = NoStr + "," + dtr("No")
                        End If


                    Next

                Else
                    DDuplicateNo.Text = "" '�p�G�S�����дN�M��

                End If

                If NoStr <> "" Then

                    DDuplicateNo.Text = NoStr
                    Message = Message + "\n" + "�����Ш̿઺PANTONE�⸹�A���i����?!"
                    isOK = False


                End If
            End If
        End If





        'Dim Q As Integer = 0
        ''jessica 20150326
        'If wStep = 70 Then
        '    '�ˬd�O�_��Q 
        '    Q = InStr(1, DYKKColorCode.Value, "Q", 1)
        '    If DYKKColorType.SelectedValue = "�å�" Then


        '        SQL = " select * from   m_referp"
        '        SQL = SQL + " where dkey  = 'LightCode'"
        '        SQL = SQL + " and cat = 5001"
        '        SQL = SQL + " and data  ='" + DYKKColorCode.Value + "'"

        '        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        '        If DBAdapter1.Rows.Count > 0 Then

        '        Else
        '            If Q = 0 Then
        '                Message = Message + "\n" + "�u��Q�r�˿ﶵ�~���å�!"
        '                DYKKColorCode.Value = ""
        '                isOK = False
        '            End If
        '        End If
        '    Else
        '        If Q > 0 Then
        '            Message = Message + "\n" + "��Q�r�˿ﶵ�u���å�!"
        '            DYKKColorCode.Value = ""
        '            isOK = False
        '        End If
        '    End If


        'End If






        If Not isOK Then

            uJavaScript.PopMsg(Me, Message)



        End If




        Return isOK


    End Function

    Sub CheckDuplicateNo()

        '�ˬd�O�_�����Ш̿઺�s��
        Dim SQL As String
        SQL = "select no,customer,customercode,buyer,buyercode,customercolor,customercolorcode,overseaykkcode,pantonecode devyear,devseason "
        SQL = SQL + " from F_NewColorUAKIPLING"
        SQL = SQL + " where  customercode='" + DCustomerCode.Text + "'"
        SQL = SQL + " and  buyercode='" + DBuyerCode.Text + "'"
        SQL = SQL + " and customercolor='" + YKK.ReplaceString(DCustomerColor.Text) + "'"
        SQL = SQL + " and  customercolorcode='" + YKK.ReplaceString(DCustomerColorCode.Text) + "'"
        SQL = SQL + " and overseaykkcode='" + YKK.ReplaceString(DOverseaYKKCode.Text) + "'"
        SQL = SQL + " and pantonecode='" + YKK.ReplaceString(DPANTONECode.Text) + "'"
        SQL = SQL + " and  devyear='" + DDevYear.SelectedValue + "'"
        SQL = SQL + " and devseason='" + DDevSeason.SelectedValue + "'"

        Dim NoStr As String = ""
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            For Each dtr As Data.DataRow In DBAdapter1.Rows
                If NoStr = "" Then
                    NoStr = dtr("No")
                Else
                    NoStr = NoStr + "," + dtr("No")
                End If

            Next
        End If

        If NoStr <> "" Then
            DDuplicateNo.Text = NoStr
            uJavaScript.PopMsg(Me, "�����Ш̿઺�s��!")
            ' If MsgBox("�����Ш̿઺�s���O�_�~�����?", MsgBoxStyle.OkCancel + MsgBoxStyle.MsgBoxSetForeground, "�ˬd") <> MsgBoxResult.Ok Then
            '������
            'Message = "�����Ш̿઺�s��"
            isOK = False

            'End If
            'MsgBox("�����Ш̿઺�s���O�_�~�����?")

        End If





    End Sub

    Sub MsgBox(ByVal text As String)
        'Dim scriptstr As String
        'scriptstr = "<script language=javascript>" + Chr(10) _
        '+ "confirm(""" + text + """)" + Chr(10) _
        '+ "</script>"
        'Response.Write(scriptstr)
        'Response.Write("<script language='javascript'>if(confirm(""" + text + """)==true)function1();else function2();</script>")
        'Response.Write("<script language=javascript>if (confirm('�A�T�w�n�~���X��r�H')==false) {return fale;};</script>")
    End Sub


    Protected Sub DCustomerCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCustomerCheck.CheckedChanged
        If DCustomerCheck.Checked Then
            DFactoryCheck.Checked = False
            DVCACheck.Checked = False
            CheckNGColor()


        End If
    End Sub

    Protected Sub DNOCCS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNOCCS.CheckedChanged
        CheckDnoCCS()

    End Sub

    Protected Sub DFactoryCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFactoryCheck.CheckedChanged
        If DFactoryCheck.Checked Then
            DCustomerCheck.Checked = False
            DVCACheck.Checked = False
            CheckNGColor()

        End If
    End Sub

    Protected Sub DVCACheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DVCACheck.CheckedChanged
        If DVCACheck.Checked Then
            DCustomerCheck.Checked = False
            DFactoryCheck.Checked = False
            CheckNGColor()


        End If
    End Sub

    Sub CheckDnoCCS()
        If DNOCCS.Checked = True Then
            DNOCCSReason.BackColor = Color.GreenYellow
            ShowRequiredFieldValidator("DNOCCSReasonRqd", "DNOCCSReason", "���`�G�ݿ�J�L�kCCS��]")
            DNOCCSReason.ReadOnly = False
        Else
            DNOCCSReason.Text = ""
            DNOCCSReason.BackColor = Color.LightGray
            DNOCCSReason.ReadOnly = True


        End If


    End Sub

    Sub CheckNGColor()
        If DVCACheck.Checked = True Then

            DCustomerNGColor.BackColor = Color.Yellow
        Else


            DCustomerNGColor.BackColor = Color.GreenYellow
            '  ShowRequiredFieldValidator("DCustomerNGColorRqd", "DCustomerNGColor", "���`�G�ݿ�J�Ȥ�_�M��YKK�⸹")
            DCustomerNGColor.Visible = True


        End If


    End Sub

    Sub CopyNo()
        'COPY ���
        DFactoryCheck.Checked = False
        DVCACheck.Checked = False
        DCustomerCheck.Checked = False

        Dim NO1, sql As String
        NO1 = DNO1.Text
        If NO1 <> "" Then
            sql = "Select * From V_NewColor "
            sql = sql & " Where no  =  '" & NO1 & "'"

            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            If DBAdapter1.Rows.Count > 0 Then


                If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then '�Ȥ�ۮ�
                    DCustomerCheck.Checked = True


                End If
                If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then '�u�t�ۮ�
                    DFactoryCheck.Checked = True


                End If
                If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then '�u�t�ۮ�
                    DVCACheck.Checked = True

                End If



                DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
                DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
                DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
                DBuyerCode.Text = DBAdapter1.Rows(0).Item("BuyerCode")
                DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
                DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
                DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
                DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
                SetFieldData("DevYear", DBAdapter1.Rows(0).Item("DevYear"))    '�~
                SetFieldData("DevSeason", DBAdapter1.Rows(0).Item("DevSeason"))    '�u

                If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                    DReceiveDate.Text = ""
                Else
                    DReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '�U��ɶ�
                End If


                SetFieldData("ColorLight1", DBAdapter1.Rows(0).Item("ColorLight1"))    '���O1
                SetFieldData("ColorLight2", DBAdapter1.Rows(0).Item("ColorLight2"))    '���O2
                SetFieldData("ColorLight3", DBAdapter1.Rows(0).Item("ColorLight3"))    '���O2

                If Mid(DBAdapter1.Rows(0).Item("DeliveryDate").ToString, 1, 4) = "1900" Then
                    DDeliveryDate.Value = ""
                Else
                    DDeliveryDate.Value = DBAdapter1.Rows(0).Item("DeliveryDate")               '�U��ɶ�
                End If

                DDeliveryDate.Value = ""


                DDuplicateNo.Text = DBAdapter1.Rows(0).Item("DuplicateNo")
                If DBAdapter1.Rows(0).Item("NOCCS") = 1 Then
                    DNOCCS.Checked = True
                End If
                DNOCCSReason.Text = DBAdapter1.Rows(0).Item("NOCCSReason")
                DCustomerNGColor.Text = DBAdapter1.Rows(0).Item("CustomerNGColor")
                DRemark.Text = DBAdapter1.Rows(0).Item("Remark")
                SetFieldData("CustomerSample", DBAdapter1.Rows(0).Item("CustomerSample"))    '���O2
            End If


        End If
    End Sub


    Protected Sub DNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNO1.TextChanged
        CopyNo()
    End Sub



End Class

