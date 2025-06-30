Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet01_UAKIPLINGDT
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

        wUserID = Request.QueryString("pUserID")

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

        SQL = "Select * From F_NewColorUAKIPLINGDT "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DNO1.Text = DBAdapter1.Rows(0).Item("No1")

            DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")

            DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
            DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")

            SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))

            DYKKColorCode.Value = DBAdapter1.Rows(0).Item("YKKColorCode")

            SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))

            DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
            DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

            SetFieldData("DTSheet", DBAdapter1.Rows(0).Item("DTSheet"))




        End If


        '�a�J�̷s���
        Dim sqlhistory As String
        sqlhistory = " SELECT"
        sqlhistory = sqlhistory + " a.No  As Field1,"
        sqlhistory = sqlhistory + " case b.Sts When '0' Then '�֩w��' When '1' Then '����' When '2' Then '����' Else '����' End As Field2,"
        sqlhistory = sqlhistory + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3,"
        sqlhistory = sqlhistory + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate,"
        sqlhistory = sqlhistory + " a.FormName as Field4,"
        sqlhistory = sqlhistory + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
        sqlhistory = sqlhistory + " YKKColorType As Field9,YKKColorCode as Field10,SLDColor As Field11,VFColor As Field12,NewOldColor,"
        sqlhistory = sqlhistory + " '....' as WorkFlow, ViewURL,"
        sqlhistory = sqlhistory + " 'http://10.245.1.10/WorkFlow/BefOPList.aspx?' +"
        sqlhistory = sqlhistory + " 'pFormNo='   + a.FormNo +"
        sqlhistory = sqlhistory + " '&pFormSno=' + str(a.FormSno,Len(a.FormSno)) +"
        sqlhistory = sqlhistory + " '&pStep='    + str(a.Step,Len(a.Step)) +"
        sqlhistory = sqlhistory + " '&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) +"
        sqlhistory = sqlhistory + " '&pApplyID=' + a.ApplyID"
        sqlhistory = sqlhistory + " As OPURL, "
        sqlhistory = sqlhistory + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,"
        sqlhistory = sqlhistory + " customerColorCode,overSeaYkkCode,pantonecode,substring(stepnamedesc,7,len(stepnamedesc )-1)stepnamedesc,a.FormSno "
        sqlhistory = sqlhistory + " from V_WaitHandle_01 a,V_NewColor b"
        sqlhistory = sqlhistory + " Where a.formno=b.formno and a.formsno =b.formsno and active  = '1' "
        sqlhistory = sqlhistory + " and a.no ='" & DNO1.Text & "'"
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(sqlhistory)

        If DBAdapter3.Rows.Count > 0 Then

           
            HyperLink1.NavigateUrl = DBAdapter3.Rows(0).Item("OPURL")  '�а��ҩ�
            HyperLink1.Visible = True
            MNOSts.Text = DBAdapter3.Rows(0).Item("stepnamedesc")
            MNOSts.Visible = True
            DOFormSno.Text = DBAdapter3.Rows(0).Item("FormSno")
            HyperLink2.NavigateUrl = "DTMW_NewColorSheet02_UAKIPLING.aspx?pFormNo=005011" & "&pFormSno=" & DOFormSno.Text
            HyperLink2.Visible = True
            Dim i As Integer

            MNOSts.Text = ""

            For i = 0 To DBAdapter3.Rows.Count - 1
                If MNOSts.Text = "" Then
                    MNOSts.Text = DBAdapter3.Rows(i).Item("stepnamedesc")
                Else
                    MNOSts.Text = MNOSts.Text + "," + DBAdapter3.Rows(i).Item("stepnamedesc")
                End If

            Next


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
                    Top = 580
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    If DBAdapter2.Rows(0).Item("Flowtype") = "1" Then
                        Top = 580
                    Else
                        Top = 580
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
            Top = 580
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

        BCopySheet.Attributes.Add("onClick", "CopyNewColor('" + wFormNo + "')") '��buyer


        BYKKColorCode.Attributes.Add("onClick", "GetYKKColorCode('DYKKColorCode')") '��buyer
        Me.DColorSystem.Attributes.CssStyle.Add("text-transform", "uppercase")

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
                    Top = 520
                Else
                    If DDelay.Visible = True Then
                        Top = 520
                    Else
                        Top = 520
                    End If
                End If
            End If
        Else
            Top = 520

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


        'Dim SQL As String


        'MNo
        Select Case FindFieldInf("NO1")
            Case 0  '���
                DNO1.BackColor = Color.LightGray
                DNO1.ReadOnly = True
                DNO1.Visible = True
                BCopySheet.Visible = False
            Case 1  '�ק�+�ˬd
                DNO1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNO1Rqd", "DNO1", "���`�G�ݿ�J�ܢ�")
                DNO1.Visible = True
                BCopySheet.Visible = True
            Case 2  '�ק�
                DNO1.BackColor = Color.Yellow
                DNO1.Visible = True
                BCopySheet.Visible = True
            Case Else   '����
                DNO1.Visible = False
                BCopySheet.Visible = False
        End Select
        If pPost = "New" Then DNO1.Text = ""





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





        'SLDColor
        Select Case FindFieldInf("SLDColor")
            Case 0  '���
                DSLDColor.BackColor = Color.LightGray
                DSLDColor.ReadOnly = True
                DSLDColor.Visible = True

            Case 1  '�ק�+�ˬd
                DSLDColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSLDColorRqd", "DSLDColor", "���`�G�ݿ�J�T�{���Y�ݥΦ�")
                DSLDColor.Visible = True

            Case 2  '�ק�
                DSLDColor.BackColor = Color.Yellow
                DSLDColor.Visible = True
            Case Else   '����
                DSLDColor.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("SLDColor", "ZZZZZZ")

        'VFColor 
        Select Case FindFieldInf("VFColor")
            Case 0  '���
                DVFColor.BackColor = Color.LightGray
                DVFColor.ReadOnly = True
                DVFColor.Visible = True
            Case 1  '�ק�+�ˬd
                DVFColor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVFColorRqd", "DVFColor", "���`�G�ݿ�J�T�{VF�ݥΦ�")
                DVFColor.Visible = True

            Case 2  '�ק�
                DVFColor.BackColor = Color.Yellow
                DVFColor.Visible = True
            Case Else   '����
                DVFColor.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("VFColor", "ZZZZZZ")







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



        'PFBWire


        'PFBWire 
        Select Case FindFieldInf("PFBWire")
            Case 0  '���
                DPFBWire.BackColor = Color.LightGray
                DPFBWire.ReadOnly = True
                DPFBWire.Visible = True
            Case 1  '�ק�+�ˬd
                DPFBWire.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFBWireRqd", "DPFBWire", "���`�G�ݿ�J�T�{VF�ݥΦ�")
                DPFBWire.Visible = True

            Case 2  '�ק�
                DPFBWire.BackColor = Color.Yellow
                DPFBWire.Visible = True
            Case Else   '����
                DPFBWire.Visible = False
        End Select
        If pPost = "New" Then DPFBWire.Text = ""



        'DPFOpenParts 
        Select Case FindFieldInf("PFOpenParts")
            Case 0  '���
                DPFOpenParts.BackColor = Color.LightGray
                DPFOpenParts.ReadOnly = True
                DPFOpenParts.Visible = True
            Case 1  '�ק�+�ˬd
                DPFOpenParts.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPFOpenPartsRqd", "DPFOpenParts", "���`�G�ݿ�J�T�{VF�ݥΦ�")
                DPFOpenParts.Visible = True

            Case 2  '�ק�
                DPFOpenParts.BackColor = Color.Yellow
                DPFOpenParts.Visible = True
            Case Else   '����
                DPFOpenParts.Visible = False
        End Select
        If pPost = "New" Then DPFOpenParts.Text = ""




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

        '�֥i�����
        Select Case FindFieldInf("DTSheet")
            Case 0  '���
                DDTSheet.BackColor = Color.LightGray
                DDTSheet.Visible = True

            Case 1  '�ק�+�ˬd
                DDTSheet.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDTSheetRqd", "DDTSheet", "���`�G�ݿ�֥i�����")
                DDTSheet.Visible = True
            Case 2  '�ק�
                DDTSheet.BackColor = Color.Yellow
                DDTSheet.Visible = True
            Case Else   '����
                DDTSheet.Visible = False
        End Select

        If pPost = "New" Then SetFieldData("DTSheet", "ZZZZZZ")


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


        '�֥i�����
        If pFieldName = "DTSheet" Then
            DDTSheet.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDTSheet.Items.Add(ListItem1)
                End If
            Else
                SQL = "  Select  * from M_referp"
                SQL = SQL & " where  cat = '5001'"
                SQL = SQL & " and dkey = 'DTSheet'"


                Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
                DDTSheet.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDTSheet.Items.Add(ListItem1)
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
            If DNO1.Text <> "" Then
                ErrCode = oCommon.CommissionNo("005012", wFormSno, wStep, DNO1.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                            ModifyTranData(pFun, pSts)

                            '�y�{��Ƶ���
                            'Modify-Start by Joy 2009/11/20(2010��ƾ����)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo,wUserID, pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, wUserID, pRunNextStep)
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

                    If wStep = 10 Then

                        If DDTSheet.SelectedValue = "���a����" Then
                            pAction = 2
                        Else

                            If DNewOldColor.SelectedValue = "�s��" Then
                                If DDTSheet.SelectedValue = "���a" Then
                                    pAction = 0
                                Else
                                    pAction = 3
                                End If
                            Else
                                If DYKKColorType.SelectedValue = "�å�" Then
                                    pAction = 1  '151
                                Else
                                    pAction = 2
                                End If
                            End If

                        End If




                    End If

                    If wStep = 20 Then
                        If DYKKColorType.SelectedValue = "�å�" Then
                            pAction = 0
                        Else
                            pAction = 1
                        End If
                    End If

                    If wStep = 152 Then
                        If DDTSheet.SelectedValue = "���a����" Then
                            pAction = 1
                        ElseIf DDTSheet.SelectedValue = "���Y" Then
                            pAction = 2
                        Else
                            pAction = 0
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


                    DNO1.Text = ""
                    DNo.Text = ""
                    HyperLink1.Text = ""
                    MNOSts.Text = ""

                    wFormSno = wbFormSno
                    wStep = wbStep


                Else



                    URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                "&pUserID=" & wUserID & "&pApplyID=" & wApplyID

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
        SQl = "Insert into F_NewColorUAKIPLINGDT "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, DepoName,Date,Name,colorsystem,"
        SQl = SQl + "PFBwire,PFOpenParts,YKKColorType,YKKColorCode,"
        SQl = SQl + "SLDColor,VFCOLOR,NewOldColor,oFormNo,oFormSno,DTSheet,MNOSts,No1,"
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '005012', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����
        SQl = SQl + " N'" + DNo.Text + "', "                '����2
        SQl = SQl + " N'" + DDepoName.Text + "', "                '���3
        SQl = SQl + " N'" + DDate.Text + "', "                '�m�W4
        SQl = SQl + " N'" + DName.Text + "', "                '����2
        SQl = SQl + " N'" + DColorSystem.Text + "', "                '���3
        SQl = SQl + " N'" + DPFBWire.Text + "', "
        SQl = SQl + " N'" + DPFOpenParts.Text + "', "
        SQl = SQl + " N'" + DYKKColorType.Text + "', "
        SQl = SQl + " N'" + DYKKColorCode.Value + "', "
        SQl = SQl + " N'" + DSLDColor.Text.ToUpper + "', "
        SQl = SQl + " N'" + DVFColor.Text.ToUpper + "', "
        SQl = SQl + " N'" + DNewOldColor.Text.ToUpper + "', "
        SQl = SQl + " '005011',"
        SQl = SQl + " N'" + DOFormSno.Text + "', "
        SQl = SQl + " N'" + DDTSheet.Text + "', "
        SQl = SQl + " N'" + MNOSts.Text + "', "
        SQl = SQl + " N'" + DNO1.Text + "', "
        SQl = SQl + " '" + wUserID + "', "     '�@����
        SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                       '�ק��
        SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

        SQl = " Update F_NewColorUAKIPLING"
        SQl = SQl + " Set "
        SQl = SQl + " DTNO =1 "
        SQl = SQl + " ,OFormNo ='005012' "
        SQl = SQl + " ,OFormsNo ='" + CStr(NewFormSno) + "'"
        SQl = SQl + " ,DTSheet = '" + DDTSheet.Text + "'"
        SQl = SQl + " where formsno ='" + DOFormSno.Text + "'"
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

        SQL = " Update F_NewColorUAKIPLINGDT"
        SQL = SQL + " Set "
        If pFun <> "SAVE" Then
            SQL = SQL + " Sts = '" & pSts & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQL = SQL + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL = SQL + " Date = N'" & DDate.Text & "',"
        SQL = SQL + " DepoName = N'" & DDepoName.Text & "',"
        SQL = SQL + " Name = N'" & DName.Text & "',"
        SQL = SQL + " ColorSystem = N'" & DColorSystem.Text.ToUpper & "',"
        SQL = SQL + " PFBWire = N'" & DPFBWire.Text & "',"
        SQL = SQL + " PFOpenParts = N'" & DPFOpenParts.Text & "',"
        SQL = SQL + " YKKColorType = N'" & DYKKColorType.Text & "',"
        SQL = SQL + " YKKColorCode = N'" & DYKKColorCode.Value & "',"
        SQL = SQL + " SLDColor = N'" & DSLDColor.Text.ToUpper & "',"
        SQL = SQL + " VFColor = N'" & DVFColor.Text.ToUpper & "',"
        SQL = SQL + " NewOldColor = N'" & DNewOldColor.Text & "',"
        SQL = SQL + " oFormNo= '005011',"
        SQL = SQL + " oFormSno = N'" & DOFormSno.Text & "',"
        SQL = SQL + " DTSheet = N'" & DDTSheet.Text & "',"
        SQL = SQL + " MNoSts = N'" & MNOSts.Text & "',"
        SQL = SQL + " No1 = N'" & DNO1.Text & "',"
        SQL = SQL + " ModifyUser = '" & wUserID & "',"
        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
        SQL = SQL + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL)




        SQL = " Update F_NewColorUAKIPLING"
        SQL = SQL + " Set "
        SQL = SQL + " DTNO =1 "
        SQL = SQL + " ,DTSheet = '" + DDTSheet.Text + "'"
        SQL = SQL + " ,NewoldColor='" + DNewOldColor.SelectedValue + "'"

        If DDTSheet.SelectedValue = "���a" Then
            SQL = SQL + " ,YKKColorCode = '" + DYKKColorCode.Value + "'"
        ElseIf DDTSheet.SelectedValue = "���a����" Then
            SQL = SQL + " ,YKKColorCode = '" + DYKKColorCode.Value + "'"
            SQL = SQL + " ,YKKColorCodeVF = '" + DYKKColorCode.Value + "'"

        ElseIf DDTSheet.SelectedValue = "����" Then
            SQL = SQL + " ,YKKColorCodeVF = '" + DYKKColorCode.Value + "'"
        Else
            SQL = SQL + " ,YKKColorCodeSLD = '" + DYKKColorCode.Value + "'"
        End If

        SQL = SQL + " where formsno ='" + DOFormSno.Text + "'"
        uDataBase.ExecuteNonQuery(SQL)


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
                If DNO1.Text <> DBAdapter1.Rows(0).Item("No") Then  'No
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
            DDescSheet.Style("top") = Top - 30 & "px"
            DDecideDesc.Style("top") = Top - 30 + 6 & "px"
            Top = Top + 70
        End If

        If DDelay.Visible Then                                          ' ���𻡩�
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"


        Top = Top + 30
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
        Str = "K" + CStr(DateTime.Now.Year) + Str2
        '�~
        'Set�渹
        '�������Ʀ��X��  20150414 Modify by Jessica
        SQL = " select   isnull(max(convert(int,substring(no,8,4))),0) cun from  F_NewColorUAKIPLINGDT"
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
        Top = 600
        SetControlPosition()
        '  ShowFormData()
        Dim SQL As String
        Dim DOUBLENo As Integer = 0

        Dim Errcode As Integer = 0





        'Dim Q As Integer = 0
        ''jessica 20150326
        'If wStep = 10 Then
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


    Sub MsgBox(ByVal text As String)
        'Dim scriptstr As String
        'scriptstr = "<script language=javascript>" + Chr(10) _
        '+ "confirm(""" + text + """)" + Chr(10) _
        '+ "</script>"
        'Response.Write(scriptstr)
        'Response.Write("<script language='javascript'>if(confirm(""" + text + """)==true)function1();else function2();</script>")
        'Response.Write("<script language=javascript>if (confirm('�A�T�w�n�~���X��r�H')==false) {return fale;};</script>")
    End Sub

    Sub CopyNo()
        'COPY ���


        Dim NO1, sql, sqlhistory As String
        NO1 = DNO1.Text
        If NO1 <> "" Then
            sql = "Select * From V_NewColor "
            sql = sql & " Where no  =  '" & NO1 & "'"

            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(sql)

            If DBAdapter1.Rows.Count > 0 Then
                DColorSystem.Text = DBAdapter1.Rows(0).Item("ColorSystem")
                DPFBWire.Text = DBAdapter1.Rows(0).Item("PFBWire")
                DPFOpenParts.Text = DBAdapter1.Rows(0).Item("PFOpenParts")
                DSLDColor.Text = DBAdapter1.Rows(0).Item("SLDColor")
                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")
                DVFColor.Text = DBAdapter1.Rows(0).Item("VFColor")

                SetFieldData("NewOldColor", DBAdapter1.Rows(0).Item("NewOldColor"))    '���O2
                SetFieldData("YKKColorType", DBAdapter1.Rows(0).Item("YKKColorType"))    '���O2

            End If

            sqlhistory = " SELECT"
            sqlhistory = sqlhistory + " a.No  As Field1,"
            sqlhistory = sqlhistory + " case b.Sts When '0' Then '�֩w��' When '1' Then '����' When '2' Then '����' Else '����' End As Field2,"
            sqlhistory = sqlhistory + " Convert(VARCHAR(10), a.ApplyTime, 111) as Field3,"
            sqlhistory = sqlhistory + " case when b.sts ='0' then '' else Convert(VARCHAR(10), b.completedtime, 111)  end as completedate,"
            sqlhistory = sqlhistory + " a.FormName as Field4,"
            sqlhistory = sqlhistory + " a.Division As Field5,a.ApplyName as Field6,b.Customer As Field7,b.Buyer As Field8,"
            sqlhistory = sqlhistory + " YKKColorType As Field9,YKKColorCode as Field10,SLDColor As Field11,VFColor As Field12,NewOldColor,"
            sqlhistory = sqlhistory + " '....' as WorkFlow, ViewURL,"
            sqlhistory = sqlhistory + " 'http://10.245.1.10/WorkFlow/BefOPList.aspx?' +"
            sqlhistory = sqlhistory + " 'pFormNo='   + a.FormNo +"
            sqlhistory = sqlhistory + " '&pFormSno=' + str(a.FormSno,Len(a.FormSno)) +"
            sqlhistory = sqlhistory + " '&pStep='    + str(a.Step,Len(a.Step)) +"
            sqlhistory = sqlhistory + " '&pSeqNo='   + str(a.SeqNo,Len(a.SeqNo)) +"
            sqlhistory = sqlhistory + " '&pApplyID=' + a.ApplyID"
            sqlhistory = sqlhistory + " As OPURL, "
            sqlhistory = sqlhistory + " Case when b.sts = 1 then convert(char(10),b.completedTime,111) else '' end as completedDate,customerColor,"
            sqlhistory = sqlhistory + " customerColorCode,overSeaYkkCode,pantonecode,substring(stepnamedesc,7,len(stepnamedesc )-1)stepnamedesc,a.FormSno "
            sqlhistory = sqlhistory + " from V_WaitHandle_01 a,V_NewColor b"
            sqlhistory = sqlhistory + " Where a.formno=b.formno and a.formsno =b.formsno and active  = '1' "
            sqlhistory = sqlhistory + " and a.no ='" & NO1 & "'"
            Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(sqlhistory)
            Dim i As Integer
            If DBAdapter2.Rows.Count > 0 Then
                HyperLink1.NavigateUrl = DBAdapter2.Rows(0).Item("OPURL")  ' 
                HyperLink1.Visible = True


                MNOSts.Text = ""

                For i = 0 To DBAdapter2.Rows.Count - 1
                    If MNOSts.Text = "" Then
                        MNOSts.Text = DBAdapter2.Rows(i).Item("stepnamedesc")
                    Else
                        MNOSts.Text = MNOSts.Text + "," + DBAdapter2.Rows(i).Item("stepnamedesc")
                    End If

                Next



                'MNOSts.Text = DBAdapter2.Rows(0).Item("stepnamedesc")
                MNOSts.Visible = True
                DOFormSno.Text = DBAdapter2.Rows(0).Item("FormSno")
                DDTSheet.Text = ""
                ' DYKKColorType.SelectedValue = ""
                ' DNewOldColor.SelectedValue = ""
            End If

        End If
    End Sub


    Protected Sub DNO1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNO1.TextChanged
        CopyNo()
    End Sub


End Class

