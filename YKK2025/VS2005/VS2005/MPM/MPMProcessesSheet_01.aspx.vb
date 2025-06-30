Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class MPMProcessesSheet_01
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

            '�W�Ǹ���ˬd����ܰT��

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
        If wStep = 40 Or wStep = 500 Then
            Dim SQL As String

            SQL = "Select * From T_waitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & " and active=1  and step ='" + CStr(wStep) + "'"
            Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
            If DBAdapter1.Rows.Count > 0 Then
                wUserID = DBAdapter1.Rows(0).Item("Workid")
            End If
        Else
            wUserID = Request.QueryString("pUserID")
        End If



        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)

        Response.Cookies("PGM").Value = "MPMProcessesSheet_01.aspx"
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
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("MPMProcessesFilePath")



        SQL = "Select * From F_MPMProcessesSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            If wStep = 1 Then
                LNO.Visible = False
            Else
                LNO.Visible = True
                LNO.NavigateUrl = "MPMProcessesReport_02.aspx?&pNo=" + DBAdapter1.Rows(0).Item("No")
            End If


            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Value = DBAdapter1.Rows(0).Item("AppDate")              'AppDate           
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DClinter.Text = DBAdapter1.Rows(0).Item("Clinter")              'Clinter
            ' DDivisionCode.SelectedValue = DBAdapter1.Rows(0).Item("DivisionCode") + "-" + DBAdapter1.Rows(0).Item("Division")
            'DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            '  DDivisionCode.Text = DBAdapter1.Rows(0).Item("DivisionCode")              'Division
            SetFieldData("DivisionCode", DBAdapter1.Rows(0).Item("DivisionCode") + "-" + DBAdapter1.Rows(0).Item("Division"))    '���O1

            SetFieldData("Type1", DBAdapter1.Rows(0).Item("Type1"))    '���O1
            SetFieldData("Type2", DBAdapter1.Rows(0).Item("Type2"))    '���O1
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'Code
            DQty.Text = DBAdapter1.Rows(0).Item("Qty")              'Qty
            DAQty.Text = DBAdapter1.Rows(0).Item("AQty")              'Qty
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'MapNo
            DWeight.Text = DBAdapter1.Rows(0).Item("Weight")              'Weight
            DMaterial.Text = DBAdapter1.Rows(0).Item("Material")              'Weight
            DDevNo.Text = DBAdapter1.Rows(0).Item("DevNo")              'Weight
            DManufout.Text = DBAdapter1.Rows(0).Item("ManufOut")              'Weight

            If Mid(DBAdapter1.Rows(0).Item("FinishDate").ToString, 1, 4) = "1900" Then
                DFinishDate.Value = ""
            Else
                DFinishDate.Value = DBAdapter1.Rows(0).Item("FinishDate")               '�U��ɶ�
            End If

            '  If Mid(DBAdapter1.Rows(0).Item("AFinishDate").ToString, 1, 4) = "1900" Then
            If wStep = 40 Then
                DAFinishDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")
            End If




            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '������
                LMapFile1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  '�~��̿��
                LMapFile.Visible = True
                LMapFile1.Visible = True
            Else
                LMapFile.Visible = False
                LMapFile1.Visible = False
            End If
        End If


        Dim EngineStr As String

        SQL = "Select Engine,RecDate,StartDate,EndDate,WorkTime,Starter,remark,SeqNo From F_MPMProcessesSheetDT "
        SQL = SQL & " Where delmark =0 and seqno1 =0 and FormSno =  '" & CStr(wFormSno) & "' order by SeqNo"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

        For Each dtr As Data.DataRow In dt.Rows
            EngineStr = "DEngine" + Trim(Str(dtr("SeqNo")))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = dtr("Engine")
        Next



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
                    Top = 1100
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 800
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
            Top = 600
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
        BDate1.Attributes.Add("onClick", "calendarPicker('DFinishDate')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DAFinishDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DAppDate')")

        BClienter.Attributes.Add("onClick", "EmpDatePicker('" + wApplyID + "')") '����u
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
        Dim SQL As String
        Dim Empid As String

        Empid = D1.Text




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
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "lightgrey")
                DAppDate.Attributes.Add("readonly", "true")
                BDate3.Visible = False

            Case 1  '�ק�+�ˬd
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "greenyellow")
                DAppDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DAppDateRqd", "DAppDate", "���`�G�ݿ�J�w�w������")
                BDate3.Visible = False

            Case 2  '�ק�
                DAppDate.Visible = True
                DAppDate.Style.Add("background-color", "yellow")
                DAppDate.Attributes.Add("readonly", "true")
                BDate3.Visible = False

            Case Else   '����
                DAppDate.Visible = False
                BDate3.Visible = False

        End Select

        If pPost = "New" Then DAppDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")








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


        '�̿��
        Select Case FindFieldInf("Clinter")
            Case 0  '���
                DClinter.BackColor = Color.LightGray
                DClinter.Visible = True
                DClinter.ReadOnly = True
                BClienter.Visible = False
            Case 1  '�ק�+�ˬd
                DClinter.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DClinterRqd", "DClinter", "���`�G�ݿ�J�̿��")
                DClinter.Visible = True
                '    DClinter.ReadOnly = True
                BClienter.Visible = False
            Case 2  '�ק�
                DClinter.BackColor = Color.Yellow
                DClinter.Visible = True
                DClinter.ReadOnly = True
                BClienter.Visible = False
            Case Else   '����
                DClinter.Visible = False
                BClienter.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Clinter", "ZZZZZZ")


        '����
        Select Case FindFieldInf("DivisionCode")
            Case 0  '���
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.Visible = True

            Case 1  '�ק�+�ˬd
                DDivisionCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "���`�G�ݿ�J���O1")
                DDivisionCode.Visible = True
            Case 2  '�ק�
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.Visible = True
            Case Else   '����
                DDivisionCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DivisionCode", "ZZZZZZ")




        '����
        '  Select Case FindFieldInf("Division")
        '      Case 0  '���
        ' DDivision.BackColor = Color.LightGray
        ' DDivision.Visible = True
        ' DDivision.ReadOnly = True
        '     Case 1  '�ק�+�ˬd
        ' DDivision.BackColor = Color.GreenYellow
        'ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "���`�G�ݿ�J����")
        'DDivision.Visible = True
        'DDivision.ReadOnly = True
        '    Case 2  '�ק�
        'DDivision.BackColor = Color.Yellow
        'DDivision.Visible = True
        '    Case Else   '����
        'DDivision.Visible = False
        'End Select
        'If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")



        '���O1
        Select Case FindFieldInf("Type1")
            Case 0  '���
                DType1.BackColor = Color.LightGray
                DType1.Visible = True

            Case 1  '�ק�+�ˬd
                DType1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DType1Rqd", "DType1", "���`�G�ݿ�J���O1")
                DType1.Visible = True
            Case 2  '�ק�
                DType1.BackColor = Color.Yellow
                DType1.Visible = True
            Case Else   '����
                DType1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Type1", "ZZZZZZ")


        '���O2
        Select Case FindFieldInf("Type2")
            Case 0  '���
                DType2.BackColor = Color.LightGray
                DType2.Visible = True

            Case 1  '�ק�+�ˬd
                DType2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DType2Rqd", "DType2", "���`�G�ݿ�J���O2")
                DType2.Visible = True
            Case 2  '�ק�
                DType2.BackColor = Color.Yellow
                DType2.Visible = True
            Case Else   '����
                DType2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Type2", "ZZZZZZ")


        '���O2
        Select Case FindFieldInf("Engine")
            Case 0  '���
                DEngine.BackColor = Color.LightGray
                DEngine.Visible = True


            Case 1  '�ק�+�ˬd
                DEngine.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DEngineRqd", "DEngine", "���`�G�ݿ�J�u�{")
                DEngine.Visible = True

            Case 2  '�ק�
                DEngine.BackColor = Color.Yellow
                DEngine.Visible = True

            Case Else   '����
                DEngine.Visible = False


        End Select
        If pPost = "New" Then SetFieldData("Engine", "ZZZZZZ")



        '�ϸ�
        Select Case FindFieldInf("MapNo")
            Case 0  '���
                DMapNo.BackColor = Color.LightGray
                DMapNo.Visible = True
                DMapNo.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "���`�G�ݿ�J�ϸ�")
                DMapNo.Visible = True
            Case 2  '�ק�
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
            Case Else   '����
                DMapNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MapNo", "ZZZZZZ")

        '�ƶq
        Select Case FindFieldInf("Qty")
            Case 0  '���
                DQty.BackColor = Color.LightGray
                DQty.Visible = True
                DQty.ReadOnly = True
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
        If pPost = "New" Then SetFieldData("Qty", "ZZZZZZ")

        '�ƶq
        Select Case FindFieldInf("AQty")
            Case 0  '���
                DAQty.BackColor = Color.LightGray
                DAQty.Visible = True
                DAQty.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DAQty.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAQtyRqd", "DAQty", "���`�G�ݿ�J�ƶq")
                DAQty.Visible = True
            Case 2  '�ק�
                DAQty.BackColor = Color.Yellow
                DAQty.Visible = True
            Case Else   '����
                DAQty.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AQty", "ZZZZZZ")




        'Code
        Select Case FindFieldInf("Code")
            Case 0  '���
                DCode.BackColor = Color.LightGray
                DCode.Visible = True
                DCode.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeRqd", "DCode", "���`�G�ݿ�J�ƶq")
                DCode.Visible = True
            Case 2  '�ק�
                DCode.BackColor = Color.Yellow
                DCode.Visible = True
            Case Else   '����
                DCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Code", "ZZZZZZ")





        '�~�W
        Select Case FindFieldInf("Weight")
            Case 0  '���
                DWeight.BackColor = Color.LightGray
                DWeight.Visible = True
                DWeight.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DWeight.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWeightRqd", "DWeight", "���`�G�ݿ�J�~�W")
                DWeight.Visible = True
            Case 2  '�ק�
                DWeight.BackColor = Color.Yellow
                DWeight.Visible = True
            Case Else   '����
                DWeight.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Weight", "ZZZZZZ")



        '����
        Select Case FindFieldInf("Material")
            Case 0  '���
                DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = True
                DMaterial.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "���`�G�ݿ�J�~�W")
                DMaterial.Visible = True
            Case 2  '�ק�
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '����
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")




        '����
        Select Case FindFieldInf("DevNo")
            Case 0  '���
                DDevNo.BackColor = Color.LightGray
                DDevNo.Visible = True
                DDevNo.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DDevNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevNoRqd", "DDevNo", "���`�G�ݿ�J�~�W")
                DDevNo.Visible = True
            Case 2  '�ק�
                DDevNo.BackColor = Color.Yellow
                DDevNo.Visible = True
            Case Else   '����
                DDevNo.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("DevNo", "ZZZZZZ")


        '����
        Select Case FindFieldInf("Manufout")
            Case 0  '���
                DManufout.BackColor = Color.LightGray
                DManufout.Visible = True
                DManufout.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DManufout.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufoutRqd", "DManufout", "���`�G�ݿ�J�~�W")
                DManufout.Visible = True
            Case 2  '�ק�
                DManufout.BackColor = Color.Yellow
                DManufout.Visible = True
            Case Else   '����
                DManufout.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Manufout", "ZZZZZZ")



        '�w�w������
        Select Case FindFieldInf("FinishDate")
            Case 0  '���
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "lightgrey")
                DFinishDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = True


            Case 1  '�ק�+�ˬd
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "greenyellow")
                DFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DFinishDateRqd", "DFinishDate", "���`�G�ݿ�J�w�w������")
                BDate1.Disabled = False

            Case 2  '�ק�
                DFinishDate.Visible = True
                DFinishDate.Style.Add("background-color", "yellow")
                DFinishDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = False

            Case Else   '����
                DFinishDate.Visible = False
                BDate1.Disabled = True

        End Select
        If pPost = "New" Then DFinishDate.Value = ""



        '������
        Select Case FindFieldInf("AFinishDate")
            Case 0  '���
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "lightgrey")
                DAFinishDate.Attributes.Add("readonly", "true")
                BDate2.Visible = False

            Case 1  '�ק�+�ˬd
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "greenyellow")
                DAFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DAFinishDateRqd", "DAFinishDate", "���`�G�ݿ�J�w�w������")
                BDate2.Visible = True


            Case 2  '�ק�
                DAFinishDate.Visible = True
                DAFinishDate.Style.Add("background-color", "yellow")
                DAFinishDate.Attributes.Add("readonly", "true")
                BDate2.Visible = True


            Case Else   '����
                DAFinishDate.Visible = False
                BDate2.Visible = False

        End Select

        If pPost = "New" Then DAFinishDate.Value = CDate(NowDateTime).ToString("yyyy/MM/dd")




        '²��
        Select Case FindFieldInf("MapFile")
            Case 0  '���
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "���`�G�ݿ�J²��")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DMapFile.Visible = False
        End Select
        LMapFile.Visible = False
        LMapFile1.Visible = False

        '�u�{1
        Select Case FindFieldInf("Engine1")
            Case 0  '���
                DEngine1.BackColor = Color.LightGray
                DEngine1.Visible = True
                DEngine1.ReadOnly = True


            Case 1  '�ק�+�ˬd
                DEngine1.BackColor = Color.GreenYellow
                '   ShowRequiredFieldValidator("DEngine1Rqd", "DEngine1", "���`�G�ݿ�J�~�W")
                DEngine1.Visible = True

            Case 2  '�ק�
                DEngine1.BackColor = Color.Yellow
                DEngine1.Visible = True
                DEngine1.ReadOnly = True

            Case Else   '����
                DEngine1.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("Engine1", "ZZZZZZ")

        '�u�{2
        Select Case FindFieldInf("Engine2")
            Case 0  '���
                DEngine2.BackColor = Color.LightGray
                DEngine2.Visible = True
                DEngine2.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine2Rqd", "DEngine2", "���`�G�ݿ�J�~�W")
                DEngine2.Visible = True
            Case 2  '�ק�
                DEngine2.BackColor = Color.Yellow
                DEngine2.Visible = True
                DEngine2.ReadOnly = True
            Case Else   '����
                DEngine2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine2", "ZZZZZZ")

        '�u�{3
        Select Case FindFieldInf("Engine3")
            Case 0  '���
                DEngine3.BackColor = Color.LightGray
                DEngine3.Visible = True
                DEngine3.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine3Rqd", "DEngine3", "���`�G�ݿ�J�~�W")
                DEngine3.Visible = True
            Case 2  '�ק�
                DEngine3.BackColor = Color.Yellow
                DEngine3.Visible = True
                DEngine3.ReadOnly = True
            Case Else   '����
                DEngine3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine3", "ZZZZZZ")


        '�u�{3
        Select Case FindFieldInf("Engine4")
            Case 0  '���
                DEngine4.BackColor = Color.LightGray
                DEngine4.Visible = True
                DEngine4.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine4Rqd", "DEngine4", "���`�G�ݿ�J�~�W")
                DEngine4.Visible = True
            Case 2  '�ק�
                DEngine4.BackColor = Color.Yellow
                DEngine4.Visible = True
                DEngine4.ReadOnly = True
            Case Else   '����
                DEngine4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine4", "ZZZZZZ")



        '�u�{3
        Select Case FindFieldInf("Engine5")
            Case 0  '���
                DEngine5.BackColor = Color.LightGray
                DEngine5.Visible = True
                DEngine5.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine5Rqd", "DEngine5", "���`�G�ݿ�J�~�W")
                DEngine5.Visible = True
            Case 2  '�ק�
                DEngine5.BackColor = Color.Yellow
                DEngine5.Visible = True
                DEngine5.ReadOnly = True
            Case Else   '����
                DEngine5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine5", "ZZZZZZ")


        '�u�{3
        Select Case FindFieldInf("Engine6")
            Case 0  '���
                DEngine6.BackColor = Color.LightGray
                DEngine6.Visible = True
                DEngine6.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine6Rqd", "DEngine6", "���`�G�ݿ�J�~�W")
                DEngine6.Visible = True
            Case 2  '�ק�
                DEngine6.BackColor = Color.Yellow
                DEngine6.Visible = True
                DEngine6.ReadOnly = True
            Case Else   '����
                DEngine6.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine6", "ZZZZZZ")


        '�u�{3
        Select Case FindFieldInf("Engine7")
            Case 0  '���
                DEngine7.BackColor = Color.LightGray
                DEngine7.Visible = True
                DEngine7.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine7Rqd", "DEngine7", "���`�G�ݿ�J�~�W")
                DEngine7.Visible = True
            Case 2  '�ק�
                DEngine7.BackColor = Color.Yellow
                DEngine7.Visible = True
                DEngine7.ReadOnly = True
            Case Else   '����
                DEngine7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine7", "ZZZZZZ")



        '�u�{3
        Select Case FindFieldInf("Engine8")
            Case 0  '���
                DEngine8.BackColor = Color.LightGray
                DEngine8.Visible = True
                DEngine8.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine8.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine8Rqd", "DEngine8", "���`�G�ݿ�J�~�W")
                DEngine8.Visible = True
            Case 2  '�ק�
                DEngine8.BackColor = Color.Yellow
                DEngine8.Visible = True
                DEngine8.ReadOnly = True
            Case Else   '����
                DEngine8.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine8", "ZZZZZZ")


        '�u�{3
        Select Case FindFieldInf("Engine9")
            Case 0  '���
                DEngine9.BackColor = Color.LightGray
                DEngine9.Visible = True
                DEngine9.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine9.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine9Rqd", "DEngine9", "���`�G�ݿ�J�~�W")
                DEngine9.Visible = True
            Case 2  '�ק�
                DEngine9.BackColor = Color.Yellow
                DEngine9.Visible = True
                DEngine9.ReadOnly = True
            Case Else   '����
                DEngine9.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine9", "ZZZZZZ")

        '�u�{3
        Select Case FindFieldInf("Engine10")
            Case 0  '���
                DEngine10.BackColor = Color.LightGray
                DEngine10.Visible = True
                DEngine10.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine10.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine10Rqd", "DEngine10", "���`�G�ݿ�J�~�W")
                DEngine10.Visible = True
            Case 2  '�ק�
                DEngine10.BackColor = Color.Yellow
                DEngine10.Visible = True
                DEngine10.ReadOnly = True
            Case Else   '����
                DEngine10.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine10", "ZZZZZZ")

        '�u�{3
        Select Case FindFieldInf("Engine11")
            Case 0  '���
                DEngine11.BackColor = Color.LightGray
                DEngine11.Visible = True
                DEngine11.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine11.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine11Rqd", "DEngine11", "���`�G�ݿ�J�~�W")
                DEngine11.Visible = True
            Case 2  '�ק�
                DEngine11.BackColor = Color.Yellow
                DEngine11.Visible = True
                DEngine11.ReadOnly = True
            Case Else   '����
                DEngine11.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine11", "ZZZZZZ")

        '�u�{3
        Select Case FindFieldInf("Engine12")
            Case 0  '���
                DEngine12.BackColor = Color.LightGray
                DEngine12.Visible = True
                DEngine12.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine12.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine12Rqd", "DEngine12", "���`�G�ݿ�J�~�W")
                DEngine12.Visible = True
            Case 2  '�ק�
                DEngine12.BackColor = Color.Yellow
                DEngine12.Visible = True
                DEngine12.ReadOnly = True
            Case Else   '����
                DEngine12.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine12", "ZZZZZZ")

        '�u�{3
        Select Case FindFieldInf("Engine13")
            Case 0  '���
                DEngine13.BackColor = Color.LightGray
                DEngine13.Visible = True
                DEngine13.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine13.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine13Rqd", "DEngine13", "���`�G�ݿ�J�~�W")
                DEngine13.Visible = True
            Case 2  '�ק�
                DEngine13.BackColor = Color.Yellow
                DEngine13.Visible = True
                DEngine13.ReadOnly = True
            Case Else   '����
                DEngine13.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine13", "ZZZZZZ")



        '�u�{3
        Select Case FindFieldInf("Engine14")
            Case 0  '���
                DEngine14.BackColor = Color.LightGray
                DEngine14.Visible = True
                DEngine14.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DEngine14.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngine14Rqd", "DEngine14", "���`�G�ݿ�J�~�W")
                DEngine14.Visible = True
            Case 2  '�ק�
                DEngine14.BackColor = Color.Yellow
                DEngine14.Visible = True
                DEngine14.ReadOnly = True
            Case Else   '����
                DEngine14.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engine14", "ZZZZZZ")


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

        '���̤γ��� 

        SQL = "Select Divname,Username From M_Users "
        SQL = SQL & " Where UserID = '" & wApplyID & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(SQL)
        DAppper.Text = DBUser.Rows(0).Item("Username")

        'DNo.Text = SetNo(wFormSno)




        If pFieldName = "DivisionCode" Then
            DDivisionCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then


                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivisionCode.Items.Add(ListItem1)
                End If
            Else
                SQL = " select  distinct dep_code+'-'+dep_name Data  from M_emp"
                SQL = SQL + " where com_code  in ('01','60','65','70','71')"
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010010-�u�t�@�q' Data from M_emp"
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010331-�u��(SPD)' Data from M_emp "
                SQL = SQL + " union all"
                SQL = SQL + " select top 1 '1010331-�u��(���})' Data from M_emp "
                SQL = SQL + " order by data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DDivisionCode.Items.Add("")
                'DDivisionCode.Items.Add("1010010-�u�t�@�q")
                'DDivisionCode.Items.Add("1010331-�u��(SPD)")
                'DDivisionCode.Items.Add("1010331-�u��(���})")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then
                        ListItem1.Selected = True
                    End If

                    DDivisionCode.Items.Add(ListItem1)
                Next
            End If

        End If





        If pFieldName = "Type1" Then
            DType1.Items.Clear()

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DType1.Items.Add(ListItem1)

                End If
            Else
                SQL = "Select * From M_Referp Where Cat='4001' and dkey='TYPE1' Order by DKey, Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DType1.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DType1.Items.Add(ListItem1)
                Next
            End If

        End If

        If pFieldName = "Type2" Then
            DType2.Items.Clear()

            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DType2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='4001' and dkey='TYPE2' Order by DKey, Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DType2.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DType2.Items.Add(ListItem1)
                Next
            End If



        End If

        If pFieldName = "Engine" Then
            DEngine.Items.Clear()

            If pName <> "ZZZZZZ" Then
                Dim ListItem1 As New ListItem
                ListItem1.Text = pName
                ListItem1.Value = pName
                DEngine.Items.Add(ListItem1)

            Else
                SQL = "Select distinct substring(dkey,14,len(dkey)-1)data  From M_Referp Where Cat='4001' and substring(dkey,1,12) =  'EngineSelect' Order by Data "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DEngine.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEngine.Items.Add(ListItem1)
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
            DisabledButton()   '����Button�B�@
            FlowControl("NG1", 1, "2")

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK���s�I���ƥ�
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick

        AQty = DAQty.Text
        DevNo = DDevNo.Text
        Manufout = DManufout.Text

        If Request.Cookies("RunBOK").Value = True Then

            If (wStep = 10) Or (wStep = 20) Or (wStep = 30) Or (wStep = 40) Then
                If OK() = True Then
                    DisabledButton()   '����Button�B�@
                    FlowControl("OK", 0, "1")
                End If
            Else
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
                ErrCode = oCommon.CommissionNo("004001", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, wUserID, wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '��渹�X,�u�{���d���X,ñ�֪�,�ӽЪ�,�Q�N�z��,�Q���w��,�h�H�֩w�u�{No,
                    '�U�@�u�{��, ���X, ����, �Q�N�z��, �H��, �B�z��k, �ʧ@(0:OK,1:NG,2:Save) 

                    If wStep = 15 Then

                    End If

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
                        oCommon.Send("mcs006a", wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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
                                                                 "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
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
                   System.Configuration.ConfigurationManager.AppSettings("MPMProcessesFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String


        SQl = "Insert into F_MPMProcessesSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate,APPPER,Clinter,DivisionCode,"                  '1~5
        SQl = SQl + "Division,Type1,Type2,MapNo,Qty,AQty,Code,"                  '6~10
        SQl = SQl + "Weight,Material,DevNo,Manufout, FinishDate,"
        If wStep = 40 Then
            SQl = SQl + "AFinishDate, "
        End If

        SQl = SQl + " MapFile, Aranger, Arang, "                  '11~13"
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '004001', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����
        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Value + "', "                '���
        SQl = SQl + " '" + DAppper.Text + "', "   '����
        SQl = SQl + " '" + DClinter.Text + "', "     '���
        SQl = SQl + " '" + Mid(DDivisionCode.SelectedValue, 1, 7) + "', "  'buyer
        SQl = SQl + " '" + Mid(DDivisionCode.SelectedValue, 9, Len(DDivisionCode.SelectedValue) - 1) + "',"
        SQl = SQl + " '" + DType1.SelectedValue + "', "  'buyer
        SQl = SQl + " '" + DType2.SelectedValue + "', "  'buyer
        SQl = SQl + " '" + DMapNo.Text + "', "  'buyer
        SQl = SQl + " '" + DQty.Text + "', "  'buyer
        SQl = SQl + " '" + DAQty.Text + "', "  'buyer
        SQl = SQl + " '" + DCode.Text + "', "  'buyer
        SQl = SQl + " '" + DWeight.Text + "', "  'buyer
        SQl = SQl + " '" + DMaterial.Text + "', "  'buyer
        SQl = SQl + " '" + DDevNo.Text + "', "  'buyer
        SQl = SQl + " '" + DManufout.Text + "', "  'buyer


        If DFinishDate.Value = "" Then
            SQl = SQl + " '', "  'buyer
        Else
            SQl = SQl + " '" + DFinishDate.Value + "', "  'buyer
        End If

        If wStep = 40 Then
            If DAFinishDate.Value = "" Then
                SQl = SQl + " '', "  'buyer
            Else
                SQl = SQl + " '" + DAFinishDate.Value + "', "  'buyer
            End If


        End If



        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)

                DMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + FileName + "'," '�~��̿�� 
        SQl = SQl + " '', "  'Aranger
        SQl = SQl + " 0," 'Arange
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
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                   System.Configuration.ConfigurationManager.AppSettings("MPMProcessesFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim SQL1 As String = ""
        Dim FileName As String = ""
        Dim i As Integer


        If wStep = 15 Then
            SQL1 = "Insert into F_MPMProcessesSheetDT "
            SQL1 = SQL1 + "( "
            SQL1 = SQL1 + "FormSno,Seqno,No,Engine)"
            SQL1 = SQL1 + "Values('" + Str(wFormSno) + "','1','" + YKK.ReplaceString(DNo.Text) + "','CAM')"
            uDataBase.ExecuteNonQuery(SQL1)

        End If


        '�[�J�u�{�Ӷ�
        Dim EngineStr As String

        If wStep = 10 Or wStep = 20 Then
            'SQL1 = "Delete F_MPMProcessesSheetDT "
            'SQL1 = SQL1 + " where Formsno = '" + Str(wFormSno) + "'"
            'SQL1 = SQL1 + " and recdate is   null"
            'uDataBase.ExecuteNonQuery(SQL1)

            '20221104 Jessica �p�G�Ĥ@������O���N�q2�}�l
            Dim seqno As Integer = 0
            SQL1 = " select isnull(max(seqno),'0')Data from  F_MPMProcessesSheetDT "
            SQL1 = SQL1 + " where Formsno = '" + Str(wFormSno) + "'"

            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL1)
            If dt.Rows.Count > 0 Then
                seqno = dt.Rows(0)("Data")
            End If





            For i = 1 + seqno To 14
                EngineStr = "DEngine" + Trim(Str(i))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                If DText.Text <> "" Then

                    SQL1 = "Insert into F_MPMProcessesSheetDT "
                    SQL1 = SQL1 + "( "
                    SQL1 = SQL1 + "FormSno,Seqno,No,Engine)"
                    SQL1 = SQL1 + "Values('" + Str(wFormSno) + "',N'" + Str(i) + "','" + YKK.ReplaceString(DNo.Text) + "','" + DText.Text + "')"
                    uDataBase.ExecuteNonQuery(SQL1)

                    ' ��s�w�Ƥu�{�H��()

                End If

            Next

        End If


  




        SQL1 = " Update F_MPMProcessesSheet"
        SQL1 = SQL1 + " Set "
        If pFun <> "SAVE" Then
            SQL1 = SQL1 + " Sts = '" & pSts & "',"
            SQL1 = SQL1 + " CompletedTime = '" & NowDateTime & "',"
        End If

        SQL1 = SQL1 + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        SQL1 = SQL1 + " AppDate = '" & DAppDate.Value & "',"
        SQL1 = SQL1 + " Appper = '" & DAppper.Text & "',"
        SQL1 = SQL1 + " Clinter = '" & DClinter.Text & "',"
        SQL1 = SQL1 + " DivisionCode = '" & Mid(DDivisionCode.SelectedValue, 1, 7) & "',"
        SQL1 = SQL1 + " Division = '" & Mid(DDivisionCode.SelectedValue, 9, Len(DDivisionCode.SelectedValue) - 1) & "',"
        SQL1 = SQL1 + " Type1 = '" & DType1.SelectedValue & "',"
        SQL1 = SQL1 + " Type2 = '" & DType2.SelectedValue & "',"
        SQL1 = SQL1 + " MapNo = '" & DMapNo.Text & "',"
        SQL1 = SQL1 + " Qty = '" & Trim(DQty.Text) & "',"
        SQL1 = SQL1 + " AQty = '" & AQty & "',"
        SQL1 = SQL1 + " Code = '" & DCode.Text & "',"
        SQL1 = SQL1 + " Weight = '" & DWeight.Text & "',"
        SQL1 = SQL1 + " Material = '" & DMaterial.Text & "',"
        SQL1 = SQL1 + " DevNo = '" & DevNo & "',"
        SQL1 = SQL1 + " Manufout = '" & Manufout & "',"

        SQL1 = SQL1 + " FinishDate = '" & DFinishDate.Value & "',"
        If wStep = 40 Then
            SQL1 = SQL1 + " AFinishDate = '" & DAFinishDate.Value & "',"
        End If


       
        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                DMapFile.PostedFile.SaveAs(Path + FileName)
                SQL1 = SQL1 + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '�~��̿��
            Else
                FileName = ""
            End If
        End If

        If wStep = 10 Or wStep = 20 Or wStep = 30 Then
            SQL1 = SQL1 + " Aranger = '" & Request.QueryString("pUserID") & "',"
        End If

        SQL1 = SQL1 + " Arang = 1" & ","
        SQL1 = SQL1 + " ModifyUser = '" & wUserID & "',"
        SQL1 = SQL1 + " ModifyTime = '" & NowDateTime & "' "

        SQL1 = SQL1 + " Where Formsno ='" + Str(wFormSno) + "'"
        uDataBase.ExecuteNonQuery(SQL1)






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

        If wStep = 40 Then
            Top = Top - 250
        End If

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

        If wStep = 40 Then
            Top = Top - 100
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
        '  Str2 = CStr(DateTime.Now.Month)  '��
        '  For i = Str2.Length To 1
        ' Str2 = "0" + Str2
        ' Next
        ' Str = Str2
        ' Str2 = CStr(DateTime.Now.Day)    '��
        'For i = Str2.Length To 1
        'Str2 = "0" + Str2
        'Next
        'Str = Str + Str2
        Str = CStr(DateTime.Now.Year)  '�~
        'Set�渹
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function

    Protected Sub DEngine_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DEngine.SelectedIndexChanged
        Dim i As Integer
        i = 30
        BEngineAdd.Visible = True
        BEngineDel.Visible = True

        '��ܤu�{���e

        If DEngine.SelectedValue = "��W���" Then
            '�i�}�u�{�ﶵ 
            'Engine-A ��W���

            DPanel.Visible = True
            CheckBoxList1.RepeatColumns = 8
            CheckBoxList1.RepeatDirection = RepeatDirection.Horizontal
            CheckBoxList1.Visible = True
            CheckBoxList1.Items.Clear()
            RadioButtonList1.Visible = False
            DPanel.BackColor = Color.Aquamarine


            Dim SQL As String

            SQL = "select * from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + "and dkey ='EngineSelect-��W���'"
            SQL = SQL + " Order by Unique_ID"

            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)

            For Each dtr As Data.DataRow In dt.Rows
                Dim EngineA As New ListItem(dtr("Data"), dtr("Data"))
                'New ListItem(dtr("RUserName"), dtr("RUserID"))
                CheckBoxList1.Items.Add(EngineA)
            Next
            i = 90
            '��ܰ���
            DPanel.Height = i
            Top = 1000
        ElseIf DEngine.SelectedValue <> "" Then

            '�Ҩ�ϡB�s���
            DPanel.Visible = True
            CheckBoxList1.Visible = False
            RadioButtonList1.Visible = True
            If DEngine.SelectedValue = "�s���" Then
                DPanel.BackColor = Color.Violet
                Top = 1200
            Else
                DPanel.BackColor = Color.LightGreen
                Top = 1100
            End If



            '�զX�u�{�r��
            Dim SQL As String


            SQL = "select * from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + "and dkey ='EngineSelect-" + DEngine.SelectedValue + "'"
            SQL = SQL + " Order by Unique_ID"
            Dim dt As Data.DataTable = uDataBase.GetDataTable(SQL)
            RadioButtonList1.Items.Clear()
            For Each dtr As Data.DataRow In dt.Rows
                Dim EngineStr As String = ""
                SQL = "select * from M_referp "
                SQL = SQL + "where cat = '4001'"
                SQL = SQL + "and  dkey = '" + dtr("Data") + "'"
                SQL = SQL + " Order by Unique_ID"
                Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
                For Each dtr1 As Data.DataRow In dt1.Rows
                    If EngineStr = "" Then
                        EngineStr = dtr1("Data")
                    Else
                        EngineStr = EngineStr + "," + dtr1("Data")
                    End If

                Next
                Dim EngineB As New ListItem(dtr("Data") + "�i" + EngineStr + "�j", dtr("Data"))
                'New ListItem(dtr("RUserName"), dtr("RUserID"))
                RadioButtonList1.Items.Add(EngineB)
                i = i + 30
            Next

            DPanel.Height = i
        Else
            BEngineAdd.Visible = False
            BEngineDel.Visible = False
            DPanel.Visible = False
        End If


        SetControlPosition()
        ShowFormData()
    End Sub


    Protected Sub BEngineAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEngineAdd.Click
        Dim i, j, k As Integer
        Dim CheckStr, EngineStr As String

        '��J�u�{
        If DEngine1.Text = "CAM" Then
            k = 2
        Else
            k = 1
        End If

        For i = k To 14


            EngineStr = "DEngine" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = ""

        Next


        If DEngine.SelectedValue = "��W���" Then
            j = k

            For i = 0 To (CheckBoxList1.Items.Count - 1)
                If (CheckBoxList1.Items(i).Selected) Then
                    CheckStr = Trim(CheckBoxList1.Items(i).Text)

                    EngineStr = "DEngine" + Trim(Str(j))
                    Dim DText As TextBox = Me.FindControl(EngineStr)
                    'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)
                    If DText IsNot Nothing Then
                        DText.Text = CheckStr
                    End If
                    j = j + 1
                End If

            Next

            Top = 1000
        Else
            CheckStr = ""
            For i = 0 To (RadioButtonList1.Items.Count - 1)
                If RadioButtonList1.Items(i).Selected Then
                    CheckStr = Trim(RadioButtonList1.Items(i).Value)
                End If
            Next

            Dim SQL As String
            SQL = "select Data from M_referp "
            SQL = SQL + "where cat = '4001'"
            SQL = SQL + " and dkey = '" + CheckStr + "'"
            SQL = SQL + " Order by Unique_ID"
            Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
            k = k - 1
            For i = 0 To dt1.Rows.Count - 1

                CheckStr = dt1.Rows(i).Item("Data")
                EngineStr = "DEngine" + Trim(Str(i + 1 + k))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                'Dim DText As TextBox = mDirectCast(Page.FindControl("TextBox1"), TextBox)
                If DText IsNot Nothing Then
                    DText.Text = CheckStr
                End If
            Next

            If DEngine.SelectedValue = "�s���" Then

                Top = 1200
            Else

                Top = 1100
            End If


        End If

        SetControlPosition()
        ShowFormData()
    End Sub

    Protected Sub BEngineDel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BEngineDel.Click
        '�M�����
        Dim i As Integer
        Dim EngineStr As String
        For i = 1 To 14
            EngineStr = "DEngine" + Trim(Str(i))
            Dim DText As TextBox = Me.FindControl(EngineStr)
            DText.Text = ""

        Next

        For i = 0 To (CheckBoxList1.Items.Count - 1)
            CheckBoxList1.Items(i).Selected = False

        Next


        Top = 1200
        SetControlPosition()
        ShowFormData()
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function OK() As Boolean
        Top = 1200
        SetControlPosition()
        ShowFormData()

        Dim isOK As Boolean = True
        Dim Errcode As Integer = 0



        Dim Message As String = ""


        Dim i, j As Integer
        Dim EngineStr As String

        If wStep = 10 Or wStep = 20 Or wStep = 30 Then
            j = 0
            For i = 1 To 14
                EngineStr = "DEngine" + Trim(Str(i))
                Dim DText As TextBox = Me.FindControl(EngineStr)
                If DText.Text <> "" Then
                    j = j + 1
                End If


            Next

            If j = 0 Then
                isOK = False
                Message = "���`�G�ݿ�J�u�{!"
            End If

        End If
     
        If wStep = 40 Then '�ˬd�O�_�i����
            Dim SQL As String
            SQL = " select * from  F_MPMProcessesSheetdt "
            SQL = SQL + "  where  no ='" + DNo.Text + "'"
            SQL = SQL + " and engine <> '' "
            SQL = SQL + " and (starter ='' "
            SQL = SQL + " or ender  ='') "
            SQL = SQL + " and delmark = 0 "
            Dim dt1 As Data.DataTable = uDataBase.GetDataTable(SQL)
            If dt1.Rows.Count > 0 Then
                j = 1
            End If

            'SQL = " select * from  F_MPMProcessesSheetdt "
            'SQL = SQL + "  where  no ='" + DNo.Text + "'"
            'SQL = SQL + "and ender  <>'' "
            'SQL = SQL + "and delmark = 0 "
            'Dim dt3 As Data.DataTable = uDataBase.GetDataTable(SQL)
            'If dt3.Rows.Count > 0 Then
            '    j = 0
            'End If


            SQL = " select * from  F_MPMProcessesSheetdt "
            SQL = SQL + "  where  no ='" + DNo.Text + "'"
            SQL = SQL + "and ( AlwaysStoper <>'')"
            SQL = SQL + "and delmark = 0 "
            Dim dt2 As Data.DataTable = uDataBase.GetDataTable(SQL)
            If dt2.Rows.Count > 0 Then
                j = 0
            End If



            If j = 1 Then
                isOK = False
                Message = "���`�G�|���u�{�������A���i����!"
            End If

        End If

    

     
        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If


        Return isOK

      
    End Function

 
 

    Protected Sub DAppper_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAppper.TextChanged

    End Sub

    Protected Sub DNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.TextChanged

    End Sub
End Class

