Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SCD_SampleSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSampleSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DSampleSheet2 As System.Web.UI.WebControls.Image
    Protected WithEvents DDEVPRD As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDEVNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTAWIDTH As System.Web.UI.WebControls.TextBox
    Protected WithEvents LSAMPLEFILE As System.Web.UI.WebControls.Image
    Protected WithEvents DCODENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSIZENO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTACOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DECOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTHCOL As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSAMPLEFILE As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DTNLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDLITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCNITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTNRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTSRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTDRITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCSITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCDITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF5 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF6 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF7 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DTALINE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAPPBUYER As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDATE As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWF3Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF4Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF5Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF6Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DWF7Name As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BRNO As System.Web.UI.WebControls.Button
    Protected WithEvents DRNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO2ITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DO1ITEM As System.Web.UI.WebControls.TextBox
    Protected WithEvents D1Other As System.Web.UI.WebControls.Label
    Protected WithEvents D2Other As System.Web.UI.WebControls.Label
    Protected WithEvents LQCFILE1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LQCFILE4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DQCFILE1 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE3 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE2 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents DQCFILE4 As System.Web.UI.HtmlControls.HtmlInputFile
    
    Protected WithEvents Hidden1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Hidden2 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DQCFILE5 As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents LQCFILE5 As System.Web.UI.WebControls.HyperLink

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
    'Dim wDepo As String = "CL"      '�x�_��ƾ�(CL->���c, TP->�x�_)
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

        Response.Cookies("CommissionNo").Value = ""    'CodeNo, DevelopDataPicker�ϥ�

        Response.Cookies("PGM").Value = "SCD_SampleSheet_01.aspx"
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

        '��ڼ˫~
        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "��ڼ˫~"
                Else
                    Message = Message & ", " & "��ڼ˫~"
                End If
            End If
        End If

        'update by alin
        '�~�����i1
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~�����i"
                Else
                    Message = Message & ", " & "�~�����i"
                End If
            End If
        End If
        '�~�����i2
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~�����i"
                Else
                    Message = Message & ", " & "�~�����i"
                End If
            End If
        End If
        '�~�����i3
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~�����i"
                Else
                    Message = Message & ", " & "�~�����i"
                End If
            End If
        End If
        '�~�����i4
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~�����i"
                Else
                    Message = Message & ", " & "�~�����i"
                End If
            End If
        End If

        '�~�����i5
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "�~�����i"
                Else
                    Message = Message & ", " & "�~�����i"
                End If
            End If
        End If
        '--------------------------------------------------


        If Message <> "" Then
            Message = "�U�C�ҳ]�w�����[�ɮױN�|�� (" & Message & ") " & ",�Э��s�]�w!"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SampleFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_SampleSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SampleSheet")
        If DBDataSet1.Tables("F_SampleSheet").Rows.Count > 0 Then
            '�����
            DRNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Rno")                   'Commission-No
            DNo.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("No")                     'No
            DAPPBUYER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("AppBuyer")         'Customer
            DDATE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Date")                 '�o���
            DSIZENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SizeNo")             'Size
            DITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Item")                 'Item
            DCODENO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CodeNo")             'Code No
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile") <> "" Then          '��ڼ˫~
                LSAMPLEFILE.ImageUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("SampleFile")
            Else
                LSAMPLEFILE.Visible = False
            End If
            DTAWIDTH.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TAWidth")           '���a�e��
            DDEVNO.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevNo")               '�}�oNo
            DDEVPRD.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("DevPrd")             '�}�o����
            DTACOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TACol")               '���aColor
            DTALINE.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TALine")             '�����uColor
            DECOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("ECol")                 '�Ⱦ�
            DCCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CCol")                 '�Y��
            DTHCOL.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("THCol")               '�_�u�u
            DOTHER.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other")               '��L

            'update by alin
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1") <> "" Then              '�~�����i1
                LQCFILE1.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile1")
            Else
                LQCFILE1.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2") <> "" Then              '�~�����i2
                LQCFILE2.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile2")
            Else
                LQCFILE2.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3") <> "" Then              '�~�����i3
                LQCFILE3.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile3")
            Else
                LQCFILE3.Visible = False
            End If
            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4") <> "" Then              '�~�����i4
                LQCFILE4.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile4")
            Else
                LQCFILE4.Visible = False
            End If

            If DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5") <> "" Then              '�~�����i5
                LQCFILE5.NavigateUrl = Path & DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("QCFile5")
            Else
                LQCFILE5.Visible = False
            End If

            D1Other.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other1")           'Other1
            Hidden1.Value = D1Other.Text
            D2Other.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("Other2")           'Other2
            Hidden2.Value = D2Other.Text
            DO1ITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O1Item")           'O1Item
            DO2ITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("O2Item")           'O2Item

            '----------------------------------

            DTNLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNLItem")           'TAPE NAT-��
            DTNRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TNRItem")           'TAPE NAT-�k
            DTSLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSLItem")           'TAPE SET-��
            DTSRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TSRItem")           'TAPE SET-�k
            DTDLITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDLItem")           'TAPE DYED-��
            DTDRITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("TDRItem")           'TAPE DYED-�k
            DCNITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CNItem")             'CHAIN NAT
            DCSITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CSItem")             'CHAIN SET
            DCDITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CDItem")             'CHAIN DYED
            DCITEM.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("CItem")               'CORD
            DOP1.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP1")                   '�u�{1
            DOP2.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP2")                   '�u�{2
            DOP3.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP3")                   '�u�{3
            DOP4.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP4")                   '�u�{4
            DOP5.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP5")                   '�u�{5
            DOP6.Text = DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("OP6")                   '�u�{6
            SetFieldData("WF1", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF1"))          '�ӻ{WF1
            SetFieldData("WF2", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF2"))          '�ӻ{WF2
            SetFieldData("WF3", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3"))          '�ӻ{WF3
            SetFieldData("WF4", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4"))          '�ӻ{WF4
            SetFieldData("WF5", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5"))          '�ӻ{WF5
            SetFieldData("WF6", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6"))          '�ӻ{WF6
            SetFieldData("WF7", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7"))          '�ӻ{WF7
            SetFieldData("WF3NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF3Name"))      '�ӻ{�̳���WF3
            SetFieldData("WF4NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF4Name"))      '�ӻ{�̳���WF4
            SetFieldData("WF5NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF5Name"))      '�ӻ{�̳���WF5
            SetFieldData("WF6NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF6Name"))      '�ӻ{�̳���WF6
            SetFieldData("WF7NAME", DBDataSet1.Tables("F_SampleSheet").Rows(0).Item("WF7Name"))      '�ӻ{�̳���WF7
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
            SQL = SQL + "Order by Unique_ID Desc "
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
                DSampleSheet1.Visible = True     '���Sheet-1
                DSampleSheet2.Visible = True     '���Sheet-2
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
                    Top = 1285
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 1157
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
                LSAMPLEFILE.Visible = True     'Sample File

                'update by alin
                LQCFILE1.Visible = True         'QC File1
                LQCFILE2.Visible = True         'QC File2
                LQCFILE3.Visible = True         'QC File3
                LQCFILE4.Visible = True         'QC File4
                LQCFILE5.Visible = True         'QC File5
                '---------------------------

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
            Top = 1085
            'Sheet����
            DDescSheet.Visible = False  '����Sheet
            DDelay.Visible = False      '����Sheet
            '�������
            DDecideDesc.Visible = False    '����
            DReasonCode.Visible = False    '����z�ѥN�X
            DReason.Visible = False        '����z��
            DReasonDesc.Visible = False    '�����L����
            DHistoryLabel.Visible = False  '�֩w�i��
            '�s������
            LSAMPLEFILE.Visible = False    'Sample File

            'update by alin
            LQCFILE1.Visible = False        'QC File1
            LQCFILE2.Visible = False        'QC File2
            LQCFILE3.Visible = False        'QC File3
            LQCFILE4.Visible = False        'QC File4
            LQCFILE5.Visible = False        'QC File5
            '------------------------------

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
        '��ƿ��
        BRNO.Attributes("onclick") = "DataPicker('Develop');"
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
                    Top = 1157
                Else
                    If DDelay.Visible = True Then
                        Top = 1285
                    Else
                        Top = 1157
                    End If
                End If
            End If
        Else
            Top = 1085
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
        '-Commission-No�S�O�B�z-------------------------------------------------------------------
        If pPost = "New" Then
            DRNO.Text = ""
            DRNO.ForeColor = Color.White
            DRNO.BackColor = Color.White
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
        'AppBuyer
        Select Case FindFieldInf("AppBuyer")
            Case 0  '���
                DAPPBUYER.BackColor = Color.LightGray
                DAPPBUYER.ReadOnly = True
                DAPPBUYER.Visible = True
            Case 1  '�ק�+�ˬd
                DAPPBUYER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAppBuyerRqd", "DAppBuyer", "���`�G�ݿ�JCustomer")
                DAPPBUYER.Visible = True
            Case 2  '�ק�
                DAPPBUYER.BackColor = Color.Yellow
                DAPPBUYER.Visible = True
            Case Else   '����
                DAPPBUYER.Visible = False
        End Select
        If pPost = "New" Then DAPPBUYER.Text = ""
        '�o���
        Select Case FindFieldInf("Date")
            Case 0  '���
                DDATE.BackColor = Color.LightGray
                DDATE.ReadOnly = True
                DDATE.Visible = True
            Case 1  '�ק�+�ˬd
                DDATE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "���`�G�ݿ�J�o���")
                DDATE.Visible = True
            Case 2  '�ק�
                DDATE.BackColor = Color.Yellow
                DDATE.Visible = True
            Case Else   '����
                DDATE.Visible = False
        End Select
        If pPost = "New" Then DDATE.Text = CStr(DateTime.Now.Today)
        'Size
        Select Case FindFieldInf("SizeNo")
            Case 0  '���
                DSIZENO.BackColor = Color.LightGray
                DSIZENO.ReadOnly = True
                DSIZENO.Visible = True
            Case 1  '�ק�+�ˬd
                DSIZENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSizeNoRqd", "DSizeNo", "���`�G�ݿ�JSize")
                DSIZENO.Visible = True
            Case 2  '�ק�
                DSIZENO.BackColor = Color.Yellow
                DSIZENO.Visible = True
            Case Else   '����
                DSIZENO.Visible = False
        End Select
        If pPost = "New" Then DSIZENO.Text = ""
        'Item
        Select Case FindFieldInf("Item")
            Case 0  '���
                DITEM.BackColor = Color.LightGray
                DITEM.ReadOnly = True
                DITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DItemRqd", "DItem", "���`�G�ݿ�JItem")
                DITEM.Visible = True
            Case 2  '�ק�
                DITEM.BackColor = Color.Yellow
                DITEM.Visible = True
            Case Else   '����
                DITEM.Visible = False
        End Select
        If pPost = "New" Then DITEM.Text = ""
        'Tape
        Select Case FindFieldInf("CodeNo")
            Case 0  '���
                DCODENO.BackColor = Color.LightGray
                DCODENO.ReadOnly = True
                DCODENO.Visible = True
            Case 1  '�ק�+�ˬd
                DCODENO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCodeNoRqd", "DCodeNo", "���`�G�ݿ�JTape")
                DCODENO.Visible = True
            Case 2  '�ק�
                DCODENO.BackColor = Color.Yellow
                DCODENO.Visible = True
            Case Else   '����
                DCODENO.Visible = False
        End Select
        If pPost = "New" Then DCODENO.Text = ""
        '��ڼ˫~
        Select Case FindFieldInf("SampleFile")
            Case 0  '���
                DSAMPLEFILE.Visible = False
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "���`�G�ݿ�J��ڼ˫~��")
                DSAMPLEFILE.Visible = True
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DSAMPLEFILE.Visible = True
                DSAMPLEFILE.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DSAMPLEFILE.Visible = False
        End Select
        '���a�e��
        Select Case FindFieldInf("TAWidth")
            Case 0  '���
                DTAWIDTH.BackColor = Color.LightGray
                DTAWIDTH.ReadOnly = True
                DTAWIDTH.Visible = True
            Case 1  '�ק�+�ˬd
                DTAWIDTH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTAWidthRqd", "DTAWidth", "���`�G�ݿ�J���a�e��")
                DTAWIDTH.Visible = True
            Case 2  '�ק�
                DTAWIDTH.BackColor = Color.Yellow
                DTAWIDTH.Visible = True
            Case Else   '����
                DTAWIDTH.Visible = False
        End Select
        If pPost = "New" Then DTAWIDTH.Text = ""
        '�}�oNo
        Select Case FindFieldInf("DevNo")
            Case 0  '���
                DDEVNO.BackColor = Color.LightGray
                DDEVNO.ReadOnly = True
                DDEVNO.Visible = True
            Case 1  '�ק�+�ˬd
                DDEVNO.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevNoRqd", "DDevNo", "���`�G�ݿ�J�}�oNo")
                DDEVNO.Visible = True
            Case 2  '�ק�
                DDEVNO.BackColor = Color.Yellow
                DDEVNO.Visible = True
            Case Else   '����
                DDEVNO.Visible = False
        End Select
        If pPost = "New" Then DDEVNO.Text = ""
        '�}�o����
        Select Case FindFieldInf("DevPrd")
            Case 0  '���
                DDEVPRD.BackColor = Color.LightGray
                DDEVPRD.ReadOnly = True
                DDEVPRD.Visible = True
            Case 1  '�ק�+�ˬd
                DDEVPRD.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDevPrdRqd", "DDevPrd", "���`�G�ݿ�J�}�o����")
                DDEVPRD.Visible = True
            Case 2  '�ק�
                DDEVPRD.BackColor = Color.Yellow
                DDEVPRD.Visible = True
            Case Else   '����
                DDEVPRD.Visible = False
        End Select
        If pPost = "New" Then DDEVPRD.Text = ""
        '���a
        Select Case FindFieldInf("TACol")
            Case 0  '���
                DTACOL.BackColor = Color.LightGray
                DTACOL.ReadOnly = True
                DTACOL.Visible = True
            Case 1  '�ק�+�ˬd
                DTACOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTAColRqd", "DTACol", "���`�G�ݿ�J���a")
                DTACOL.Visible = True
            Case 2  '�ק�
                DTACOL.BackColor = Color.Yellow
                DTACOL.Visible = True
            Case Else   '����
                DTACOL.Visible = False
        End Select
        If pPost = "New" Then DTACOL.Text = ""
        '�����u
        Select Case FindFieldInf("TALine")
            Case 0  '���
                DTALINE.BackColor = Color.LightGray
                DTALINE.ReadOnly = True
                DTALINE.Visible = True
            Case 1  '�ק�+�ˬd
                DTALINE.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTALineRqd", "DTALine", "���`�G�ݿ�J�����u")
                DTALINE.Visible = True
            Case 2  '�ק�
                DTALINE.BackColor = Color.Yellow
                DTALINE.Visible = True
            Case Else   '����
                DTALINE.Visible = False
        End Select
        If pPost = "New" Then DTALINE.Text = ""
        '�Ⱦ�
        Select Case FindFieldInf("ECol")
            Case 0  '���
                DECOL.BackColor = Color.LightGray
                DECOL.ReadOnly = True
                DECOL.Visible = True
            Case 1  '�ק�+�ˬd
                DECOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEColRqd", "DECol", "���`�G�ݿ�J�Ⱦ�")
                DECOL.Visible = True
            Case 2  '�ק�
                DECOL.BackColor = Color.Yellow
                DECOL.Visible = True
            Case Else   '����
                DECOL.Visible = False
        End Select
        If pPost = "New" Then DECOL.Text = ""
        '�Y��
        Select Case FindFieldInf("CCol")
            Case 0  '���
                DCCOL.BackColor = Color.LightGray
                DCCOL.ReadOnly = True
                DCCOL.Visible = True
            Case 1  '�ק�+�ˬd
                DCCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCColRqd", "DCCol", "���`�G�ݿ�J�Y��")
                DCCOL.Visible = True
            Case 2  '�ק�
                DCCOL.BackColor = Color.Yellow
                DCCOL.Visible = True
            Case Else   '����
                DCCOL.Visible = False
        End Select
        If pPost = "New" Then DCCOL.Text = ""
        '�_�u�u
        Select Case FindFieldInf("THCol")
            Case 0  '���
                DTHCOL.BackColor = Color.LightGray
                DTHCOL.ReadOnly = True
                DTHCOL.Visible = True
            Case 1  '�ק�+�ˬd
                DTHCOL.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTHColRqd", "DTHCol", "���`�G�ݿ�J�_�u�u")
                DTHCOL.Visible = True
            Case 2  '�ק�
                DTHCOL.BackColor = Color.Yellow
                DTHCOL.Visible = True
            Case Else   '����
                DTHCOL.Visible = False
        End Select
        If pPost = "New" Then DTHCOL.Text = ""
        '��L
        Select Case FindFieldInf("Other")
            Case 0  '���
                DOTHER.BackColor = Color.LightGray
                DOTHER.ReadOnly = True
                DOTHER.Visible = True
            Case 1  '�ק�+�ˬd
                DOTHER.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOtherRqd", "DOther", "���`�G�ݿ�J��L")
                DOTHER.Visible = True
            Case 2  '�ק�
                DOTHER.BackColor = Color.Yellow
                DOTHER.Visible = True
            Case Else   '����
                DOTHER.Visible = False
        End Select
        If pPost = "New" Then DOTHER.Text = ""

        'update by alin
        '�~�����i1
        Select Case FindFieldInf("QCFile1")
            Case 0  '���
                DQCFILE1.Visible = False
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFileRqd1", "DQCFile1", "���`�G�ݿ�J�~�����i��1")
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFILE1.Visible = True
                DQCFILE1.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFILE1.Visible = False
        End Select
        '�~�����i2
        Select Case FindFieldInf("QCFile2")
            Case 0  '���
                DQCFILE2.Visible = False
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFileRqd2", "DQCFile2", "���`�G�ݿ�J�~�����i��2")
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFILE2.Visible = True
                DQCFILE2.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFILE2.Visible = False
        End Select
        '�~�����i3
        Select Case FindFieldInf("QCFile3")
            Case 0  '���
                DQCFILE3.Visible = False
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFileRqd3", "DQCFile3", "���`�G�ݿ�J�~�����i��3")
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFILE3.Visible = True
                DQCFILE3.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFILE3.Visible = False
        End Select
        '�~�����i4
        Select Case FindFieldInf("QCFile4")
            Case 0  '���
                DQCFILE4.Visible = False
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFileRqd4", "DQCFile4", "���`�G�ݿ�J�~�����i��4")
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFILE4.Visible = True
                DQCFILE4.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFILE4.Visible = False
        End Select

        '�~�����i5
        Select Case FindFieldInf("QCFile5")
            Case 0  '���
                DQCFILE5.Visible = False
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCFileRqd5", "DQCFile5", "���`�G�ݿ�J�~�����i��4")
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCFILE5.Visible = True
                DQCFILE5.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCFILE5.Visible = False
        End Select
        '--------------------------------

        'Tape Nat-��
        Select Case FindFieldInf("TNLItem")
            Case 0  '���
                DTNLITEM.BackColor = Color.LightGray
                DTNLITEM.ReadOnly = True
                DTNLITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTNLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTNLItemRqd", "DTNLItem", "���`�G�ݿ�JTape Nat-��")
                DTNLITEM.Visible = True
            Case 2  '�ק�
                DTNLITEM.BackColor = Color.Yellow
                DTNLITEM.Visible = True
            Case Else   '����
                DTNLITEM.Visible = False
        End Select
        If pPost = "New" Then DTNLITEM.Text = ""
        'Tape Nat-�k
        Select Case FindFieldInf("TNRItem")
            Case 0  '���
                DTNRITEM.BackColor = Color.LightGray
                DTNRITEM.ReadOnly = True
                DTNRITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTNRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTNRItemRqd", "DTNRItem", "���`�G�ݿ�JTape Nat-�k")
                DTNRITEM.Visible = True
            Case 2  '�ק�
                DTNRITEM.BackColor = Color.Yellow
                DTNRITEM.Visible = True
            Case Else   '����
                DTNRITEM.Visible = False
        End Select
        If pPost = "New" Then DTNRITEM.Text = ""
        'Tape Set-��
        Select Case FindFieldInf("TSLItem")
            Case 0  '���
                DTSLITEM.BackColor = Color.LightGray
                DTSLITEM.ReadOnly = True
                DTSLITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTSLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTSLItemRqd", "DTSLItem", "���`�G�ݿ�JTape Set-��")
                DTSLITEM.Visible = True
            Case 2  '�ק�
                DTSLITEM.BackColor = Color.Yellow
                DTSLITEM.Visible = True
            Case Else   '����
                DTSLITEM.Visible = False
        End Select
        If pPost = "New" Then DTSLITEM.Text = ""
        'Tape Set-�k
        Select Case FindFieldInf("TSRItem")
            Case 0  '���
                DTSRITEM.BackColor = Color.LightGray
                DTSRITEM.ReadOnly = True
                DTSRITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTSRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTSRItemRqd", "DTSRItem", "���`�G�ݿ�JTape Set-�k")
                DTSRITEM.Visible = True
            Case 2  '�ק�
                DTSRITEM.BackColor = Color.Yellow
                DTSRITEM.Visible = True
            Case Else   '����
                DTSRITEM.Visible = False
        End Select
        If pPost = "New" Then DTSRITEM.Text = ""
        'Tape Dyed-��
        Select Case FindFieldInf("TDLItem")
            Case 0  '���
                DTDLITEM.BackColor = Color.LightGray
                DTDLITEM.ReadOnly = True
                DTDLITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTDLITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTDLItemRqd", "DTDLItem", "���`�G�ݿ�JTape Dyed-��")
                DTDLITEM.Visible = True
            Case 2  '�ק�
                DTDLITEM.BackColor = Color.Yellow
                DTDLITEM.Visible = True
            Case Else   '����
                DTDLITEM.Visible = False
        End Select
        If pPost = "New" Then DTDLITEM.Text = ""
        'Tape Dyed-�k
        Select Case FindFieldInf("TDRItem")
            Case 0  '���
                DTDRITEM.BackColor = Color.LightGray
                DTDRITEM.ReadOnly = True
                DTDRITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DTDRITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTDRItemRqd", "DTDRItem", "���`�G�ݿ�JTape Dyed-�k")
                DTDRITEM.Visible = True
            Case 2  '�ק�
                DTDRITEM.BackColor = Color.Yellow
                DTDRITEM.Visible = True
            Case Else   '����
                DTDRITEM.Visible = False
        End Select
        If pPost = "New" Then DTDRITEM.Text = ""
        'Chain Nat
        Select Case FindFieldInf("CNItem")
            Case 0  '���
                DCNITEM.BackColor = Color.LightGray
                DCNITEM.ReadOnly = True
                DCNITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DCNITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCNItemRqd", "DCNItem", "���`�G�ݿ�JChain Nat")
                DCNITEM.Visible = True
            Case 2  '�ק�
                DCNITEM.BackColor = Color.Yellow
                DCNITEM.Visible = True
            Case Else   '����
                DCNITEM.Visible = False
        End Select
        If pPost = "New" Then DCNITEM.Text = ""
        'Chain Set
        Select Case FindFieldInf("CSItem")
            Case 0  '���
                DCSITEM.BackColor = Color.LightGray
                DCSITEM.ReadOnly = True
                DCSITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DCSITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCSItemRqd", "DCSItem", "���`�G�ݿ�JChain Set")
                DCSITEM.Visible = True
            Case 2  '�ק�
                DCSITEM.BackColor = Color.Yellow
                DCSITEM.Visible = True
            Case Else   '����
                DCSITEM.Visible = False
        End Select
        If pPost = "New" Then DCSITEM.Text = ""
        'Chain Dyed
        Select Case FindFieldInf("CDItem")
            Case 0  '���
                DCDITEM.BackColor = Color.LightGray
                DCDITEM.ReadOnly = True
                DCDITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DCDITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCDItemRqd", "DCDItem", "���`�G�ݿ�JChain Dyed")
                DCDITEM.Visible = True
            Case 2  '�ק�
                DCDITEM.BackColor = Color.Yellow
                DCDITEM.Visible = True
            Case Else   '����
                DCDITEM.Visible = False
        End Select

        'update by alin
        'O1ITEM
        Select Case FindFieldInf("O1Item")
            Case 0  '���
                DO1ITEM.BackColor = Color.LightGray
                DO1ITEM.ReadOnly = True
                DO1ITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DO1ITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DO1ItemRqd", "DO1Item", "���`�G�ݿ�JOther Item1")
                DO1ITEM.Visible = True
            Case 2  '�ק�
                DO1ITEM.BackColor = Color.Yellow
                DO1ITEM.Visible = True
            Case Else   '����
                DO1ITEM.Visible = False
        End Select

        'O2ITEM
        Select Case FindFieldInf("O2Item")
            Case 0  '���
                DO2ITEM.BackColor = Color.LightGray
                DO2ITEM.ReadOnly = True
                DO2ITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DO2ITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DO2ItemRqd", "DO2Item", "���`�G�ݿ�JOther Item2")
                DO2ITEM.Visible = True
            Case 2  '�ק�
                DO2ITEM.BackColor = Color.Yellow
                DO2ITEM.Visible = True
            Case Else   '����
                DO2ITEM.Visible = False
        End Select
        '----------------------------------


        If pPost = "New" Then DCDITEM.Text = ""
        'Cord
        Select Case FindFieldInf("CItem")
            Case 0  '���
                DCITEM.BackColor = Color.LightGray
                DCITEM.ReadOnly = True
                DCITEM.Visible = True
            Case 1  '�ק�+�ˬd
                DCITEM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCItemRqd", "DCItem", "���`�G�ݿ�JCord")
                DCITEM.Visible = True
            Case 2  '�ק�
                DCITEM.BackColor = Color.Yellow
                DCITEM.Visible = True
            Case Else   '����
                DCITEM.Visible = False
        End Select
        If pPost = "New" Then DCITEM.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP1")
            Case 0  '���
                DOP1.BackColor = Color.LightGray
                DOP1.ReadOnly = True
                DOP1.Visible = True
            Case 1  '�ק�+�ˬd
                DOP1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP1Rqd", "DOP1", "���`�G�ݿ�J�u�{��")
                DOP1.Visible = True
            Case 2  '�ק�
                DOP1.BackColor = Color.Yellow
                DOP1.Visible = True
            Case Else   '����
                DOP1.Visible = False
        End Select
        If pPost = "New" Then DOP1.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP2")
            Case 0  '���
                DOP2.BackColor = Color.LightGray
                DOP2.ReadOnly = True
                DOP2.Visible = True
            Case 1  '�ק�+�ˬd
                DOP2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP2Rqd", "DOP2", "���`�G�ݿ�J�u�{��")
                DOP2.Visible = True
            Case 2  '�ק�
                DOP2.BackColor = Color.Yellow
                DOP2.Visible = True
            Case Else   '����
                DOP2.Visible = False
        End Select
        If pPost = "New" Then DOP2.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP3")
            Case 0  '���
                DOP3.BackColor = Color.LightGray
                DOP3.ReadOnly = True
                DOP3.Visible = True
            Case 1  '�ק�+�ˬd
                DOP3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP3Rqd", "DOP3", "���`�G�ݿ�J�u�{��")
                DOP3.Visible = True
            Case 2  '�ק�
                DOP3.BackColor = Color.Yellow
                DOP3.Visible = True
            Case Else   '����
                DOP3.Visible = False
        End Select
        If pPost = "New" Then DOP3.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP4")
            Case 0  '���
                DOP4.BackColor = Color.LightGray
                DOP4.ReadOnly = True
                DOP4.Visible = True
            Case 1  '�ק�+�ˬd
                DOP4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP4Rqd", "DOP4", "���`�G�ݿ�J�u�{��")
                DOP4.Visible = True
            Case 2  '�ק�
                DOP4.BackColor = Color.Yellow
                DOP4.Visible = True
            Case Else   '����
                DOP4.Visible = False
        End Select
        If pPost = "New" Then DOP4.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP5")
            Case 0  '���
                DOP5.BackColor = Color.LightGray
                DOP5.ReadOnly = True
                DOP5.Visible = True
            Case 1  '�ק�+�ˬd
                DOP5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP5Rqd", "DOP5", "���`�G�ݿ�J�u�{��")
                DOP5.Visible = True
            Case 2  '�ק�
                DOP5.BackColor = Color.Yellow
                DOP5.Visible = True
            Case Else   '����
                DOP5.Visible = False
        End Select
        If pPost = "New" Then DOP5.Text = ""
        '�u�{��
        Select Case FindFieldInf("OP6")
            Case 0  '���
                DOP6.BackColor = Color.LightGray
                DOP6.ReadOnly = True
                DOP6.Visible = True
            Case 1  '�ק�+�ˬd
                DOP6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOP6Rqd", "DOP6", "���`�G�ݿ�J�u�{��")
                DOP6.Visible = True
            Case 2  '�ק�
                DOP6.BackColor = Color.Yellow
                DOP6.Visible = True
            Case Else   '����
                DOP6.Visible = False
        End Select
        If pPost = "New" Then DOP6.Text = ""
        '�@����
        Select Case FindFieldInf("WF1")
            Case 0  '���
                DWF1.BackColor = Color.LightGray
                DWF1.Visible = True
            Case 1  '�ק�+�ˬd
                DWF1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF1Rqd", "DWF1", "���`�G�ݿ�J�@����")
                DWF1.Visible = True
            Case 2  '�ק�
                DWF1.BackColor = Color.Yellow
                DWF1.Visible = True
            Case Else   '����
                DWF1.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF1", "ZZZZZZ")
        'EA�d����
        Select Case FindFieldInf("WF2")
            Case 0  '���
                DWF2.BackColor = Color.LightGray
                DWF2.Visible = True
            Case 1  '�ק�+�ˬd
                DWF2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF2Rqd", "DWF2", "���`�G�ݿ�JEA�d����")
                DWF2.Visible = True
            Case 2  '�ק�
                DWF2.BackColor = Color.Yellow
                DWF2.Visible = True
            Case Else   '����
                DWF2.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF2", "ZZZZZZ")
        '�s�y��
        Select Case FindFieldInf("WF3")
            Case 0  '���
                DWF3.BackColor = Color.LightGray
                DWF3.Visible = True
            Case 1  '�ק�+�ˬd
                DWF3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF3Rqd", "DWF3", "���`�G�ݿ�J�s�y��")
                DWF3.Visible = True
            Case 2  '�ק�
                DWF3.BackColor = Color.Yellow
                DWF3.Visible = True
            Case Else   '����
                DWF3.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF3", "ZZZZZZ")
        '�s�y��
        Select Case FindFieldInf("WF4")
            Case 0  '���
                DWF4.BackColor = Color.LightGray
                DWF4.Visible = True
            Case 1  '�ק�+�ˬd
                DWF4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF4Rqd", "DWF4", "���`�G�ݿ�J�s�y��")
                DWF4.Visible = True
            Case 2  '�ק�
                DWF4.BackColor = Color.Yellow
                DWF4.Visible = True
            Case Else   '����
                DWF4.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF4", "ZZZZZZ")
        '�s�y��
        Select Case FindFieldInf("WF5")
            Case 0  '���
                DWF5.BackColor = Color.LightGray
                DWF5.Visible = True
            Case 1  '�ק�+�ˬd
                DWF5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF5Rqd", "DWF5", "���`�G�ݿ�J�s�y��")
                DWF5.Visible = True
            Case 2  '�ק�
                DWF5.BackColor = Color.Yellow
                DWF5.Visible = True
            Case Else   '����
                DWF5.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF5", "ZZZZZZ")
        '�s�y��
        Select Case FindFieldInf("WF6")
            Case 0  '���
                DWF6.BackColor = Color.LightGray
                DWF6.Visible = True
            Case 1  '�ק�+�ˬd
                DWF6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF6Rqd", "DWF6", "���`�G�ݿ�J�s�y��")
                DWF6.Visible = True
            Case 2  '�ק�
                DWF6.BackColor = Color.Yellow
                DWF6.Visible = True
            Case Else   '����
                DWF6.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF6", "ZZZZZZ")
        '�t��
        Select Case FindFieldInf("WF7")
            Case 0  '���
                DWF7.BackColor = Color.LightGray
                DWF7.Visible = True
            Case 1  '�ק�+�ˬd
                DWF7.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF7Rqd", "DWF7", "���`�G�ݿ�J�t��")
                DWF7.Visible = True
            Case 2  '�ק�
                DWF7.BackColor = Color.Yellow
                DWF7.Visible = True
            Case Else   '����
                DWF7.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF7", "ZZZZZZ")
        '�s�y���ӻ{�̳���
        Select Case FindFieldInf("WF3Name")
            Case 0  '���
                DWF3Name.BackColor = Color.LightGray
                DWF3Name.Visible = True
            Case 1  '�ק�+�ˬd
                DWF3Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF3NameRqd", "DWF3Name", "���`�G�ݿ�J�s�y������")
                DWF3Name.Visible = True
            Case 2  '�ק�
                DWF3Name.BackColor = Color.Yellow
                DWF3Name.Visible = True
            Case Else   '����
                DWF3Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF3NAME", "ZZZZZZ")
        '�s�y���ӻ{�̳���
        Select Case FindFieldInf("WF4Name")
            Case 0  '���
                DWF4Name.BackColor = Color.LightGray
                DWF4Name.Visible = True
            Case 1  '�ק�+�ˬd
                DWF4Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF4NameRqd", "DWF4Name", "���`�G�ݿ�J�s�y������")
                DWF4Name.Visible = True
            Case 2  '�ק�
                DWF4Name.BackColor = Color.Yellow
                DWF4Name.Visible = True
            Case Else   '����
                DWF4Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF4NAME", "ZZZZZZ")
        '�s�y���ӻ{�̳���
        Select Case FindFieldInf("WF5Name")
            Case 0  '���
                DWF5Name.BackColor = Color.LightGray
                DWF5Name.Visible = True
            Case 1  '�ק�+�ˬd
                DWF5Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF5NameRqd", "DWF5Name", "���`�G�ݿ�J�s�y������")
                DWF5Name.Visible = True
            Case 2  '�ק�
                DWF5Name.BackColor = Color.Yellow
                DWF5Name.Visible = True
            Case Else   '����
                DWF5Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF5NAME", "ZZZZZZ")
        '�s�y���ӻ{�̳���
        Select Case FindFieldInf("WF6Name")
            Case 0  '���
                DWF6Name.BackColor = Color.LightGray
                DWF6Name.Visible = True
            Case 1  '�ק�+�ˬd
                DWF6Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF6NameRqd", "DWF6Name", "���`�G�ݿ�J�s�y������")
                DWF6Name.Visible = True
            Case 2  '�ק�
                DWF6Name.BackColor = Color.Yellow
                DWF6Name.Visible = True
            Case Else   '����
                DWF6Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF6NAME", "ZZZZZZ")
        '�t���ӻ{�̳���
        Select Case FindFieldInf("WF7Name")
            Case 0  '���
                DWF7Name.BackColor = Color.LightGray
                DWF7Name.Visible = True
            Case 1  '�ק�+�ˬd
                DWF7Name.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DWF7NameRqd", "DWF7Name", "���`�G�ݿ�J�t������")
                DWF7Name.Visible = True
            Case 2  '�ק�
                DWF7Name.BackColor = Color.Yellow
                DWF7Name.Visible = True
            Case Else   '����
                DWF7Name.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("WF7NAME", "ZZZZZZ")

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

        '�@����
        If pFieldName = "WF1" Then
            DWF1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where Active = 1 "
                SQL = SQL & "   And UserID =  '" & Request.QueryString("pUserID") & "'"
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = LTrim(RTrim(DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")))
                    ListItem1.Value = LTrim(RTrim(DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")))
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF1.Items.Add(ListItem1)
                End If
            End If
        End If
        'EA�d����
        If pFieldName = "WF2" Then
            DWF2.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF2.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF2' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF2.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y��
        If pFieldName = "WF3" Then
            DWF3.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF3.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF3.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y��
        If pFieldName = "WF4" Then
            DWF4.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF4.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF4.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y��
        If pFieldName = "WF5" Then
            DWF5.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF5.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF5.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y��
        If pFieldName = "WF6" Then
            DWF6.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF6.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF6.Items.Add(ListItem1)
                Next
            End If
        End If
        '�t��
        If pFieldName = "WF7" Then
            DWF7.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF7.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WF' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF7.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y������
        If pFieldName = "WF3NAME" Then
            DWF3Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF3Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF3Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y������
        If pFieldName = "WF4NAME" Then
            DWF4Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF4Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF4Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y������
        If pFieldName = "WF5NAME" Then
            DWF5Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF5Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF5Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '�s�y������
        If pFieldName = "WF6NAME" Then
            DWF6Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF6Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF6Name.Items.Add(ListItem1)
                Next
            End If
        End If
        '�t������
        If pFieldName = "WF7NAME" Then
            DWF7Name.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DWF7Name.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='2001' and DKey='WFNAME' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DWF7Name.Items.Add(ListItem1)
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
    'Joy-090113
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

        'Check�u�{�֩w
        If ErrCode = 0 Then
            If DWF1.SelectedValue = "" Then ErrCode = 9010
            If DWF2.SelectedValue = "" Then ErrCode = 9010

            If DWF3Name.SelectedValue <> "" Then
                If DWF3.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF4Name.SelectedValue <> "" Then
                If DWF4.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF5Name.SelectedValue <> "" Then
                If DWF5.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF6Name.SelectedValue <> "" Then
                If DWF6.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF7Name.SelectedValue <> "" Then
                If DWF7.SelectedValue = "" Then ErrCode = 9010
            End If

            If DWF3.SelectedValue <> "" Then
                If DWF3Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF4.SelectedValue <> "" Then
                If DWF4Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF5.SelectedValue <> "" Then
                If DWF5Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF6.SelectedValue <> "" Then
                If DWF6Name.SelectedValue = "" Then ErrCode = 9010
            End If
            If DWF7.SelectedValue <> "" Then
                If DWF7Name.SelectedValue = "" Then ErrCode = 9010
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
        '�ˬd�e�U��No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("002001", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                    '��Xñ�֪�User ID
                    Dim wAllocateID As String = ""
                    Dim SQL As String = ""
                    Dim DBDataSet1 As New DataSet
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
                    OleDbConnection1.Open()

                    If pAction = 0 Then
                        SQL = "Select UserID From M_Users "
                        SQL = SQL & " Where Active = 1 "
                        Select Case wStep
                            Case 1, 500
                                SQL = SQL & "   And UserName =  '" & DWF2.SelectedValue & "'"
                            Case 10
                                SQL = SQL & "   And UserName =  '" & DWF3.SelectedValue & "'"
                            Case 20
                                SQL = SQL & "   And UserName =  '" & DWF4.SelectedValue & "'"
                            Case 30
                                SQL = SQL & "   And UserName =  '" & DWF5.SelectedValue & "'"
                            Case 40
                                SQL = SQL & "   And UserName =  '" & DWF6.SelectedValue & "'"
                            Case 50
                                SQL = SQL & "   And UserName =  '" & DWF7.SelectedValue & "'"
                            Case Else
                        End Select
                        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        DBAdapter1.Fill(DBDataSet1, "M_Users")
                        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")
                        OleDbConnection1.Close()
                        '�p�G���L�ĨϥΪ̮� Action=2 (������999����)
                        If wAllocateID = "" Then pAction = 2
                    End If
                    '���oñ�֪�
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
                    '�u�{����(999)��OK���s�ɳB�z
                    If pNextStep = 999 Then
                        If pFun = "OK" Then
                            InterfaceProc()
                        End If
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
            If ErrCode = 9010 Then Message = "�ӻ{�̫��w���~,�нT�{!"
            If ErrCode = 9040 Then Message = "����z�Ѭ���L�ɻݶ�g����,�нT�{!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_SampleSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "            '1~5
        SQl = SQl + "Date, AppBuyer, SizeNo, Item, CodeNo, "               '6~10
        SQl = SQl + "SampleFile, TAWidth, DevNo, DevPrd, TACol, "          '11~15
        SQl = SQl + "TALine, ECol, CCol, THCol, Other, "                   '16~20

        'update by alin
        SQl = SQl + "QCFile1,QCFile2,QCFile3,QCFile4,QCFile5,TNLItem, TNRItem, TSLItem, TSRItem, "         '21~25
        SQl = SQl + "TDLItem, TDRItem, CNITem, CSItem, CDItem, "           '26~30
        SQl = SQl + "CItem, "                                              '31
        SQl = SQl + "OP1, OP2, OP3, OP4, OP5, OP6, "                       '32~37
        SQl = SQl + "WF1, WF2, WF3, WF4, WF5 , WF6, WF7, "                 '38~44
        SQl = SQl + "WF3Name, WF4Name, WF5Name, WF6Name, WF7Name, "        '45~49
        SQl = SQl + "Rno, "                                                '50
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime ,Other1,Other2,O1Item,O2Item"      '51~54
        '------------------------

        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                        '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "                      '���פ�
        SQl = SQl + " '002001', "                                   '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "                 '���y����
        SQl = SQl + " '" + DNo.Text + "', "                         'NO
        '6~10
        SQl = SQl + " '" + DDATE.Text + "', "                       '�o����
        SQl = SQl + " '" + DAPPBUYER.Text + "', "                   '�Ȥ�
        SQl = SQl + " '" + DSIZENO.Text + "', "                     'Size
        SQl = SQl + " '" + DITEM.Text + "', "                       'Item
        SQl = SQl + " '" + DCODENO.Text + "', "                     'Code-No
        '11~15
        FileName = ""
        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSAMPLEFILE.PostedFile.FileName, InStr(StrReverse(DSAMPLEFILE.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DSAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'Sample-File

        SQl = SQl + " '" + DTAWIDTH.Text + "', "                    '���a�e��
        SQl = SQl + " '" + DDEVNO.Text + "', "                      '�}�oNo
        SQl = SQl + " '" + DDEVPRD.Text + "', "                     '�}�o����
        SQl = SQl + " '" + DTACOL.Text + "', "                      '���a
        '16~20
        SQl = SQl + " '" + DTALINE.Text + "', "                     '�����u
        SQl = SQl + " '" + DECOL.Text + "', "                       '�Ⱦ�
        SQl = SQl + " '" + DCCOL.Text + "', "                       '�Y��
        SQl = SQl + " '" + DTHCOL.Text + "', "                      '�_�u�u
        SQl = SQl + " N'" + DOTHER.Text + "', "                     '��L
        '21~25

        'update by alin
        '�~���ɮ�1
        FileName = ""
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & Right(DQCFILE1.PostedFile.FileName, InStr(StrReverse(DQCFILE1.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File1

        '�~���ɮ�2
        FileName = ""
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & Right(DQCFILE2.PostedFile.FileName, InStr(StrReverse(DQCFILE2.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File2

        '�~���ɮ�3
        FileName = ""
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & Right(DQCFILE3.PostedFile.FileName, InStr(StrReverse(DQCFILE3.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFILE3.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File3

        '�~���ɮ�4
        FileName = ""
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & Right(DQCFILE4.PostedFile.FileName, InStr(StrReverse(DQCFILE4.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFILE4.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File4

        '�~���ɮ�5
        FileName = ""
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & Right(DQCFILE5.PostedFile.FileName, InStr(StrReverse(DQCFILE5.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǹ���
                    DQCFILE5.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " '" + FileName + "', "                         'QC-File5
        '------------------------

        SQl = SQl + " '" + DTNLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTNRITEM.Text + "', "                    '
        SQl = SQl + " '" + DTSLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTSRITEM.Text + "', "                    '
        '26~30
        SQl = SQl + " '" + DTDLITEM.Text + "', "                    '
        SQl = SQl + " '" + DTDRITEM.Text + "', "                    '
        SQl = SQl + " '" + DCNITEM.Text + "', "                     '
        SQl = SQl + " '" + DCSITEM.Text + "', "                     '
        SQl = SQl + " '" + DCDITEM.Text + "', "                     '
        '31
        SQl = SQl + " '" + DCITEM.Text + "', "                      '
        '32~37
        SQl = SQl + " '" + DOP1.Text + "', "                        '
        SQl = SQl + " '" + DOP2.Text + "', "                        '
        SQl = SQl + " '" + DOP3.Text + "', "                        '
        SQl = SQl + " '" + DOP4.Text + "', "                        '
        SQl = SQl + " '" + DOP5.Text + "', "                        '
        SQl = SQl + " '" + DOP6.Text + "', "                        '
        '38~44
        SQl = SQl + " '" + DWF1.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF2.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF3.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF4.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF5.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF6.SelectedValue + "', "               '
        SQl = SQl + " '" + DWF7.SelectedValue + "', "               '
        '45~49
        SQl = SQl + " '" + DWF3Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF4Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF5Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF6Name.SelectedValue + "', "           '
        SQl = SQl + " '" + DWF7Name.SelectedValue + "', "           '
        '50
        SQl = SQl + " '" + DRNO.Text + "', "                    '
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "  '�@����
        SQl = SQl + " '" + NowDateTime + "', "                      '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                               '�ק��
        SQl = SQl + " '" + NowDateTime + "', "                       '�ק�ɶ�

        'update by alin
        SQl = SQl + " '" + Hidden1.Value + "', "                       'Other1
        SQl = SQl + " '" + Hidden2.Value + "', "                       'Other2
        SQl = SQl + " '" + DO1ITEM.Text + "', "                       'Item1
        SQl = SQl + " '" + DO2ITEM.Text + "' "                        'Item2
        '------------------------



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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SampleFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_SampleSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If

        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDATE.Text & "',"
        SQl = SQl + " AppBuyer = '" & DAPPBUYER.Text & "',"
        SQl = SQl + " SizeNo = '" & DSIZENO.Text & "',"
        SQl = SQl + " Item = '" & DITEM.Text & "',"
        SQl = SQl + " CodeNo = '" & DCODENO.Text & "',"

        If DSAMPLEFILE.Visible Then
            If DSAMPLEFILE.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSAMPLEFILE.PostedFile.FileName, InStr(StrReverse(DSAMPLEFILE.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DSAMPLEFILE.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DSAMPLEFILE.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " SampleFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " TAWidth = '" & DTAWIDTH.Text & "',"
        SQl = SQl + " DevNo = '" & DDEVNO.Text & "',"
        SQl = SQl + " DevPrd = '" & DDEVPRD.Text & "',"
        SQl = SQl + " TACol = '" & DTACOL.Text & "',"
        SQl = SQl + " TALine = '" & DTALINE.Text & "',"

        SQl = SQl + " ECol = '" & DECOL.Text & "',"
        SQl = SQl + " CCol = '" & DCCOL.Text & "',"
        SQl = SQl + " THCol = '" & DTHCOL.Text & "',"
        SQl = SQl + " Other = N'" & DOTHER.Text & "',"

        'update by alin
        If DQCFILE1.Visible Then
            If DQCFILE1.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & Right(DQCFILE1.PostedFile.FileName, InStr(StrReverse(DQCFILE1.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile1" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE1.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DQCFILE1.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile1 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE2.Visible Then
            If DQCFILE2.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & Right(DQCFILE2.PostedFile.FileName, InStr(StrReverse(DQCFILE2.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile2" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE2.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DQCFILE2.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile2 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE3.Visible Then
            If DQCFILE3.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & Right(DQCFILE3.PostedFile.FileName, InStr(StrReverse(DQCFILE3.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile3" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE3.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DQCFILE3.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile3 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE4.Visible Then
            If DQCFILE4.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & Right(DQCFILE4.PostedFile.FileName, InStr(StrReverse(DQCFILE4.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile4" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE4.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DQCFILE4.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile4 = N'" & FileName & "',"
            End If
        End If
        If DQCFILE5.Visible Then
            If DQCFILE5.PostedFile.FileName <> "" Then  '�P�_���ɮפW��
                '*** IE8����-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & Right(DQCFILE5.PostedFile.FileName, InStr(StrReverse(DQCFILE5.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "QCFile5" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DQCFILE5.PostedFile.FileName)
                '*** IE8����-End
                Try    '�W�ǯ��
                    DQCFILE5.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " QCFile5 = N'" & FileName & "',"
            End If
        End If
        '------------------------

        SQl = SQl + " TNLItem = '" & DTNLITEM.Text & "',"
        SQl = SQl + " TNRItem = '" & DTNRITEM.Text & "',"
        SQl = SQl + " TSLItem = '" & DTSLITEM.Text & "',"
        SQl = SQl + " TSRItem = '" & DTSRITEM.Text & "',"
        SQl = SQl + " TDLItem = '" & DTDLITEM.Text & "',"
        SQl = SQl + " TDRItem = '" & DTDRITEM.Text & "',"
        SQl = SQl + " CNITem = '" & DCNITEM.Text & "',"
        SQl = SQl + " CSItem = '" & DCSITEM.Text & "',"
        SQl = SQl + " CDItem = '" & DCDITEM.Text & "',"
        SQl = SQl + " CItem = '" & DCITEM.Text & "',"

        SQl = SQl + " OP1 = '" & DOP1.Text & "',"
        SQl = SQl + " OP2 = '" & DOP2.Text & "',"
        SQl = SQl + " OP3 = '" & DOP3.Text & "',"
        SQl = SQl + " OP4 = '" & DOP4.Text & "',"
        SQl = SQl + " OP5 = '" & DOP5.Text & "',"
        SQl = SQl + " OP6 = '" & DOP6.Text & "',"

        SQl = SQl + " WF1 = '" & DWF1.SelectedValue & "',"
        SQl = SQl + " WF2 = '" & DWF2.SelectedValue & "',"
        SQl = SQl + " WF3 = '" & DWF3.SelectedValue & "',"
        SQl = SQl + " WF4 = '" & DWF4.SelectedValue & "',"
        SQl = SQl + " WF5 = '" & DWF5.SelectedValue & "',"
        SQl = SQl + " WF6 = '" & DWF6.SelectedValue & "',"
        SQl = SQl + " WF7 = '" & DWF7.SelectedValue & "',"

        SQl = SQl + " WF3Name = '" & DWF3Name.SelectedValue & "',"
        SQl = SQl + " WF4Name = '" & DWF4Name.SelectedValue & "',"
        SQl = SQl + " WF5Name = '" & DWF5Name.SelectedValue & "',"
        SQl = SQl + " WF6Name = '" & DWF6Name.SelectedValue & "',"
        SQl = SQl + " WF7Name = '" & DWF7Name.SelectedValue & "',"

        SQl = SQl + " Rno = '" & DRNO.Text & "',"

        'update by alin
        SQl = SQl + " Other1 = '" & Hidden1.Value & "',"
        SQl = SQl + " Other2 = '" & Hidden2.Value & "',"
        SQl = SQl + " O1Item = '" & DO1ITEM.Text & "',"
        SQl = SQl + " O2Item = '" & DO2ITEM.Text & "',"
        '-----------------------------

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
    '**(InterfaceProc)
    '**     ��s�t�ζ�I/F���
    '**
    '*****************************************************************
    Sub InterfaceProc()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SCDSqlConn")  'SQL�s���]�w
        Dim OleDBCommand1 As New OleDbCommand

        Dim SQl As String
        SQl = "Update Commission Set "
        SQl = SQl + " SAMFN = '" & "�s��" & "',"
        SQl = SQl + " SAMFormNo = '" & wFormNo & "',"
        SQl = SQl + " SAMFormSno = '" & CStr(wFormSno) & "',"
        SQl = SQl + " ModOp = '" & "WF" & "',"
        SQl = SQl + " ModDate = '" & NowDateTime & "' "
        SQl = SQl + " Where Rno  =  '" & DRNO.Text & "'"
        SQl = SQl + "   And DevNo =  '" & DDEVNO.Text & "'"
        SQl = SQl + "   And CodeNo =  '" & DCODENO.Text & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
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
        '�o���
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDATE.Text = "" Then InputCheck = 1
            End If
        End If
        'AppBuyer
        If InputCheck = 0 Then
            If FindFieldInf("AppBuyer") = 1 Then
                If DAPPBUYER.Text = "" Then InputCheck = 1
            End If
        End If
        'Size
        If InputCheck = 0 Then
            If FindFieldInf("SizeNo") = 1 Then
                If DSIZENO.Text = "" Then InputCheck = 1
            End If
        End If
        'Item
        If InputCheck = 0 Then
            If FindFieldInf("Item") = 1 Then
                If DITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape
        If InputCheck = 0 Then
            If FindFieldInf("CodeNo") = 1 Then
                If DCODENO.Text = "" Then InputCheck = 1
            End If
        End If
        '��ڼ˫~
        If InputCheck = 0 Then
            If FindFieldInf("SampleFile") = 1 Then
                If DSAMPLEFILE.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '���a�e��
        If InputCheck = 0 Then
            If FindFieldInf("TAWidth") = 1 Then
                If DTAWIDTH.Text = "" Then InputCheck = 1
            End If
        End If
        '�}�oNo
        If InputCheck = 0 Then
            If FindFieldInf("DevNo") = 1 Then
                If DDEVNO.Text = "" Then InputCheck = 1
            End If
        End If
        '�}�o����
        If InputCheck = 0 Then
            If FindFieldInf("DevPrd") = 1 Then
                If DDEVPRD.Text = "" Then InputCheck = 1
            End If
        End If
        '���a
        If InputCheck = 0 Then
            If FindFieldInf("TACol") = 1 Then
                If DTACOL.Text = "" Then InputCheck = 1
            End If
        End If
        '�����u
        If InputCheck = 0 Then
            If FindFieldInf("TALine") = 1 Then
                If DTALINE.Text = "" Then InputCheck = 1
            End If
        End If
        '�Ⱦ�
        If InputCheck = 0 Then
            If FindFieldInf("ECol") = 1 Then
                If DECOL.Text = "" Then InputCheck = 1
            End If
        End If
        '�Y��
        If InputCheck = 0 Then
            If FindFieldInf("CCol") = 1 Then
                If DCCOL.Text = "" Then InputCheck = 1
            End If
        End If
        '�_�u�u
        If InputCheck = 0 Then
            If FindFieldInf("THCol") = 1 Then
                If DTHCOL.Text = "" Then InputCheck = 1
            End If
        End If
        '��L
        If InputCheck = 0 Then
            If FindFieldInf("Other") = 1 Then
                If DOTHER.Text = "" Then InputCheck = 1
            End If
        End If
        '�~�����i1
        If InputCheck = 0 Then
            If FindFieldInf("QCFile1") = 1 Then
                If DQCFILE1.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�~�����i2
        If InputCheck = 0 Then
            If FindFieldInf("QCFile2") = 1 Then
                If DQCFILE2.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�~�����i3
        If InputCheck = 0 Then
            If FindFieldInf("QCFile3") = 1 Then
                If DQCFILE3.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�~�����i4
        If InputCheck = 0 Then
            If FindFieldInf("QCFile4") = 1 Then
                If DQCFILE4.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '�~�����i5
        If InputCheck = 0 Then
            If FindFieldInf("QCFile5") = 1 Then
                If DQCFILE5.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        'Tape Nat-��
        If InputCheck = 0 Then
            If FindFieldInf("TNLItem") = 1 Then
                If DTNLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Nat-�k
        If InputCheck = 0 Then
            If FindFieldInf("TNRItem") = 1 Then
                If DTNRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Set-��
        If InputCheck = 0 Then
            If FindFieldInf("TSLItem") = 1 Then
                If DTSLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Set-�k
        If InputCheck = 0 Then
            If FindFieldInf("TSRItem") = 1 Then
                If DTSRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Dyed-��
        If InputCheck = 0 Then
            If FindFieldInf("TDLItem") = 1 Then
                If DTDLITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Tape Dyed-�k
        If InputCheck = 0 Then
            If FindFieldInf("TDRItem") = 1 Then
                If DTDRITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Nat
        If InputCheck = 0 Then
            If FindFieldInf("CNItem") = 1 Then
                If DCNITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Set
        If InputCheck = 0 Then
            If FindFieldInf("CSItem") = 1 Then
                If DCSITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Chain Dyed
        If InputCheck = 0 Then
            If FindFieldInf("CDItem") = 1 Then
                If DCDITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'O1ITEM
        If InputCheck = 0 Then
            If FindFieldInf("O1Item") = 1 Then
                If DO1ITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'O2ITEM
        If InputCheck = 0 Then
            If FindFieldInf("O2Item") = 1 Then
                If DO2ITEM.Text = "" Then InputCheck = 1
            End If
        End If
        'Cord
        If InputCheck = 0 Then
            If FindFieldInf("CItem") = 1 Then
                If DCITEM.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP1") = 1 Then
                If DOP1.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP2") = 1 Then
                If DOP2.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP3") = 1 Then
                If DOP3.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP4") = 1 Then
                If DOP4.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP5") = 1 Then
                If DOP5.Text = "" Then InputCheck = 1
            End If
        End If
        '�u�{��
        If InputCheck = 0 Then
            If FindFieldInf("OP6") = 1 Then
                If DOP6.Text = "" Then InputCheck = 1
            End If
        End If
        '�@����
        If InputCheck = 0 Then
            If FindFieldInf("WF1") = 1 Then
                If DWF1.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'EA�d����
        If InputCheck = 0 Then
            If FindFieldInf("WF2") = 1 Then
                If DWF2.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y��
        If InputCheck = 0 Then
            If FindFieldInf("WF3") = 1 Then
                If DWF3.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y��
        If InputCheck = 0 Then
            If FindFieldInf("WF4") = 1 Then
                If DWF4.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y��
        If InputCheck = 0 Then
            If FindFieldInf("WF5") = 1 Then
                If DWF5.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y��
        If InputCheck = 0 Then
            If FindFieldInf("WF6") = 1 Then
                If DWF6.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�t��
        If InputCheck = 0 Then
            If FindFieldInf("WF7") = 1 Then
                If DWF7.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y���ӻ{�̳���
        If InputCheck = 0 Then
            If FindFieldInf("WF3Name") = 1 Then
                If DWF3Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y���ӻ{�̳���
        If InputCheck = 0 Then
            If FindFieldInf("WF4Name") = 1 Then
                If DWF4Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y���ӻ{�̳���
        If InputCheck = 0 Then
            If FindFieldInf("WF5Name") = 1 Then
                If DWF5Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�s�y���ӻ{�̳���
        If InputCheck = 0 Then
            If FindFieldInf("WF6Name") = 1 Then
                If DWF6Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '�t���ӻ{�̳���
        If InputCheck = 0 Then
            If FindFieldInf("WF7Name") = 1 Then
                If DWF7Name.SelectedValue = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
