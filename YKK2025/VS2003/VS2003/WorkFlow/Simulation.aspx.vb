Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class Simulation
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP1D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP8 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP7 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP8D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP7D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP10 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP10D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP8I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP4I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP5I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP6I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP7I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP3I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP2I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP1I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP10I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP14I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP14 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP14D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP12 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP15 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP11 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP12D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP11D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP12I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP11I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP15D As System.Web.UI.WebControls.TextBox
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents DFormName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP15I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP20 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP20D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP17 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP17D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP17I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP18 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP18D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP18I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP19 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP19D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP19I As System.Web.UI.WebControls.Image
    Protected WithEvents LOP6 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP5 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP7 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP8 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP9 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP10 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP11 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP12 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP13 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP14 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP15 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP16 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP17 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP18 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP19 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP20 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOP21 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP24I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP20I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP21D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP21I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP22 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP22D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP22I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP23 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP23D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP23I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP24 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP24D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP25 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP25D As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOP25 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP24 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP23 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP22 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP21 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOP25I As System.Web.UI.WebControls.Image

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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer = 0     '�渹
    Dim wStep As Integer            '�u�{�N�X
    Dim wUserID As String           '�֩w��ID
    Dim wLevel As String            '������
    Dim wAllocateID As String       '���wID
    Dim wAgentID As String = ""
    Dim wMultiJob As Integer = 0
    Dim wQCLT As Integer = 0
    'Modify-Start by Joy 2009/11/20(2010��ƾ����)
    'Dim wDepo As String             '���c��ƾ�(CL->���c, TP->�x�_)
    '
    '�s�զ�ƾ�
    Dim wDepo As String = "CL1"      '���c��ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
    'Modify-End

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "Simulation.aspx"

        If Not Me.IsPostBack Then
            SetParameter()  '�]�w�@�ΰѼ�
            SetFormItem()   '�]�w�e���U���ت�l��
            SetContent()    '�]�w�e���u�{���e
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�@�ΰѼ�
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '�渹
        wStep = Request.QueryString("pStep")        '�u�{�N�X
        wUserID = Request.QueryString("pUserID")    '�֩w��ID
        wLevel = Request.QueryString("pLevel")      '������
        wAllocateID = Request.QueryString("pAllocateID") '���wID
        wDepo = Request.QueryString("pDepo")        '��ƾ�
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�e���u�{���e
    '**
    '*****************************************************************
    Sub SetContent()
        Dim SQL, Str As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        '--Start OP---------
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step =  '" & wStep & "'"
        SQL = SQL & "   And Action = 0 "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FLOW_1")
        If DBDataSet1.Tables("M_FLOW_1").Rows.Count > 0 Then
            DFormName.Text = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("FormName")
            DOP1.Text = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("StepName")

            SQL = "Select UserName From M_Users "
            SQL = SQL & " Where Active = 1 "
            If wUserID <> "" Then
                SQL = SQL & "   And UserID =  '" & wUserID & "'"
            Else
                SQL = SQL & "   And UserID =  '" & Request.Cookies("UserID").Value & "'"
            End If
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users_1")
            If DBDataSet1.Tables("M_Users_1").Rows.Count > 0 Then
                Str = DBDataSet1.Tables("M_Users_1").Rows(0).Item("UserName")
            End If
        End If

        If wStep = 1 Then
            DOP1D.Text = "�e�U�H�G" & Str & Chr(13) & _
                         "�e�U�ɶ��G" & FormatDateTime(Now(), DateFormat.GeneralDate)
        Else
            DOP1D.Text = "���G" & Str & Chr(13) & _
                         "�w�w�����ɶ��G" & FormatDateTime(Now(), DateFormat.GeneralDate)
        End If

        SetLayout(1)

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

        Dim xStep As Integer = wStep    '�u�{�N�X
        Dim xUserID As String = Request.Cookies("UserID").Value  'ñ�֪�ID
        Dim xApplyID As String = Request.Cookies("UserID").Value '�ӽЪ�ID
        Dim xOP As String = ""          '�u�{�W
        Dim xUser As String = ""       '���W
        Dim xNow As DateTime = Now()    '���

        Dim pAction As Integer = 0
        Dim RtnCode As Integer = 0
        Dim i As Integer = 2
        Dim j As Integer

        If wFormSno <> 0 Then
            DBDataSet1.Clear()
            SQL = "Select ApplyID From T_WaitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '1' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle_1")
            If DBDataSet1.Tables("T_WaitHandle_1").Rows.Count > 0 Then
                xApplyID = DBDataSet1.Tables("T_WaitHandle_1").Rows(0).Item("ApplyID")
            End If
        End If

        While pNextStep <> 999
            DBDataSet1.Clear()
            For j = 1 To 10
                pNextGate(j) = ""
            Next j

            '���o�U�@��
            RtnCode = oCommon.GetNextGate(wFormNo, xStep, xUserID, xApplyID, wAgentID, wAllocateID, wMultiJob, _
                                          pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
            '��渹�X,�u�{���d���X,ñ�֪�,�ӽЪ�,�Q�N�z��,�Q���w��,�h�H�֩w�u�{No,
            '�U�@�u�{��, ���X, ����, �Q�N�z��, �H��, �B�z��k, �ʧ@(0:OK,1:NG,2:Save) 

            SQL = "Select * From M_Flow "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And Step =  '" & pNextStep & "'"
            SQL = SQL & "   And Action =  '" & pAction & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_FLOW_1")
            If DBDataSet1.Tables("M_FLOW_1").Rows.Count > 0 Then
                xOP = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("StepName")
            Else
                xOP = "Not Found"
            End If

            If pFlowType = 0 Then xOP = xOP & "(�q��)"
            'If pFlowType = 1 Then xOP = xOP & "(�u�{)"
            'If pFlowType = 2 Then xOP = xOP & "(�e�U)"

            If pNextStep <> 999 Then
                '���o�u�{�t���̫���
                oSchedule.GetLastTime(pNextGate(1), wFormNo, pNextStep, pFlowType, xNow, pLastTime, pCount1)
                '��渹�X, �u�{���X, ���O(0:�q��,1:�֩w), �}�l���, �̫���, ���

                '���o�w�w�}�l,������{�p��
                oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                '��渹�X,�u�{���X,������,QC-L/T,�{�b�ɶ�, �w�w�}�l���, �w�w�������, ��ƾ�

                xUser = ""
                For j = 1 To pCount
                    DBDataSet2.Clear()
                    SQL = "Select UserName From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    SQL = SQL & "   And UserID =  '" & pNextGate(j) & "'"
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet2, "M_Users_1")
                    If DBDataSet2.Tables("M_Users_1").Rows.Count > 0 Then
                        If xUser = "" Then
                            xUser = DBDataSet2.Tables("M_Users_1").Rows(0).Item("UserName")
                        Else
                            xUser = xUser & ", " & DBDataSet2.Tables("M_Users_1").Rows(0).Item("UserName")
                        End If
                    End If
                Next j

                Select Case i
                    Case 2
                        DOP2.Text = xOP
                        DOP2D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP2.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 3
                        DOP3.Text = xOP
                        DOP3D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP3.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 4
                        DOP4.Text = xOP
                        DOP4D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP4.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 5
                        DOP5.Text = xOP
                        DOP5D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP5.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 6
                        DOP6.Text = xOP
                        DOP6D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP6.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 7
                        DOP7.Text = xOP
                        DOP7D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP7.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 8
                        DOP8.Text = xOP
                        DOP8D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP8.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 9
                        DOP9.Text = xOP
                        DOP9D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP9.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 10
                        DOP10.Text = xOP
                        DOP10D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP10.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 11
                        DOP11.Text = xOP
                        DOP11D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP11.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 12
                        DOP12.Text = xOP
                        DOP12D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP12.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 13
                        DOP13.Text = xOP
                        DOP13D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP13.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 14
                        DOP14.Text = xOP
                        DOP14D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP14.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 15
                        DOP15.Text = xOP
                        DOP15D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP15.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 16
                        DOP16.Text = xOP
                        DOP16D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP16.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 17
                        DOP17.Text = xOP
                        DOP17D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP17.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 18
                        DOP18.Text = xOP
                        DOP18D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP18.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 19
                        DOP19.Text = xOP
                        DOP19D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP19.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 20
                        DOP20.Text = xOP
                        DOP20D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP20.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 21
                        DOP21.Text = xOP
                        DOP21D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP21.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 22
                        DOP22.Text = xOP
                        DOP22D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP22.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 23
                        DOP23.Text = xOP
                        DOP23D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP23.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 24
                        DOP24.Text = xOP
                        DOP24D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP24.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 25
                        DOP25.Text = xOP
                        DOP25D.Text = "���G" & xUser & Chr(13) & _
                                     "�ݳB�z��ơG" & CStr(pCount1) & Chr(13) & _
                                     "�̫᧹���w�w�G" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�}�l�w�w�G" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP25.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case Else
                        DOP25.Text = "���󤣨�"
                End Select
            Else
                Select Case i   'NextStep=999
                    Case 2
                        DOP2.Text = xOP
                        DOP2D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 3
                        DOP3.Text = xOP
                        DOP3D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 4
                        DOP4.Text = xOP
                        DOP4D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 5
                        DOP5.Text = xOP
                        DOP5D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 6
                        DOP6.Text = xOP
                        DOP6D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 7
                        DOP7.Text = xOP
                        DOP7D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 8
                        DOP8.Text = xOP
                        DOP8D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 9
                        DOP9.Text = xOP
                        DOP9D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 10
                        DOP10.Text = xOP
                        DOP10D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 11
                        DOP11.Text = xOP
                        DOP11D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 12
                        DOP12.Text = xOP
                        DOP12D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 13
                        DOP13.Text = xOP
                        DOP13D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 14
                        DOP14.Text = xOP
                        DOP14D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 15
                        DOP15.Text = xOP
                        DOP15D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 16
                        DOP16.Text = xOP
                        DOP16D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 17
                        DOP17.Text = xOP
                        DOP17D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 18
                        DOP18.Text = xOP
                        DOP18D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 19
                        DOP19.Text = xOP
                        DOP19D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 20
                        DOP20.Text = xOP
                        DOP20D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 21
                        DOP21.Text = xOP
                        DOP21D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 22
                        DOP22.Text = xOP
                        DOP22D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 23
                        DOP23.Text = xOP
                        DOP23D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 24
                        DOP24.Text = xOP
                        DOP24D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 25
                        DOP25.Text = xOP
                        DOP25D.Text = "�����w�w�G" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case Else
                        DOP25.Text = "���󤣨�"
                End Select
            End If
            SetLayout(i)
            i = i + 1

            xStep = pNextStep  '�u�{���d���X
            xUserID = pNextGate(1)     'ñ�֪�ID
            If pFlowType <> 0 Then xNow = pEndTime '���
        End While

        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�e���u�{��
    '**
    '*****************************************************************
    Sub SetLayout(ByVal idx As Integer)
        Select Case idx
            Case 1
                DOP1.Visible = True
                DOP1D.Visible = True
                If LOP1.NavigateUrl <> "" Then LOP1.Visible = True
            Case 2
                DOP1I.Visible = True
                DOP2.Visible = True
                DOP2D.Visible = True
                If LOP2.NavigateUrl <> "" Then LOP2.Visible = True
            Case 3
                DOP2I.Visible = True
                DOP3.Visible = True
                DOP3D.Visible = True
                If LOP3.NavigateUrl <> "" Then LOP3.Visible = True
            Case 4
                DOP3I.Visible = True
                DOP4.Visible = True
                DOP4D.Visible = True
                If LOP4.NavigateUrl <> "" Then LOP4.Visible = True
            Case 5
                DOP4I.Visible = True
                DOP5.Visible = True
                DOP5D.Visible = True
                If LOP5.NavigateUrl <> "" Then LOP5.Visible = True
            Case 6
                DOP5I.Visible = True
                DOP6.Visible = True
                DOP6D.Visible = True
                If LOP6.NavigateUrl <> "" Then LOP6.Visible = True
            Case 7
                DOP6I.Visible = True
                DOP7.Visible = True
                DOP7D.Visible = True
                If LOP7.NavigateUrl <> "" Then LOP7.Visible = True
            Case 8
                DOP7I.Visible = True
                DOP8.Visible = True
                DOP8D.Visible = True
                If LOP8.NavigateUrl <> "" Then LOP8.Visible = True
            Case 9
                DOP8I.Visible = True
                DOP9.Visible = True
                DOP9D.Visible = True
                If LOP9.NavigateUrl <> "" Then LOP9.Visible = True
            Case 10
                DOP9I.Visible = True
                DOP10.Visible = True
                DOP10D.Visible = True
                If LOP10.NavigateUrl <> "" Then LOP10.Visible = True
            Case 11
                DOP10I.Visible = True
                DOP11.Visible = True
                DOP11D.Visible = True
                If LOP11.NavigateUrl <> "" Then LOP11.Visible = True
            Case 12
                DOP11I.Visible = True
                DOP12.Visible = True
                DOP12D.Visible = True
                If LOP12.NavigateUrl <> "" Then LOP12.Visible = True
            Case 13
                DOP12I.Visible = True
                DOP13.Visible = True
                DOP13D.Visible = True
                If LOP13.NavigateUrl <> "" Then LOP13.Visible = True
            Case 14
                DOP13I.Visible = True
                DOP14.Visible = True
                DOP14D.Visible = True
                If LOP14.NavigateUrl <> "" Then LOP14.Visible = True
            Case 15
                DOP14I.Visible = True
                DOP15.Visible = True
                DOP15D.Visible = True
                If LOP15.NavigateUrl <> "" Then LOP15.Visible = True
            Case 16
                DOP15I.Visible = True
                DOP16.Visible = True
                DOP16D.Visible = True
                If LOP16.NavigateUrl <> "" Then LOP16.Visible = True
            Case 17
                DOP16I.Visible = True
                DOP17.Visible = True
                DOP17D.Visible = True
                If LOP17.NavigateUrl <> "" Then LOP17.Visible = True
            Case 18
                DOP17I.Visible = True
                DOP18.Visible = True
                DOP18D.Visible = True
                If LOP18.NavigateUrl <> "" Then LOP18.Visible = True
            Case 19
                DOP18I.Visible = True
                DOP19.Visible = True
                DOP19D.Visible = True
                If LOP19.NavigateUrl <> "" Then LOP19.Visible = True
            Case 20
                DOP19I.Visible = True
                DOP20.Visible = True
                DOP20D.Visible = True
                If LOP20.NavigateUrl <> "" Then LOP20.Visible = True
            Case 21
                DOP20I.Visible = True
                DOP21.Visible = True
                DOP21D.Visible = True
                If LOP21.NavigateUrl <> "" Then LOP21.Visible = True
            Case 22
                DOP21I.Visible = True
                DOP22.Visible = True
                DOP22D.Visible = True
                If LOP22.NavigateUrl <> "" Then LOP22.Visible = True
            Case 23
                DOP22I.Visible = True
                DOP23.Visible = True
                DOP23D.Visible = True
                If LOP23.NavigateUrl <> "" Then LOP23.Visible = True
            Case 24
                DOP23I.Visible = True
                DOP24.Visible = True
                DOP24D.Visible = True
                If LOP24.NavigateUrl <> "" Then LOP24.Visible = True
            Case 25
                DOP24I.Visible = True
                DOP25.Visible = True
                DOP25D.Visible = True
                If LOP25.NavigateUrl <> "" Then LOP25.Visible = True
            Case Else
        End Select
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w�e���U���ت�l��
    '**
    '*****************************************************************
    Sub SetFormItem()
        '�e���U���ت�ܳB�z
        DOP1.Visible = False
        DOP1D.Visible = False
        DOP1I.Visible = False
        DOP2.Visible = False
        DOP2D.Visible = False
        DOP2I.Visible = False
        DOP3.Visible = False
        DOP3D.Visible = False
        DOP3I.Visible = False
        DOP4.Visible = False
        DOP4D.Visible = False
        DOP4I.Visible = False
        DOP5.Visible = False
        DOP5D.Visible = False
        DOP5I.Visible = False
        DOP6.Visible = False
        DOP6D.Visible = False
        DOP6I.Visible = False
        DOP7.Visible = False
        DOP7D.Visible = False
        DOP7I.Visible = False
        DOP8.Visible = False
        DOP8D.Visible = False
        DOP8I.Visible = False
        DOP9.Visible = False
        DOP9D.Visible = False
        DOP9I.Visible = False
        DOP10.Visible = False
        DOP10D.Visible = False
        DOP10I.Visible = False
        DOP11.Visible = False
        DOP11D.Visible = False
        DOP11I.Visible = False
        DOP12.Visible = False
        DOP12D.Visible = False
        DOP12I.Visible = False
        DOP13.Visible = False
        DOP13D.Visible = False
        DOP13I.Visible = False
        DOP14.Visible = False
        DOP14D.Visible = False
        DOP14I.Visible = False
        DOP15.Visible = False
        DOP15D.Visible = False
        DOP15I.Visible = False
        DOP16.Visible = False
        DOP16D.Visible = False
        DOP16I.Visible = False
        DOP17.Visible = False
        DOP17D.Visible = False
        DOP17I.Visible = False
        DOP18.Visible = False
        DOP18D.Visible = False
        DOP18I.Visible = False
        DOP19.Visible = False
        DOP19D.Visible = False
        DOP19I.Visible = False
        DOP20.Visible = False
        DOP20D.Visible = False
        DOP20I.Visible = False
        DOP21.Visible = False
        DOP21D.Visible = False
        DOP21I.Visible = False
        DOP22.Visible = False
        DOP22D.Visible = False
        DOP22I.Visible = False
        DOP23.Visible = False
        DOP23D.Visible = False
        DOP23I.Visible = False
        DOP24.Visible = False
        DOP24D.Visible = False
        DOP24I.Visible = False
        DOP25.Visible = False
        DOP25D.Visible = False
        DOP25I.Visible = False

        LOP1.Visible = False
        LOP2.Visible = False
        LOP3.Visible = False
        LOP4.Visible = False
        LOP5.Visible = False
        LOP6.Visible = False
        LOP7.Visible = False
        LOP8.Visible = False
        LOP9.Visible = False
        LOP10.Visible = False
        LOP11.Visible = False
        LOP12.Visible = False
        LOP13.Visible = False
        LOP14.Visible = False
        LOP15.Visible = False
        LOP16.Visible = False
        LOP17.Visible = False
        LOP18.Visible = False
        LOP19.Visible = False
        LOP20.Visible = False
        LOP21.Visible = False
        LOP22.Visible = False
        LOP23.Visible = False
        LOP24.Visible = False
        LOP25.Visible = False

    End Sub

    Private Sub Textbox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DOP22.TextChanged

    End Sub
End Class
