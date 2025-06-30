Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDAppendSpecSheet_02
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
        TopPosition()           '���s��RequestedField��Top��m
        If Not Me.IsPostBack Then   '���OPostBack
            ShowFormData()      '��ܪ����
        End If

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

                End If
            End If
        Else
            Top = 696
        End If
        '----


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

        Response.Cookies("PGM").Value = "SBDAppendSpecSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X
    End Sub





    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim SQL As String
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDAppendSpecPath")



        SQL = "Select * From F_SBDAppendSpecSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Value = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBuyer.Value = DBAdapter1.Rows(0).Item("Buyer")
            DSupplier.Value = DBAdapter1.Rows(0).Item("Supplier") '�~�`�t��
            DVendor.Value = DBAdapter1.Rows(0).Item("Vendor")              'Vendor

            DSurfSheetNo.Value = DBAdapter1.Rows(0).Item("SurfSheetNo") ' �s���B�zNO
            DSurfSupplier.Value = DBAdapter1.Rows(0).Item("SurfSupplier") '�~�`�t��

            DCap.Value = DBAdapter1.Rows(0).Item("Cap")              '�鲣��





            DSchedule.Value = DBAdapter1.Rows(0).Item("Schedule")              '��¦��{

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


            '�˴����G
            DQCResult.Text = DBAdapter1.Rows(0).Item("QCResult")

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



            If Trim(DBAdapter1.Rows(0).Item("FinalSampleFile")) <> "" Then
                LFinalSampleFile.ImageUrl = Path & DBAdapter1.Rows(0).Item("FinalSampleFile")  '������
                LFinalSampleFile.Visible = True
            Else
                LFinalSampleFile.Visible = False
            End If


        End If








    End Sub


    Protected Sub DNo_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles DNo.ServerChange

    End Sub

    Protected Sub DDivision_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDivision.TextChanged

    End Sub
End Class

