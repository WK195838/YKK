Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls
Imports System.Drawing
Imports System.Windows.forms
 


Partial Class DASW_DISPOSALReasonURL
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
    Dim isOK As Boolean = True
    Dim Message As String = ""
    Dim OpenDir2 As String = ""
    Dim SignDate, OpenDate As String
    Dim UploadName As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim SQL1 As String
        SQL1 = " select ITEM as '�d������',CHINESE as '����',JAPAN as  '���' from  F_DISPOSALReason "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL1)
        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()

    End Sub






End Class

