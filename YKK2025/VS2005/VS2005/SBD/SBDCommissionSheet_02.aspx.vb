Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDCommissionSheet_02
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
                    Top = 1600
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

        Response.Cookies("PGM").Value = "SBDCommissionSheet_01.aspx"
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

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDCommissionFilePath")



        SQL = "Select * From F_SBDCommissionSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")              'Buyer
            DVendor.Text = DBAdapter1.Rows(0).Item("Vendor")              'Vendor
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBackground.Text = DBAdapter1.Rows(0).Item("Background")              'Backround
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'Code
            DMapDate.Text = DBAdapter1.Rows(0).Item("MapDate")              'MapDate
            DSampleDate.Text = DBAdapter1.Rows(0).Item("SampleDate")              'SampleDate
            DLight.Text = DBAdapter1.Rows(0).Item("Light")


            DHalfFinishNo.Text = DBAdapter1.Rows(0).Item("HalfFinishNo")              'HalfFinishNot
            DMaterial.Text = DBAdapter1.Rows(0).Item("Material")
            DMaterialDetail.Text = DBAdapter1.Rows(0).Item("MaterialDetail")
            DMaterialDetail_1.Text = DBAdapter1.Rows(0).Item("MaterialDetail_1")              'MaterialDetail_1
            ' LMapFile
            If DBAdapter1.Rows(0).Item("RefMapFile") <> "" Then
                LRefMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("RefMapFile")  'LCertifcateFile
                LRefMapFile.Visible = True
            Else
                LRefMapFile.Visible = False
            End If

            DSample.Text = DBAdapter1.Rows(0).Item("Sample") 'Sample



            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")              'Remark
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'MapNo
            DMakeMap.Text = DBAdapter1.Rows(0).Item("MakeMap")         'MakeMap
            DLevel.Text = DBAdapter1.Rows(0).Item("Level")             'Level
            'LMapFile

            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  'LMapFile
                LMapFile.Visible = True
            Else
                LMapFile.Visible = False
            End If

            DReason1.Text = DBAdapter1.Rows(0).Item("Reason1") 'Reason

            DContent1.Text = DBAdapter1.Rows(0).Item("Content1")              'DContent1

            If DBAdapter1.Rows(0).Item("ContentFile1") <> "" Then
                LContentFile1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile1")  'LContentFile1 
                LContentFile1.Visible = True
            Else
                LContentFile1.Visible = False
            End If
            DContent2.Text = DBAdapter1.Rows(0).Item("Content2")              'DContent2

            If DBAdapter1.Rows(0).Item("ContentFile2") <> "" Then
                LContentFile2.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile2")  'LContentFile2 
                LContentFile2.Visible = True
            Else
                LContentFile2.Visible = False
            End If

            DContent3.Text = DBAdapter1.Rows(0).Item("Content3")              'DContent3

            If DBAdapter1.Rows(0).Item("ContentFile3") <> "" Then
                LContentFile3.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile3")  'LContentFile3 
                LContentFile3.Visible = True
            Else
                LContentFile3.Visible = False
            End If

            DContent4.Text = DBAdapter1.Rows(0).Item("Content4")              'DContent4

            If DBAdapter1.Rows(0).Item("ContentFile4") <> "" Then
                LContentFile4.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile4")  'LContentFile4 
                LContentFile4.Visible = True
            Else
                LContentFile4.Visible = False
            End If

            DContent5.Text = DBAdapter1.Rows(0).Item("Content5")              'DContent5

            If DBAdapter1.Rows(0).Item("ContentFile5") <> "" Then
                LContentFile5.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile5")  'LContentFile5 
                LContentFile5.Visible = True
            Else
                LContentFile5.Visible = False
            End If

            DContent6.Text = DBAdapter1.Rows(0).Item("Content6")              'DContent6

            If DBAdapter1.Rows(0).Item("ContentFile6") <> "" Then
                LContentFile6.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContentFile6")  'LContentFile6 
                LContentFile6.Visible = True
            Else
                LContentFile6.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map1") <> "" Then
                LMap1.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map1")  'LMap1
                LMap1.Visible = True
            Else
                LMap1.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map2") <> "" Then
                LMap2.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map2")  'LMap1
                LMap2.Visible = True
            Else
                LMap2.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map3") <> "" Then
                LMap3.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map3")  'LMap1
                LMap3.Visible = True
            Else
                LMap3.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map4") <> "" Then
                LMap4.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map4")  'LMap1
                LMap4.Visible = True
            Else
                LMap4.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map5") <> "" Then
                LMap5.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map5")  'LMap1
                LMap5.Visible = True
            Else
                LMap5.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("Map6") <> "" Then
                LMap6.NavigateUrl = Path & DBAdapter1.Rows(0).Item("Map6")  'LMap1
                LMap6.Visible = True
            Else
                LMap6.Visible = False
            End If

            Makemap1 = 0
            Makemap2 = 0
            Makemap3 = 0
            Makemap4 = 0
            Makemap5 = 0
            Makemap6 = 0

            DMakemap1.Text = DBAdapter1.Rows(0).Item("Makemap1")   'Makemap1
            DMakemap2.Text = DBAdapter1.Rows(0).Item("Makemap2")   'Makemap2
            DMakemap3.Text = DBAdapter1.Rows(0).Item("Makemap3")   'Makemap3
            DMakemap4.Text = DBAdapter1.Rows(0).Item("Makemap4")   'Makemap4
            DMakemap5.Text = DBAdapter1.Rows(0).Item("Makemap5")   'Makemap5
            DMakemap6.Text = DBAdapter1.Rows(0).Item("Makemap6")   'Makemap6

            
            DSupplier.Text = DBAdapter1.Rows(0).Item("Supplier") 'Suppiler

            DHalfFinishOrderNo.Text = DBAdapter1.Rows(0).Item("HalfFinishOrderNo")              'HalfFinishOrderNo
            If Mid(DBAdapter1.Rows(0).Item("HalfFinishDate").ToString, 1, 4) = "1900" Then
                DHalfFinishDate.Text = ""
            Else
                DHalfFinishDate.Text = DBAdapter1.Rows(0).Item("HalfFinishDate")              'HalfFinishDdate
            End If

            DMold.Text = DBAdapter1.Rows(0).Item("Mold")              'Mold

            DMoldPoint.Text = DBAdapter1.Rows(0).Item("MoldPoint")              'MoldPoint
            DSurfcolor.Text = DBAdapter1.Rows(0).Item("Surfcolor")              'Surfcolor
            DSampleqty.Text = DBAdapter1.Rows(0).Item("Sampleqty")              'Sampleqty

            If DBAdapter1.Rows(0).Item("QCReqFile") <> "" Then
                LQCReqFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QCReqFile")  'LQCReqFile
                LQCReqFile.Visible = True
            Else
                LQCReqFile.Visible = False
            End If

            DFQA.Text = DBAdapter1.Rows(0).Item("FQA")              'FQA
            DQARemark.Text = DBAdapter1.Rows(0).Item("QARemark")              'QARemark



            If DBAdapter1.Rows(0).Item("ForcastFile") <> "" Then
                LForcastFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ForcastFile")  ' LForcastFile
                LForcastFile.Visible = True
            Else
                LForcastFile.Visible = False
            End If



            If DBAdapter1.Rows(0).Item("QAFile") <> "" Then
                LQAFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("QAFile")  ' LQAFile
                LQAFile.Visible = True
            Else
                LQAFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("AuthorizeFile") <> "" Then
                LAuthorizeFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("AuthorizeFile")  ' LAuthorizeFile
                LAuthorizeFile.Visible = True
            Else
                LAuthorizeFile.Visible = False
            End If


            If DBAdapter1.Rows(0).Item("SampleFile") <> "" Then
                LSampleFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("SampleFile")  ' LSampleFile
                LSampleFile.Visible = True
            Else
                LSampleFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("ContactFile") <> "" Then
                LContactFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("ContactFile")  'LContactFile
                LContactFile.Visible = True
            Else
                LContactFile.Visible = False
            End If

            If DBAdapter1.Rows(0).Item("RefFile") <> "" Then
                LRefFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("RefFile")  'LContactFile
                LRefFile.Visible = True
            Else
                LRefFile.Visible = False
            End If

        End If

 
    End Sub

   
    Protected Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMakemap1.TextChanged

    End Sub

    Protected Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMakemap5.TextChanged

    End Sub

    Protected Sub DBackground_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBackground.TextChanged

    End Sub

    Protected Sub TextBox1_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DSampleDate.TextChanged

    End Sub

    Protected Sub DMapNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMapNo.TextChanged

    End Sub

    Protected Sub DContent5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DContent5.TextChanged

    End Sub

    Protected Sub DMakemap3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMakemap3.TextChanged

    End Sub
End Class

