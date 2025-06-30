Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
 


Partial Class SBDCommissionSheet_01
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
    Dim flag As Integer
    Dim PAction1 As Integer



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�
        ' TopPosition()           '���s��RequestedField��Top��m
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

        Response.Cookies("PGM").Value = "SBDCommissionSheet_01.aspx"
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

        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("SBDCommissionFilePath")



        SQL = "Select * From F_SBDCommissionSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")              'AppDate
            SetFieldData("Buyer", DBAdapter1.Rows(0).Item("Buyer"))    'Buyer
            DVendor.Text = DBAdapter1.Rows(0).Item("Vendor")              'Vendor
            DDivision.Text = DBAdapter1.Rows(0).Item("Division")              'Division
            DAppper.Text = DBAdapter1.Rows(0).Item("Appper")              'Appper
            DBackground.Text = DBAdapter1.Rows(0).Item("Background")              'Backround
            DCode.Text = DBAdapter1.Rows(0).Item("Code")              'Code
            DMapDate.Value = DBAdapter1.Rows(0).Item("MapDate")              'MapDate
            DSampleDate.Value = DBAdapter1.Rows(0).Item("SampleDate")              'SampleDate
            SetFieldData("Light", DBAdapter1.Rows(0).Item("Light")) 'Light    

            DHalfFinishNo.Text = DBAdapter1.Rows(0).Item("HalfFinishNo")              'HalfFinishNot
            SetFieldData("Material", DBAdapter1.Rows(0).Item("Material"))    'Material
            SetFieldData("MaterialDetail", DBAdapter1.Rows(0).Item("MaterialDetail"))    'MaterialDetail
            DMaterialDetail_1.Text = DBAdapter1.Rows(0).Item("MaterialDetail_1")              'MaterialDetail_1



            If DBAdapter1.Rows(0).Item("RefMapFile") <> "" Then
                LRefMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("RefMapFile")  'LCertifcateFile
                LRefMapFile.Visible = True
            Else
                LRefMapFile.Visible = False
            End If




            SetFieldData("Sample", DBAdapter1.Rows(0).Item("Sample"))    'Sample


            DRemark.Text = DBAdapter1.Rows(0).Item("Remark")              'Remark
            DMapNo.Text = DBAdapter1.Rows(0).Item("MapNo")              'MapNo
            SetFieldData("MakeMap", "")    'MakeMap
            DMakeMapU.Text = DBAdapter1.Rows(0).Item("MakeMap")              'MapNo
            SetFieldData("Level", DBAdapter1.Rows(0).Item("Level"))    'Level
            'LMapFile

            If Trim(DBAdapter1.Rows(0).Item("MapFile")) <> "" Then
                LMapFile.NavigateUrl = Path & DBAdapter1.Rows(0).Item("MapFile")  'LMapFile
                LMapFile.Visible = True
            Else
                LMapFile.Visible = False
            End If


            'SetFieldData("2-Makemap1", DBAdapter1.Rows(0).Item("MakeMap1"))    'MakeMap
            'SetFieldData("2-Makemap2", DBAdapter1.Rows(0).Item("MakeMap2"))    'MakeMap
            'SetFieldData("2-Makemap3", DBAdapter1.Rows(0).Item("MakeMap3"))    'MakeMap
            'SetFieldData("2-Makemap4", DBAdapter1.Rows(0).Item("MakeMap4"))    'MakeMap
            'SetFieldData("2-Makemap5", DBAdapter1.Rows(0).Item("MakeMap5"))    'MakeMap
            'SetFieldData("2-Makemap6", DBAdapter1.Rows(0).Item("MakeMap6"))    'MakeMap
            DMakeMap1.Text = DBAdapter1.Rows(0).Item("MakeMap1")
            DMakeMap2.Text = DBAdapter1.Rows(0).Item("MakeMap2")
            DMakeMap3.Text = DBAdapter1.Rows(0).Item("MakeMap3")
            DMakeMap4.Text = DBAdapter1.Rows(0).Item("MakeMap4")
            DMakeMap5.Text = DBAdapter1.Rows(0).Item("MakeMap5")
            DMakeMap6.Text = DBAdapter1.Rows(0).Item("MakeMap6")




            '  Dim a As String
            'a = DBAdapter1.Rows(0).Item("MakeMap1")
            ' DMakemap1.Text = DBAdapter1.Rows(0).Item("MakeMap1")


            SetFieldData("Reason1", DBAdapter1.Rows(0).Item("Reason1"))    'Reason

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








            SetFieldData("Supplier", DBAdapter1.Rows(0).Item("Supplier"))    'Suppiler

            DHalfFinishOrderNo.Text = DBAdapter1.Rows(0).Item("HalfFinishOrderNo")              'HalfFinishOrderNo
            If Mid(DBAdapter1.Rows(0).Item("HalfFinishDate").ToString, 1, 4) = "1900" Then
                DHalfFinishDate.Value = ""
            Else
                DHalfFinishDate.Value = DBAdapter1.Rows(0).Item("HalfFinishDate")              'HalfFinishDdate
            End If

            DMold.Text = DBAdapter1.Rows(0).Item("Mold")              'Mold

            DMoldPoint.Text = DBAdapter1.Rows(0).Item("MoldPoint")              'MoldPoint

            If DBAdapter1.Rows(0).Item("Surfcolor") <> "" Then
                DSurfcolor.Value = DBAdapter1.Rows(0).Item("Surfcolor")              'Surfcolor
                LSurfColor.Visible = True
                LSurfColor.Text = "________"
                LSurfColor.NavigateUrl = "SBDSurfaceSheet_02.aspx?pFormNo=" + RTrim(CStr(DBAdapter1.Rows(0).Item("SurfFormNo"))) + "&pFormSno=" + RTrim(CStr(DBAdapter1.Rows(0).Item("SurfFormSno")))

            End If
            'DSurfcolor.Value = DBAdapter1.Rows(0).Item("Surfcolor")              'Surfcolor
            DSurfcolor1.Text = DBAdapter1.Rows(0).Item("Surfcolor1")              'Surfcolor
            'DSampleQty.Text = DBAdapter1.Rows(0).Item("SampleQty")


            SetFieldData("SampleQty", DBAdapter1.Rows(0).Item("SampleQty"))   'Suppiler

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
                    Top = 1200
                Else
                    DReasonCode.Visible = False     '����z�ѥN�X
                    DReason.Visible = False         '����z��
                    DReasonDesc.Visible = False     '�����L����
                    Top = 1500
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
            Top = 1400
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
        BDate1.Attributes.Add("onClick", "calendarPicker('DMapDate')")
        BDate2.Attributes.Add("onClick", "calendarPicker('DSampleDate')")
        BDate3.Attributes.Add("onClick", "calendarPicker('DHalfFinishDate')")
        BColor.Attributes.Add("onClick", "openNoPicker('DSurfcolor')")
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
                    Top = 1600
                Else
                    If DDelay.Visible = True Then
                        Top = 1600
                    Else
                        Top = 1600
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

        'Vendor
        Select Case FindFieldInf("Vendor")
            Case 0  '���
                DVendor.BackColor = Color.LightGray
                DVendor.ReadOnly = False
                DVendor.Visible = True
            Case 1  '�ק�+�ˬd
                DVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DVendorRqd", "DVendor", "���`�G�ݿ�JVendor")
                DVendor.Visible = True
            Case 2  '�ק�
                DVendor.BackColor = Color.Yellow
                DVendor.Visible = True
            Case Else   '����
                DVendor.Visible = False
        End Select
        If pPost = "New" Then DVendor.Text = ""

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

        '�}�o�I��
        Select Case FindFieldInf("Background")
            Case 0  '���
                DBackground.BackColor = Color.LightGray
                DBackground.Visible = True
                DBackground.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DBackground.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBackgroundtRqd", "DBackground", "���`�G�ݿ�J�}�o�I��")
                DBackground.Visible = True
            Case 2  '�ק�
                DBackground.BackColor = Color.Yellow
                DBackground.Visible = True
            Case Else   '����
                DBackground.Visible = False
        End Select
        If pPost = "New" Then DBackground.Text = ""

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

        '�ϭ��Ʊ���
        Select Case FindFieldInf("MapDate")
            Case 0  '���
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "lightgrey")
                DMapDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = True
                BDate1.Disabled = True
            Case 1  '�ק�+�ˬd
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "greenyellow")
                DMapDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DMapDateRqd", "DMapDate", "���`�G�ݿ�J�ϭ��Ʊ���")
                BDate1.Disabled = False
                BDate1.Disabled = False
            Case 2  '�ק�
                DMapDate.Visible = True
                DMapDate.Style.Add("background-color", "yellow")
                DMapDate.Attributes.Add("readonly", "true")
                BDate1.Disabled = False
                BDate1.Disabled = False
            Case Else   '����
                DMapDate.Visible = False
                BDate1.Disabled = True
                BDate1.Disabled = True
        End Select
        If pPost = "New" Then DMapDate.Value = ""

        '�˫~�Ʊ���

        Select Case FindFieldInf("SampleDate")
            Case 0  '���
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "lightgrey")
                DSampleDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = True
                BDate2.Disabled = True
            Case 1  '�ק�+�ˬd
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "greenyellow")
                DSampleDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DSampledateRqd", "Dsampledate", "���`�G�ݿ�J�˫~�Ʊ���")
                BDate2.Disabled = False
                BDate2.Disabled = False
            Case 2  '�ק�
                DSampleDate.Visible = True
                DSampleDate.Style.Add("background-color", "yellow")
                DSampleDate.Attributes.Add("readonly", "true")
                BDate2.Disabled = False
                BDate2.Disabled = False
            Case Else   '����
                DSampleDate.Visible = False
                BDate2.Disabled = True
                BDate2.Disabled = True
        End Select
        If pPost = "New" Then DSampleDate.Value = ""

        '���y��
        Select Case FindFieldInf("Light")
            Case 0  '���
                DLight.BackColor = Color.LightGray
                DLight.Visible = True

            Case 1  '�ק�+�ˬd
                DLight.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLightRqd", "DLight", "���`�G�ݿ�J���y��")
                DLight.Visible = True
            Case 2  '�ק�
                DLight.BackColor = Color.Yellow
                DLight.Visible = True
            Case Else   '����
                DLight.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Light", "ZZZZZZ")

        '�b���~�Ƹ�
        Select Case FindFieldInf("HalffinishNo")
            Case 0  '���
                DHalfFinishNo.BackColor = Color.LightGray
                DHalfFinishNo.Visible = True
                DHalfFinishNo.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DHalfFinishNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DHalfFinishNoRqd", "DHalfFinishNo", "���`�G�ݿ�J�b���~�Ƹ�")
                DHalfFinishNo.Visible = True
            Case 2  '�ק�
                DHalfFinishNo.BackColor = Color.Yellow
                DHalfFinishNo.Visible = True
            Case Else   '����
                DHalfFinishNo.Visible = False
        End Select
        If pPost = "New" Then DHalfFinishNo.Text = ""

        '�[�u���� Material

        Select Case FindFieldInf("Material")
            Case 0  '���
                DMaterial.BackColor = Color.LightGray
                DMaterial.Visible = True
            Case 1  '�ק�+�ˬd
                DMaterial.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialRqd", "DMaterial", "���`�G�ݿ�J�b���~�Ƹ�")
                DMaterial.Visible = True
            Case 2  '�ק�
                DMaterial.BackColor = Color.Yellow
                DMaterial.Visible = True
            Case Else   '����
                DMaterial.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Material", "ZZZZZZ")


        '�[�u���� DMaterialDetailDetail 
        Select Case FindFieldInf("MaterialDetail")
            Case 0  '���
                DMaterialDetail.BackColor = Color.LightGray
                DMaterialDetail.Visible = True
            Case 1  '�ק�+�ˬd
                DMaterialDetail.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetailRqd", "DMaterialDetail", "���`�G�ݿ�J�[�u����")
                DMaterialDetail.Visible = True
            Case 2  '�ק�
                DMaterialDetail.BackColor = Color.Yellow
                DMaterialDetail.Visible = True
            Case Else   '����
                DMaterialDetail.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MaterialDetail", "ZZZZZZ")

        '�[�u���� DMaterialDetailDetail 
        Select Case FindFieldInf("MaterialDetail_1")
            Case 0  '���
                DMaterialDetail_1.BackColor = Color.LightGray
                DMaterialDetail_1.Visible = True
                DMaterialDetail_1.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DMaterialDetail_1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMaterialDetail_1Rqd", "DMaterialDetail_1", "���`�G�[�u�����Ƶ�")
                DMaterialDetail_1.Visible = True
            Case 2  '�ק�
                DMaterialDetail_1.BackColor = Color.Yellow
                DMaterialDetail_1.Visible = True
            Case Else   '����
                DMaterialDetail_1.Visible = False
        End Select
        If pPost = "New" Then DMaterialDetail_1.Text = ""


        '���
        Select Case FindFieldInf("RefMapFile")

            Case 0  '���
                DRefMapFile.Visible = False
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DRefMapFileRqd", "DRefMapFile", "���`�G�ݿ�J���")
                DRefMapFile.Visible = True
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DRefMapFile.Visible = True
                DRefMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DRefMapFile.Visible = False
        End Select
        LRefMapFile.Visible = False


        '�˫~
        Select Case FindFieldInf("Sample")
            Case 0  '���DSample
                DSample.BackColor = Color.LightGray
                DSample.Visible = True
            Case 1  '�ק�+�ˬd
                DSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleRqd", "DSample", "���`�G�ݿ�J�˫~")
                DSample.Visible = True
            Case 2  '�ק�
                DSample.BackColor = Color.Yellow
                DSample.Visible = True
            Case Else   '����
                DSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sample", "ZZZZZZ")

        '�Ƶ�
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
        If pPost = "New" Then DRemark.Text = ""


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
        If pPost = "New" Then DMapNo.Text = ""


        '�s�Ϫ�
        Select Case FindFieldInf("MakeMapU")
            Case 0  '���
                DMakeMapU.BackColor = Color.LightGray
                DMakeMapU.Visible = True
            Case 1  '�ק�+�ˬd
                DMakeMapU.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeMapURqd", "DMakeMapU", "���`�G�ݿ�J�s�Ϫ�")
                DMakeMapU.Visible = True
            Case 2  '�ק�
                DMakeMapU.BackColor = Color.Yellow
                DMakeMapU.Visible = True
            Case Else   '����
                DMakeMapU.Visible = False
        End Select


        '�s�Ϫ�
        Select Case FindFieldInf("MakeMap")
            Case 0  '���
                DMakeMap.BackColor = Color.LightGray
                DMakeMap.Visible = True
            Case 1  '�ק�+�ˬd
                DMakeMap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeMapRqd", "DMakeMap", "���`�G�ݿ�J�s�Ϫ�")
                DMakeMap.Visible = True
            Case 2  '�ק�
                DMakeMap.BackColor = Color.Yellow
                DMakeMap.Visible = True
            Case Else   '����
                DMakeMap.Visible = False
        End Select


        If pPost = "New" Then SetFieldData("MakeMap", "ZZZZZZ")


        '������
        Select Case FindFieldInf("Level")
            Case 0  '���
                DLevel.BackColor = Color.LightGray
                DLevel.Visible = True

            Case 1  '�ק�+�ˬd
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "���`�G�ݿ�J������")
                DLevel.Visible = True
            Case 2  '�ק�
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '����
                DLevel.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")



        '����
        Select Case FindFieldInf("MapFile")
            Case 0  '���
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DMapFile.Attributes.Add("readonly", "true")
            Case 1  '�ק�+�ˬd
                If LMapFile.NavigateUrl <> "" Then
                    ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "���`�G�ݿ�J����")
                End If

                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DMapFile.Visible = False
        End Select

        LMapFile.Visible = False







        '��]���O
        Select Case FindFieldInf("Reason1")


            Case 0  '���
                DReason1.BackColor = Color.LightGray
                DReason1.Visible = True
            Case 1  '�ק�+�ˬd
                DReason1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DReason1Rqd", "DReason1", "���`�G�ݿ�J��]���O")
                DReason1.Visible = True
            Case 2  '�ק�
                DReason1.BackColor = Color.Yellow
                DReason1.Visible = True
            Case Else   '����
                DReason1.Visible = False

        End Select
        If pPost = "New" Then SetFieldData("Reason1", "ZZZZZZ")


        '�ק鷺�e-1

        Dim sql As String
        Sql = "Select * From F_SBDCommissionSheet "
        Sql = Sql & " Where FormNo =  '" & wFormNo & "'"
        Sql = Sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(Sql)

        If DBAdapter1.Rows.Count > 0 Then
            flag = DBAdapter1.Rows(0).Item("Flag")
        End If




        Dim idx As Integer
        idx = FindFieldInf("1-Content")
        If flag > 0 And (wStep = 30) Then
            idx = 2
        End If

        Select Case idx
            Case 0  '���
                DContent1.BackColor = Color.LightGray
                DContent1.Visible = True
                DContent1.ReadOnly = True
                DContent2.BackColor = Color.LightGray
                DContent2.Visible = True
                DContent2.ReadOnly = True
                DContent3.BackColor = Color.LightGray
                DContent3.Visible = True
                DContent3.ReadOnly = True
                DContent4.BackColor = Color.LightGray
                DContent4.Visible = True
                DContent4.ReadOnly = True
                DContent5.BackColor = Color.LightGray
                DContent5.Visible = True
                DContent5.ReadOnly = True
                DContent6.BackColor = Color.LightGray
                DContent6.Visible = True
                DContent6.ReadOnly = True


            Case 1  '�ק�+�ˬd
                DContent1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent1", "���`�G�ݿ�J�ק鷺�e-1")
                DContent1.Visible = True
                DContent2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent2", "���`�G�ݿ�J�ק鷺�e-2")
                DContent2.Visible = True
                DContent3.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent3", "���`�G�ݿ�J�ק鷺�e-3")
                DContent3.Visible = True
                DContent4.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent4", "���`�G�ݿ�J�ק鷺�e-4")
                DContent4.Visible = True
                DContent5.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent5", "���`�G�ݿ�J�ק鷺�e-5")
                DContent5.Visible = True
                DContent6.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContent1Rqd", "DContent6", "���`�G�ݿ�J�ק鷺�e-6")
                DContent6.Visible = True

            Case 2  '�ק�
                DContent1.BackColor = Color.Yellow
                DContent1.Visible = True
                DContent2.BackColor = Color.Yellow
                DContent2.Visible = True
                DContent3.BackColor = Color.Yellow
                DContent3.Visible = True
                DContent4.BackColor = Color.Yellow
                DContent4.Visible = True
                DContent5.BackColor = Color.Yellow
                DContent5.Visible = True
                DContent6.BackColor = Color.Yellow
                DContent6.Visible = True

            Case Else   '����
                DContent1.Visible = False
                DContent2.Visible = False
                DContent3.Visible = False
                DContent4.Visible = False
                DContent5.Visible = False
                DContent6.Visible = False

        End Select
        If pPost = "New" Then DContent1.Text = ""
        If pPost = "New" Then DContent2.Text = ""
        If pPost = "New" Then DContent3.Text = ""
        If pPost = "New" Then DContent4.Text = ""
        If pPost = "New" Then DContent5.Text = ""
        If pPost = "New" Then DContent6.Text = ""





        idx = FindFieldInf("1-Contenfile")
        If flag > 0 And (wStep = 30) Then
            idx = 2
        End If
        '�ק鷺�e����
        Select Case idx
            Case 0  '���
                DContentFile1.Visible = False
                DContentFile1.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile1.Attributes.Add("readonly", "true")
                DContentFile2.Visible = False
                DContentFile2.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile2.Attributes.Add("readonly", "true")
                DContentFile3.Visible = False
                DContentFile3.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile3.Attributes.Add("readonly", "true")
                DContentFile4.Visible = False
                DContentFile4.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile4.Attributes.Add("readonly", "true")
                DContentFile5.Visible = False
                DContentFile5.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile5.Attributes.Add("readonly", "true")
                DContentFile6.Visible = False
                DContentFile6.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DContentFile6.Attributes.Add("readonly", "true")
                LContentFile1.Visible = True
                LContentFile1.BackColor = Color.LightGray
                LContentFile2.Visible = True
                LContentFile2.BackColor = Color.LightGray
                LContentFile3.Visible = True
                LContentFile3.BackColor = Color.LightGray
                LContentFile4.Visible = True
                LContentFile4.BackColor = Color.LightGray
                LContentFile5.Visible = True
                LContentFile5.BackColor = Color.LightGray
                LContentFile6.Visible = True
                LContentFile6.BackColor = Color.LightGray


            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DContentFile1Rqd", "DContentFile1", "���`�G�ݿ�J����")
                DContentFile1.Visible = True
                DContentFile1.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile2Rqd", "DContentFile2", "���`�G�ݿ�J����")
                DContentFile2.Visible = True
                DContentFile2.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile3Rqd", "DContentFile3", "���`�G�ݿ�J����")
                DContentFile3.Visible = True
                DContentFile3.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile4Rqd", "DContentFile4", "���`�G�ݿ�J����")
                DContentFile4.Visible = True
                DContentFile4.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile5Rqd", "DContentFile5", "���`�G�ݿ�J����")
                DContentFile5.Visible = True
                DContentFile5.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DContentFile6Rqd", "DContentFile6", "���`�G�ݿ�J����")
                DContentFile6.Visible = True
                DContentFile6.Style.Add("BACKGROUND-COLOR", "GreenYellow")

            Case 2  '�ק�
                DContentFile1.Visible = True
                DContentFile1.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile2.Visible = True
                DContentFile2.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile3.Visible = True
                DContentFile3.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile4.Visible = True
                DContentFile4.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile5.Visible = True
                DContentFile5.Style.Add("BACKGROUND-COLOR", "Yellow")
                DContentFile6.Visible = True
                DContentFile6.Style.Add("BACKGROUND-COLOR", "Yellow")

            Case Else   '����
                DContentFile1.Visible = False
                DContentFile2.Visible = False
                DContentFile3.Visible = False
                DContentFile4.Visible = False
                DContentFile5.Visible = False
                DContentFile6.Visible = False
        End Select
        LContentFile1.Visible = False
        LContentFile2.Visible = False
        LContentFile3.Visible = False
        LContentFile4.Visible = False
        LContentFile5.Visible = False
        LContentFile6.Visible = False









        Select Case FindFieldInf("2-Makemap")
            Case 0  '���
                DMakeMap1.BackColor = Color.LightGray
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.LightGray
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.LightGray
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.LightGray
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.LightGray
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.LightGray
                DMakeMap6.Visible = True

                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True
            Case 1  '�ק�+�ˬd
                DMakeMap1.BackColor = Color.GreenYellow
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.GreenYellow
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.GreenYellow
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.GreenYellow
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.GreenYellow
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.GreenYellow
                DMakeMap6.Visible = True
                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True
            Case 2  '�ק�
                DMakeMap1.BackColor = Color.Yellow
                DMakeMap1.Visible = True
                DMakeMap2.BackColor = Color.Yellow
                DMakeMap2.Visible = True
                DMakeMap3.BackColor = Color.Yellow
                DMakeMap3.Visible = True
                DMakeMap4.BackColor = Color.Yellow
                DMakeMap4.Visible = True
                DMakeMap5.BackColor = Color.Yellow
                DMakeMap5.Visible = True
                DMakeMap6.BackColor = Color.Yellow
                DMakeMap6.Visible = True
                LMap1.Visible = True
                LMap2.Visible = True
                LMap3.Visible = True
                LMap4.Visible = True
                LMap5.Visible = True
                LMap6.Visible = True

            Case Else   '����

                DMakeMap1.Visible = False
                DMakeMap2.Visible = False
                DMakeMap3.Visible = False
                DMakeMap4.Visible = False
                DMakeMap5.Visible = False
                DMakeMap6.Visible = False

                LMap1.Visible = False
                LMap2.Visible = False
                LMap3.Visible = False
                LMap4.Visible = False
                LMap5.Visible = False
                LMap6.Visible = False
        End Select











        '�~�`��
        Select Case FindFieldInf("Supplier")
            Case 0  '���
                DSupplier.BackColor = Color.LightGray
                DSupplier.Visible = True

            Case 1  '�ק�+�ˬd
                DSupplier.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSupplierRqd", "DSupplier", "���`�G�ݿ�J�~�`��")
                DSupplier.Visible = True
            Case 2  '�ק�
                DSupplier.BackColor = Color.Yellow
                DSupplier.Visible = True
            Case Else   '����
                DSupplier.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Supplier", "ZZZZZZ")

        '�b���~�q��NO.
        Select Case FindFieldInf("HalfFinishOrderNo")
            Case 0  '���
                DHalfFinishOrderNo.BackColor = Color.LightGray
                DHalfFinishOrderNo.Visible = True
                DHalfFinishOrderNo.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DHalfFinishOrderNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DHalfFinishOrderNoRqd", "DHalfFinishOrderNo", "���`�G�ݿ�J�b���~�q��NO.")
                DHalfFinishOrderNo.Visible = True
            Case 2  '�ק�
                DHalfFinishOrderNo.BackColor = Color.Yellow
                DHalfFinishOrderNo.Visible = True
            Case Else   '����
                DHalfFinishOrderNo.Visible = False
        End Select
        If pPost = "New" Then DHalfFinishOrderNo.Text = ""

        '�b���~�w�q������

        Select Case FindFieldInf("HalfFinishDate")
            Case 0  '���
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "lightgrey")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = True

            Case 1  '�ק�+�ˬd
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "greenyellow")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DHalfFinishDateRqd", "DHalfFinishDate", "���`�G�ݿ�J�b���~�w�q������")
                BDate3.Disabled = False

            Case 2  '�ק�
                DHalfFinishDate.Visible = True
                DHalfFinishDate.Style.Add("background-color", "yellow")
                DHalfFinishDate.Attributes.Add("readonly", "true")
                BDate3.Disabled = False

            Case Else   '����
                DHalfFinishDate.Visible = False
                BDate3.Disabled = True

        End Select
        If pPost = "New" Then DHalfFinishDate.Value = ""

        '�Ҩ�
        Select Case FindFieldInf("Mold")
            Case 0  '���
                DMold.BackColor = Color.LightGray
                DMold.Visible = True
                DMold.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DMold.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldRqd", "DMold", "���`�G�ݿ�J�Ҩ�")
                DMold.Visible = True
            Case 2  '�ק�
                DMold.BackColor = Color.Yellow
                DMold.Visible = True
            Case Else   '����
                DMold.Visible = False
        End Select
        If pPost = "New" Then DMold.Text = ""



        '�ި�
        Select Case FindFieldInf("MoldPoint")
            Case 0  '���
                DMoldPoint.BackColor = Color.LightGray
                DMoldPoint.Visible = True
                DMoldPoint.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DMoldPoint.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMoldPointRqd", "DMoldPoint", "���`�G�ݿ�J�ި�")
                DMoldPoint.Visible = True
            Case 2  '�ק�
                DMoldPoint.BackColor = Color.Yellow
                DMoldPoint.Visible = True
            Case Else   '����
                DMoldPoint.Visible = False
        End Select
        If pPost = "New" Then DMoldPoint.Text = ""

        '���C��

        Select Case FindFieldInf("Surfcolor")
            Case 0  '���
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "lightgrey")
                DSurfcolor.Attributes.Add("readonly", "true")
                BColor.Disabled = True
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If


            Case 1  '�ק�+�ˬd
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "greenyellow")
                DSurfcolor.Attributes.Add("readonly", "true")
                ShowRequiredFieldValidator("DSurfcolorRqd", "DSurfcolor", "���`�G�ݿ�J���C��")
                BColor.Disabled = False
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If



            Case 2  '�ק�
                DSurfcolor.Visible = True
                DSurfcolor.Style.Add("background-color", "yellow")
                DSurfcolor.Attributes.Add("readonly", "true")
                BColor.Disabled = False
                If DSurfcolor.Value = "" Then
                    LSurfColor.Visible = False
                Else
                    LSurfColor.Visible = True
                End If


            Case Else   '����
                DSurfcolor.Visible = False
                BColor.Disabled = True
                LSurfColor.Visible = False

        End Select
        If pPost = "New" Then DSurfcolor.Value = ""


        '���C�� 1
        Select Case FindFieldInf("Surfcolor")
            Case 0  '���
                DSurfcolor1.BackColor = Color.LightGray
                DSurfcolor1.Visible = True
                DSurfcolor1.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DSurfcolor1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSurfcolor1Rqd", "DSurfcolor1", "���`�G���C��")
                DSurfcolor1.Visible = True
            Case 2  '�ק�
                DSurfcolor1.BackColor = Color.Yellow
                DSurfcolor1.Visible = True
            Case Else   '����
                DSurfcolor1.Visible = False
        End Select
        If pPost = "New" Then DSurfcolor1.Text = ""



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



        '���R�̿��
        Select Case FindFieldInf("QCReqFile")
            Case 0  '���
                DQCReqFile.Visible = False
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQCReqFile.Attributes.Add("readonly", "true")
                LQCReqFile.Visible = True
                LQCReqFile.BackColor = Color.LightGray
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DQCReqFileRqd", "DQCReqFile", "���`�G�ݿ�J���R�̿��")
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DQCReqFile.Visible = True
                DQCReqFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DQCReqFile.Visible = False
        End Select
        LQCReqFile.Visible = False

        'EA�P�w
        Select Case FindFieldInf("FQA")
            Case 0  '���
                DFQA.BackColor = Color.LightGray
                DFQA.Visible = True
                DFQA.ReadOnly = True

            Case 1  '�ק�+�ˬd
                DFQA.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFQARqd", "DFQA", "���`�G�ݿ�JEA�P�w")
                DFQA.Visible = True
            Case 2  '�ק�
                DFQA.BackColor = Color.Yellow
                DFQA.Visible = True
            Case Else   '����
                DFQA.Visible = False
        End Select
        If pPost = "New" Then DFQA.Text = ""

        '�~��Ƶ�
        Select Case FindFieldInf("QARemark")
            Case 0  '���
                DQARemark.BackColor = Color.LightGray
                DQARemark.Visible = True
                DQARemark.ReadOnly = True
            Case 1  '�ק�+�ˬd
                DQARemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DQARemarkRqd", "DQARemark", "���`�G�ݿ�J�~��Ƶ�")
                DQARemark.Visible = True
            Case 2  '�ק�
                DQARemark.BackColor = Color.Yellow
                DQARemark.Visible = True
            Case Else   '����
                DQARemark.Visible = False
        End Select
        If pPost = "New" Then DQARemark.Text = ""

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

        '�~�����i��
        Select Case FindFieldInf("QAFile")
            Case 0  '���
                DQAFile.Visible = False
                DQAFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DQAFile.Attributes.Add("readonly", "true")
                LQAFile.Visible = True
                LQAFile.BackColor = Color.LightGray
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


        '�T�{��
        Select Case FindFieldInf("AuthorizeFile")
            Case 0  '���
                DAuthorizeFile.Visible = False
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DAuthorizeFile.Attributes.Add("readonly", "true")
                LAuthorizeFile.Visible = True
                LAuthorizeFile.BackColor = Color.LightGray
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DAuthorizeFileRqd", "DAuthorizeFile", "���`�G�ݿ�J�T�{��")
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DAuthorizeFile.Visible = True
                DAuthorizeFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DAuthorizeFile.Visible = False
        End Select
        LAuthorizeFile.Visible = False


        '�̲׼˫~��
        Select Case FindFieldInf("SampleFile")
            Case 0  '���
                DSampleFile.Visible = False
                DSampleFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DSampleFile.Attributes.Add("readonly", "true")
                LSampleFile.Visible = True
                LSampleFile.BackColor = Color.LightGray
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DSampleFileRqd", "DSampleFile", "���`�G�ݿ�J�̲׼˫~��")
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DSampleFile.Visible = True
                DSampleFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DSampleFile.Visible = False
        End Select
        LSampleFile.Visible = False


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


        '�䥦���
        Select Case FindFieldInf("RefFile")
            Case 0  '���
                DRefFile.Visible = False
                DRefFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DRefFile.Attributes.Add("readonly", "true")
                LRefFile.Visible = True
                LRefFile.BackColor = Color.LightGray
            Case 1  '�ק�+�ˬd
                ShowRequiredFieldValidator("DRefFileRqd", "DRefFile", "���`�G�ݿ�J�䥦���")
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '�ק�
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '����
                DRefFile.Visible = False
        End Select
        LRefFile.Visible = False



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

        '���y��
        If pFieldName = "Light" Then
            DLight.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLight.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Light' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DLight.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLight.Items.Add(ListItem1)
                Next
            End If
        End If

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



        '�[�u����
        If pFieldName = "Material" Then
            DMaterial.Items.Clear()
            DMaterialDetail.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMaterial.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'MATERIAL' "
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMaterial.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMaterial.Items.Add(ListItem1)
                Next
            End If
        End If



        '�[�u����
        If pFieldName = "MaterialDetail" Then
            DMaterialDetail.Items.Clear()
            DMaterialDetail.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMaterialDetail.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3002' "
                SQL = SQL & "   And dkey = '" + DMaterial.SelectedValue + "'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMaterialDetail.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMaterialDetail.Items.Add(ListItem1)
                Next
            End If
        End If






        '���˫~
        If pFieldName = "Sample" Then
            DSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Sample'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DSample.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSample.Items.Add(ListItem1)
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


        '�s�Ϫ�
        If pFieldName = "MakeMap" Then

            DMakeMap.Items.Clear()


            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMakeMap.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'MakeMap'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DMakeMap.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMakeMap.Items.Add(ListItem1)


                Next
            End If
        End If






        'LEVEL
        If pFieldName = "Level" Then
            DLevel.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLevel.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Level'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DLevel.Items.Add("")
                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLevel.Items.Add(ListItem1)
                Next
            End If
        End If

        '��]���O
        If pFieldName = "Reason1" Then
            DReason1.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReason1.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select Data From M_referp "
                SQL = SQL & " Where cat  = '3001' "
                SQL = SQL & "   And dkey = 'Reason'"
                Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
                DReason1.Items.Add("")

                For i = 0 To DBAdapter1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
                    ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReason1.Items.Add(ListItem1)
                Next
            End If
        End If

        '�~�`��
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

                    Dim URL As String = uCommon.GetAppSetting("RedirectURL") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
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
    Private Sub BOK_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If wStep = 70 Then
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
                ErrCode = oCommon.CommissionNo("003001", wFormSno, wStep, DNo.Text) '��渹�X, ���y����, �u�{, �e�U��No
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
                    PAction1 = pAction '���F60�u�{
                    Select Case wStep
                        Case 30
                            If pAction = 0 Then  '2013/01/10   jessica
                                wAllocateID = oCommon.GetUserID(DMakeMap.Text)
                            End If
                        Case 50
                            If pAction = 1 Then  '2013/01/10   jessica
                                If DMakeMap.SelectedValue = "" Then
                                    If flag = 0 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMapU.Text)
                                    ElseIf flag = 1 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap1.Text)
                                    ElseIf flag = 2 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap2.Text)
                                    ElseIf flag = 3 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap3.Text)
                                    ElseIf flag = 4 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap4.Text)
                                    ElseIf flag = 5 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap5.Text)
                                    ElseIf flag = 6 Then
                                        wAllocateID = oCommon.GetUserID(DMakeMap5.Text)
                                    End If
                                Else
                                    wAllocateID = oCommon.GetUserID(DMakeMap.SelectedValue)
                                End If
                               
                            End If
                        Case Else

                    End Select

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
                            Select Case wStep
                                Case 60
                                    If PAction1 = 1 Then
                                        '��sflag ���s�}�l
                                        SQL = " update f_sbdcommissionsheet "
                                        SQL = SQL + " Set flag = 0"
                                        SQL = SQL + " ,mapno=''"
                                        SQL = SQL + " ,makemap =''"
                                        SQL = SQL + " ,mapfile = ''"
                                        SQL = SQL + " ,map1 =''"
                                        SQL = SQL + " ,makemap1 =''"
                                        SQL = SQL + " ,map2 = ''"
                                        SQL = SQL + " ,makemap2 = ''"
                                        SQL = SQL + " ,map3 = ''"
                                        SQL = SQL + " ,makemap3 = ''"
                                        SQL = SQL + " ,map4 = ''"
                                        SQL = SQL + " ,makemap4 = ''"
                                        SQL = SQL + " ,map5 = ''"
                                        SQL = SQL + " ,makemap5 = ''"
                                        SQL = SQL + " ,ModifyUser = '" & Request.QueryString("pUserID") & "',"
                                        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                                        SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                                        uDataBase.ExecuteNonQuery(SQL)

                                    End If
                                Case Else

                            End Select
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
            If ErrCode = 45 Then Message = "���`�G�ݿ�J�~�`��!"
            If ErrCode = 46 Then Message = "���`�G�ݿ�J�b���~�q��No!"
            If ErrCode = 47 Then Message = "���`�G�ݿ�J�b���~�w�q������!"
            If ErrCode = 48 Then Message = "���`�G�ݿ�J�Ҩ�!"
            If ErrCode = 49 Then Message = "���`�G�ݿ�J�ҫ�!"
            If ErrCode = 50 Then Message = "���`�G�ݿ�J�ި�!"

            If ErrCode = 51 Then Message = "���`�G�ݿ�J���C��!"
            If ErrCode = 52 Then Message = "���`�G�ݿ�J�˫~�ƶq!"
            If ErrCode = 53 Then Message = "���`�G�ݿ�J���R�̿��!"
            If ErrCode = 54 Then Message = "���`�G�ݿ�JEA�P�w!"
            If ErrCode = 55 Then Message = "���`�G�ݿ�J�~��Ƶ�!"

            If ErrCode = 56 Then Message = "���`�G�ݿ�J������!"
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
                   System.Configuration.ConfigurationManager.AppSettings("SBDCommissionFilePath")

        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String



        SQl = "Insert into F_SBDCommissionSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Appdate, Buyer, Vendor, Division, "                  '1~5
        SQl = SQl + "APPPER, Background,Code,Mapdate, Sampledate, "               '6~10
        SQl = SQl + "Light,halffinishNo, material, MaterialDetail,MaterialDetail_1, "                 '11~15
        SQl = SQl + "RefMapFile, Sample, Remark, Mapno, MakeMap,"   '16~20
        SQl = SQl + "level,MapFile,Reason1, "                                                       '21~23
        SQl = SQl + "Content1,Content2,Content3,Content4,Content5,Content6, "            '24~29
        SQl = SQl + "Contentfile1,Contentfile2,Contentfile3,Contentfile4,Contentfile5,Contentfile6, "            '30~35
        SQl = SQl + "Map1,Map2,Map3,Map4,Map5,Map6, "            '36~41
        SQl = SQl + "Makemap1,Makemap2,Makemap3,Makemap4,Makemap5,Makemap6, "            '42~47
        SQl = SQl + "Supplier,HalfFinishOrderNo,HalfFinishDate,Mold,"            '48~52
        SQl = SQl + "MoldPoint,Surfcolor,Surfcolor1,SurfFormNo,SurfFormSno,SampleQty,QCReqFile,FQA, "            '53~57
        SQl = SQl + "QARemark,ForcastFile,QAFile,AuthorizeFile,SampleFile, "            '58~62
        SQl = SQl + "ContactFile,RefFile, "            '63~64
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "


        SQl = SQl + ")  "

        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '���A(0:����,1:�w��NG,2:�w��OK)
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '003001', "                     '���N��
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '���y����

        SQl = SQl + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO
        SQl = SQl + " '" + DAppDate.Text + "', "                '���
        SQl = SQl + " '" + DBuyer.SelectedValue + "', "   'BUYER
        SQl = SQl + " '" + DVendor.Text + "', "     'VENDOR
        SQl = SQl + " '" + DDivision.Text + "', "  '����

        SQl = SQl + " '" + DAppper.Text + "', "    '���
        SQl = SQl + " '" + DBackground.Text + "', "      '�}�o�I��
        SQl = SQl + " N'" + YKK.ReplaceString(DCode.Text) + "', "   '�Ƹ�
        SQl = SQl + " '" + DMapDate.Value + "', "     '�ϭ��Ʊ���
        SQl = SQl + " '" + DSampleDate.Value + "', "  '�˫~�Ʊ���


        SQl = SQl + " '" + DLight.SelectedValue + "', "    '���y��
        SQl = SQl + " '" + DHalfFinishNo.Text + "',"   '�b���~��
        SQl = SQl + " '" + DMaterial.SelectedValue + "', "   '�[�u����-1
        SQl = SQl + " '" + DMaterialDetail.SelectedValue + "', "   '�[�u����-2
        SQl = SQl + " '" + DMaterialDetail_1.Text + "', "   '�[�u����-3
        'FileName = ""
        ' If DRefMapFile.PostedFile.FileName <> "" Then

        'FileName = UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)
        'DRefMapFile.PostedFile.SaveAs(Path + FileName)
        'Else
        'FileName = ""
        'End If


        FileName = ""
        If DRefMapFile.Visible Then
            If DRefMapFile.PostedFile.FileName <> "" Then
                ' Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                '                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '�W�Ǥ��
                'Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                'System.Configuration.ConfigurationManager.AppSettings("TimeOffSheetPath")

                FileName = CStr(NewFormSno) & "-" & "RefMapFile" & "-" & UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)

                DRefMapFile.PostedFile.SaveAs(Path + FileName)

            Else
                FileName = ""
            End If
        End If




        SQl = SQl + " '" + FileName + "'," '���
        SQl = SQl + " '" + DSample.Text + "', "   '�˫~
        SQl = SQl + " '" + DRemark.Text + "', "    '�Ƶ� 
        SQl = SQl + " '" + DMapNo.Text + "', "    '�ϸ�
        SQl = SQl + " '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�    16~20

        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                DMapFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If
        SQl = SQl + " '" + DLevel.Text + "', "    '������
        SQl = SQl + " '" + FileName + "'," '����
        SQl = SQl + " '" + DReason1.SelectedValue + "', "    '��]���O 

        SQl = SQl + " '" + DContent1.Text + "', "    '�ק鷺�e1
        SQl = SQl + " '" + DContent2.Text + "', "    '�ק鷺�e2
        SQl = SQl + " '" + DContent3.Text + "', "    '�ק鷺�e3
        SQl = SQl + " '" + DContent4.Text + "', "    '�ק鷺�e4
        SQl = SQl + " '" + DContent5.Text + "', "    '�ק鷺�e5
        SQl = SQl + " '" + DContent6.Text + "', "    '�ק鷺�e6

        FileName = ""

        If DContentFile1.Visible Then
            If DContentFile1.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile1" & "-" & UploadDateTime & "-" & Right(DContentFile1.PostedFile.FileName, InStr(StrReverse(DContentFile1.PostedFile.FileName), "\") - 1)
                DContentFile1.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����1

        FileName = ""
        If DContentFile2.Visible Then
            If DContentFile2.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile2" & "-" & UploadDateTime & "-" & Right(DContentFile2.PostedFile.FileName, InStr(StrReverse(DContentFile2.PostedFile.FileName), "\") - 1)
                DContentFile2.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����2

        FileName = ""
        If DContentFile3.Visible Then
            If DContentFile3.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile3" & "-" & UploadDateTime & "-" & Right(DContentFile3.PostedFile.FileName, InStr(StrReverse(DContentFile3.PostedFile.FileName), "\") - 1)
                DContentFile3.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����3

        FileName = ""

        If DContentFile4.Visible Then
            If DContentFile4.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile4" & "-" & UploadDateTime & "-" & Right(DContentFile4.PostedFile.FileName, InStr(StrReverse(DContentFile4.PostedFile.FileName), "\") - 1)
                DContentFile4.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����4

        FileName = ""

        If DContentFile5.Visible Then
            If DContentFile5.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile5" & "-" & UploadDateTime & "-" & Right(DContentFile5.PostedFile.FileName, InStr(StrReverse(DContentFile5.PostedFile.FileName), "\") - 1)
                DContentFile5.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����5

        FileName = ""

        If DContentFile6.Visible Then
            If DContentFile6.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ContentFile6" & "-" & UploadDateTime & "-" & Right(DContentFile6.PostedFile.FileName, InStr(StrReverse(DContentFile6.PostedFile.FileName), "\") - 1)
                DContentFile6.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "'," '�ק鷺�e����6



        SQl = SQl + " ''," '����1
        SQl = SQl + " ''," '����2
        SQl = SQl + " ''," '����3
        SQl = SQl + " ''," '����4
        SQl = SQl + " ''," '����5
        SQl = SQl + " ''," '����6

        SQl = SQl + " ''," '�s��1
        SQl = SQl + " ''," '�s��2
        SQl = SQl + " ''," '�s��3
        SQl = SQl + " ''," '�s��4
        SQl = SQl + " ''," '�s��5
        SQl = SQl + " ''," '�s��6



        SQl = SQl + " '" + DSupplier.SelectedValue + "', "    '�~�`��
        SQl = SQl + " '" + DHalfFinishOrderNo.Text + "', "    '�b���~�q��NO
        SQl = SQl + " '" + DHalfFinishDate.Value + "', "      '�b���~�w�q�������
        SQl = SQl + " '" + DMold.Text + "', "                 '�Ҩ�
        '
        SQl = SQl + " '" + DMoldPoint.Text + "', "            '�ި�
        SQl = SQl + " '" + DSurfcolor.Value + "', "            '���C��
        SQl = SQl + " '" + DSurfcolor1.Text + "', "            '���C��
        SQl = SQl + " '', "            '���C��
        SQl = SQl + " '', "            '���C��

        SQl = SQl + " '" + DSampleQty.SelectedValue + "', "            '�˫~�ƶq

        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '���R�̿��

        SQl = SQl + " '" + DFQA.Text + "', "                  'EA�P�w
        SQl = SQl + " '" + DQARemark.Text + "', "             '�~��Ƶ�

        FileName = ""
        If DForcastFile.Visible Then
            If DForcastFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "ForcastFile" & "-" & UploadDateTime & "-" & Right(DForcastFile.PostedFile.FileName, InStr(StrReverse(DForcastFile.PostedFile.FileName), "\") - 1)
                DForcastFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If


        SQl = SQl + " '" + FileName + "',"                    '������

        FileName = ""

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�~�����i��

        FileName = ""
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "AuthorizeFile" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�T�{��

        FileName = ""

        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                DSampleFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�̲׼˫~��


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
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then

                FileName = CStr(NewFormSno) & "-" & "RefFile" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                DRefFile.PostedFile.SaveAs(Path + FileName)
            Else
                FileName = ""
            End If
        End If

        SQl = SQl + " '" + FileName + "',"                    '�䥦���



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
               System.Configuration.ConfigurationManager.AppSettings("SBDCommissionFilePath")
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '�W�Ǥ��
        Dim FileName As String
        Dim SQl As String



        ' SQl = "Select Divname,Username From M_Users ")
        ' SQl = SQl & " Where UserID = '" & Request.QueryString("pUserID") & "'"
        ' SQl = SQl & "   And Active = '1' "
        ' Dim DBUser As DataTable = uDataBase.GetDataTable(SQl)
        'DMakeMap.SelectedValue = DBUser.Rows(0).Item("Username")



        '��w�g�ק�X��
        Dim sql2 As String = ""
        sql2 = " select flag from F_SBDCommissionSheet "
        sql2 = sql2 + " where formno =  '" & wFormNo & "'"
        sql2 = sql2 + " and formsno = '" & CStr(wFormSno) & "'"
        Dim DBData As DataTable = uDataBase.GetDataTable(sql2)
        Dim sflag As String = ""
        flag = DBData.Rows(0).Item("flag")





        SQl = "Update F_SBDCommissionSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"

        SQl = SQl + " AppDate = '" & DAppDate.Text & "',"
        SQl = SQl + " Buyer = '" & DBuyer.SelectedValue & "',"
        SQl = SQl + " Vendor = '" & DVendor.Text & "',"
        SQl = SQl + " Division = '" & DDivision.Text & "',"
        SQl = SQl + " APPPER = '" & DAppper.Text & "',"

        SQl = SQl + " Background = '" & DBackground.Text & "',"
        SQl = SQl + " Code = '" & DCode.Text & "',"
        SQl = SQl + " MapDate = '" & DMapDate.Value & "',"
        SQl = SQl + " SampleDate = '" & DSampleDate.Value & "',"
        SQl = SQl + " Light = '" & DLight.SelectedValue & "',"

        SQl = SQl + " HalfFinishNo = '" & DHalfFinishNo.Text & "',"
        SQl = SQl + " Material= '" & DMaterial.SelectedValue & "',"
        SQl = SQl + " MaterialDetail= '" & DMaterialDetail.SelectedValue & "',"
        SQl = SQl + " MaterialDetail_1= '" & DMaterialDetail_1.Text & "',"

        If wStep = 50 And PAction1 = 1 Then
            SQl = SQl + " flag = flag +1" & ","
        End If

        FileName = ""
        If DRefMapFile.Visible Then
            If DRefMapFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "RefMapFile" & "-" & UploadDateTime & "-" & Right(DRefMapFile.PostedFile.FileName, InStr(StrReverse(DRefMapFile.PostedFile.FileName), "\") - 1)
                DRefMapFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " RefMapFile= N'" + YKK.ReplaceString(FileName) + "',"           '���
            Else
                FileName = ""
            End If
        End If



        SQl = SQl + " Sample= '" & DSample.SelectedValue & "',"                      '�˫~
        SQl = SQl + " Remark= '" & DRemark.Text & "',"                               '�Ƶ�
        SQl = SQl + " MapNo= '" & DMapNo.Text & "',"                                 '�ϸ�
        If wStep = 30 Then
            If DMakeMap.SelectedValue <> "" Then
                SQl = SQl + " MakeMap= '" & DMakeMap.SelectedValue & "',"
            End If
            '�s�Ϫ�
        End If


        SQl = SQl + " Level= '" & DLevel.Text & "',"                                 '������



        If flag = 0 Then
            FileName = ""
            If DMapFile.Visible Then
                If DMapFile.PostedFile.FileName <> "" Then

                    FileName = CStr(wFormSno) & "-" & "MapFile" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " MapFile= N'" + YKK.ReplaceString(FileName) + "',"           '����
                Else
                    FileName = ""
                End If
            End If

        End If




        SQl = SQl + " Reason1= '" & DReason1.SelectedValue & "',"                   '��]���O

        SQl = SQl + " Content1= '" + DContent1.Text + "', "    '�ק鷺�e1
        SQl = SQl + " Content2= '" + DContent2.Text + "', "    '�ק鷺�e2
        SQl = SQl + " Content3= '" + DContent3.Text + "', "    '�ק鷺�e3
        SQl = SQl + " Content4= '" + DContent4.Text + "', "    '�ק鷺�e4
        SQl = SQl + " Content5= '" + DContent5.Text + "', "    '�ק鷺�e5
        SQl = SQl + " Content6= '" + DContent6.Text + "', "    '�ק鷺�e6

        FileName = ""
        If DContentFile1.Visible Then
            If DContentFile1.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile1" & "-" & UploadDateTime & "-" & Right(DContentFile1.PostedFile.FileName, InStr(StrReverse(DContentFile1.PostedFile.FileName), "\") - 1)
                DContentFile1.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile1=  N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����1
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile2.Visible Then
            If DContentFile2.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile2" & "-" & UploadDateTime & "-" & Right(DContentFile2.PostedFile.FileName, InStr(StrReverse(DContentFile2.PostedFile.FileName), "\") - 1)
                DContentFile2.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile2=  N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����2
            Else
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile3.Visible Then
            If DContentFile3.PostedFile.FileName <> "" Then
                FileName = CStr(wFormSno) & "-" & "ContentFile3" & "-" & UploadDateTime & "-" & Right(DContentFile3.PostedFile.FileName, InStr(StrReverse(DContentFile3.PostedFile.FileName), "\") - 1)
                DContentFile3.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile3=  N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����3
                FileName = ""
            End If
        End If



        FileName = ""
        If DContentFile4.Visible Then
            If DContentFile4.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile4" & "-" & UploadDateTime & "-" & Right(DContentFile4.PostedFile.FileName, InStr(StrReverse(DContentFile4.PostedFile.FileName), "\") - 1)
                DContentFile4.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile4=  N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����4
            Else
                FileName = ""
            End If
        End If


        FileName = ""

        If DContentFile5.Visible Then
            If DContentFile5.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "ContentFile5" & "-" & UploadDateTime & "-" & Right(DContentFile5.PostedFile.FileName, InStr(StrReverse(DContentFile5.PostedFile.FileName), "\") - 1)
                DContentFile5.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " ContentFile5=  N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����5
            Else
                FileName = ""
            End If
        End If



        FileName = ""

        If DContentFile6.Visible Then
            If DContentFile6.PostedFile.FileName <> "" Then
                FileName = CStr(wFormSno) & "-" & "ContentFile6" & "-" & UploadDateTime & "-" & Right(DContentFile6.PostedFile.FileName, InStr(StrReverse(DContentFile6.PostedFile.FileName), "\") - 1)
                DContentFile6.PostedFile.SaveAs(Path + FileName)

                SQl = SQl + " ContentFile6= N'" + YKK.ReplaceString(FileName) + "'," '�ק鷺�e����6
            Else
                FileName = ""
            End If
        End If


        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                If flag = 1 Then
                    FileName = CStr(wFormSno) & "-" & "Map1" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map1= N'" + YKK.ReplaceString(FileName) + "'," '����2

                ElseIf flag = 2 Then
                    FileName = CStr(wFormSno) & "-" & "Map2" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map2= N'" + YKK.ReplaceString(FileName) + "'," '����2

                ElseIf flag = 3 Then
                    FileName = CStr(wFormSno) & "-" & "Map3" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map3= N'" + YKK.ReplaceString(FileName) + "'," '����3

                ElseIf flag = 4 Then
                    FileName = CStr(wFormSno) & "-" & "Map4" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map4= N'" + YKK.ReplaceString(FileName) + "'," '����4

                ElseIf flag = 5 Then
                    FileName = CStr(wFormSno) & "-" & "Map5" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map5= N'" + YKK.ReplaceString(FileName) + "'," '����5

                ElseIf flag = 6 Then
                    FileName = CStr(wFormSno) & "-" & "Map6" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                    SQl = SQl + " Map6= N'" + YKK.ReplaceString(FileName) + "'," '����6


                End If

            End If
        End If


        If wStep = 50 And PAction1 = 1 Then
            If DMakeMap.SelectedValue = "" Then
                If flag = 0 Then
                    SQl = SQl + " Makemap1= '" + DMakeMapU.Text + "', "    '�s�Ϫ�1 
                ElseIf flag = 1 Then
                    SQl = SQl + " Makemap2= '" + DMakeMap1.Text + "', "    '�s�Ϫ�1 
                ElseIf flag = 2 Then
                    SQl = SQl + " Makemap3= '" + DMakeMap2.Text + "', "    '�s�Ϫ�1 
                ElseIf flag = 3 Then
                    SQl = SQl + " Makemap4= '" + DMakeMap3.Text + "', "    '�s�Ϫ�1 
                ElseIf flag = 4 Then
                    SQl = SQl + " Makemap5= '" + DMakeMap4.Text + "', "    '�s�Ϫ�1 
                ElseIf flag = 5 Then
                    SQl = SQl + " Makemap6= '" + DMakeMap5.Text + "', "    '�s�Ϫ�1 
                End If
            Else
                If flag = 0 Then
                    SQl = SQl + " Makemap1= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                ElseIf flag = 1 Then
                    SQl = SQl + " Makemap2= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                ElseIf flag = 2 Then
                    SQl = SQl + " Makemap3= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                ElseIf flag = 3 Then
                    SQl = SQl + " Makemap4= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                ElseIf flag = 4 Then
                    SQl = SQl + " Makemap5= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                ElseIf flag = 5 Then
                    SQl = SQl + " Makemap6= '" + DMakeMap.SelectedValue + "', "    '�s�Ϫ�1 
                End If
            End If


        End If


        SQl = SQl + " Supplier= '" + DSupplier.SelectedValue + "', "    '�~�`��
        SQl = SQl + " HalfFinishOrderNo= '" + DHalfFinishOrderNo.Text + "', "    '�b���~�q��NO
        SQl = SQl + " HalfFinishdate= '" + DHalfFinishDate.Value + "', "      '�b���~�w�q�������
        SQl = SQl + " Mold= '" + DMold.Text + "', "                 '�Ҩ�
        '
        SQl = SQl + " MoldPoint= '" + DMoldPoint.Text + "', "            '�ި�



        SQl = SQl + " Surfcolor= '" + DSurfcolor.Value + "', "            '���C��


        If DSurfFormNo.Text <> "" Then
            SQl = SQl + " SurfFormNo= '" + DSurfFormNo.Text + "', "            '�� formno
            SQl = SQl + " SurfFormSno= '" + DSurfFormSno.Text + "', "            '���C�� formsno
        End If

        SQl = SQl + " Surfcolor1= '" + DSurfcolor1.Text + "', "            '���C��1

        SQl = SQl + " SampleQty= '" + DSampleQty.SelectedValue + "', "            '�˫~�ƶq



        FileName = ""
        If DQCReqFile.Visible Then
            If DQCReqFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QCReqFile" & "-" & UploadDateTime & "-" & Right(DQCReqFile.PostedFile.FileName, InStr(StrReverse(DQCReqFile.PostedFile.FileName), "\") - 1)
                DQCReqFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QCReqFile=   N'" + YKK.ReplaceString(FileName) + "'," '���R�̿��
            Else
                FileName = ""
            End If
        End If



        SQl = SQl + " FQA= '" + DFQA.Text + "', "                  'EA�P�w
        SQl = SQl + " QARemark= '" + DQARemark.Text + "', "             '�~��Ƶ�

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

        If DQAFile.Visible Then
            If DQAFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "QAFile" & "-" & UploadDateTime & "-" & Right(DQAFile.PostedFile.FileName, InStr(StrReverse(DQAFile.PostedFile.FileName), "\") - 1)
                DQAFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " QAFile= N'" + YKK.ReplaceString(FileName) + "',"                 '�~�����i��

            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DAuthorizeFile.Visible Then
            If DAuthorizeFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "AuthorizeFile" & "-" & UploadDateTime & "-" & Right(DAuthorizeFile.PostedFile.FileName, InStr(StrReverse(DAuthorizeFile.PostedFile.FileName), "\") - 1)
                DAuthorizeFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " AuthorizeFile=  N'" + YKK.ReplaceString(FileName) + "',"                   '�T�{��
            Else
                FileName = ""
            End If
        End If




        FileName = ""
        If DSampleFile.Visible Then
            If DSampleFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "SampleFile" & "-" & UploadDateTime & "-" & Right(DSampleFile.PostedFile.FileName, InStr(StrReverse(DSampleFile.PostedFile.FileName), "\") - 1)
                DSampleFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " SampleFile=  N'" + YKK.ReplaceString(FileName) + "',"                   '�̲׼˫~��
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
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then

                FileName = CStr(wFormSno) & "-" & "RefFile" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                DRefFile.PostedFile.SaveAs(Path + FileName)
                SQl = SQl + " RefFile= N'" + YKK.ReplaceString(FileName) + "',"                     '�䥦���
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
    '**
    '**     ��J�ˬd
    '**
    '*****************************************************************
    Function OK() As Boolean

        Dim isOK As Boolean = True
        Dim Errcode As Integer = 0



        Dim Message As String = ""


        '�~�`��
        If Errcode = 0 Then
            If DSupplier.Text = "" Then Errcode = 45
        End If


        '�w�q������
        If Errcode = 0 Then
            If DHalfFinishDate.Value = "" Then Errcode = 47
        End If

        '�Ҩ�
        If Errcode = 0 Then

            If DMold.Text = "" Then Errcode = 48

        End If



        '�ި�
        If Errcode = 0 Then
            If DMoldPoint.Text = "" Then Errcode = 50

        End If

        '���C��
        If Errcode = 0 Then
            If DSurfcolor.Value = "" And DSurfcolor1.Text = "" Then Errcode = 51

        End If

        '�˫~�ƶq
        If Errcode = 0 Then

            If DSampleQty.SelectedValue = "" Then Errcode = 52

        End If



        '������ 
        If Errcode = 0 Then

            If LForcastFile.Text = "" Then Errcode = 56

        End If






        If Errcode > 0 Then
            isOK = False
            If Errcode = 45 Then Message = "���`�G�ݿ�J�~�`��!"
            If Errcode = 46 Then Message = "���`�G�ݿ�J�b���~�q��No!"
            If Errcode = 47 Then Message = "���`�G�ݿ�J�b���~�w�q������!"
            If Errcode = 48 Then Message = "���`�G�ݿ�J�Ҩ�!"
            If Errcode = 49 Then Message = "���`�G�ݿ�J�ҫ�!"
            If Errcode = 50 Then Message = "���`�G�ݿ�J�ި�!"

            If Errcode = 51 Then Message = "���`�G�ݿ�J���C��!"
            If Errcode = 52 Then Message = "���`�G�ݿ�J�˫~�ƶq!"
            If Errcode = 53 Then Message = "���`�G�ݿ�J���R�̿��!"
            If Errcode = 54 Then Message = "���`�G�ݿ�JEA�P�w!"
            If Errcode = 55 Then Message = "���`�G�ݿ�J�~��Ƶ�!"

            If Errcode = 56 Then Message = "���`�G�ݿ�J������!"

        End If

        If Not isOK Then
            uJavaScript.PopMsg(Me, Message)
        End If


        Return isOK


    End Function

    Protected Sub DMaterial_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMaterial.SelectedIndexChanged
        '�[�u����2
        Dim Sql As String
        Dim i As Integer

        DMaterialDetail.Items.Clear()

        Sql = "Select Data From M_referp "
        Sql = Sql & " Where cat  = '3002' "
        Sql = Sql & "   And dkey = '" + DMaterial.SelectedValue + "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(Sql)
        DMaterialDetail.Items.Add("")

        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("Data")
            ListItem1.Value = DBAdapter1.Rows(i).Item("Data")
            DMaterialDetail.Items.Add(ListItem1)
        Next


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

    Function InputCheck() As Integer

    End Function

    Protected Sub DMakeMap_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DMakeMap.SelectedIndexChanged

    End Sub
End Class

