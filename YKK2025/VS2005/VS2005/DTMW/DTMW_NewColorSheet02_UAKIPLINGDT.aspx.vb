Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorSheet02_UAKIPLINGDT
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

        If Not Me.IsPostBack Then   '���OPostBack
            ShowFormData()      '��ܪ����

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
        Dim i As Integer

        If DBAdapter3.Rows.Count > 0 Then
            HyperLink1.NavigateUrl = DBAdapter3.Rows(0).Item("OPURL")  '�а��ҩ�
            HyperLink1.Visible = True


            MNOSts.Text = ""

            For i = 0 To DBAdapter3.Rows.Count - 1
                If MNOSts.Text = "" Then
                    MNOSts.Text = DBAdapter3.Rows(i).Item("stepnamedesc")
                Else
                    MNOSts.Text = MNOSts.Text + "," + DBAdapter3.Rows(i).Item("stepnamedesc")
                End If

            Next
            '    MNOSts.Text = DBAdapter3.Rows(0).Item("stepnamedesc")
            MNOSts.Visible = True
            DOFormSno.Text = DBAdapter3.Rows(0).Item("FormSno")

        End If




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

            Case 1  '�ק�+�ˬd
                DNO1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNO1Rqd", "DNO1", "���`�G�ݿ�J�ܢ�")
                DNO1.Visible = True

            Case 2  '�ק�
                DNO1.BackColor = Color.Yellow
                DNO1.Visible = True

            Case Else   '����
                DNO1.Visible = False

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
            
            Case 1  '�ק�+�ˬd
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "greenyellow")
                '  DYKKColorCode.Attributes.Add("readonly", "true")
                '  ShowRequiredFieldValidator("DYKKColorCodeRqd", "DYKKColorCode", "���`�G�ݿ�JYKK�⸹")


            Case 2  '�ק�
                DYKKColorCode.Visible = True
                DYKKColorCode.Style.Add("background-color", "yellow")
                DYKKColorCode.Attributes.Add("readonly", "true")

            Case Else   '����
                DYKKColorCode.Visible = False


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

 
End Class

