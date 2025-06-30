Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DTMW_NewColorAgain_01
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
    Dim isOK As Boolean = True
    Dim Message As String = ""




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '�]�w�@�ΰѼ�


        If Not Me.IsPostBack Then   '���OPostBack
            SetFieldAttribute("Post")
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
        wSeqNo = Request.QueryString("pSeqNo")      '�Ǹ�
        wApplyID = Request.QueryString("pApplyID")  '�ӽЪ�ID
        wAgentID = Request.QueryString("pAgentID")  '�Q�N�z�HID
        wUserID = Request.QueryString("pUserID")  '�Q�N�z�HID
        'Add-Start by Joy 2009/11/20(2010��ƾ����)
        '�ӽЪ�-�s�զ�ƾ�(TP1->�x�_-����, TP2->�x�_-�ا�, CL1->���c-����, CL2->���c-�s�y)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(wUserID)
        Response.Cookies("PGM").Value = "DTMW_NewColorComplete.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '���y����
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '�u�{�N�X

    End Sub



    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Table As String = ""

        Dim SQL As String
        SQL = "Select * From M_referp "
        SQL = SQL & " Where Cat='5001' "
        SQL = SQL & " and dkey  ='Sheet-" + wFormNo + "'"
        Dim DBTable As DataTable = uDataBase.GetDataTable(SQL)
        If DBTable.Rows.Count > 0 Then
            LSheetName.Text = DBTable.Rows(0).Item("Data")
        End If

        SQL = "Select * From v_NewColor "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then

            If DBAdapter1.Rows(0).Item("CustomerCheck") = 1 Then
                DCustomerCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("FactoryCheck") = 1 Then
                DFactoryCheck.Checked = True
            End If
            If DBAdapter1.Rows(0).Item("VCACheck") = 1 Then
                DVCACheck.Checked = True
            End If

            DNo.Text = DBAdapter1.Rows(0).Item("No")                         'No
            '  DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = DBAdapter1.Rows(0).Item("Date")                         'No 
            DDepoName.Text = DBAdapter1.Rows(0).Item("DepoName")                         'No
            DName.Text = DBAdapter1.Rows(0).Item("Name")                         'No
            DCustomer.Text = DBAdapter1.Rows(0).Item("Customer")                         'No
            DCustomerCode.Text = DBAdapter1.Rows(0).Item("CustomerCode")                         'No
            DBuyer.Text = DBAdapter1.Rows(0).Item("Buyer")                         'No
            DCustomerColor.Text = DBAdapter1.Rows(0).Item("CustomerColor")                         'No3
            DCustomerColorCode.Text = DBAdapter1.Rows(0).Item("CustomerColorCode")
            DOverseaYKKCode.Text = DBAdapter1.Rows(0).Item("OverseaYKKCode")
            DPANTONECode.Text = DBAdapter1.Rows(0).Item("PANTONECode")
            DDevYear.Text = DBAdapter1.Rows(0).Item("DevYear")
            DDevSeason.Text = DBAdapter1.Rows(0).Item("DevSeason")

            If Mid(DBAdapter1.Rows(0).Item("ReceiveDate").ToString, 1, 4) = "1900" Then
                DDReceiveDate.Text = ""
            Else
                DDReceiveDate.Text = DBAdapter1.Rows(0).Item("ReceiveDate")               '�U��ɶ�
            End If
            DColorLight1.Text = DBAdapter1.Rows(0).Item("ColorLight1")
            DColorLight1.Text = DBAdapter1.Rows(0).Item("ColorLight2")
            DYKKColorType.Text = DBAdapter1.Rows(0).Item("YKKColorType")
            DYKKColorCode.Text = DBAdapter1.Rows(0).Item("YKKColorCode")

        End If

        'DTM�}�l��
        SQL = " select  convert(char(10), max(astarttime),111) astarttime from t_waithandle"
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " and step =60 "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter2.Rows.Count > 0 Then
            DAgainSdate.Text = DBAdapter2.Rows(0).Item("Astarttime")
        End If


        'DTM������
        SQL = " select  convert(char(10),max(AEndtime),111) AEndtime from t_waithandle "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & " And FormSno =  '" & CStr(wFormSno) & "'"
        SQL = SQL & " and step= 60  "
        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter3.Rows.Count > 0 Then
            DAgainEDate.Text = DBAdapter3.Rows(0).Item("AEndtime")
        End If

        '�p��DTM�Ѽ�
        Dim date1 As DateTime
        Dim date2 As DateTime
        date1 = DAgainSdate.Text
        date2 = DAgainEDate.Text
        Dim ts As TimeSpan
        ts = date2 - date1
        Dim iDays As Integer
        iDays = ts.Days
        DAgainDays.Text = iDays

        SetFieldData("CheckType", "ZZZZZZ")
        SetFieldData("Color", "ZZZZZZ")




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
    '**     ���U����ݩʤ�����J�ˬd���]�w
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '���U����ݩʤ�����J�ˬd���]�w


        ShowRequiredFieldValidator("DDyeTimesRqd", "DDyeTimes", "���`�G�ݿ�J�V�⦸��")

    End Sub



    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)




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
    '**(AppendData)
    '**     �s�W�����
    '**
    '*****************************************************************
    Sub AppendData()


        Dim SQl As String

        SQl = "Select * From M_referp "
        SQl = SQl & " Where Cat='5001' "
        SQl = SQl & "and dkey  ='Table-" + wFormNo + "'"
        Dim DBTable1 As DataTable = uDataBase.GetDataTable(SQl)
        Dim wTable As String = ""
        If DBTable1.Rows.Count > 0 Then
            wTable = DBTable1.Rows(0).Item("Data")
        End If


        SQl = "Insert into F_NewColorAgain "
        SQl = SQl + "( "
        SQl = SQl + "CompletedTime, FormNo, FormSno,"
        SQl = SQl + "No, AgainSDate,AgainEDate,AgainDays,DyeTimes,"  '1~5                
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '" + NowDateTime + "', "        '���פ�
        SQl = SQl + " '" + CStr(wFormNo) + "', "   '���y����
        SQl = SQl + " '" + CStr(wFormSno) + "', "   '���y����
        SQl = SQl + " N'" + DNo.Text + "', "   'NO  1     
        SQl = SQl + " '" + DAgainSdate.Text + "', "                '�Ȥ�8
        SQl = SQl + " '" + DAgainEDate.Text + "', "                '�Ȥ�9
        SQl = SQl + " '" + DAgainDays.Text + "', "                '�Ȥ�9
        SQl = SQl + " '" + DDyeTimes.Text + "', "                'BUYER11   
        SQl = SQl + " '" + wUserID + "', "     '�@����
        SQl = SQl + " '" + NowDateTime + "', "       '�@���ɶ�
        SQl = SQl + " '" + "" + "', "                       '�ק��
        SQl = SQl + " '" + NowDateTime + "' "       '�ק�ɶ�
        SQl = SQl + " ) "
        uDataBase.ExecuteNonQuery(SQl)

   
        SQl = " Update  " + wTable
        SQl = SQl & " Set Again = 1"
        SQl = SQl & " Where FormNo =  '" & wFormNo & "'"
        SQl = SQl & " And FormSno =  '" & CStr(wFormSno) & "'"
        uDataBase.ExecuteNonQuery(SQl)



    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim SQL As String  '�ˬd�O�_���ۦP�������s��̿৹����
        SQL = " select  *  from F_NewColorAgain "
        SQL = SQL & " where No='" + DNo.Text + "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count = 0 Then '�p�G�S���~�i�H�s�W
            AppendData()  'Insert
            Dim URL As String
            URL = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                                                              "&pUserID=" & wUserID & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            uJavaScript.PopMsg(Me, "�渹�G" + DNo.Text + "-�s��A�{������w����!")
        End If




    End Sub
End Class

