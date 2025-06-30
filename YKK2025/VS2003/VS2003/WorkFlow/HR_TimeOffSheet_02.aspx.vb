Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_TimeOffSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BVARecord As System.Web.UI.WebControls.Button
    Protected WithEvents DAEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOTNo4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo5 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOTHours5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours As System.Web.UI.WebControls.TextBox
    Protected WithEvents DADays As System.Web.UI.WebControls.TextBox
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DVDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalary As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEvidence As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTimeOffAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTimeOffSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAfter As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVacation As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDieType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVDaysBlank As System.Web.UI.WebControls.TextBox

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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �����ܼ�
    '**
    '*****************************************************************
    Dim FieldName(50) As String     '�U���
    Dim Attribute(50) As Integer    '�U����ݩ�    
    Dim Top As Integer              '�ʺA����Top��m
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim wStep As Integer            '�u�{�N�X
    Dim wSeqNo As Integer           '�Ǹ�
    Dim wApplyID As String          '�ӽЪ�ID
    Dim NowDateTime As String       '�{�b����ɶ�

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "HR_TimeOffSheet_02.aspx"

        SetParameter()          '�]�w�@�ΰѼ�
        TopPosition()           '���s��RequestedField��Top��m
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
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
        wStep = 999                                 '�u�{�N�X
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     �]�w�u�X�����ƥ�
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("TimeOffFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_TimeOffSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_TimeOffSheet")
        If DBDataSet1.Tables("F_TimeOffSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Date")                   '�ӽФ��
            DSalaryYM.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("SalaryYM")           '���ݦ~��

            DName.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Name")                   '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DivisionCode")   '�����N�X

            DAfter.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("After")                 '�ƫe,�ƫ�
            DJobAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobAgent")           '¾�ȥN�z�H
            DTimeOffAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("TimeOffAgent")   '�N�а��H
            DVacation.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VacationCode") + _
                           ":" + _
                           DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Vacation")             '���O

            DEvidence.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Evidence")           '����
            DSalary.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Salary")               '�~��
            DDieType.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DieType")             '�ల�O
            DVDays.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VDays").ToString        '�i�ФѼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1") <> "" Then                 '�[�ZNo1 
                LOTNo1.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
                LOTNo1.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
            Else
                LOTNo1.Visible = False
                DOTHours1.Visible = False
            End If
            DOTHours1.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours1"), 1)  '�[�ZNo1-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2") <> "" Then                 '�[�ZNo2 
                LOTNo2.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
                LOTNo2.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
            Else
                LOTNo2.Visible = False
                DOTHours2.Visible = False
            End If
            DOTHours2.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours2"), 1)  '�[�ZNo2-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3") <> "" Then                 '�[�ZNo3
                LOTNo3.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
                LOTNo3.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
            Else
                LOTNo3.Visible = False
                DOTHours3.Visible = False
            End If
            DOTHours3.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours3"), 1)  '�[�ZNo3-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4") <> "" Then                         '�[�ZNo4
                LOTNo4.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
                LOTNo4.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
            Else
                LOTNo4.Visible = False
                DOTHours4.Visible = False
            End If
            DOTHours4.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours4"), 1)  '�[�ZNo4-�ɼ�

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5") <> "" Then                         '�[�ZNo5
                LOTNo5.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
                LOTNo5.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
            Else
                LOTNo5.Visible = False
                DOTHours5.Visible = False
            End If
            DOTHours5.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours5"), 1)  '�[�ZNo5-�ɼ�

            DOTHours.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours"), 1)    '�[�Z�`�ɼ�

            DBStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartDate")       '�w�w�}�l���
            DBStartH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartH").ToString    '�w�w�}�l��
            DBEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndDate")           '�w�w�������
            DBEndH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndH").ToString        '�w�w������
            DBDays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BDays"), 1)    '�w�w��

            DAStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartDate")       '��ڶ}�l���
            DAStartH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartH").ToString    '��ڶ}�l��
            DAEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndDate")           '��ڵ������
            DAEndH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndH").ToString        '��ڵ�����
            DADays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("ADays"), 1)    '��ڤ�

            DFReason.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("FReason")             '�а��z��

            BCardTime.Attributes("onclick") = "ShowCardTime();"
            BVARecord.Attributes("onclick") = "ShowVacation();"    '�а��O��
        End If

        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     ���s��RequestedField��Top��m
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 432
    End Sub

End Class
