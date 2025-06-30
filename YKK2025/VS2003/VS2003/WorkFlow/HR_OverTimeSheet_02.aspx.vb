Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_OverTimeSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDateType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOverTimeSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DOverTimeDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFood As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTraffic As System.Web.UI.WebControls.TextBox
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DCVacation As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "HR_OverTimeSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("OverTimeFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_OverTimeSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_OverTimeSheet")
        If DBDataSet1.Tables("F_OverTimeSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("No")                     'No
            DOverTimeDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("OverTimeDate")   '�[�Z���
            DDateType.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DateType")           '�[�Z�O
            DSalaryYM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")           '���ݦ~��
            DDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Date")                   '�ӽФ��
            DName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Name")                   '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoName")           '���q
            DDepoCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoCode")           '���q�N�X
            DDivision.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DivisionCode")   '�����N�X
            DCVacation.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("CVacation")         '�հ�
            DFood.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Food")                   '�뭹
            DTraffic.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Traffic")             '��q

            DBStartH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartH").ToString    '�w�w�}�l-��
            DBStartM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartM").ToString    '�w�w�}�l-��
            DBEndH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndH").ToString        '�w�w�פ�-��
            DBEndM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndM").ToString        '�w�w�פ�-��
            DBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BH").ToString              '�p�⵲�G-��
            DBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BM").ToString              '�p�⵲�G-��

            DAStartH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartH").ToString    '��ڶ}�l-��
            DAStartM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartM").ToString    '��ڶ}�l-��
            DAEndH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndH").ToString        '��ڲפ�-��
            DAEndM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndM").ToString        '��ڲפ�-��
            DAH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AH").ToString              '�p�⵲�G-��
            DAM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AM").ToString              '�p�⵲�G-��

            DFAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH1").ToString          '�֩w����2��-��
            DFAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM1").ToString          '�֩w����2��-��
            DFAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH2").ToString          '�֩w����2�~-��
            DFAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM2").ToString          '�֩w����2�~-��

            DFBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBH").ToString            '�֩w����-��
            DFBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBM").ToString            '�֩w����-��
            DFCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCH").ToString            '�֩w��w����-��
            DFCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCM").ToString            '�֩w��w����-��

            DFReason.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FReason")             '�[�Z�z��

            BCardTime.Attributes("onclick") = "ShowCardTime();"
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
