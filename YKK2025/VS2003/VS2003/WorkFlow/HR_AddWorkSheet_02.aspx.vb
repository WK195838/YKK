Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_AddWorkSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DWM As System.Web.UI.WebControls.TextBox
    Protected WithEvents BACardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DAWorkDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOverTime As System.Web.UI.WebControls.Button
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAddWorkSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DWorkDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DWEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAddWorkType As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "HR_AddWorkSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("AddWorkFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_AddWorkSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_AddWorkSheet")
        If DBDataSet1.Tables("F_AddWorkSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Date")                   '�ӽФ��
            DSalaryYM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("SalaryYM")           '���ݦ~��

            DName.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Name")                   '�m�W
            DEmpID.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("JobTitle")           '¾��
            DJobCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("JobCode")             '¾�٥N�X
            DDepoName.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DepoName")           '���q�O
            DDepoCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DepoCode")           '���q�O�N�X
            DDivision.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("Division")           '����
            DDivisionCode.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("DivisionCode")   '�����N�X

            DWorkDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WorkDate")           '�ʶԤ��
            DWStartH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WStartH").ToString    '�ʶԶ}�l-��
            DWStartM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WStartM").ToString    '�ʶԶ}�l-��
            DWEndH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WEndH").ToString        '�ʶԲפ�-��
            DWEndM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WEndM").ToString        '�ʶԲפ�-��
            DWH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WH").ToString              '�p�⵲�G-��
            DWM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("WM").ToString              '�p�⵲�G-��

            DAWorkDate.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AWorkDate")         '�ɤu���
            DAStartH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AStartH").ToString    '�ɤu�}�l-��
            DAStartM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AStartM").ToString    '�ɤu�}�l-��
            DAEndH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AEndH").ToString        '�ɤu�פ�-��
            DAEndM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AEndM").ToString        '�ɤu�פ�-��
            DAH.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AH").ToString              '�p�⵲�G-��
            DAM.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AM").ToString              '�p�⵲�G-��

            DFReason.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("FReason")             '�z��
            DAddWorkType.Text = DBDataSet1.Tables("F_AddWorkSheet").Rows(0).Item("AddWorkType")     '�ɤu���O

            BCardTime.Attributes("onclick") = "ShowCardTime();"    '�ʶԨ�d�O��
            BACardTime.Attributes("onclick") = "ShowCardTimeA();"  '�ɤu��d�O��
            BOverTime.Attributes("onclick") = "ShowOverTime();"    '�[�Z�O��
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
