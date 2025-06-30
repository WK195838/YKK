Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ImportCTSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOPContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox

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
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "ImportCTSheet_02.aspx"

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

    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ImportCTFilePath")

        Dim PathOld1 As String = ConfigurationSettings.AppSettings.Get("HttpOld")
        Dim PathOld2 As String = ConfigurationSettings.AppSettings.Get("ImportCTFilePath")
        Dim RtnCode As Integer = 0

        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ImportCTSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ImportCTSheet")
        If DBDataSet1.Tables("F_ImportCTSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Date")                   '���
            DDivision.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Division") '����
            DPerson.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Person")   '���
            DSliderCode.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("SliderCode")       'Slider Code

            DOFormNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("OFormNo")             '�ϸ�
            DOFormSno.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("OFormSno")             '�ϸ�
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                LOFormNo.NavigateUrl = "ImportSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormNo")             '�ϸ�
            If DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("NFormSno")             '�ϸ�
            End If
            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ImportModSheet_02.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
            Else
                LNFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Target")             '�ϸ�
            DSpec.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Spec")             '�ϸ�

            DContent.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Content")             '�ϸ�
            DDReason.Text = DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("Reason")             '�ϸ�
            If DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile") <> "" Then          '����1
                RtnCode = oCommon.RecoveryArchiveFile(PathOld1, PathOld2, DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile"))

                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ImportCTSheet").Rows(0).Item("AttachFile")
            Else
                LAttachFile.Visible = False
            End If
            DFormSno.Text = "�渹�G" & CStr(wFormSno)    '�渹
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
