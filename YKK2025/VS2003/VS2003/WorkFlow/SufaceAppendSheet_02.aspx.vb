Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class SufaceAppendSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DAppendSpecSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DSpec As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCNeed As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents LONo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DQCRemark As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAppendReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DONo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList

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
    Dim wFormNo As String           '��渹�X
    Dim wFormSno As Integer         '���y����
    Dim NowDateTime As String       '�{�b����ɶ�

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "SufaceAppendSheet_02.aspx"

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
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '�{�b���
        wFormNo = Request.QueryString("pFormNo")    '��渹�X
        wFormSno = Request.QueryString("pFormSno")  '���y����
    End Sub


    '*****************************************************************
    '**
    '**     ��ܪ����
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SufaceAppendFilePath")
        Dim i As Integer
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_SufaceAppendSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceAppendSheet")
        If DBDataSet1.Tables("F_SufaceAppendSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Date")                   '���
            SetFieldData("Division", DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Division"))  '����
            SetFieldData("Person", DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Person"))      '���

            DCode.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Code")                   'Code
            DONo.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("ONo")                     '��No
            DOFormNo.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("OFormNo")             '����
            DOFormSno.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("OFormSno")           '��渹
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                LONo.NavigateUrl = "SufaceSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
            Else
                LONo.Visible = False
            End If

            DBuyer.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Buyer")                 'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("SellVendor")       '�e�U�t��

            DSpec.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("Spec")                   '���O
            DAppendReason.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("AppendReason")   '�z��
            SetFieldData("QCNeed", DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("QCNeed"))      'QC�ݧ_
            DQCRemark.Text = DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("QCRemark")           'QC�Ƶ�

            If DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("AttachFile1") <> "" Then           '�ѦҪ���
                LAttachFile1.NavigateUrl = Path & DBDataSet1.Tables("F_SufaceAppendSheet").Rows(0).Item("AttachFile1")
            Else
                LAttachFile1.Visible = False
            End If
        End If
        DFormSno.Text = "�渹�G" & CStr(wFormSno)
        'DB�s������
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
    '**     �ظm�U�Ԧ�����l��
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        '����
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DDivision.Items.Add(ListItem1)
        End If
        '���
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DPerson.Items.Add(ListItem1)
        End If

        'QC���ջݧ_
        If pFieldName = "QCNeed" Then
            DQCNeed.Items.Clear()
            Dim ListItem1 As New ListItem
            ListItem1.Text = pName
            ListItem1.Value = pName
            DQCNeed.Items.Add(ListItem1)
        End If
    End Sub

End Class
