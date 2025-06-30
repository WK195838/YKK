Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManufInCTSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOPContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm

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
        Response.Cookies("PGM").Value = "ManufInCTSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInCTFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInCTSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInCTSheet")
        If DBDataSet1.Tables("F_ManufInCTSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Date")                   '���
            DDivision.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Division") '����
            DPerson.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Person")   '���
            DSliderCode.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("SliderCode")       'Slider Code

            DLevel.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Level")              '������
            DMapNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("MapNo")             '�ϸ�
            If DMapNo.Text <> "" Then
                SQL = "Select FormNo, FormSno From F_MapSheet "
                SQL = SQL & " Where Sts = 1 "
                SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "MapSheet")
                If DBDataSet1.Tables("MapSheet").Rows.Count > 0 Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormNo") & _
                                                         "&pFormSno=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormSno")
                Else
                    SQL = "Select FormNo, FormSno From F_MapModSheet "
                    SQL = SQL & " Where Sts = 1 "
                    SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                    Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter4.Fill(DBDataSet1, "MapModSheet")
                    If DBDataSet1.Tables("MapModSheet").Rows.Count > 0 Then
                        LMapNo.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapModSheet").Rows(0).Item("FormNo") & _
                                                                "&pFormSno=" & DBDataSet1.Tables("MapModSheet").Rows(0).Item("FormSno")
                    End If
                End If
            Else
                LMapNo.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("OFormNo")             '�ϸ�
            DOFormSno.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("OFormSno")             '�ϸ�
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000003" Then
                    LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                End If
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormNo")             '�ϸ�
            If DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormSno")             '�ϸ�
            End If
            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ManufInModSheet_02.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
            Else
                LNFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Target")             '�ϸ�
            DContent.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Content")             '�ϸ�
            DDReason.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Reason")             '�ϸ�
            If DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("AttachFile") <> "" Then          '����1
                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("AttachFile")
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
        Top = 464
    End Sub

End Class
