Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class IT_NewSystemSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSystem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEffect As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevelopItem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DEffect As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFinishDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEngineer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNewSystemSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "IT_NewSystemSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("NewSystemFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB�s���]�w
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w
        OleDbConnection1.Open()

        DEngineer.Items.Clear()
        DDevelopItem.Items.Clear()
        DSystem.Items.Clear()

        SQL = "Select * From F_NewSystemSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_NewSystemSheet")
        If DBDataSet1.Tables("F_NewSystemSheet").Rows.Count > 0 Then
            '�����
            DNo.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("No")
            DDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Date")
            DName.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Name")
            DEmpID.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("JobTitle")
            DJobCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("JobCode")
            DDepoName.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DepoName")
            DDepoCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DepoCode")
            DDivision.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Division")
            DDivisionCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DivisionCode")
            DFinishDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("FinishDate")
            DTarget.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Target")
            DEffect.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Effect")

            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("RefFile") <> "" Then
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If
            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BStartDate") = "1900/1/1" Then
                DBStartDate.Text = ""
            Else
                DBStartDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BStartDate")
            End If

            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BEndDate") = "1900/1/1" Then
                DBEndDate.Text = ""
            Else
                DBEndDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BEndDate")
            End If

            DBDays.Text = CStr(DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BDays"))

            Dim ListItem1 As New ListItem
            ListItem1.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Engineer")
            ListItem1.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Engineer")
            DEngineer.Items.Add(ListItem1)

            Dim ListItem2 As New ListItem
            ListItem2.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DevelopItem")
            ListItem2.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DevelopItem")
            DDevelopItem.Items.Add(ListItem2)

            Dim ListItem3 As New ListItem
            ListItem3.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("System")
            ListItem3.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("System")
            DSystem.Items.Add(ListItem3)

            DITEffect.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("ITEffect")
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
