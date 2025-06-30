Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class _3S_DownloadPage
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents DSystemTitle As System.Web.UI.WebControls.Label
    Protected WithEvents LFun16 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun15 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun14 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun33 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun32 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun31 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun21 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun13 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun12 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun11 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun04 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun03 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun02 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun01 As System.Web.UI.WebControls.HyperLink

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
    Dim NowDateTime As String       '�{�b����ɶ�

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()        '�]�w�@�ΰѼ�
        SetMainMenu()         '�]�w�D�e��

        If Not Me.IsPostBack Then
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
        Response.Cookies("PGM").Value = "3S_DownloadPage.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w���U�{�� 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & "/WorkFlow/Images/"

        LFun01.Enabled = True
        LFun01.NavigateUrl = Path & "Blank_FunctionSpecification.xls"

        LFun11.Enabled = True
        LFun11.NavigateUrl = Path & "Sample_FunctionSpecification.xls"

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ���U�\��ƹ��B�z
    '**
    '*****************************************************************
    Private Sub LFun01_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun01.Init
        LFun01.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun01.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun11_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun11.Init
        LFun11.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun11.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub

End Class

