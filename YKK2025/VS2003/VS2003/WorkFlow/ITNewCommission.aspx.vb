Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ITNewCommission
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents LFun14 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
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
    Protected WithEvents DSystemTitle As System.Web.UI.WebControls.Label
    Protected WithEvents LFun15 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun16 As System.Web.UI.WebControls.HyperLink

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
        Response.Cookies("PGM").Value = "ITNewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �]�w���U�{�� 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL�s���]�w

        '���
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '001101' And FormNo <= '001199' "
        SQL = SQL + "  And (IniAuthority = '0' "
        SQL = SQL + "       Or (IniAuthority = '1' "
        SQL = SQL + "           And (IniUser like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "                Or IniUserName like '" + "%" + Request.QueryString("pUserID") + "%" + "' "
        SQL = SQL + "               ) "
        SQL = SQL + "          ) "
        SQL = SQL + "      ) "
        SQL = SQL + "Order by FormNo "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Form")
        DBTable1 = DBDataSet1.Tables("M_Form")
        For i = 0 To DBTable1.Rows.Count - 1
            If DBTable1.Rows(i).Item("FormNo") = "001101" Then
                LFun11.Enabled = True
                LFun11.NavigateUrl = "IT_NewSystemSheet_01.aspx?pFormNo=001101" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001102" Then
                LFun12.Enabled = True
                LFun12.NavigateUrl = "IT_ChangeSystemSheet_01.aspx?pFormNo=001102" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001103" Then
                LFun13.Enabled = True
                LFun13.NavigateUrl = "IT_ChangeDataSheet_01.aspx?pFormNo=001103" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001104" Then
                LFun14.Enabled = True
                LFun14.NavigateUrl = "IT_SelectDataSheet_01.aspx?pFormNo=001104" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001105" Then
                LFun15.Enabled = True
                LFun15.NavigateUrl = "IT_TroubleShooterSheet_01.aspx?pFormNo=001105" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001106" Then
                LFun16.Enabled = True
                LFun16.NavigateUrl = "IT_ChangeAccountSheet_01.aspx?pFormNo=001106" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
        Next
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     ���U�\��ƹ��B�z
    '**
    '*****************************************************************
    Private Sub LFun11_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun11.Init
        LFun11.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun11.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun12_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun12.Init
        LFun12.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun12.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun13_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun13.Init
        LFun13.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun13.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun14_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun14.Init
        LFun14.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun14.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun15_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun15.Init
        LFun15.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun15.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub
    Private Sub LFun16_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun16.Init
        LFun16.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '�ƹ������ܦ� 
        LFun16.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '�ƹ����}�����_
    End Sub

End Class
