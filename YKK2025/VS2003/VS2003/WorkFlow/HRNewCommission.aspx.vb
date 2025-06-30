Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HRNewCommission
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents LFun01 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun02 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun03 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun04 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun11 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun12 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun13 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun21 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun31 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun32 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LFun33 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DSystemTitle As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents LFun14 As System.Web.UI.WebControls.HyperLink

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()        '設定共用參數
        SetMainMenu()         '設定主畫面

        If Not Me.IsPostBack Then
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        Response.Cookies("PGM").Value = "HRNewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定菜單各程式 
    '**
    '*****************************************************************
    Private Sub SetMainMenu()
        Dim i As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        '表單
        SQL = "SELECT FormName, FormNo FROM M_FORM "
        SQL = SQL + "Where Active = '1' "
        SQL = SQL + "  And FormNo >= '001001' And FormNo <= '001050' "
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
            If DBTable1.Rows(i).Item("FormNo") = "001001" Then
                LFun11.Enabled = True
                LFun11.NavigateUrl = "HR_OverTimeSheet_01.aspx?pFormNo=001001" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pUserID=" & Request.QueryString("pUserID") & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001002" Then
                LFun12.Enabled = True
                LFun12.NavigateUrl = "HR_TimeOffSheet_01.aspx?pFormNo=001002" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pUserID=" & Request.QueryString("pUserID") & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001003" Then
                LFun13.Enabled = True
                LFun13.NavigateUrl = "HR_AwaySheet_01.aspx?pFormNo=001003" & _
                                                          "&pFormSno=0" & _
                                                          "&pStep=1" & _
                                                          "&pSeqNo=0" & _
                                                             "&pUserID=" & Request.QueryString("pUserID") & _
                                                          "&pApplyID=" & Request.QueryString("pUserID")
            End If
            If DBTable1.Rows(i).Item("FormNo") = "001004" Then
                LFun14.Enabled = False
                LFun14.NavigateUrl = "HR_AddWorkSheet_01.aspx?pFormNo=001004" & _
                                                             "&pFormSno=0" & _
                                                             "&pStep=1" & _
                                                             "&pSeqNo=0" & _
                                                             "&pUserID=" & Request.QueryString("pUserID") & _
                                                             "&pApplyID=" & Request.QueryString("pUserID")
            End If
        Next
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     菜單各功能滑鼠處理
    '**
    '*****************************************************************
    Private Sub LFun11_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun11.Init
        LFun11.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun11.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Private Sub LFun12_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun12.Init
        LFun12.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun12.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Private Sub LFun13_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun13.Init
        LFun13.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun13.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

    Private Sub LFun14_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles LFun14.Init
        LFun14.Attributes.Add("onmouseover", "this.style.backgroundColor='GreenYellow'")   '滑鼠移到變色 
        LFun14.Attributes.Add("onmouseout", "this.style.backgroundColor='White'")  '滑鼠移開底色恢復
    End Sub

End Class
