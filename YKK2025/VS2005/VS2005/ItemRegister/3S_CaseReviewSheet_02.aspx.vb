Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class _3S_CaseReviewSheet_02
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代號
    Dim wUserID As String           '簽核者
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            ShowFormData()          ' 顯示表單資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代號
        wUserID = Request.QueryString("pUserID")    '簽核者
        '
        Response.Cookies("PGM").Value = "3S_CaseReviewSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("CaseReviewFilePath")
        Dim SQL As String
        SQL = "Select * From F_CaseReviewSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtItemRegisterSheet.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtItemRegisterSheet.Rows(0).Item("No")                           ' No
            DDate.Text = dtItemRegisterSheet.Rows(0).Item("Date")                       ' 申請日
            DName.Text = dtItemRegisterSheet.Rows(0).Item("Name")                       ' 申請人姓名
            DJobTitle.Text = dtItemRegisterSheet.Rows(0).Item("JobTitle")               ' 職稱
            DDivision.Text = dtItemRegisterSheet.Rows(0).Item("Division")               ' 部門
            DRemark.Text = dtItemRegisterSheet.Rows(0).Item("Remark")                   ' Remark
            If dtItemRegisterSheet.Rows(0).Item("Attachfile1") <> "" Then
                LAttachfile1.NavigateUrl = Path & dtItemRegisterSheet.Rows(0).Item("Attachfile1")   '附件
                LAttachfile1.Visible = True
            Else
                LAttachfile1.Visible = False
            End If
        End If
    End Sub

End Class
