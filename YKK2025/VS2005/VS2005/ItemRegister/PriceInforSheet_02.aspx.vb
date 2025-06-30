Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class PriceInforSheet_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim NowDateTime As String       '現在日期時間
    '
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim HourList, MinList, AdjList As New ListItemCollection
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            Server.ScriptTimeout = 900  ' 設定逾時時間
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

        Response.Cookies("PGM").Value = "PriceInforSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim SQL As String
        SQL = "Select * From  F_PriceInforSheetH "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)

        DDate.Text = CDate(dt.Rows(0)("Date").ToString).ToString("yyyy/MM/dd")
        DDepoName.Text = dt.Rows(0)("DepoName").ToString.Trim
        DDepoCode.Text = dt.Rows(0)("DepoCode").ToString.Trim
        DDivision.Text = dt.Rows(0)("Division").ToString.Trim
        DDivisionCode.Text = dt.Rows(0)("DivisionCode").ToString.Trim
        DName.Text = dt.Rows(0)("Name").ToString.Trim
        DEmpID.Text = dt.Rows(0)("EMPID").ToString.Trim
        DJobTitle.Text = dt.Rows(0)("JobTitle").ToString.Trim
        DJobCode.Text = dt.Rows(0)("JobCode").ToString.Trim
        DNo.Text = dt.Rows(0)("No").ToString.Trim()
        Dremark.Text = dt.Rows(0)("Remark").ToString.Trim()

        SQL = " Select * From  F_PriceInforSheetDt "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        SQL &= " Order by Code, Version "
        Dim dtPriceInforDT As DataTable = uDataBase.GetDataTable(SQL)
        grd.DataSource = dtPriceInforDT
        grd.DataBind()
    End Sub
End Class
