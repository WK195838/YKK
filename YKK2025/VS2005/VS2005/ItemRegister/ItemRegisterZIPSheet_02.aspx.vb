Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class ItemRegisterZIPSheet_02
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim NowDateTime As String       '現在日期時間

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        TopPosition()                   ' 設定top
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
        '
        Response.Cookies("PGM").Value = "ItemRegisterZIPSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL, wNo As String
        Top = 129
        SQL = "Select No From F_ItemRegisterZIPSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtItemRegisterSheet.Rows.Count > 0 Then
            wNo = dtItemRegisterSheet.Rows(0).Item("No")
        End If
        '
        SQL = "Select "
        SQL = SQL & "'ItemRegisterSheet_02.aspx?pFormNo=001151&pFormSno=' + LTrim(RTrim(str(FormSno))) As FNoURL, "
        SQL = SQL + "Case No When '' Then '未編號' Else No End + "
        SQL = SQL + "Case AttachFile1 When '' Then '' Else '-@' End As FNo, "
        SQL = SQL & "Case DTApplyType When 0 Then '申請' When 1 Then '啟用' Else '無' End As ApplyType, "
        SQL = SQL & "Code, ItemName1+' '+ItemName2+' '+Itemname3 As ItemName, "
        SQL = SQL & "DTCode, DTItemName1+' '+DTItemName2+' '+DTItemname3 As DTItemName "
        SQL = SQL & "From V_ItemRegisterSheet_02 "
        SQL = SQL & "Where DTDataType  = '0' "
        SQL = SQL & "  And ZIPApply = '1' "
        SQL = SQL & "  And DTApplyType Between '0' and '1' "
        SQL = SQL & "  And DTFormNo =  '" & wFormNo & "'"
        SQL = SQL & "  And DTNo     =  '" & wNo & "'"
        SQL = SQL & "Order by DTCode "
        Dim dt_ItemRegister As DataTable = uDataBase.GetDataTable(SQL)
        Top = Top + (dt_ItemRegister.Rows.Count + 1) * 40
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
        DRemarkSheet.Style("top") = Top & "px"
        DRemark.Style("top") = Top + 7 & "px"
        LAttachfile1.Style("top") = Top + 61 & "px"
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = uCommon.GetAppSetting("Http") & uCommon.GetAppSetting("ItemRegisterZIPFilePath")
        Dim SQL As String
        SQL = "Select * From F_ItemRegisterZIPSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtItemRegisterSheet As DataTable = uDataBase.GetDataTable(Sql)
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
