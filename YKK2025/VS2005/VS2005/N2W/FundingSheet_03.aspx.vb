Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls



Partial Class FundingSheet_03
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()          '設定共用參數



        ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查



        ShowFormData()      '顯示表單資料



        ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

        '上傳資料檢查及顯示訊息


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
        wStep = Request.QueryString("pStep")        '工程代碼
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        '
        Response.Cookies("PGM").Value = "FundingSheet_03.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
        
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetField)
    '**     表單欄位顯示及欄位輸入檢查
    '**
    '*****************************************************************
    Sub ShowSheetField(ByVal pPost As String)
        '建置欄位及屬性陣列
        oCommon.GetFieldAttribute(wFormNo, wStep, FieldName, Attribute)
        '表單號碼,工程關卡號碼,欄位名,欄位屬性
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        ' wFormNo = "003110"
        ' wFormSno = "508"
        Dim SQL As String
        SQL = "Select * From F_FundingSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")                         'No 
            DDivision1.Text = dtData.Rows(0).Item("Division1")
            DDivision2.Text = dtData.Rows(0).Item("Division2")
            DDivision3.Text = dtData.Rows(0).Item("Division3")
            DDivisionCode1.Text = dtData.Rows(0).Item("Divisioncode1")
            DDivisionCode2.Text = dtData.Rows(0).Item("Divisioncode2")
            DDivisionCode3.Text = dtData.Rows(0).Item("Divisioncode3")
            DEmpID1.Text = dtData.Rows(0).Item("EmpID1")
            DEmpID2.Text = dtData.Rows(0).Item("EmpID2")
            DEmpName1.Text = dtData.Rows(0).Item("EmpName1")
            DEmpName2.Text = dtData.Rows(0).Item("EmpName2")
            DJobTitle1.Text = dtData.Rows(0).Item("Jobtitle1")
            DJobTitle2.Text = dtData.Rows(0).Item("Jobtitle2")
            DApplyAmt.Text = dtData.Rows(0).Item("ApplyAmt")
            DDebitAmt.Text = dtData.Rows(0).Item("DebitAmt")
            DSumAmt.Text = dtData.Rows(0).Item("SumAmt")


            SQL = "Select ExpCat+'--'+ExpItem ExpItem,ACID, ADate, TaxType, NetAmt, TaxAmt, Amt, Content, Remark,case when InvoiceNo ='' then Taxno else InvoiceNo end as InvoiceNo,GUINo  From F_FundingSheetdt "
            SQL &= " Where FormNo =  '" & wFormNo & "'"
            SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
            Dim dtData1 As DataTable = uDataBase.GetDataTable(SQL)
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()


        End If
    End Sub



End Class
