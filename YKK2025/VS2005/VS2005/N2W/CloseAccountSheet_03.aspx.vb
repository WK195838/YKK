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



Partial Class CloseAccountSheet_03
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
    Dim wUserIP As String = ""
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"

        SetParameter()          '設定共用參數

        ShowFormData()      '顯示表單資料

    



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
        wUserIP = Request.ServerVariables("REMOTE_ADDR")
        '
        Response.Cookies("PGM").Value = "CloseAccountSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼





    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FindFieldInf)
    '**     尋找表單欄位屬性
    '**
    '*****************************************************************
    Function FindFieldInf(ByVal pFieldName As String) As Integer
        Dim Run As Boolean
        Dim i As Integer
        Run = True
        FindFieldInf = 9
        While i <= 60 And Run
            If FieldName(i) = pFieldName Then
                FindFieldInf = Attribute(i)
                Run = False
            End If
            i = i + 1
        End While
    End Function
 


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

  
        Dim SQL As String


        SQL = "Select * From F_CloseAccountSheet "
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
            DObject.Text = dtData.Rows(0).Item("Object")
            DSDate.Text = dtData.Rows(0).Item("SDate")
            DEDate.Text = dtData.Rows(0).Item("EDate")
            DASDate.Text = dtData.Rows(0).Item("ASDate")
            DAEDate.Text = dtData.Rows(0).Item("AEDate")
            DTripNo.Text = dtData.Rows(0).Item("TripNo")
            DLocation.Text = dtData.Rows(0).Item("Location")
            SQL = " select * from F_ComplaintOutSheet"
            SQL = SQL + "  where no ='" & dtData.Rows(0).Item("QCNO1") & "'"
            Dim Ddata1 As DataTable = uDataBase.GetDataTable(SQL)
            If Ddata1.Rows.Count > 0 Then
                DQCNO.Text = Ddata1.Rows(0).Item("QCNO")
            End If
        End If


        '出差迄日
        Dim AEDateT As DateTime
        AEDateT = DAEDate.Text
        Dim AEdateS As String
        AEdateS = AEDateT.ToString("yyyy/MM/dd")

        '取明細資料

        If AEdateS > "2024/08/01" Then
            SQL = "  select  UNIQUE_ID,TYPE,TYPENO,CURRENCY,SDATE,PAY,MONEY,DAYS,case when pay ='公司卡' then 0 else  AVERAGERATE end AS RATE,REMARK,EXPENSENO,case when pay ='公司卡' then 0 else  round(money*days*AVERAGERATE,0) end  as SumAmt from ("
            SQL = SQL + " select * from ("
            SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
            SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
            SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
            SQL = SQL + " from  F_CloseAccountSheetDT "
            SQL = SQL + "  where 1=1 and No='" + DNo.Text + "')a "
            SQL = SQL + " ,("
            SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
            SQL = SQL + " )b where a.currency =b.currencycode "
            SQL = SQL + " and '" + AEdateS + "' between strD and EndD "
            SQL = SQL + " )a"
            SQL = SQL + " order by typeno,sdate  "
        Else

            SQL = " select *, round(money*days*rate,0) as SumAmt from ("
            SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
            SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
            SQL = SQL + " ,convert(decimal(9,6),rate) as rate,remark,expenseno"
            SQL = SQL + " from  F_CloseAccountSheetDT "
            SQL = SQL + "  where 1=1 and No='" + DNo.Text + "')a "
            SQL = SQL + " order by typeno,sdate  "




        End If

 
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()


        SQL = " select  isnull(sum(SumAmt),0) sumamt   from ("
        SQL = SQL + " select *,case when pay ='公司卡' then 0 else  round(money*days*AVERAGERATE,0) end  as SumAmt from ("
        SQL = SQL + " select Unique_ID,type,typeno,currency,convert(char(10),sdate,111)sdate,pay,case when money = '' then '0' else  convert(decimal(9,2),money) end as money,"
        SQL = SQL + "  case when days='' then '0' else convert(decimal(9,0),days) end as days"
        SQL = SQL + " ,case when pay = '公司卡' then 0 when currency ='TWD' then 1 when Rate ='' then 1  else  convert(decimal(9,6),rate) end as rate,remark,expenseno"
        SQL = SQL + " from  F_CloseAccountSheetDT "
        SQL = SQL + "  where 1=1 and No='" + DNo.Text + "')a "
        SQL = SQL + " ,("
        SQL = SQL + " select currencycode,syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,AVERAGERATE  from F_exchangerate"
        SQL = SQL + " )b where a.currency =b.currencycode "
        SQL = SQL + " and '" + AEdateS + "' between strD and EndD "
        SQL = SQL + " )a"

        Dim dtExpData1 As DataTable = uDataBase.GetDataTable(SQL)
        If dtExpData1.Rows.Count > 0 Then
            If dtExpData1.Rows(0).Item("SumAmt") <> 0 Then
                DSumAmt.Text = String.Format("{0:N0}", Val(dtExpData1.Rows(0).Item("SumAmt")))
            End If
        End If

    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub
End Class
