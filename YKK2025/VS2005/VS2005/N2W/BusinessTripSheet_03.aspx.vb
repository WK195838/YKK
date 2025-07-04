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



Partial Class BusinessTripSheet_03
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
        Response.Cookies("PGM").Value = "BusinessTripSheet_01.aspx"
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

        'wFormNo = "003114"
        'wFormSno = "5"

        Dim SQL As String
        SQL = "Select * From F_BusinessTripSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'" '
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")                         'No 
            DDivision1.Text = dtData.Rows(0).Item("Division1")
            DDivision2.Text = dtData.Rows(0).Item("Division2")
            DDivision3.Text = dtData.Rows(0).Item("Division3")
            DDivision4.Text = dtData.Rows(0).Item("Division4")
            DDivisionCode1.Text = dtData.Rows(0).Item("Divisioncode1")
            DDivisionCode2.Text = dtData.Rows(0).Item("Divisioncode2")
            DDivisionCode3.Text = dtData.Rows(0).Item("Divisioncode3")
            DDivisionCode4.Text = dtData.Rows(0).Item("Divisioncode4")
            DEmpID1.Text = dtData.Rows(0).Item("EmpID1")
            DEmpID2.Text = dtData.Rows(0).Item("EmpID2")
            DEmpID4.Text = dtData.Rows(0).Item("EmpID4")
            DEmpName1.Text = dtData.Rows(0).Item("EmpName1")
            DEmpName2.Text = dtData.Rows(0).Item("EmpName2")
            DEmpName4.Text = dtData.Rows(0).Item("EmpName4")
            DJobTitle1.Text = dtData.Rows(0).Item("Jobtitle1")
            DJobTitle2.Text = dtData.Rows(0).Item("Jobtitle2")
            DJobTitle4.Text = dtData.Rows(0).Item("Jobtitle4")
            DPassDate.Text = dtData.Rows(0).Item("PassDate")
            DPhoneNo.Text = dtData.Rows(0).Item("PhoneNo")
            DQCNO.Text = dtData.Rows(0).Item("QCNO")
            DQCNO1.Text = dtData.Rows(0).Item("QCNO1")

            If dtData.Rows(0).Item("ChkPhone") = 1 Then
                DChkPhone.Checked = True
            Else
                DChkPhone.Checked = False
            End If

            '  DPhoneNo.Text = dtData.Rows(0).Item("PhoneNo")

            If dtData.Rows(0).Item("chkInsurance") = 1 Then
                DchkInsurance.Checked = True
            Else
                DchkInsurance.Checked = False
            End If

            If dtData.Rows(0).Item("chkVisa") = 1 Then
                DchkVisa.Checked = True
            Else
                DchkVisa.Checked = False
            End If




            DAirTickets.Text = dtData.Rows(0).Item("AirTickets")
            DObject.Text = dtData.Rows(0).Item("Object")
            DLocation.Text = dtData.Rows(0).Item("Location")
            DSDate.Text = dtData.Rows(0).Item("SDate")
            DEDate.Text = dtData.Rows(0).Item("EDate")
            DRemark.Text = dtData.Rows(0).Item("Remark")


            '明細
            SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
            SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
            SQL = SQL + " Days,Money,Currency, FlyInf,SFly,EFly,HotelInf, Remark  from  F_BusinessTripSheetDT where 1=1 "
            SQL = SQL + " and NO = '" + DNo.Text + "'"
            SQL = SQL + " order by unique_id "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
       
            

 
        End If
    End Sub


    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub
End Class
