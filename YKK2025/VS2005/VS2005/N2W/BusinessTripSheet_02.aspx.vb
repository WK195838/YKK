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



Partial Class BusinessTripSheet_02
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

        LPrint.NavigateUrl = "BusinessTripSheet_03.aspx?pformno=" & wFormNo & "&pformsno=" & wFormSno



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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(CheckTempFile)
    '**     同步化 F_BusinessTripSheetDT & F_BusinessTripSheetDTTemp
    '**
    '*****************************************************************
    Sub CheckTempFile()
        Dim sql As String = "Exec sp_FundingReplication "
        uDataBase.ExecuteNonQuery(sql)
    End Sub


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()



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
            DQCNO.Text = dtData.Rows(0).Item("QCNO")
            DQCNO1.Text = dtData.Rows(0).Item("QCNO1")

         
            If dtData.Rows(0).Item("ChkPhone") = 1 Then
                DChkPhone.Checked = True
            Else
                DChkPhone.Checked = False
            End If

            DPhoneNo.Text = dtData.Rows(0).Item("PhoneNo")

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



            '檢查是否有清算資料
            SQL = " Select  * from F_CloseAccountSheet "
            SQL = SQL + " where sts in (1,0) and tripno = '" + DNo.Text + "'"
            Dim dtTrip As DataTable = uDataBase.GetDataTable(SQL)
            If dtTrip.Rows.Count > 0 Then
                LTripNo.Visible = True
                LTripNo.Text = "有清算申請"
                '  LTripNo.NavigateUrl = "http://10.245.1.10/N2W/BusinessTripSheet_02.aspx?pFormNo=003115&pFormSno=" + Trim(Ddata.Rows(0).Item("formsno"))
                LTripNo.NavigateUrl = "CloseAccountSheet_02.aspx?&pFormno=003115&pFormsno=" + Trim(dtTrip.Rows(0).Item("formsno"))

            End If




            ' 客訴連結
            SQL = " select * from F_ComplaintOutSheet"
            SQL = SQL + "  where no ='" & dtData.Rows(0).Item("QCNO1") & "'"
            Dim Ddata1 As DataTable = uDataBase.GetDataTable(SQL)
            If Ddata1.Rows.Count > 0 Then
                DQCNO.Text = Ddata1.Rows(0).Item("QCNO")
                LQCNo.Visible = True
                LQCNo.Text = "Link"
                '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
                LQCNo.NavigateUrl = "ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
            End If



            '明細
            SQL = " Select  Unique_ID,No, Type, Appoint,convert(char(10),SDate,111)+' '+stime1+':'+stime2 as SDate,"
            SQL = SQL + " convert(char(10),EDate,111)+' '+etime1+':'+etime2 as EDate,"
            SQL = SQL + " Days,Money,Currency, FlyInf,SFly,EFly,HotelInf, Remark  from  F_BusinessTripSheetDT where 1=1 "
            SQL = SQL + " and NO = '" + DNo.Text + "'"
            SQL = SQL + " order by unique_id "
            GridView1.DataSource = uDataBase.GetDataTable(SQL)
            GridView1.DataBind()
            If GridView1.Rows.Count > 0 Then
                DHaveData.Text = "1"
            Else
                DHaveData.Text = "0"
            End If


            '總務備註


            SQL = " Select Unique_ID,No,FlyType,Pay, Money1,Currency1,Remark1,convert(char(10),SDate,111)Sdate,SNo "
            SQL = SQL + "   from  F_BusinessFlyDT  where 1=1 "
            SQL = SQL + " and NO = '" + DNo.Text + "'"
            SQL = SQL + " order by unique_id "

            GridView3.DataSource = uDataBase.GetDataTable(SQL)
            GridView3.DataBind()

            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '3114'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/AttachFile"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/AttachFile"
            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")



            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")


            '核定履歷資料
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            GridView2.DataSource = uDataBase.GetDataTable(SQL)
            GridView2.DataBind()
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑 1
    '**
    '*****************************************************************
    Sub NewAttachFilePath()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '3114'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If






        OpenDir1 = OpenDir1 + DNo.Text + "/AttachFile"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text + "\AttachFile"
        Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
        If dInfo.Exists Then

        Else
            'dInfo.Create()
            My.Computer.FileSystem.CreateDirectory(tempFolderPath)
        End If


        Dim dirInfo As New System.IO.DirectoryInfo(tempFolderPath)


        Dim FileDir As Integer  '資料夾
        FileDir = dirInfo.GetDirectories("*").Length
        Dim FileCount As Integer '檔案
        FileCount = dirInfo.GetFiles("*.*").Length
 

        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub

    Protected Sub GridView3_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowCreated
        e.Row.Cells(0).Visible = False
    End Sub
End Class
