Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class IRWNoOrderReportSheet_02
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間
        SetParameter()                  ' 設定共用參數
        ' 設定top
        If Not IsPostBack Then
            ShowFormData()          ' 顯示表單資料
            NewAttachFilePath()
            NewAttachFilePath2()

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
        Response.Cookies("PGM").Value = "IRWNoOrderReportSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
    End Sub
  
 
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                    CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("Http") & _
                    System.Configuration.ConfigurationManager.AppSettings("ComplaintInPath")
        Dim FileName As String = ""

        Dim SQL As String
        SQL = "Select * From F_NoOrderReortSheet"
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNO.Text = dtData.Rows(0).Item("No")                         'No
            DDepName.Text = dtData.Rows(0).Item("DepName")
            DLeadDate.Text = dtData.Rows(0).Item("LeadDate")
            DAppID.Text = dtData.Rows(0).Item("AppID")
            DAppName.Text = dtData.Rows(0).Item("AppName")


            DApplyQty.Text = dtData.Rows(0).Item("ApplyQty")
            DOrderQty.Text = dtData.Rows(0).Item("OrderQty")

            DNoOrderQty.Text = dtData.Rows(0).Item("NoOrderQty")
            DPercen.Text = dtData.Rows(0).Item("Percen")
            DRemark.Text = dtData.Rows(0).Item("Percen")
            DNoworkHour1.Text = dtData.Rows(0).Item("NoworkHour1")
            DNoworkHour2.Text = dtData.Rows(0).Item("NoworkHour2")

            If dtData.Rows(0).Item("CContent") <> "" Then
                LCContent.NavigateUrl = dtData.Rows(0).Item("CContent")  '附件

            End If


            If dtData.Rows(0).Item("ApplyHistory") <> "" Then
                LApplyHistory.NavigateUrl = dtData.Rows(0).Item("ApplyHistory") '附件
            End If

            If dtData.Rows(0).Item("RRule") <> "" Then
                LRRule.NavigateUrl = dtData.Rows(0).Item("RRule")  '附件
            End If


            DRReason.Text = dtData.Rows(0).Item("RReason")
            DMethod.Text = dtData.Rows(0).Item("Method")
            DComment.Text = dtData.Rows(0).Item("Comment")
            DRemark.Text = dtData.Rows(0).Item("Remark")
            DDescription.Text = dtData.Rows(0).Item("Description")



            SQL = "Select Data from M_referp"
            SQL = SQL + " where dkey = 'text'"
            SQL = SQL + " and cat =1171"
            Dim dtdata1 As DataTable = uDataBase.GetDataTable(SQL)
            DText.Text = dtdata1.Rows(0)("Data")





            '核定履歷資料
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            SQL = SQL + "Order by   CreateTime Desc, Unique_ID Desc "
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
        SQL = SQL + " where cat = '1171'"
        SQL = SQL + " and dkey ='AttachfilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If





        OpenDir1 = OpenDir1 + DNO.Text + "/AttachFile1"   '開啟附檔資料夾路徑
 
        '開啟附檔資料夾路徑
        DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定附檔路徑
    '**
    '*****************************************************************
    Sub NewAttachFilePath2()
        Dim SQL As String
        '主檔資料
        SQL = " select  data,replace(data,'/','\')data1  from M_referp"
        SQL = SQL + " where cat = '1171'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If




        OpenDir1 = OpenDir1 + DNO.Text + "/AttachFile2"   '開啟附檔資料夾路徑
 

        '開啟附檔資料夾路徑
        DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")
    End Sub


End Class
