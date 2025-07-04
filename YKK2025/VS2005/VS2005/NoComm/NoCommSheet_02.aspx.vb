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


Partial Class NoCommSheet_02
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
    Dim OpenDir1 As String = ""
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SetParameter()
        ShowFormData()


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
        Response.Cookies("PGM").Value = "EApprovalRDSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub


    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        'Dim Path As String = "http://localhost:60679/EApproval/Document/001172/"  '測試環境

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http1") & _
        System.Configuration.ConfigurationManager.AppSettings("EApprovalPath")

        Dim SQL As String
        SQL = "Select * From F_NoCommSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            DDate.Text = dtData.Rows(0).Item("Date")
            DDivision.Text = dtData.Rows(0).Item("Division")
            DAppName.Text = dtData.Rows(0).Item("AppName")
            DEmpName.Text = dtData.Rows(0).Item("EmpName")
            DCust.Text = dtData.Rows(0).Item("Cust")
            DCustName.Text = dtData.Rows(0).Item("CustName")

            If dtData.Rows(0).Item("Supplier") = "1" Then
                DSUPPLIER.Checked = True
            Else
                DSUPPLIER.Checked = False
            End If

            '細項連結
            LNoCommList.NavigateUrl = "PNoCommList.aspx?pCust=" + DCust.Text + "&pcode=" + dtData.Rows(0).Item("Item") + "&pNo=" + dtData.Rows(0).Item("No")


            DComplainNo.Text = dtData.Rows(0).Item("ComplainNo")
            DQty.Text = dtData.Rows(0).Item("Qty")
            DAmount.Text = dtData.Rows(0).Item("Amount")
            SetFieldData("SReason", dtData.Rows(0).Item("SReason"))
            DCDescrption.Text = dtData.Rows(0).Item("CDescrption")
            DEDescrption.Text = dtData.Rows(0).Item("EDescrption")
            DComment1.Text = dtData.Rows(0).Item("Comment1")
            DComment2.Text = dtData.Rows(0).Item("Comment2")
            DComment3.Text = dtData.Rows(0).Item("Comment3")

            '客訴內容
            If DComplainNo.Text <> "" Then
                SQL = " Select * from F_ComplaintOutSheet "
                SQL = SQL & "  where no='" + DComplainNo.Text + "'"

                Dim DBUser3 As DataTable = uDataBase.GetDataTable(SQL)
                If DBUser3.Rows.Count > 0 Then

                    LQCNo.Visible = True
                    LQCNo.Text = "LINK"
                    '  LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(Ddata1.Rows(0).Item("formsno"))
                    LQCNo.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutSheet_02.aspx?&pFormno=003109&pFormsno=" + Trim(DBUser3.Rows(0).Item("formsno"))
 
                End If
            End If
       

            NewAttachFilePath()



            '主檔資料
            SQL = " select  data,replace(data,'/','\')data1  from M_referp"
            SQL = SQL + " where cat = '1172'"
            SQL = SQL + " and dkey ='AttachfilePath1'"
            Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
            Dim OpenDir As String = ""
            If DBAdapter3.Rows.Count > 0 Then
                OpenDir = "file://" + DBAdapter3.Rows(0).Item("Data") + DNo.Text + "/附件"
            End If
            Dim OpenDir2 As String = ""
            OpenDir2 = "\\" + DBAdapter3.Rows(0).Item("Data1") + DNo.Text + "\附件"

            Dim dirInfo As New System.IO.DirectoryInfo(OpenDir2)

            Dim FileDir As Integer  '資料夾
            FileDir = dirInfo.GetDirectories("*").Length
            Dim FileCount As Integer '檔案
            FileCount = dirInfo.GetFiles("*.*").Length

            DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir + "','_blank');return false;")
            '  DAttachfile2.Attributes.Add("onclick", "window.open('" + OpenDir2 + "','_blank');return false;")



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
    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim sql As String = ""
        Dim idx As Integer = 1
        Dim i As Integer
        '擔當者及部門 

        '製造1
        If pFieldName = "SReason" Then
            DSReason.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSReason.Items.Add(ListItem1)
                End If
            Else
                sql = "Select  Data  From M_Referp Where Cat='1172' and dkey ='SReason'  Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)

                DSReason.Items.Add("")
                For i = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = Trim(dtReasonCode.Rows(i)("Data"))
                    ListItem1.Value = Trim(dtReasonCode.Rows(i)("Data"))
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSReason.Items.Add(ListItem1)
                Next
            End If
        End If




    End Sub


    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**
    ''**     設定附檔路徑 1
    ''**
    ''*****************************************************************
    'Sub NewAttachFilePath()
    '    Dim SQL As String
    '    '主檔資料
    '    SQL = " select  data,replace(data,'/','\')data1  from M_referp"
    '    SQL = SQL + " where cat = '1172'"
    '    SQL = SQL + " and dkey ='AttachfilePath1'"
    '    Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

    '    If DBAdapter1.Rows.Count > 0 Then
    '        OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
    '        OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
    '    End If




    '    D3.Text = Now.ToString("yyyyMMddHHmmss")




    '    OpenDir1 = OpenDir1 + D3.Text + "/附件"   '開啟附檔資料夾路徑


    '    '開啟附檔資料夾路徑
    '    DAttachfile1.Attributes.Add("onclick", "window.open('" + OpenDir1 + "','_blank');return true;")

    'End Sub



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
        SQL = SQL + " where cat = '1172'"
        SQL = SQL + " and dkey ='AttachfilePath1'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        If DBAdapter1.Rows.Count > 0 Then
            OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If







        OpenDir1 = OpenDir1 + DNo.Text + "/附件"   '開啟附檔資料夾路徑

        '檢查目錄是否存，不存在就新建一個
        Dim tempFolderPath As String
        tempFolderPath = OpenDir2 + DNo.Text + "\附件"
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




End Class
