Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class ISMSSheet_02
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
        '
        Response.Cookies("PGM").Value = "ISMSSheet_01.aspx"
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
        SQL = "Select * From F_ISMSSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtISMS As DataTable = uDataBase.GetDataTable(SQL)
        If dtISMS.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtISMS.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtISMS.Rows(0).Item("Date")                         'No 
            DDepName.Text = dtISMS.Rows(0).Item("DepName")                         'No
            DAppName.Text = dtISMS.Rows(0).Item("AppName")                         'No
            DCheckDate.Value = dtISMS.Rows(0).Item("Checkdate")
            DITName.Items.Add(dtISMS.Rows(0).Item("ITName"))
            DITName.SelectedValue = dtISMS.Rows(0).Item("ITName")
            DPlace.Items.Add(dtISMS.Rows(0).Item("Place"))
            DPlace.SelectedValue = dtISMS.Rows(0).Item("Place")
            DEroom.Items.Add(dtISMS.Rows(0).Item("Eroom"))
            DEroom.SelectedValue = dtISMS.Rows(0).Item("Eroom")
            DTemp.Text = dtISMS.Rows(0).Item("Temp")
            DHumidity.Text = dtISMS.Rows(0).Item("Humidity")
            DOther.Text = dtISMS.Rows(0).Item("Other")
            DState.Text = dtISMS.Rows(0).Item("State")
            DCause.Text = dtISMS.Rows(0).Item("Cause")
            DDeal.Text = dtISMS.Rows(0).Item("Deal")
            DRemark.Text = dtISMS.Rows(0).Item("Remark")
        End If
     

    End Sub

    

End Class
