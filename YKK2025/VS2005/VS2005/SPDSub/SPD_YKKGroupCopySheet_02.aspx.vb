Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class SPD_YKKGroupCopySheet_02
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
    Dim wNo As String = ""          '表單-委託單No
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""    '姓名代理用

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
        TopPosition()
        If Not IsPostBack Then
            ShowFormData()          ' 顯示表單資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Top = 450
        Top = Top + 85
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
        Response.Cookies("PGM").Value = "SPD_YKKGroupSheet_02.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.Cookies("UserID").Value)
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
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
        Dim Path As String = uCommon.GetAppSetting("Http")
        Dim SQL As String
        SQL = "Select * From F_YKKGroupCopySheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dt.Rows(0).Item("No")                           ' No
            DDate.Text = dt.Rows(0).Item("Date")                       ' 申請日
            DPerson.Text = dt.Rows(0).Item("Person")                       ' 申請人姓名
            DDivision.Text = dt.Rows(0).Item("Division")               ' 部門
            DBuyer.Text = dt.Rows(0).Item("Buyer")                   ' BUYER
            DProvideDate.Text = dt.Rows(0).Item("ProvideDate")                   ' ProvideDate
            DMapNo.Text = dt.Rows(0).Item("MapNo")                   ' MAPNO
            DSliderCode.Text = dt.Rows(0).Item("SliderCode")                   ' SliderCode
            DWaveCode.Text = dt.Rows(0).Item("WaveCode")                   ' WAVECODE
            DYKKGroup.Text = dt.Rows(0).Item("YKKGROUP") 'YKKGROUP

            DCopyReason.Text = dt.Rows(0).Item("CopyReason")                   ' copyReason
            DForcast.Text = dt.Rows(0).Item("Forcast")                   ' Forcast
            DRemark.Text = dt.Rows(0).Item("Remark")                   ' Remark
            If dt.Rows(0).Item("OFormNo") = "000003" Then
                LFormsno.NavigateUrl = Path & "WorkFlow/ManufInSheet_02.aspx?pFormNo=" & dt.Rows(0).Item("OFormNo") & "&pFormSno=" & dt.Rows(0).Item("OFormSNo")
                LFormsno.Visible = True
            ElseIf dt.Rows(0).Item("OFormNo") = "000007" Then
                LFormsno.NavigateUrl = Path & "WorkFlow/ManufOutSheet_02.aspx?pFormNo=" & dt.Rows(0).Item("OFormNo") & "&pFormSno=" & dt.Rows(0).Item("OFormSNo")
                LFormsno.Visible = True
            Else
                LFormsno.Visible = False
            End If
            If dt.Rows(0).Item("CopyCheck1") = 1 Then
                DCopyCheck1.Checked = True
            ElseIf dt.Rows(0).Item("CopyCheck2") = 1 Then
                DCopyCheck2.Checked = True
            Else
                DCopyCheck3.Checked = True
            End If
        End If
    End Sub
 
End Class
