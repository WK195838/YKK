Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class Simulation
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOP1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP1D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP2D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP3D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP8 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP7 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP8D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP7D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP6D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP5D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP4D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP10 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP10D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP9I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP8I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP4I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP5I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP6I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP7I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP3I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP2I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP1I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP10I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP14I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP14 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP14D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP12 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP15 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP11 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP12D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP11D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP13I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP12I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP11I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP15D As System.Web.UI.WebControls.TextBox
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents DFormName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP15I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP20 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP20D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP16I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP17 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP17D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP17I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP18 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP18D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP18I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP19 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP19D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP19I As System.Web.UI.WebControls.Image
    Protected WithEvents LOP6 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP5 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP7 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP8 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP9 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP10 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP11 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP12 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP13 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP14 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP15 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP16 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP17 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP18 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP19 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP20 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOP21 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP24I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP20I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP21D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP21I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP22 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP22D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP22I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP23 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP23D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP23I As System.Web.UI.WebControls.Image
    Protected WithEvents DOP24 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP24D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP25 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOP25D As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOP25 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP24 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP23 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP22 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOP21 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOP25I As System.Web.UI.WebControls.Image

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer = 0     '單號
    Dim wStep As Integer            '工程代碼
    Dim wUserID As String           '核定者ID
    Dim wLevel As String            '難易度
    Dim wAllocateID As String       '指定ID
    Dim wAgentID As String = ""
    Dim wMultiJob As Integer = 0
    Dim wQCLT As Integer = 0
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String             '中壢行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wDepo As String = "CL1"      '中壢行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
    'Modify-End

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "Simulation.aspx"

        If Not Me.IsPostBack Then
            SetParameter()  '設定共用參數
            SetFormItem()   '設定畫面各項目初始值
            SetContent()    '設定畫面工程內容
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '單號
        wStep = Request.QueryString("pStep")        '工程代碼
        wUserID = Request.QueryString("pUserID")    '核定者ID
        wLevel = Request.QueryString("pLevel")      '難易度
        wAllocateID = Request.QueryString("pAllocateID") '指定ID
        wDepo = Request.QueryString("pDepo")        '行事曆
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定畫面工程內容
    '**
    '*****************************************************************
    Sub SetContent()
        Dim SQL, Str As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        '--Start OP---------
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step =  '" & wStep & "'"
        SQL = SQL & "   And Action = 0 "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_FLOW_1")
        If DBDataSet1.Tables("M_FLOW_1").Rows.Count > 0 Then
            DFormName.Text = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("FormName")
            DOP1.Text = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("StepName")

            SQL = "Select UserName From M_Users "
            SQL = SQL & " Where Active = 1 "
            If wUserID <> "" Then
                SQL = SQL & "   And UserID =  '" & wUserID & "'"
            Else
                SQL = SQL & "   And UserID =  '" & Request.Cookies("UserID").Value & "'"
            End If
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_Users_1")
            If DBDataSet1.Tables("M_Users_1").Rows.Count > 0 Then
                Str = DBDataSet1.Tables("M_Users_1").Rows(0).Item("UserName")
            End If
        End If

        If wStep = 1 Then
            DOP1D.Text = "委託人：" & Str & Chr(13) & _
                         "委託時間：" & FormatDateTime(Now(), DateFormat.GeneralDate)
        Else
            DOP1D.Text = "擔當：" & Str & Chr(13) & _
                         "預定完成時間：" & FormatDateTime(Now(), DateFormat.GeneralDate)
        End If

        SetLayout(1)

        '--取得下一關參數設定---------
        Dim pNextGate(10) As String
        Dim pAgentGate(10) As String
        Dim pNextStep As Integer = 0
        Dim pFlowType As Integer = 0    '0=通知
        Dim pCount As Integer
        '--取得LeadTime參數設定---------
        Dim pCTime, pStartTime, pEndTime As DateTime
        '--取得工程負荷參數設定---------
        Dim pLastTime As DateTime
        Dim pCount1 As Integer

        Dim xStep As Integer = wStep    '工程代碼
        Dim xUserID As String = Request.Cookies("UserID").Value  '簽核者ID
        Dim xApplyID As String = Request.Cookies("UserID").Value '申請者ID
        Dim xOP As String = ""          '工程名
        Dim xUser As String = ""       '擔當名
        Dim xNow As DateTime = Now()    '日時

        Dim pAction As Integer = 0
        Dim RtnCode As Integer = 0
        Dim i As Integer = 2
        Dim j As Integer

        If wFormSno <> 0 Then
            DBDataSet1.Clear()
            SQL = "Select ApplyID From T_WaitHandle "
            SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '1' "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle_1")
            If DBDataSet1.Tables("T_WaitHandle_1").Rows.Count > 0 Then
                xApplyID = DBDataSet1.Tables("T_WaitHandle_1").Rows(0).Item("ApplyID")
            End If
        End If

        While pNextStep <> 999
            DBDataSet1.Clear()
            For j = 1 To 10
                pNextGate(j) = ""
            Next j

            '取得下一關
            RtnCode = oCommon.GetNextGate(wFormNo, xStep, xUserID, xApplyID, wAgentID, wAllocateID, wMultiJob, _
                                          pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
            '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
            '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

            SQL = "Select * From M_Flow "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And Step =  '" & pNextStep & "'"
            SQL = SQL & "   And Action =  '" & pAction & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "M_FLOW_1")
            If DBDataSet1.Tables("M_FLOW_1").Rows.Count > 0 Then
                xOP = DBDataSet1.Tables("M_FLOW_1").Rows(0).Item("StepName")
            Else
                xOP = "Not Found"
            End If

            If pFlowType = 0 Then xOP = xOP & "(通知)"
            'If pFlowType = 1 Then xOP = xOP & "(工程)"
            'If pFlowType = 2 Then xOP = xOP & "(委託)"

            If pNextStep <> 999 Then
                '取得工程負荷最後日時
                oSchedule.GetLastTime(pNextGate(1), wFormNo, pNextStep, pFlowType, xNow, pLastTime, pCount1)
                '表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                '取得預定開始,完成日程計算
                oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

                xUser = ""
                For j = 1 To pCount
                    DBDataSet2.Clear()
                    SQL = "Select UserName From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    SQL = SQL & "   And UserID =  '" & pNextGate(j) & "'"
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter3.Fill(DBDataSet2, "M_Users_1")
                    If DBDataSet2.Tables("M_Users_1").Rows.Count > 0 Then
                        If xUser = "" Then
                            xUser = DBDataSet2.Tables("M_Users_1").Rows(0).Item("UserName")
                        Else
                            xUser = xUser & ", " & DBDataSet2.Tables("M_Users_1").Rows(0).Item("UserName")
                        End If
                    End If
                Next j

                Select Case i
                    Case 2
                        DOP2.Text = xOP
                        DOP2D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP2.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 3
                        DOP3.Text = xOP
                        DOP3D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP3.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 4
                        DOP4.Text = xOP
                        DOP4D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP4.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 5
                        DOP5.Text = xOP
                        DOP5D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP5.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 6
                        DOP6.Text = xOP
                        DOP6D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP6.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 7
                        DOP7.Text = xOP
                        DOP7D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP7.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 8
                        DOP8.Text = xOP
                        DOP8D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP8.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 9
                        DOP9.Text = xOP
                        DOP9D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP9.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 10
                        DOP10.Text = xOP
                        DOP10D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP10.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 11
                        DOP11.Text = xOP
                        DOP11D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP11.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 12
                        DOP12.Text = xOP
                        DOP12D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP12.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 13
                        DOP13.Text = xOP
                        DOP13D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP13.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 14
                        DOP14.Text = xOP
                        DOP14D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP14.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 15
                        DOP15.Text = xOP
                        DOP15D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP15.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 16
                        DOP16.Text = xOP
                        DOP16D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP16.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 17
                        DOP17.Text = xOP
                        DOP17D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP17.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 18
                        DOP18.Text = xOP
                        DOP18D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP18.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 19
                        DOP19.Text = xOP
                        DOP19D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP19.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 20
                        DOP20.Text = xOP
                        DOP20D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP20.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 21
                        DOP21.Text = xOP
                        DOP21D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP21.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 22
                        DOP22.Text = xOP
                        DOP22D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP22.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 23
                        DOP23.Text = xOP
                        DOP23D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP23.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 24
                        DOP24.Text = xOP
                        DOP24D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP24.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case 25
                        DOP25.Text = xOP
                        DOP25D.Text = "擔當：" & xUser & Chr(13) & _
                                     "待處理件數：" & CStr(pCount1) & Chr(13) & _
                                     "最後完成預定：" & FormatDateTime(pLastTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "開始預定：" & FormatDateTime(pStartTime, DateFormat.GeneralDate) & Chr(13) & _
                                     "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                        LOP25.NavigateUrl = "OPSchedule_01.aspx?pUserID=" & pNextGate(1) & "&pFormNo=" & wFormNo
                    Case Else
                        DOP25.Text = "元件不足"
                End Select
            Else
                Select Case i   'NextStep=999
                    Case 2
                        DOP2.Text = xOP
                        DOP2D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 3
                        DOP3.Text = xOP
                        DOP3D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 4
                        DOP4.Text = xOP
                        DOP4D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 5
                        DOP5.Text = xOP
                        DOP5D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 6
                        DOP6.Text = xOP
                        DOP6D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 7
                        DOP7.Text = xOP
                        DOP7D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 8
                        DOP8.Text = xOP
                        DOP8D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 9
                        DOP9.Text = xOP
                        DOP9D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 10
                        DOP10.Text = xOP
                        DOP10D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 11
                        DOP11.Text = xOP
                        DOP11D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 12
                        DOP12.Text = xOP
                        DOP12D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 13
                        DOP13.Text = xOP
                        DOP13D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 14
                        DOP14.Text = xOP
                        DOP14D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 15
                        DOP15.Text = xOP
                        DOP15D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 16
                        DOP16.Text = xOP
                        DOP16D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 17
                        DOP17.Text = xOP
                        DOP17D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 18
                        DOP18.Text = xOP
                        DOP18D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 19
                        DOP19.Text = xOP
                        DOP19D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 20
                        DOP20.Text = xOP
                        DOP20D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 21
                        DOP21.Text = xOP
                        DOP21D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 22
                        DOP22.Text = xOP
                        DOP22D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 23
                        DOP23.Text = xOP
                        DOP23D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 24
                        DOP24.Text = xOP
                        DOP24D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case 25
                        DOP25.Text = xOP
                        DOP25D.Text = "完成預定：" & FormatDateTime(pEndTime, DateFormat.GeneralDate)
                    Case Else
                        DOP25.Text = "元件不足"
                End Select
            End If
            SetLayout(i)
            i = i + 1

            xStep = pNextStep  '工程關卡號碼
            xUserID = pNextGate(1)     '簽核者ID
            If pFlowType <> 0 Then xNow = pEndTime '日時
        End While

        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定畫面工程數
    '**
    '*****************************************************************
    Sub SetLayout(ByVal idx As Integer)
        Select Case idx
            Case 1
                DOP1.Visible = True
                DOP1D.Visible = True
                If LOP1.NavigateUrl <> "" Then LOP1.Visible = True
            Case 2
                DOP1I.Visible = True
                DOP2.Visible = True
                DOP2D.Visible = True
                If LOP2.NavigateUrl <> "" Then LOP2.Visible = True
            Case 3
                DOP2I.Visible = True
                DOP3.Visible = True
                DOP3D.Visible = True
                If LOP3.NavigateUrl <> "" Then LOP3.Visible = True
            Case 4
                DOP3I.Visible = True
                DOP4.Visible = True
                DOP4D.Visible = True
                If LOP4.NavigateUrl <> "" Then LOP4.Visible = True
            Case 5
                DOP4I.Visible = True
                DOP5.Visible = True
                DOP5D.Visible = True
                If LOP5.NavigateUrl <> "" Then LOP5.Visible = True
            Case 6
                DOP5I.Visible = True
                DOP6.Visible = True
                DOP6D.Visible = True
                If LOP6.NavigateUrl <> "" Then LOP6.Visible = True
            Case 7
                DOP6I.Visible = True
                DOP7.Visible = True
                DOP7D.Visible = True
                If LOP7.NavigateUrl <> "" Then LOP7.Visible = True
            Case 8
                DOP7I.Visible = True
                DOP8.Visible = True
                DOP8D.Visible = True
                If LOP8.NavigateUrl <> "" Then LOP8.Visible = True
            Case 9
                DOP8I.Visible = True
                DOP9.Visible = True
                DOP9D.Visible = True
                If LOP9.NavigateUrl <> "" Then LOP9.Visible = True
            Case 10
                DOP9I.Visible = True
                DOP10.Visible = True
                DOP10D.Visible = True
                If LOP10.NavigateUrl <> "" Then LOP10.Visible = True
            Case 11
                DOP10I.Visible = True
                DOP11.Visible = True
                DOP11D.Visible = True
                If LOP11.NavigateUrl <> "" Then LOP11.Visible = True
            Case 12
                DOP11I.Visible = True
                DOP12.Visible = True
                DOP12D.Visible = True
                If LOP12.NavigateUrl <> "" Then LOP12.Visible = True
            Case 13
                DOP12I.Visible = True
                DOP13.Visible = True
                DOP13D.Visible = True
                If LOP13.NavigateUrl <> "" Then LOP13.Visible = True
            Case 14
                DOP13I.Visible = True
                DOP14.Visible = True
                DOP14D.Visible = True
                If LOP14.NavigateUrl <> "" Then LOP14.Visible = True
            Case 15
                DOP14I.Visible = True
                DOP15.Visible = True
                DOP15D.Visible = True
                If LOP15.NavigateUrl <> "" Then LOP15.Visible = True
            Case 16
                DOP15I.Visible = True
                DOP16.Visible = True
                DOP16D.Visible = True
                If LOP16.NavigateUrl <> "" Then LOP16.Visible = True
            Case 17
                DOP16I.Visible = True
                DOP17.Visible = True
                DOP17D.Visible = True
                If LOP17.NavigateUrl <> "" Then LOP17.Visible = True
            Case 18
                DOP17I.Visible = True
                DOP18.Visible = True
                DOP18D.Visible = True
                If LOP18.NavigateUrl <> "" Then LOP18.Visible = True
            Case 19
                DOP18I.Visible = True
                DOP19.Visible = True
                DOP19D.Visible = True
                If LOP19.NavigateUrl <> "" Then LOP19.Visible = True
            Case 20
                DOP19I.Visible = True
                DOP20.Visible = True
                DOP20D.Visible = True
                If LOP20.NavigateUrl <> "" Then LOP20.Visible = True
            Case 21
                DOP20I.Visible = True
                DOP21.Visible = True
                DOP21D.Visible = True
                If LOP21.NavigateUrl <> "" Then LOP21.Visible = True
            Case 22
                DOP21I.Visible = True
                DOP22.Visible = True
                DOP22D.Visible = True
                If LOP22.NavigateUrl <> "" Then LOP22.Visible = True
            Case 23
                DOP22I.Visible = True
                DOP23.Visible = True
                DOP23D.Visible = True
                If LOP23.NavigateUrl <> "" Then LOP23.Visible = True
            Case 24
                DOP23I.Visible = True
                DOP24.Visible = True
                DOP24D.Visible = True
                If LOP24.NavigateUrl <> "" Then LOP24.Visible = True
            Case 25
                DOP24I.Visible = True
                DOP25.Visible = True
                DOP25D.Visible = True
                If LOP25.NavigateUrl <> "" Then LOP25.Visible = True
            Case Else
        End Select
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定畫面各項目初始值
    '**
    '*****************************************************************
    Sub SetFormItem()
        '畫面各項目表示處理
        DOP1.Visible = False
        DOP1D.Visible = False
        DOP1I.Visible = False
        DOP2.Visible = False
        DOP2D.Visible = False
        DOP2I.Visible = False
        DOP3.Visible = False
        DOP3D.Visible = False
        DOP3I.Visible = False
        DOP4.Visible = False
        DOP4D.Visible = False
        DOP4I.Visible = False
        DOP5.Visible = False
        DOP5D.Visible = False
        DOP5I.Visible = False
        DOP6.Visible = False
        DOP6D.Visible = False
        DOP6I.Visible = False
        DOP7.Visible = False
        DOP7D.Visible = False
        DOP7I.Visible = False
        DOP8.Visible = False
        DOP8D.Visible = False
        DOP8I.Visible = False
        DOP9.Visible = False
        DOP9D.Visible = False
        DOP9I.Visible = False
        DOP10.Visible = False
        DOP10D.Visible = False
        DOP10I.Visible = False
        DOP11.Visible = False
        DOP11D.Visible = False
        DOP11I.Visible = False
        DOP12.Visible = False
        DOP12D.Visible = False
        DOP12I.Visible = False
        DOP13.Visible = False
        DOP13D.Visible = False
        DOP13I.Visible = False
        DOP14.Visible = False
        DOP14D.Visible = False
        DOP14I.Visible = False
        DOP15.Visible = False
        DOP15D.Visible = False
        DOP15I.Visible = False
        DOP16.Visible = False
        DOP16D.Visible = False
        DOP16I.Visible = False
        DOP17.Visible = False
        DOP17D.Visible = False
        DOP17I.Visible = False
        DOP18.Visible = False
        DOP18D.Visible = False
        DOP18I.Visible = False
        DOP19.Visible = False
        DOP19D.Visible = False
        DOP19I.Visible = False
        DOP20.Visible = False
        DOP20D.Visible = False
        DOP20I.Visible = False
        DOP21.Visible = False
        DOP21D.Visible = False
        DOP21I.Visible = False
        DOP22.Visible = False
        DOP22D.Visible = False
        DOP22I.Visible = False
        DOP23.Visible = False
        DOP23D.Visible = False
        DOP23I.Visible = False
        DOP24.Visible = False
        DOP24D.Visible = False
        DOP24I.Visible = False
        DOP25.Visible = False
        DOP25D.Visible = False
        DOP25I.Visible = False

        LOP1.Visible = False
        LOP2.Visible = False
        LOP3.Visible = False
        LOP4.Visible = False
        LOP5.Visible = False
        LOP6.Visible = False
        LOP7.Visible = False
        LOP8.Visible = False
        LOP9.Visible = False
        LOP10.Visible = False
        LOP11.Visible = False
        LOP12.Visible = False
        LOP13.Visible = False
        LOP14.Visible = False
        LOP15.Visible = False
        LOP16.Visible = False
        LOP17.Visible = False
        LOP18.Visible = False
        LOP19.Visible = False
        LOP20.Visible = False
        LOP21.Visible = False
        LOP22.Visible = False
        LOP23.Visible = False
        LOP24.Visible = False
        LOP25.Visible = False

    End Sub

    Private Sub Textbox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DOP22.TextChanged

    End Sub
End Class
