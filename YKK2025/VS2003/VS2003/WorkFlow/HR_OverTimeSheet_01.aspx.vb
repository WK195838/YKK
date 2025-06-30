Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_OverTimeSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOverTimeDate As System.Web.UI.WebControls.Button
    Protected WithEvents DOverTimeDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFood As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAStartM As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DFAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOverTimeSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDateType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTraffic As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DCVacation As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DName As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BOverTime As System.Web.UI.WebControls.Button
    Protected WithEvents D46Hours As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents D As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAgentID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX15H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX167H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX20H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX267H As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX15M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX167M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX20M As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFPRTAX267M As System.Web.UI.WebControls.TextBox

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
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Dim wDepo As String = "TP"      '台北行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End
    Dim wUserName As String = ""    '姓名代理用

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示
            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
            End If
            SetPopupFunction()      '設定彈出視窗事件
        Else
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查
            ShowMessage()           '上傳資料檢查及顯示訊息
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = Request.QueryString("pStep")        '工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        'Add-Start by Joy 2009/11/20(2010行事曆對應)
        '申請者-群組行事曆(TP1->台北-拉鏈, TP2->台北-建材, CL1->中壢-間接, CL2->中壢-製造)
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        'Add-End

        ' PAYROLL FIELD-START
        DFPRAH1.Style("left") = -500 & "px"
        DFPRAM1.Style("left") = -500 & "px"
        DFPRAH2.Style("left") = -500 & "px"
        DFPRAM2.Style("left") = -500 & "px"
        DFPRBH.Style("left") = -500 & "px"
        DFPRBM.Style("left") = -500 & "px"
        DFPRCH.Style("left") = -500 & "px"
        DFPRCM.Style("left") = -500 & "px"

        DFPRTAX15H.Style("left") = -500 & "px"
        DFPRTAX15M.Style("left") = -500 & "px"
        DFPRTAX167H.Style("left") = -500 & "px"
        DFPRTAX167M.Style("left") = -500 & "px"
        DFPRTAX20H.Style("left") = -500 & "px"
        DFPRTAX20M.Style("left") = -500 & "px"
        DFPRTAX267H.Style("left") = -500 & "px"
        DFPRTAX267M.Style("left") = -500 & "px"

        DAgentID.Style("left") = -500 & "px"
        DSalaryYM1.Style("left") = -500 & "px"
        ' END

        Response.Cookies("PGM").Value = "HR_OverTimeSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        'Dim Message As String = ""

        'Check切結書
        'If DContactFile.Visible Then
        'If DContactFile.PostedFile.FileName <> "" Then
        'If Message = "" Then
        'Message = "切結書"
        'Else
        '    Message = Message & ", " & "切結書"
        'End If
        'End If
        'End If

        'If Message <> "" Then
        'Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
        'Response.Write(YKK.ShowMessage(Message))
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, Request.QueryString("pUserID"))
        '
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者
        'Modify-End
        '表單號碼,表單流水號,工程關卡號碼,序號,台北,簽核者
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("OverTimeFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_OverTimeSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_OverTimeSheet")
        If DBDataSet1.Tables("F_OverTimeSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("No")                     'No
            DOverTimeDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("OverTimeDate")   '加班日期
            DDateType.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DateType")           '加班別
            DSalaryYM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")           '所屬年月
            DSalaryYM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")          '所屬年月-檢查使用
            DDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Date")                   '申請日期
            SetFieldData("Name", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Name"))          '姓名
            DEmpID.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobTitle")           '職稱
            DJobCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobCode")             '職稱代碼
            DDepoName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoName")           '公司別
            DDepoCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoCode")           '公司別代碼
            DDivision.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Division")           '部門
            DDivisionCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DivisionCode")   '部門代碼
            SetFieldData("CVacation", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("CVacation"))    '調休
            SetFieldData("Food", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Food"))          '伙食
            SetFieldData("Traffic", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Traffic"))    '交通

            SetFieldData("BStartH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartH").ToString)   '預定開始-時
            SetFieldData("BStartM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartM").ToString)   '預定開始-分
            SetFieldData("BEndH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndH").ToString)       '預定終止-時
            SetFieldData("BEndM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndM").ToString)       '預定終止-分
            DBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BH").ToString                      '計算結果-時
            DBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BM").ToString                      '計算結果-分

            SetFieldData("AStartH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartH").ToString)   '實際開始-時
            SetFieldData("AStartM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartM").ToString)   '實際開始-分
            SetFieldData("AEndH", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndH").ToString)       '實際終止-時
            SetFieldData("AEndM", DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndM").ToString)       '實際終止-分
            DAH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AH").ToString                      '計算結果-時
            DAM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AM").ToString                      '計算結果-分

            DFAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH1").ToString                  '核定平日2內-時
            DFAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM1").ToString                  '核定平日2內-分
            DFAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH2").ToString                  '核定平日2外-時
            DFAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM2").ToString                  '核定平日2外-分

            DFBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBH").ToString                    '核定假日-時
            DFBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBM").ToString                    '核定假日-分
            DFCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCH").ToString                    '核定國定假日-時
            DFCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCM").ToString                    '核定國定假日-分

            ' PAYROLL FIELD-START
            DFPRAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAH1").ToString              'PAYROLL核定平日2內-時
            DFPRAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAM1").ToString              'PAYROLL核定平日2內-分
            DFPRAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAH2").ToString              'PAYROLL核定平日2外-時
            DFPRAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRAM2").ToString              'PAYROLL核定平日2外-分

            DFPRBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRBH").ToString                'PAYROLL核定假日-時
            DFPRBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRBM").ToString                'PAYROLL核定假日-分
            DFPRCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRCH").ToString                'PAYROLL核定國定假日-時
            DFPRCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRCM").ToString                'PAYROLL核定國定假日-分

            ' JOY BY 2017/6/19
            DFPRTAX15H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX15H").ToString        'PAYROLL核定應稅1.5
            DFPRTAX15M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX15M").ToString        '
            DFPRTAX167H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX167H").ToString      'PAYROLL核定應稅1.67
            DFPRTAX167M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX167M").ToString      '
            DFPRTAX20H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX20H").ToString        'PAYROLL核定應稅2.0
            DFPRTAX20M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX20M").ToString        '
            DFPRTAX267H.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX267H").ToString      'PAYROLL核定應稅2.67
            DFPRTAX267M.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FPRTAX267M").ToString      '

            ' END
            DFReason.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FReason")                     '加班理由
            '
            '超過46H/92H
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' 平日已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' 平日未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' 總已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' 總未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "總時數：[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "平日時數：[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-淺灰色
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-橘色
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-粉紅色
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' 白色
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
            '交易資料
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '說明

                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '延遲理由代碼
                        If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '延遲理由
                            DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '延遲其他說明
                        End If

                    End If
                End If
            End If
            '核定履歷資料
            SQL = "SELECT "
            SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
            SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
            SQL = SQL + "AEndTimeDesc As Description, "
            SQL = SQL + "URL "
            SQL = SQL + "FROM V_WaitHandle_01 "
            SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
            SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
            'SQL = SQL + "Order by Unique_ID Desc "
            SQL = SQL + "Order by CreateTime Desc, Step Desc, SeqNo Desc "
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet9, "DecideHistory")
            DataGrid9.DataSource = DBDataSet9
            DataGrid9.DataBind()
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     延遲理由代碼點選後事件
    '**
    '*****************************************************************
    Private Sub DReasonCode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DReasonCode.SelectedIndexChanged
        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            '電子簽章未使用
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Attach") = 1 Then
                'DRefMapFile.Visible = True
                'DMapFile.Visible = True
            Else
                'DRefMapFile.Visible = False
                'DMapFile.Visible = False
            End If
            '儲存按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("SaveDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
                '新
                BSAVE.Attributes("onclick") = "this.disabled = true;" & "Button('SAVE', '" + BSAVE.Value + "');"
                '--
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                '新
                BNG1.Attributes("onclick") = "this.disabled = true;" & "Button('NG1', '" + BNG1.Value + "');"
                '--
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                '新
                BNG2.Attributes("onclick") = "this.disabled = true;" & "Button('NG2', '" + BNG2.Value + "');"
                '--
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                '-- 防止按鈕按二次 JOY 16/08/17
                '舊
                'BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
                '新
                BOK.Attributes("onclick") = "this.disabled = true;" & "Button('OK', '" + BOK.Value + "');"
                '--
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("Delay") = 1 Then
                wDelay = 1
            End If
        End If

        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DOverTimeSheet1.Visible = True   '表單Sheet-1
                DDescSheet.Visible = True        '說明Sheet

                '遲納管理
                If wDelay = 1 Then
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                    Else
                        DDelay.Visible = False  '延遲Sheet
                    End If
                End If
                If DDelay.Visible = True Then
                    DReasonCode.Visible = True     '延遲理由代碼
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                    DReason.Visible = True         '延遲理由
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                    DReasonDesc.Visible = True     '延遲其他說明
                    Top = 896
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 782
                End If

                '欄位顯示
                DDecideDesc.Visible = True      '說明
                '說明需輸入
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If

                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                '核定履歷
                DHistoryLabel.Style.Add("Top", Top + 24)  '核定履歷
                DataGrid9.Style.Add("Top", Top + 48)     '核定履歷
            End If
        Else
            Top = 704
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'BOverTimeDate.Attributes("onclick") = "CalendarPicker('" + wDepo + "', 'Form1.DOverTimeDate', 'Form1.DDateType', 'Form1.DSalaryYM');"  '加班日期
        '
        BOverTimeDate.Attributes("onclick") = "CalendarPicker('" + wApplyCalendar + "', 'Form1.DOverTimeDate', 'Form1.DDateType', 'Form1.DSalaryYM');"  '加班日期
        'Modify-End
        BCardTime.Attributes("onclick") = "ShowCardTime();"    '刷卡記錄
        BOverTime.Attributes("onclick") = "ShowOverTime();"    '加班記錄
        ' 顯示通知訊息
        Dim Cmd As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From W_ReadMessageHistory "
        SQL = SQL & " Where FormNo = '" & wFormNo & "' "
        SQL = SQL & "   And ReadUser = '" & Request.QueryString("pUserID") & "' "
        SQL = SQL & "   And ReadFlag = '" & "1" & "' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "MessageHistory")
        If DBDataSet1.Tables("MessageHistory").Rows.Count <= 0 Then
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/WorkFlowSub/NoticeMessageOverTime.aspx?pUserID=" + Request.QueryString("pUserID") & "&pFormNo=" + wFormNo & "','NoticeMessage','status=0,toolbar=0,width=550,height=700,resizable=no,scrollbars=no');" + _
                  "</script>"
        End If

        OleDbConnection1.Close()
        '
        If Cmd <> "" Then
            Response.Write(Cmd)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        '按鈕及RequestedField的Top位置
        If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 782
                Else
                    If DDelay.Visible = True Then
                        Top = 896
                    Else
                        Top = 782
                    End If
                End If
            End If
        Else
            Top = 704
        End If
        '----
        'DB連結關閉
        OleDbConnection1.Close()
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

        SetFieldAttribute(pPost)     '表單各欄位屬性及欄位輸入檢查等設定
    End Sub
    '*****************************************************************
    '**(SetFieldAttribute)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim wEmpID, wJobTitle, wJobCode, wDivision, wDivisionCode, wDepoName, wDepoCode, wSalaryYM As String
        Dim wDateType As String = "平日"

        OleDbConnection1.Open()
        '取得申請者資訊
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
        SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            wDepoCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            wDepoName = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            'Delete-Start by Joy 2009/11/20(2010行事曆對應)
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "01" Then wDepo = "CL"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "10" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "11" Then wDepo = "TP"
            'If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "51" Then wDepo = "YA"
            'Delete-End
            wUserName = DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")
            wEmpID = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            wJobTitle = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            wJobCode = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            wDivision = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            wDivisionCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        '取得當日加班別
        DBDataSet1.Clear()
        SQL = "Select VacationType From M_Vacation "
        SQL = SQL + "Where Active = '1' "
        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'SQL = SQL + "  and Depo = '" + wDepo + "' "
        '
        SQL = SQL + "  and Depo = '" + wApplyCalendar + "' "
        'Modify-End
        SQL = SQL + "  and YMD  = '" + CStr(DateTime.Now.Today) + "' "
        SQL = SQL + "Order by YMD "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "Vacation")
        If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
            If DBDataSet1.Tables("Vacation").Rows(0).Item("VacationType") = 0 Then
                wDateType = "假日"
            Else
                If DBDataSet1.Tables("Vacation").Rows(0).Item("VacationType") = 1 Then
                    wDateType = "國定假日"
                Else
                    wDateType = "社休假日"
                End If
            End If
        End If
        '取得所屬年月
        If DateTime.Now.Month < 10 Then
            wSalaryYM = CStr(DateTime.Now.Year) + "/0" + CStr(DateTime.Now.Month)
        Else
            wSalaryYM = CStr(DateTime.Now.Year) + "/" + CStr(DateTime.Now.Month)
        End If
        '------------------------------------------------------------------------------------------
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                'DNo.BackColor = Color.LightGray
                DNo.BackColor = Color.White
                DNo.ReadOnly = True
                DNo.Visible = True
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        '加班日期
        Select Case FindFieldInf("OverTimeDate")
            Case 0  '顯示
                DOverTimeDate.BackColor = Color.LightGray
                DOverTimeDate.ReadOnly = True
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = False
            Case 1  '修改+檢查
                DOverTimeDate.BackColor = Color.GreenYellow
                DOverTimeDate.ReadOnly = True
                ShowRequiredFieldValidator("DOverTimeDateRqd", "DOverTimeDate", "異常：需輸入加班日期")
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = True
            Case 2  '修改
                DOverTimeDate.BackColor = Color.Yellow
                DOverTimeDate.ReadOnly = True
                DOverTimeDate.Visible = True
                BOverTimeDate.Visible = True
            Case Else   '隱藏
                DOverTimeDate.Visible = False
                BOverTimeDate.Visible = False
        End Select
        If pPost = "New" Then DOverTimeDate.Text = CStr(DateTime.Now.Today)
        '申請日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = True
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")
                DDate.Visible = True
            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = True
                DDate.Visible = True
            Case Else   '隱藏
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '加班別
        Select Case FindFieldInf("DateType")
            Case 0  '顯示
                DDateType.BackColor = Color.LightGray
                DDateType.Visible = True
            Case 1  '修改+檢查
                DDateType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateTypeRqd", "DDateType", "異常：需輸入加班別")
                DDateType.Visible = True
            Case 2  '修改
                DDateType.BackColor = Color.Yellow
                DDateType.Visible = True
            Case Else   '隱藏
                DDateType.Visible = False
        End Select
        If pPost = "New" Then DDateType.Text = wDateType
        '所屬年月
        Select Case FindFieldInf("SalaryYM")
            Case 0  '顯示
                DSalaryYM.BackColor = Color.LightGray
                DSalaryYM.ReadOnly = True
                DSalaryYM.Visible = True
            Case 1  '修改+檢查
                DSalaryYM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSalaryYMRqd", "DSalaryYM", "異常：需輸入所屬年月")
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case 2  '修改
                DSalaryYM.BackColor = Color.Yellow
                DSalaryYM.ReadOnly = False
                DSalaryYM.Visible = True
            Case Else   '隱藏
                DSalaryYM.Visible = False
        End Select
        If pPost = "New" Then DSalaryYM.Text = wSalaryYM
        '姓名
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.Visible = True
            Case 1  '修改+檢查
                DName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
                DName.Visible = True
            Case 2  '修改
                DName.BackColor = Color.Yellow
                DName.Visible = True
            Case Else   '隱藏
                DName.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Name", "ZZZZZZ")
        'EmpID
        Select Case FindFieldInf("EmpID")
            Case 0  '顯示
                DEmpID.BackColor = Color.LightGray
                DEmpID.ReadOnly = True
                DEmpID.Visible = True
            Case 1  '修改+檢查
                DEmpID.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEmpIDRqd", "DEmpID", "異常：需輸入卡號")
                DEmpID.Visible = True
            Case 2  '修改
                DEmpID.BackColor = Color.Yellow
                DEmpID.Visible = True
            Case Else   '隱藏
                DEmpID.Visible = False
        End Select
        If pPost = "New" Then DEmpID.Text = wEmpID
        '職稱
        Select Case FindFieldInf("JobTitle")
            Case 0  '顯示
                DJobTitle.BackColor = Color.LightGray
                DJobTitle.ReadOnly = True
                DJobTitle.Visible = True
            Case 1  '修改+檢查
                DJobTitle.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobTitleRqd", "DJobTitle", "異常：需輸入職稱")
                DJobTitle.Visible = True
            Case 2  '修改
                DJobTitle.BackColor = Color.Yellow
                DJobTitle.Visible = True
            Case Else   '隱藏
                DJobTitle.Visible = False
        End Select
        If pPost = "New" Then DJobTitle.Text = wJobTitle
        '職稱代碼
        Select Case FindFieldInf("JobCode")
            Case 0  '顯示
                DJobCode.BackColor = Color.LightGray
                DJobCode.ReadOnly = True
                DJobCode.Visible = True
            Case 1  '修改+檢查
                DJobCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DJobCodeRqd", "DJobCode", "異常：需輸入職稱代碼")
                DJobCode.Visible = True
            Case 2  '修改
                DJobCode.BackColor = Color.Yellow
                DJobCode.Visible = True
            Case Else   '隱藏
                DJobCode.Visible = False
        End Select
        If pPost = "New" Then DJobCode.Text = wJobCode
        'Depo Name
        Select Case FindFieldInf("DepoName")
            Case 0  '顯示
                DDepoName.BackColor = Color.LightGray
                DDepoName.ReadOnly = True
                DDepoName.Visible = True
            Case 1  '修改+檢查
                DDepoName.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoNameRqd", "DDepoName", "異常：需輸入公司")
                DDepoName.Visible = True
            Case 2  '修改
                DDepoName.BackColor = Color.Yellow
                DDepoName.Visible = True
            Case Else   '隱藏
                DDepoName.Visible = False
        End Select
        If pPost = "New" Then DDepoName.Text = wDepoName
        'Depo Code
        Select Case FindFieldInf("DepoCode")
            Case 0  '顯示
                DDepoCode.BackColor = Color.LightGray
                DDepoCode.ReadOnly = True
                DDepoCode.Visible = True
            Case 1  '修改+檢查
                DDepoCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDepoCodeRqd", "DDepoCode", "異常：需輸入公司代碼")
                DDepoCode.Visible = True
            Case 2  '修改
                DDepoCode.BackColor = Color.Yellow
                DDepoCode.Visible = True
            Case Else   '隱藏
                DDepoCode.Visible = False
        End Select
        If pPost = "New" Then DDepoCode.Text = wDepoCode
        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                DDivision.ReadOnly = True
                DDivision.Visible = True
            Case 1  '修改+檢查
                DDivision.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionRqd", "DDivision", "異常：需輸入部門")
                DDivision.Visible = True
            Case 2  '修改
                DDivision.BackColor = Color.Yellow
                DDivision.Visible = True
            Case Else   '隱藏
                DDivision.Visible = False
        End Select
        If pPost = "New" Then DDivision.Text = wDivision
        '部門代碼
        Select Case FindFieldInf("DivisionCode")
            Case 0  '顯示
                DDivisionCode.BackColor = Color.LightGray
                DDivisionCode.ReadOnly = True
                DDivisionCode.Visible = True
            Case 1  '修改+檢查
                DDivisionCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDivisionCodeRqd", "DDivisionCode", "異常：需輸入部門代碼")
                DDivisionCode.Visible = True
            Case 2  '修改
                DDivisionCode.BackColor = Color.Yellow
                DDivisionCode.Visible = True
            Case Else   '隱藏
                DDivisionCode.Visible = False
        End Select
        If pPost = "New" Then DDivisionCode.Text = wDivisionCode
        '調休
        Select Case FindFieldInf("CVacation")
            Case 0  '顯示
                DCVacation.BackColor = Color.LightGray
                DCVacation.Visible = True
            Case 1  '修改+檢查
                DCVacation.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCVacationRqd", "DCVacation", "異常：需輸入調休")
                DCVacation.Visible = True
            Case 2  '修改
                DCVacation.BackColor = Color.Yellow
                DCVacation.Visible = True
            Case Else   '隱藏
                DCVacation.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("CVacation", "ZZZZZZ")
        '伙食
        Select Case FindFieldInf("Food")
            Case 0  '顯示
                DFood.BackColor = Color.LightGray
                DFood.Visible = True
            Case 1  '修改+檢查
                DFood.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFoodRqd", "DFood", "異常：需輸入伙食")
                DFood.Visible = True
            Case 2  '修改
                DFood.BackColor = Color.Yellow
                DFood.Visible = True
            Case Else   '隱藏
                DFood.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Food", "ZZZZZZ")
        '交通
        Select Case FindFieldInf("Traffic")
            Case 0  '顯示
                DTraffic.BackColor = Color.LightGray
                DTraffic.Visible = True
            Case 1  '修改+檢查
                DTraffic.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTrafficRqd", "DTraffic", "異常：需輸入交通")
                DTraffic.Visible = True
            Case 2  '修改
                DTraffic.BackColor = Color.Yellow
                DTraffic.Visible = True
            Case Else   '隱藏
                DTraffic.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Traffic", "ZZZZZZ")
        '預定開始-時
        Select Case FindFieldInf("BStartH")
            Case 0  '顯示
                DBStartH.BackColor = Color.LightGray
                DBStartH.Visible = True
            Case 1  '修改+檢查
                DBStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartHRqd", "DBStartH", "異常：需輸入預定開始-時")
                DBStartH.Visible = True
            Case 2  '修改
                DBStartH.BackColor = Color.Yellow
                DBStartH.Visible = True
            Case Else   '隱藏
                DBStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartH", "ZZZZZZ")
        '預定開始-分
        Select Case FindFieldInf("BStartM")
            Case 0  '顯示
                DBStartM.BackColor = Color.LightGray
                DBStartM.Visible = True
            Case 1  '修改+檢查
                DBStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBStartMRqd", "DBStartM", "異常：需輸入預定開始-分")
                DBStartM.Visible = True
            Case 2  '修改
                DBStartM.BackColor = Color.Yellow
                DBStartM.Visible = True
            Case Else   '隱藏
                DBStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BStartM", "ZZZZZZ")
        '預定終止-時
        Select Case FindFieldInf("BEndH")
            Case 0  '顯示
                DBEndH.BackColor = Color.LightGray
                DBEndH.Visible = True
            Case 1  '修改+檢查
                DBEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndHRqd", "DBEndH", "異常：需輸入預定終止-時")
                DBEndH.Visible = True
            Case 2  '修改
                DBEndH.BackColor = Color.Yellow
                DBEndH.Visible = True
            Case Else   '隱藏
                DBEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndH", "ZZZZZZ")
        '預定終止-分
        Select Case FindFieldInf("BEndM")
            Case 0  '顯示
                DBEndM.BackColor = Color.LightGray
                DBEndM.Visible = True
            Case 1  '修改+檢查
                DBEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBEndMRqd", "DBEndM", "異常：需輸入預定終止-分")
                DBEndM.Visible = True
            Case 2  '修改
                DBEndM.BackColor = Color.Yellow
                DBEndM.Visible = True
            Case Else   '隱藏
                DBEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("BEndM", "ZZZZZZ")
        '預定計算-時
        Select Case FindFieldInf("BH")
            Case 0  '顯示
                DBH.BackColor = Color.LightGray
                DBH.ReadOnly = True
                DBH.Visible = True
            Case 1  '修改+檢查
                DBH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBHRqd", "DBH", "異常：需輸入預定計算-時")
                DBH.Visible = True
            Case 2  '修改
                DBH.BackColor = Color.Yellow
                DBH.Visible = True
            Case Else   '隱藏
                DBH.Visible = False
        End Select
        If pPost = "New" Then DBH.Text = "0"
        '預定計算-分
        Select Case FindFieldInf("BM")
            Case 0  '顯示
                DBM.BackColor = Color.LightGray
                DBM.ReadOnly = True
                DBM.Visible = True
            Case 1  '修改+檢查
                DBM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBMRqd", "DBM", "異常：需輸入預定計算-分")
                DBM.Visible = True
            Case 2  '修改
                DBM.BackColor = Color.Yellow
                DBM.Visible = True
            Case Else   '隱藏
                DBM.Visible = False
        End Select
        If pPost = "New" Then DBM.Text = "0"
        '實際開始-時
        Select Case FindFieldInf("AStartH")
            Case 0  '顯示
                DAStartH.BackColor = Color.LightGray
                DAStartH.Visible = True
            Case 1  '修改+檢查
                DAStartH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartHRqd", "DAStartH", "異常：需輸入實際開始-時")
                DAStartH.Visible = True
            Case 2  '修改
                DAStartH.BackColor = Color.Yellow
                DAStartH.Visible = True
            Case Else   '隱藏
                DAStartH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartH", "ZZZZZZ")
        '實際開始-分
        Select Case FindFieldInf("AStartM")
            Case 0  '顯示
                DAStartM.BackColor = Color.LightGray
                DAStartM.Visible = True
            Case 1  '修改+檢查
                DAStartM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAStartMRqd", "DAStartM", "異常：需輸入實際開始-分")
                DAStartM.Visible = True
            Case 2  '修改
                DAStartM.BackColor = Color.Yellow
                DAStartM.Visible = True
            Case Else   '隱藏
                DAStartM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AStartM", "ZZZZZZ")
        '實際終止-時
        Select Case FindFieldInf("AEndH")
            Case 0  '顯示
                DAEndH.BackColor = Color.LightGray
                DAEndH.Visible = True
            Case 1  '修改+檢查
                DAEndH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndHRqd", "DAEndH", "異常：需輸入實際終止-時")
                DAEndH.Visible = True
            Case 2  '修改
                DAEndH.BackColor = Color.Yellow
                DAEndH.Visible = True
            Case Else   '隱藏
                DAEndH.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndH", "ZZZZZZ")
        '實際終止-分
        Select Case FindFieldInf("AEndM")
            Case 0  '顯示
                DAEndM.BackColor = Color.LightGray
                DAEndM.Visible = True
            Case 1  '修改+檢查
                DAEndM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAEndMRqd", "DAEndM", "異常：需輸入實際終止-分")
                DAEndM.Visible = True
            Case 2  '修改
                DAEndM.BackColor = Color.Yellow
                DAEndM.Visible = True
            Case Else   '隱藏
                DAEndM.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("AEndM", "ZZZZZZ")
        '實際計算-時
        Select Case FindFieldInf("AH")
            Case 0  '顯示
                DAH.BackColor = Color.LightGray
                DAH.ReadOnly = True
                DAH.Visible = True
            Case 1  '修改+檢查
                DAH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAHRqd", "DAH", "異常：需輸入實際計算-時")
                DAH.Visible = True
            Case 2  '修改
                DAH.BackColor = Color.Yellow
                DAH.Visible = True
            Case Else   '隱藏
                DAH.Visible = False
        End Select
        If pPost = "New" Then DAH.Text = "0"
        '實際計算-分
        Select Case FindFieldInf("AM")
            Case 0  '顯示
                DAM.BackColor = Color.LightGray
                DAM.ReadOnly = True
                DAM.Visible = True
            Case 1  '修改+檢查
                DAM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DAMRqd", "DAM", "異常：需輸入實際計算-分")
                DAM.Visible = True
            Case 2  '修改
                DAM.BackColor = Color.Yellow
                DAM.Visible = True
            Case Else   '隱藏
                DAM.Visible = False
        End Select
        If pPost = "New" Then DAM.Text = "0"
        '核定平日2內-時
        Select Case FindFieldInf("FAH1")
            Case 0  '顯示
                DFAH1.BackColor = Color.LightGray
                DFAH1.ReadOnly = True
                DFAH1.Visible = True
            Case 1  '修改+檢查
                DFAH1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAH1Rqd", "DFAH1", "異常：需輸入核定平日2內-時")
                DFAH1.Visible = True
            Case 2  '修改
                DFAH1.BackColor = Color.Yellow
                DFAH1.Visible = True
            Case Else   '隱藏
                DFAH1.Visible = False
        End Select
        If pPost = "New" Then DFAH1.Text = "0"
        '核定平日2內-分
        Select Case FindFieldInf("FAM1")
            Case 0  '顯示
                DFAM1.BackColor = Color.LightGray
                DFAM1.ReadOnly = True
                DFAM1.Visible = True
            Case 1  '修改+檢查
                DFAM1.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAM1Rqd", "DFAM1", "異常：需輸入核定平日2內-分")
                DFAM1.Visible = True
            Case 2  '修改
                DFAM1.BackColor = Color.Yellow
                DFAM1.Visible = True
            Case Else   '隱藏
                DFAM1.Visible = False
        End Select
        If pPost = "New" Then DFAM1.Text = "0"
        '核定平日2外-時
        Select Case FindFieldInf("FAH2")
            Case 0  '顯示
                DFAH2.BackColor = Color.LightGray
                DFAH2.ReadOnly = True
                DFAH2.Visible = True
            Case 1  '修改+檢查
                DFAH2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAH2Rqd", "DFAH2", "異常：需輸入核定平日2外-時")
                DFAH2.Visible = True
            Case 2  '修改
                DFAH2.BackColor = Color.Yellow
                DFAH2.Visible = True
            Case Else   '隱藏
                DFAH2.Visible = False
        End Select
        If pPost = "New" Then DFAH2.Text = "0"
        '核定平日2外-分
        Select Case FindFieldInf("FAM2")
            Case 0  '顯示
                DFAM2.BackColor = Color.LightGray
                DFAM2.ReadOnly = True
                DFAM2.Visible = True
            Case 1  '修改+檢查
                DFAM2.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFAM2Rqd", "DFAM2", "異常：需輸入核定平日2外-分")
                DFAM2.Visible = True
            Case 2  '修改
                DFAM2.BackColor = Color.Yellow
                DFAM2.Visible = True
            Case Else   '隱藏
                DFAM2.Visible = False
        End Select
        If pPost = "New" Then DFAM2.Text = "0"
        '核定假日-時
        Select Case FindFieldInf("FBH")
            Case 0  '顯示
                DFBH.BackColor = Color.LightGray
                DFBH.ReadOnly = True
                DFBH.Visible = True
            Case 1  '修改+檢查
                DFBH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFBHRqd", "DFBH", "異常：需輸入核定假日-時")
                DFBH.Visible = True
            Case 2  '修改
                DFBH.BackColor = Color.Yellow
                DFBH.Visible = True
            Case Else   '隱藏
                DFBH.Visible = False
        End Select
        If pPost = "New" Then DFBH.Text = "0"
        '核定假日-分
        Select Case FindFieldInf("FBM")
            Case 0  '顯示
                DFBM.BackColor = Color.LightGray
                DFBM.ReadOnly = True
                DFBM.Visible = True
            Case 1  '修改+檢查
                DFBM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFBMRqd", "DFBM", "異常：需輸入核定假日-分")
                DFBM.Visible = True
            Case 2  '修改
                DFBM.BackColor = Color.Yellow
                DFBM.Visible = True
            Case Else   '隱藏
                DFBM.Visible = False
        End Select
        If pPost = "New" Then DFBM.Text = "0"
        '核定國定假日-時
        Select Case FindFieldInf("FCH")
            Case 0  '顯示
                DFCH.BackColor = Color.LightGray
                DFCH.ReadOnly = True
                DFCH.Visible = True
            Case 1  '修改+檢查
                DFCH.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCHRqd", "DFCH", "異常：需輸入核定國定假日-時")
                DFCH.Visible = True
            Case 2  '修改
                DFCH.BackColor = Color.Yellow
                DFCH.Visible = True
            Case Else   '隱藏
                DFCH.Visible = False
        End Select
        If pPost = "New" Then DFCH.Text = "0"
        '核定國定假日-分
        Select Case FindFieldInf("FCM")
            Case 0  '顯示
                DFCM.BackColor = Color.LightGray
                DFCM.ReadOnly = True
                DFCM.Visible = True
            Case 1  '修改+檢查
                DFCM.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFCMRqd", "DFCM", "異常：需輸入核定國定假日-分")
                DFCM.Visible = True
            Case 2  '修改
                DFCM.BackColor = Color.Yellow
                DFCM.Visible = True
            Case Else   '隱藏
                DFCM.Visible = False
        End Select
        If pPost = "New" Then DFCM.Text = "0"
        '加班理由
        Select Case FindFieldInf("FReason")
            Case 0  '顯示
                DFReason.BackColor = Color.LightGray
                DFReason.ReadOnly = True
                DFReason.Visible = True
            Case 1  '修改+檢查
                DFReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DFReasonRqd", "DFReason", "異常：需輸入加班理由")
                DFReason.Visible = True
            Case 2  '修改
                DFReason.BackColor = Color.Yellow
                DFReason.Visible = True
            Case Else   '隱藏
                DFReason.Visible = False
        End Select
        If pPost = "New" Then DFReason.Text = ""
        '
        ' 超過46H
        ' Step=1 起單
        If wStep = 1 Then
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' 平日已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' 平日未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' 總已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' 總未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "總時數：[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "平日時數：[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-淺灰色
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-橘色
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-粉紅色
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' 白色
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
        End If
        '
        OleDbConnection1.Close()
    End Sub

    '*****************************************************************
    '**(ShowRequiredFieldValidator)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(SetFieldData)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData(ByVal pFieldName As String, ByVal pName As String)
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim idx As Integer
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        idx = FindFieldInf(pFieldName)

        '姓名
        If pFieldName = "Name" Then
            DName.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DName.Items.Add(ListItem1)
                End If
            Else
                '登入者
                Dim ListItem1 As New ListItem
                ListItem1.Text = wUserName
                ListItem1.Value = Request.QueryString("pUserID")
                ListItem1.Selected = True
                DName.Items.Add(ListItem1)
                '全表單代理
                SQL = "Select UserName, UserID From M_Agent  "
                SQL = SQL + "Where Active = '1' "
                SQL = SQL + "  And AllForm = '0' "
                SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Agent")
                DBTable1 = DBDataSet1.Tables("M_Agent")
                If DBTable1.Rows.Count <= 0 Then
                    '單一表單代理
                    DBDataSet1.Clear()
                    SQL = "Select UserName, UserID From M_Agent  "
                    SQL = SQL + "Where Active = '1' "
                    SQL = SQL + "  And AllForm = '1' "
                    SQL = SQL + "  And AgentID = '" + Request.QueryString("pUserID") + "' "
                    SQL = SQL + "  And StartDate <= '" + NowDateTime + "' "
                    SQL = SQL + "  And EndDate >= '" + NowDateTime + "' "
                    SQL = SQL + "  And FormNo = '001001' "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2.Fill(DBDataSet1, "M_Agent")
                    DBTable1 = DBDataSet1.Tables("M_Agent")
                    If DBTable1.Rows.Count <= 0 Then
                    Else
                        For i = 0 To DBTable1.Rows.Count - 1
                            Dim ListItem2 As New ListItem
                            ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                            ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                            DName.Items.Add(ListItem2)
                        Next
                    End If
                Else
                    For i = 0 To DBTable1.Rows.Count - 1
                        Dim ListItem2 As New ListItem
                        ListItem2.Text = DBTable1.Rows(i).Item("UserName")
                        ListItem2.Value = DBTable1.Rows(i).Item("UserID")
                        DName.Items.Add(ListItem2)
                    Next
                End If
            End If
        End If
        '調休
        If pFieldName = "CVacation" Then
            DCVacation.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCVacation.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='CVACATION' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCVacation.Items.Add(ListItem1)
                Next
            End If
        End If
        '伙食
        If pFieldName = "Food" Then
            DFood.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DFood.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='FOOD' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DFood.Items.Add(ListItem1)
                Next
            End If
        End If
        '交通
        If pFieldName = "Traffic" Then
            DTraffic.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTraffic.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='TRAFFIC' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTraffic.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定開始-時
        If pFieldName = "BStartH" Then
            DBStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定開始-分
        If pFieldName = "BStartM" Then
            DBStartM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBStartM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBStartM.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定終止-時
        If pFieldName = "BEndH" Then
            DBEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '預定終止-分
        If pFieldName = "BEndM" Then
            DBEndM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DBEndM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DBEndM.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際開始-時
        If pFieldName = "AStartH" Then
            DAStartH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAStartH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartH.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際開始-分
        If pFieldName = "AStartM" Then
            DAStartM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAStartM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAStartM.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際終止-時
        If pFieldName = "AEndH" Then
            DAEndH.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAEndH.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='HOUR' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndH.Items.Add(ListItem1)
                Next
            End If
        End If
        '實際終止-分
        If pFieldName = "AEndM" Then
            DAEndM.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DAEndM.Items.Add(ListItem1)
                End If
            Else
                If pName <> "ZZZZZZ" Then
                    If CInt(pName) < 10 Then pName = "0" + pName
                End If
                SQL = "Select * From M_Referp Where Cat='1001' and DKey='MIN' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DAEndM.Items.Add(ListItem1)
                Next
            End If
        End If

        '延遲理由代碼
        If pFieldName = "ReasonCode" Then
            DReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DReasonCode.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DKey")
                    ListItem1.Value = DBTable1.Rows(i).Item("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                DReason.Text = DBTable1.Rows(i).Item("Data")
            Next
        End If
        'DB連結關閉
        OleDbConnection1.Close()
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
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BSAVE_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSAVE.ServerClick
        If Request.Cookies("RunBSAVE").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
                Dim Message As String = ""

                'Check理由
                If ErrCode = 0 Then
                    If DReasonCode.Visible = True Then
                        If DReasonCode.SelectedValue = "99" Then
                            If DReasonDesc.Text = "" Then ErrCode = 9210
                        End If
                    End If
                End If
                '儲存資料
                If ErrCode = 0 Then
                    ModifyData("SAVE", "0")           '更新表單資料 Sts=0(未結)
                    ModifyTranData("SAVE", "0")       '更新交易資料
                Else
                    If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '上傳檔案ErrCode=0

                If ErrCode = 0 Then
                    Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                        "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                    Response.Redirect(URL)
                End If
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.ServerClick
        If Request.Cookies("RunBOK").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("OK", 0, "1")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG1.ServerClick
        If Request.Cookies("RunBNG1").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG1", 1, "2")
            End If
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG2.ServerClick
        If Request.Cookies("RunBNG2").Value = True Then
            If InputCheck() = 0 Then
                DisabledButton()   '停止Button運作
                FlowControl("NG2", 2, "3")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(FlowControl)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub FlowControl(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status

        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If
        'Check薪資所屬年月變更
        If ErrCode = 0 Then
            If wStep = 50 Then
                If DSalaryYM.Text <> DSalaryYM1.Text Then
                    ErrCode = 9041
                End If
            End If
        End If
        'Check預定計算結果
        If ErrCode = 0 Then
            If wStep = 1 Or wStep = 500 Then
                If DBH.Text = "0" And DBM.Text = "0" Then ErrCode = 9050
            End If
        End If
        'Check實際計算結果
        If ErrCode = 0 Then
            If wStep = 30 Then
                If pFun = "OK" Then
                    If DAH.Text = "0" And DAM.Text = "0" Then ErrCode = 9050
                End If
            End If
        End If
        'Check調休
        If ErrCode = 0 Then
            If DCVacation.SelectedValue = "2.要" Then
                ' 預定調休
                If wStep = 1 Then
                    If CInt(DBH.Text) < 4 Then
                        ErrCode = 9051
                    End If
                End If
                If wStep = 500 Then
                    If CInt(DBH.Text) < 4 Then
                        ErrCode = 9051
                    End If
                End If
                ' 實際調休
                If wStep = 30 Then
                    If pFun = "OK" Then
                        If CInt(DAH.Text) < 4 Then
                            ErrCode = 9051
                        End If
                    End If
                End If
            End If
        End If
        'Check加班時數
        If ErrCode = 0 Then
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            Dim DBDataSet1 As New DataSet
            Dim SQL As String
            '超過46H/92H
            Dim wHours46 As Boolean = False '46H-Flag
            Dim wHours92 As Boolean = False '92H-Flag
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            Dim xHours As Integer = 0
            Dim yHours As Integer = 0
            OleDbConnection1.Open()
            '取得申請者-ID
            DAgentID.Text = ""
            SQL = "Select UserID From M_Users "
            SQL = SQL & " Where UserID = '" & DName.SelectedValue & "'"
            SQL = SQL & "   And Active = '1' "
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
                ' 被代理人ID
                If UCase(DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")) <> UCase(Request.QueryString("pUserID")) Then
                    DAgentID.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")
                End If
            End If
            ' 平日已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' 平日未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' 總已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' 總未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            xHours = CDbl(DFAH1.Text) * 60 + CDbl(DFAH2.Text) * 60 + CDbl(DFBH.Text) * 60 + CDbl(DFCH.Text) * 60 + CDbl(DFAM1.Text) + CDbl(DFAM2.Text) + CDbl(DFBM.Text) + CDbl(DFCM.Text)
            yHours = CDbl(DFAH1.Text) * 60 + CDbl(DFAH2.Text) * 60 + CDbl(DFAM1.Text) + CDbl(DFAM2.Text)
            If aHours + bHours + yHours > 2760 Then '46H
                wHours46 = True
            End If
            If cHours + dHours + xHours > 5520 Then   '92H
                wHours92 = True
            End If
            OleDbConnection1.Close()
            'Check 46, 92H
            If wStep = 500 And pFun = "NG2" Then
                ' [500]重新修正 and 按取消 時 不Check    
            Else
                If wHours46 = True Then
                    If wStep = 1 Or wStep = 500 Then
                        If UCase(Request.QueryString("pUserID")) <> "GAS007" Then
                            ErrCode = 9052
                        End If
                    End If
                End If
                If ErrCode = 0 Then
                    If wHours92 = True Then
                        If wStep = 1 Or wStep = 500 Then
                            If UCase(Request.QueryString("pUserID")) <> "GAS007" Then
                                ErrCode = 9053
                            End If
                        End If
                    End If
                End If
            End If
        End If

        ' 2016/12/23 勞基法變更對應
        ' @@@@
        ' 加班單位()
        ' 社休假日: 4H, 8H, 12H
        ' 例假日,國假日,國假日(ykk): 8H, 12H

        ' 2018/3/1 勞基法變更對應
        ' 因03/01勞基法修法, 煩請修改加班單設定,謝謝!
        '   -->休息日加班的工時，改為核實計算(但>12H時仍請出現確認訊息,暫不開放申請)
        If ErrCode = 0 Then
            ' 2018/3/1 勞基法變更對應
            If DDateType.Text = "社休假日" Or DDateType.Text = "假日" Or DDateType.Text = "國定假日" Then
                If CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) > 12 Then
                    ErrCode = 9059
                Else
                    If CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) > 0 Then
                        ErrCode = 9059
                    End If
                End If
            End If

            ' 2016/12/23 勞基法變更對應
            'If CDate(DOverTimeDate.Text) >= CDate("2016/12/23") Then
            '    If DDateType.Text = "社休假日" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 4 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9056
            '        End If
            '    End If
            '    If DDateType.Text = "假日" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9057
            '        End If
            '    End If
            '    If DDateType.Text = "國定假日" Then
            '        If (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 8 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Or _
            '           (CInt(DFAH1.Text) + CInt(DFAH2.Text) + CInt(DFBH.Text) + CInt(DFCH.Text) = 12 And CInt(DFAM1.Text) + CInt(DFAM2.Text) + CInt(DFBM.Text) + CInt(DFCM.Text) = 0) Then
            '        Else
            '            ErrCode = 9058
            '        End If
            '    End If
            'End If
        End If
        '
        '檢查委託書No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then

                ErrCode = oCommon.CommissionNo("001001", wFormSno, wStep, DNo.Text)     '表單號碼, 表單流水號, 工程, 委託書No

                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '難易度

            While Run = True
                Run = False     '執行Flag=不執行
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
                '--其他參數設定---------
                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                        '取得表單流水號
                        RtnCode = oCommon.GetFormSeqNo(wFormNo, NewFormSno) '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '申請流程資料建置
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                            '
                            oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                            '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                            'Modify-End

                            '設定委託No
                            DNo.Text = SetNo(NewFormSno)
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            ModifyTranData(pFun, pSts)

                            '流程資料結束
                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDepo, Request.QueryString("pUserID"), pRunNextStep)
                            '
                            RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)
                            '表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽)
                            'Modify-End

                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    Dim wAllocateID As String = ""

                    If DAgentID.Text <> "" Then
                        RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, DAgentID.Text, wAllocateID, MultiJob, _
                                                      pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    Else
                        RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                      pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    End If
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                    If ErrCode = 0 Then pAction = 0
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        wNextGate = ""
                        For i = 1 To pCount
                            '取得下一關人員(訊息時使用)
                            If wNextGate = "" Then
                                wNextGate = pNextGate(i)
                            Else
                                wNextGate = "," & pNextGate(i)
                            End If

                            'Modify-Start by Joy 2009/11/20(2010行事曆對應)
                            'oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            'oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wDepo)
                            'RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wDepo, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '
                            '取得核定者-群組行是曆
                            wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                            '取得工程負荷最後日時
                            oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)
                            '核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                            '取得預定開始,完成日程計算
                            oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)
                            '表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆

                            '建置流程資料
                            RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)
                            '表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
                            'Modify-End

                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDepo, Request.QueryString("pUserID"), wApplyID)
                        '
                        RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者
                        'Modify-End

                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        'Modify-Start by 2009/11/20(2010行事曆對應)
                        'RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
                        '
                        RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                        '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                        'Modify-End
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                        End If
                    Else
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                        Else
                            ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                        End If
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount

                            oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

                        Next i
                    Else

                        oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別

                    End If

                    If pFlowType <> 3 Then
                        MultiJob = 0
                    Else
                        MultiJob = 1
                    End If

                    If (pRunNextStep = 1) And (pFlowType = 0 Or pFlowType = 3) Then
                        Run = True
                        RepeatRun = True
                        pAction = 0
                    End If

                    wStep = pNextStep     '下一工程關卡號碼
                    wFormSno = NewFormSno '下一工程表單流水號
                Else
                    EnabledButton()   '起動Button運作
                    If ErrCode = 9110 Then Message = "取得表單流水號計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9120 Then Message = "流程資料更新異常,請確認或連絡系統人員!"
                    If ErrCode = 9130 Then Message = "下工程計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9131 Then Message = "無下工程管理人,請確認或連絡系統人員!"
                    If ErrCode = 9140 Then Message = "工程預定開始及完成日計算異常,請確認或連絡系統人員!"
                    If ErrCode = 9150 Then Message = "下一工程資料建置異常,請確認或連絡系統人員!"
                    If ErrCode = 9160 Then Message = "工程結束資料建置異常(999),請確認或連絡系統人員!"
                    Response.Write(YKK.ShowMessage(Message))
                End If      '儲存表單ErrCode=0
            End While       '重覆執行

            If ErrCode = 0 Then
                '--郵件傳送---------
                oCommon.SendMail()

                Dim URL As String = "http://10.245.1.6/WorkFlowSub/MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9041 Then Message = "薪資所屬年月不可變更,請確認!"
            If ErrCode = 9050 Then Message = "預定或實際加班時間需填寫,請確認!"

            If ErrCode = 9051 Then Message = "加班時間需超過 [4 H] 以上才可調休,請確認!"
            If ErrCode = 9052 Then Message = "超過平日[46H]法定警界線, 無法填寫加班申請單,請確認!"
            If ErrCode = 9053 Then Message = "超過單月[92H]法定警界線, 無法填寫加班申請單,請確認!"

            If ErrCode = 9056 Then Message = "休假日(社休假日)加班時數需符合 [4H]/[8H]/[12H],請確認!"
            If ErrCode = 9057 Then Message = "假日(例假日)加班時數需符合 [8H]/[12H],請確認!"
            If ErrCode = 9058 Then Message = "國定假日班時數需符合 [8H]/[12H],請確認!"

            If ErrCode = 9059 Then Message = "休假日加班時數不可超過 [12H],請確認!"

            If ErrCode = 9060 Then Message = "委託書No.重覆,請確認委託書No.!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetDataStatus)
    '**     取得表單進度狀態
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Flow")
        If DBDataSet1.Tables("M_Flow").Rows.Count > 0 Then
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                wOKSts = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKSts") + 1
            End If
        End If
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("OverTimeFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_OverTimeSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, No, "            '1~5
        SQl = SQl + "Date, OverTimeDate, DateType, Name, EmpID, "          '6~10
        SQl = SQl + "JobTitle, JobCode, Division, DivisionCode, CVacation, Food, Traffic, "    '11~17
        SQl = SQl + "BStartH, BStartM, BEndH, BEndM, BH, BM, "             '17~22
        SQl = SQl + "AStartH, AStartM, AEndH, AEndM, AH, AM, "             '23~28
        SQl = SQl + "FAH1, FAM1, FAH2, FAM2, "                             '29~32
        SQl = SQl + "FBH, FBM, FCH, FCM, "                                 '33~37
        ' PAYROLL FIELD-START
        SQl = SQl + "FPRAH1, FPRAM1, FPRAH2, FPRAM2, "                     '
        SQl = SQl + "FPRBH, FPRBM, FPRCH, FPRCM, "                         '

        ' JOY BY 2017/6/19
        SQl = SQl + "FPRTAX15H, FPRTAX15M, FPRTAX167H, FPRTAX167M, "
        SQl = SQl + "FPRTAX20H, FPRTAX20M, FPRTAX267H, FPRTAX267M, "

        ' END
        SQl = SQl + "FReason, SalaryYM, DepoName, DepoCode, "              '37~40
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '41~44
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        '1~5
        SQl = SQl + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "              '結案日
        SQl = SQl + " '001001', "                           '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '表單流水號
        SQl = SQl + " '" + DNo.Text + "', "                 'NO
        '6~10
        SQl = SQl + " '" + DDate.Text + "', "                   '申請日期
        SQl = SQl + " '" + DOverTimeDate.Text + "', "           '加班日期
        SQl = SQl + " N'" + DDateType.Text + "', "              '加班別
        SQl = SQl + " N'" + DName.SelectedItem.Text + "', "     '姓名
        SQl = SQl + " '" + DEmpID.Text + "', "                  'EMP-ID
        '11~15
        SQl = SQl + " N'" + DJobTitle.Text + "', "              '職稱
        SQl = SQl + " '" + DJobCode.Text + "', "                '職稱代碼
        SQl = SQl + " N'" + DDivision.Text + "', "              '部門
        SQl = SQl + " '" + DDivisionCode.Text + "', "           '部門代碼
        SQl = SQl + " N'" + DCVacation.SelectedValue + "', "    '調休
        SQl = SQl + " N'" + DFood.SelectedValue + "', "         '伙食
        SQl = SQl + " N'" + DTraffic.SelectedValue + "', "      '交通
        '16~21
        SQl = SQl + " '" + CStr(CInt(DBStartH.SelectedValue)) + "', "  '預定開始-時
        SQl = SQl + " '" + CStr(CInt(DBStartM.SelectedValue)) + "', "  '預定開始-分
        SQl = SQl + " '" + CStr(CInt(DBEndH.SelectedValue)) + "', "    '預定終止-時
        SQl = SQl + " '" + CStr(CInt(DBEndM.SelectedValue)) + "', "    '預定終止-分
        SQl = SQl + " '" + CStr(CInt(DBH.Text)) + "', "                '預定計算-時
        SQl = SQl + " '" + CStr(CInt(DBM.Text)) + "', "                '預定計算-分
        '22~27
        SQl = SQl + " '" + CStr(CInt(DAStartH.SelectedValue)) + "', "  '實際開始-時
        SQl = SQl + " '" + CStr(CInt(DAStartM.SelectedValue)) + "', "  '實際開始-分
        SQl = SQl + " '" + CStr(CInt(DAEndH.SelectedValue)) + "', "    '實際終止-時
        SQl = SQl + " '" + CStr(CInt(DAEndM.SelectedValue)) + "', "    '實際終止-分
        SQl = SQl + " '" + CStr(CInt(DAH.Text)) + "', "                '實際計算-時
        SQl = SQl + " '" + CStr(CInt(DAM.Text)) + "', "                '實際計算-分
        '28~31
        SQl = SQl + " '" + CStr(CInt(DFAH1.Text)) + "', "              '核定平日2內-時
        SQl = SQl + " '" + CStr(CInt(DFAM1.Text)) + "', "              '核定平日2內-分
        SQl = SQl + " '" + CStr(CInt(DFAH2.Text)) + "', "              '核定平日2外-時
        SQl = SQl + " '" + CStr(CInt(DFAM2.Text)) + "', "              '核定平日2外-分
        '32~36
        SQl = SQl + " '" + CStr(CInt(DFBH.Text)) + "', "               '核定假日-時
        SQl = SQl + " '" + CStr(CInt(DFBM.Text)) + "', "               '核定假日-分
        SQl = SQl + " '" + CStr(CInt(DFCH.Text)) + "', "               '核定國定假日-時
        SQl = SQl + " '" + CStr(CInt(DFCM.Text)) + "', "               '核定國定假日-分
        ' PAYROLL FIELD-START
        SQl = SQl + " '" + CStr(CInt(DFPRAH1.Text)) + "', "            'PAYROLL核定平日2內-時
        SQl = SQl + " '" + CStr(CInt(DFPRAM1.Text)) + "', "            'PAYROLL核定平日2內-分
        SQl = SQl + " '" + CStr(CInt(DFPRAH2.Text)) + "', "            'PAYROLL核定平日2外-時
        SQl = SQl + " '" + CStr(CInt(DFPRAM2.Text)) + "', "            'PAYROLL核定平日2外-分

        SQl = SQl + " '" + CStr(CInt(DFPRBH.Text)) + "', "             'PAYROLL核定假日-時
        SQl = SQl + " '" + CStr(CInt(DFPRBM.Text)) + "', "             'PAYROLL核定假日-分
        SQl = SQl + " '" + CStr(CInt(DFPRCH.Text)) + "', "             'PAYROLL核定國定假日-時
        SQl = SQl + " '" + CStr(CInt(DFPRCM.Text)) + "', "             'PAYROLL核定國定假日-分

        ' JOY BY 2017/6/19
        SQl = SQl + " '" + CStr(CInt(DFPRTAX15H.Text)) + "', "         'PAYROLL核定應稅 1.5
        SQl = SQl + " '" + CStr(CInt(DFPRTAX15M.Text)) + "', "         '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX167H.Text)) + "', "        'PAYROLL核定應稅 1.67
        SQl = SQl + " '" + CStr(CInt(DFPRTAX167M.Text)) + "', "        '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX20H.Text)) + "', "         'PAYROLL核定應稅 2.0
        SQl = SQl + " '" + CStr(CInt(DFPRTAX20M.Text)) + "', "         '
        SQl = SQl + " '" + CStr(CInt(DFPRTAX267H.Text)) + "', "        'PAYROLL核定應稅 2.67
        SQl = SQl + " '" + CStr(CInt(DFPRTAX267M.Text)) + "', "        '

        ' END
        SQl = SQl + " N'" + YKK.ReplaceString(DFReason.Text) + "', "   '加班理由
        '37~39
        SQl = SQl + " '" + DSalaryYM.Text + "', "                      '所屬年月
        SQl = SQl + " N'" + DDepoName.Text + "', "                     '公司別
        SQl = SQl + " '" + DDepoCode.Text + "', "                      '公司別Code
        '--------------------------------------------
        SQl = SQl + " '" + Request.QueryString("pUserID") + "', "     '作成者
        SQl = SQl + " '" + NowDateTime + "', "       '作成時間
        SQl = SQl + " '" + "" + "', "                       '修改者
        SQl = SQl + " '" + NowDateTime + "' "       '修改時間
        SQl = SQl + " ) "
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("OverTimeFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_OverTimeSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " OverTimeDate = '" & DOverTimeDate.Text & "',"
        SQl = SQl + " DateType = N'" & DDateType.Text & "',"
        SQl = SQl + " SalaryYM = '" & DSalaryYM.Text & "',"
        SQl = SQl + " Name = N'" & DName.SelectedItem.Text & "',"
        SQl = SQl + " EmpID = '" & DEmpID.Text & "',"
        SQl = SQl + " JobTitle = N'" & DJobTitle.Text & "',"
        SQl = SQl + " JobCode = '" & DJobCode.Text & "',"
        SQl = SQl + " DepoName = N'" & DDepoName.Text & "',"
        SQl = SQl + " DepoCode = '" & DDepoCode.Text & "',"
        SQl = SQl + " Division = N'" & DDivision.Text & "',"
        SQl = SQl + " DivisionCode = '" & DDivisionCode.Text & "',"
        SQl = SQl + " CVacation = N'" & DCVacation.SelectedValue & "',"
        SQl = SQl + " Food = N'" & DFood.SelectedValue & "',"
        SQl = SQl + " Traffic = N'" & DTraffic.SelectedValue & "',"

        SQl = SQl + " BStartH = '" & CStr(CInt(DBStartH.SelectedValue)) & "',"
        SQl = SQl + " BStartM = '" & CStr(CInt(DBStartM.SelectedValue)) & "',"
        SQl = SQl + " BEndH = '" & CStr(CInt(DBEndH.SelectedValue)) & "',"
        SQl = SQl + " BEndM = '" & CStr(CInt(DBEndM.SelectedValue)) & "',"
        SQl = SQl + " BH = '" & CStr(CInt(DBH.Text)) & "',"
        SQl = SQl + " BM = '" & CStr(CInt(DBM.Text)) & "',"

        SQl = SQl + " AStartH = '" & CStr(CInt(DAStartH.SelectedValue)) & "',"
        SQl = SQl + " AStartM = '" & CStr(CInt(DAStartM.SelectedValue)) & "',"
        SQl = SQl + " AEndH = '" & CStr(CInt(DAEndH.SelectedValue)) & "',"
        SQl = SQl + " AEndM = '" & CStr(CInt(DAEndM.SelectedValue)) & "',"
        SQl = SQl + " AH = '" & CStr(CInt(DAH.Text)) & "',"
        SQl = SQl + " AM = '" & CStr(CInt(DAM.Text)) & "',"

        SQl = SQl + " FAH1 = '" & CStr(CInt(DFAH1.Text)) & "',"
        SQl = SQl + " FAM1 = '" & CStr(CInt(DFAM1.Text)) & "',"
        SQl = SQl + " FAH2 = '" & CStr(CInt(DFAH2.Text)) & "',"
        SQl = SQl + " FAM2 = '" & CStr(CInt(DFAM2.Text)) & "',"
        SQl = SQl + " FBH = '" & CStr(CInt(DFBH.Text)) & "',"
        SQl = SQl + " FBM = '" & CStr(CInt(DFBM.Text)) & "',"
        SQl = SQl + " FCH = '" & CStr(CInt(DFCH.Text)) & "',"
        SQl = SQl + " FCM = '" & CStr(CInt(DFCM.Text)) & "',"

        ' PAYROLL FIELD-START
        SQl = SQl + " FPRAH1 = '" & CStr(CInt(DFPRAH1.Text)) & "',"
        SQl = SQl + " FPRAM1 = '" & CStr(CInt(DFPRAM1.Text)) & "',"
        SQl = SQl + " FPRAH2 = '" & CStr(CInt(DFPRAH2.Text)) & "',"
        SQl = SQl + " FPRAM2 = '" & CStr(CInt(DFPRAM2.Text)) & "',"
        SQl = SQl + " FPRBH = '" & CStr(CInt(DFPRBH.Text)) & "',"
        SQl = SQl + " FPRBM = '" & CStr(CInt(DFPRBM.Text)) & "',"
        SQl = SQl + " FPRCH = '" & CStr(CInt(DFPRCH.Text)) & "',"
        SQl = SQl + " FPRCM = '" & CStr(CInt(DFPRCM.Text)) & "',"

        ' JOY BY 2017/6/19
        SQl = SQl + " FPRTAX15H = '" & CStr(CInt(DFPRTAX15H.Text)) & "',"
        SQl = SQl + " FPRTAX15M = '" & CStr(CInt(DFPRTAX15M.Text)) & "',"
        SQl = SQl + " FPRTAX167H = '" & CStr(CInt(DFPRTAX167H.Text)) & "',"
        SQl = SQl + " FPRTAX167M = '" & CStr(CInt(DFPRTAX167M.Text)) & "',"
        SQl = SQl + " FPRTAX20H = '" & CStr(CInt(DFPRTAX20H.Text)) & "',"
        SQl = SQl + " FPRTAX20M = '" & CStr(CInt(DFPRTAX20M.Text)) & "',"
        SQl = SQl + " FPRTAX267H = '" & CStr(CInt(DFPRTAX267H.Text)) & "',"
        SQl = SQl + " FPRTAX267M = '" & CStr(CInt(DFPRTAX267M.Text)) & "',"

        ' END
        SQl = SQl + " FReason = N'" & YKK.ReplaceString(DFReason.Text) & "',"
        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Value & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Value & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Value & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step    =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
            SQl = SQl + "   And Active =  '1' "
        Else
            SQl = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & YKK.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        OleDBCommand1.Connection = OleDbConnection1
        OleDBCommand1.CommandText = SQl
        OleDbConnection1.Open()
        OleDBCommand1.ExecuteNonQuery()
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        Dim DBDataSet1 As New DataSet

        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.Open()

        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "T_CommissionNo")

        If DBDataSet1.Tables("T_CommissionNo").Rows.Count <= 0 Then
            If DNo.Text <> "" Then
                SQl = "Insert into T_CommissionNo "
                SQl = SQl + "( "
                SQl = SQl + "FormNo, FormSno, No, MapNo, CreateUser, CreateTime "    '1~5
                SQl = SQl + ")  "
                SQl = SQl + "VALUES( "
                SQl = SQl + " '" + pFormNo + "', "
                SQl = SQl + " '" + CStr(pFormSno) + "', "
                SQl = SQl + " '" + DNo.Text + "', "
                SQl = SQl + " '" + "" + "', "
                SQl = SQl + " '" + Request.QueryString("pUserID") + "', "
                SQl = SQl + " '" + NowDateTime + "' "
                SQl = SQl + " ) "
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDBCommand1.ExecuteNonQuery()
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> DBDataSet1.Tables("T_CommissionNo").Rows(0).Item("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    OleDBCommand1.Connection = OleDbConnection1
                    OleDBCommand1.CommandText = SQl
                    OleDBCommand1.ExecuteNonQuery()
                End If
            End If
        End If

        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     姓名變更 
    '**
    '*****************************************************************
    Private Sub DName_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DName.SelectedIndexChanged
        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '取得申請者資訊
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName, UserID From M_Users "
        SQL = SQL & " Where UserID = '" & DName.SelectedValue & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            DDepoCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            DDepoName.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            DEmpID.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            DJobCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            DDivision.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            DDivisionCode.Text = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        '
        ' 超過46H
        ' Step=1 起單
        If wStep = 1 Then
            Dim aHours As Integer = 0
            Dim bHours As Integer = 0
            Dim cHours As Integer = 0
            Dim dHours As Integer = 0
            ' 平日已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "OverTime1")
            If DBDataSet1.Tables("OverTime1").Rows(0).Item("Total") > 0 Then
                aHours = DBDataSet1.Tables("OverTime1").Rows(0).Item("Total")
            End If
            ' 平日未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FAM1) + Sum(FAM2), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter31 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter31.Fill(DBDataSet1, "OverTime2")
            If DBDataSet1.Tables("OverTime2").Rows(0).Item("Total") > 0 Then
                bHours = DBDataSet1.Tables("OverTime2").Rows(0).Item("Total")
            End If
            ' 總已核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '1' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter32 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter32.Fill(DBDataSet1, "OverTime3")
            If DBDataSet1.Tables("OverTime3").Rows(0).Item("Total") > 0 Then
                cHours = DBDataSet1.Tables("OverTime3").Rows(0).Item("Total")
            End If
            ' 總未核
            DBDataSet1.Clear()
            SQL = "Select IsNull(Sum(FAH1*60) + Sum(FAH2*60) + Sum(FBH*60) + Sum(FCH*60) + Sum(FAM1) + Sum(FAM2) + Sum(FBM) + Sum(FCM), 0) As Total From F_OverTimeSheet "
            SQL = SQL + "Where Sts = '0' "
            SQL = SQL + "  and EmpID = '" + DEmpID.Text + "' "
            SQL = SQL + "  and SalaryYM = '" + DSalaryYM.Text + "' "
            Dim DBAdapter33 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter33.Fill(DBDataSet1, "OverTime4")
            If DBDataSet1.Tables("OverTime4").Rows(0).Item("Total") > 0 Then
                dHours = DBDataSet1.Tables("OverTime4").Rows(0).Item("Total")
            End If
            '
            D46Hours.Text = "總時數：[" + CStr((cHours + dHours) / 60) + "H (" + CStr(cHours / 60) + " / " + CStr(dHours / 60) + ")]"
            D46Hours.Text = D46Hours.Text + " "
            D46Hours.Text = D46Hours.Text + "平日時數：[" + CStr((aHours + bHours) / 60) + "H (" + CStr(aHours / 60) + " / " + CStr(bHours / 60) + ")]"
            '
            ' 92H-淺灰色
            If cHours + dHours >= 5520 Then
                D46Hours.BackColor = Color.Gainsboro
            Else
                ' 46-橘色
                If aHours + bHours >= 2760 Then
                    D46Hours.BackColor = Color.Orange
                Else
                    ' 36-粉紅色
                    If aHours + bHours >= 2160 Then
                        D46Hours.BackColor = Color.LightPink
                    Else
                        ' 白色
                        D46Hours.BackColor = Color.White
                    End If
                End If
            End If
        End If
        OleDbConnection1.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     選取預定終止-時 
    '**
    '*****************************************************************
    Private Sub DBEndH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBEndH.SelectedIndexChanged
        Dim StartTime As String = DBStartH.SelectedValue + DBStartM.SelectedValue
        Dim EndTime As String = DBEndH.SelectedValue + DBEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DBStartH.SelectedValue, DBStartM.SelectedValue, DBEndH.SelectedValue, DBEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DBH.Text = CStr(Hour)
            DBM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DBH.Text = "0"
            DBM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     選取預定終止-分 
    '**
    '*****************************************************************
    Private Sub DBEndM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DBEndM.SelectedIndexChanged
        Dim StartTime As String = DBStartH.SelectedValue + DBStartM.SelectedValue
        Dim EndTime As String = DBEndH.SelectedValue + DBEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DBStartH.SelectedValue, DBStartM.SelectedValue, DBEndH.SelectedValue, DBEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DBH.Text = CStr(Hour)
            DBM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DBH.Text = "0"
            DBM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     選取實際終止-時 
    '**
    '*****************************************************************
    Private Sub DAEndH_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAEndH.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     選取實際終止-分 
    '**
    '*****************************************************************
    Private Sub DAEndM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAEndM.SelectedIndexChanged
        Dim StartTime As String = DAStartH.SelectedValue + DAStartM.SelectedValue
        Dim EndTime As String = DAEndH.SelectedValue + DAEndM.SelectedValue
        Dim OverTime As Integer = 0
        Dim Hour, Minute As Integer

        If EndTime > StartTime Then
            OverTime = CalOverTime(DAStartH.SelectedValue, DAStartM.SelectedValue, DAEndH.SelectedValue, DAEndM.SelectedValue)
            Hour = Int(OverTime / 60)
            Minute = OverTime - Hour * 60
            If Minute >= 30 Then
                Minute = 30
            Else
                Minute = 0
            End If
            DAH.Text = CStr(Hour)
            DAM.Text = CStr(Minute)
            '
            CalApproveOverTime(Hour, Minute)
        Else
            DAH.Text = "0"
            DAM.Text = "0"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     計算核定時間 
    '**
    '*****************************************************************
    Sub CalApproveOverTime(ByVal sHour As Integer, ByVal sMinute As Integer)
        Dim xHour As Integer = sHour
        Dim xMinute As Integer = sMinute
        '考勤計算 -----------------------------------------
        '  清初值
        DFAH1.Text = "0"
        DFAM1.Text = "0"
        DFAH2.Text = "0"
        DFAM2.Text = "0"
        DFBH.Text = "0"
        DFBM.Text = "0"
        DFCH.Text = "0"
        DFCM.Text = "0"
        '  計算
        Select Case DDateType.Text
            Case "假日"
                If sHour > 8 Then
                    DFBH.Text = "8"
                    DFBM.Text = "0"
                    sHour = sHour - 8
                    If sHour > 2 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        If sHour = 2 And sMinute > 0 Then
                            DFAH1.Text = "2"
                            DFAM1.Text = "0"
                            DFAH2.Text = CStr(sHour - 2)
                            DFAM2.Text = CStr(sMinute)
                        Else
                            DFAH1.Text = CStr(sHour)
                            DFAM1.Text = CStr(sMinute)
                        End If
                    End If
                Else
                    DFBH.Text = CStr(sHour)
                    DFBM.Text = CStr(sMinute)
                End If
            Case "國定假日"
                If sHour > 8 Then
                    DFCH.Text = "8"
                    DFCM.Text = "0"
                    sHour = sHour - 8
                    If sHour > 2 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        If sHour = 2 And sMinute > 0 Then
                            DFAH1.Text = "2"
                            DFAM1.Text = "0"
                            DFAH2.Text = CStr(sHour - 2)
                            DFAM2.Text = CStr(sMinute)
                        Else
                            DFAH1.Text = CStr(sHour)
                            DFAM1.Text = CStr(sMinute)
                        End If
                    End If
                Else
                    DFCH.Text = CStr(sHour)
                    DFCM.Text = CStr(sMinute)
                End If
            Case "社休假日"
                If sHour > 2 Then
                    DFAH1.Text = "2"
                    DFAM1.Text = "0"
                    DFAH2.Text = CStr(sHour - 2)
                    DFAM2.Text = CStr(sMinute)
                Else
                    If sHour = 2 And sMinute > 0 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        DFAH1.Text = CStr(sHour)
                        DFAM1.Text = CStr(sMinute)
                    End If
                End If
            Case Else   '平日
                If sHour > 2 Then
                    DFAH1.Text = "2"
                    DFAM1.Text = "0"
                    DFAH2.Text = CStr(sHour - 2)
                    DFAM2.Text = CStr(sMinute)
                Else
                    If sHour = 2 And sMinute > 0 Then
                        DFAH1.Text = "2"
                        DFAM1.Text = "0"
                        DFAH2.Text = CStr(sHour - 2)
                        DFAM2.Text = CStr(sMinute)
                    Else
                        DFAH1.Text = CStr(sHour)
                        DFAM1.Text = CStr(sMinute)
                    End If
                End If
        End Select
        '薪資計算 -----------------------------------------
        '清初值
        DFPRAH1.Text = "0"
        DFPRAM1.Text = "0"
        DFPRAH2.Text = "0"
        DFPRAM2.Text = "0"
        DFPRBH.Text = "0"
        DFPRBM.Text = "0"
        DFPRCH.Text = "0"
        DFPRCM.Text = "0"

        DFPRTAX15H.Text = "0"
        DFPRTAX15M.Text = "0"
        DFPRTAX167H.Text = "0"
        DFPRTAX167M.Text = "0"
        DFPRTAX20H.Text = "0"
        DFPRTAX20M.Text = "0"
        DFPRTAX267H.Text = "0"
        DFPRTAX267M.Text = "0"

        '  決定加班日-加班類別 (平日,假日,國定假日)
        Dim SQL, xDateType As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        xDateType = DDateType.Text
        SQL = "Select *  From M_Vacation "      ' 薪資專用行事曆(只有國定假日)
        SQL = SQL & " Where Active = '1' "
        SQL = SQL & " And Depo = '" & "PR1" & "'"
        SQL = SQL & " And YMD = '" & DOverTimeDate.Text & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Vacation")
        If DBDataSet1.Tables("Vacation").Rows.Count > 0 Then
            xDateType = "國定假日"
        Else
            If xDateType = "國定假日" Then
                xDateType = "假日"
            End If
        End If
        OleDbConnection1.Close()
        '  計算加班費率  JOY BY 2017/6/19
        Select Case xDateType
            Case "社休假日"
                ' 社休假日: 1~2=應稅1.5, 3~8=應稅1.67, 9~ =應稅2.67
                Select Case xHour
                    Case 1 To 2
                        DFPRTAX15H.Text = CStr(xHour)
                        If xHour = 2 Then
                            DFPRTAX167M.Text = CStr(xMinute)
                        Else
                            DFPRTAX15M.Text = CStr(xMinute)
                        End If
                    Case 3 To 8
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(xHour - 2)
                        If xHour = 8 Then
                            DFPRTAX267M.Text = CStr(xMinute)
                        Else
                            DFPRTAX167M.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(6)
                        DFPRTAX267H.Text = CStr(xHour - 8)
                        DFPRTAX267M.Text = CStr(xMinute)
                End Select
            Case "假日"
                ' 例假日: 1~8=免稅1.5, 9~10=應稅1.5, 11~ =應稅1.67
                ' 國假日: 1~8=免稅1.5, 9~10=應稅1.5, 11~ =應稅1.67
                Select Case xHour
                    Case 1 To 8
                        DFPRBH.Text = CStr(xHour)
                        If xHour = 8 Then
                            DFPRTAX15M.Text = CStr(xMinute)
                        Else
                            DFPRBM.Text = CStr(xMinute)
                        End If
                    Case 9 To 10
                        DFPRBH.Text = CStr(8)
                        DFPRTAX15H.Text = CStr(xHour - 8)
                        If xHour = 10 Then
                            DFPRTAX167M.Text = CStr(xMinute)
                        Else
                            DFPRTAX15M.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRBH.Text = CStr(8)
                        DFPRTAX15H.Text = CStr(2)
                        DFPRTAX167H.Text = CStr(xHour - 10)
                        DFPRTAX167M.Text = CStr(xMinute)
                End Select
            Case "國定假日"
                ' 國假日(ykk): 1~8=免稅2.0, , 9~ =應稅2.0
                Select Case xHour
                    Case 1 To 8
                        DFPRCH.Text = CStr(xHour)
                        If xHour = 8 Then
                            DFPRTAX20M.Text = CStr(xMinute)
                        Else
                            DFPRCM.Text = CStr(xMinute)
                        End If
                    Case Else
                        DFPRCH.Text = CStr(8)
                        DFPRTAX20H.Text = CStr(xHour - 8)
                        DFPRTAX20M.Text =  CStr(xMinute)
                End Select
            Case Else
                ' 平日: 9~10=應稅1.34, 11~ =應稅1.67
                DFPRAH1.Text = DFAH1.Text
                DFPRAM1.Text = DFAM1.Text
                DFPRAH2.Text = DFAH2.Text
                DFPRAM2.Text = DFAM2.Text
        End Select

        'Select Case xDateType
        '    Case "社休假日"     '(社休假日,休假日)

        '        ' MOD-Start 2016/12/28
        '        If CDate(DOverTimeDate.Text) >= CDate("2016/12/23") Then
        '            If xHour >= 2 Then
        '                DFPRBH.Text = CStr(2)           '2內=1.5
        '                DFPRAH2.Text = CStr(xHour - 2)  '2外=1.67
        '                DFPRAM2.Text = CStr(xMinute)
        '            Else
        '                DFPRBH.Text = CStr(xHour)
        '                DFPRBM.Text = CStr(xMinute)
        '            End If
        '        Else
        '            DFPRBH.Text = CStr(xHour)
        '            DFPRBM.Text = CStr(xMinute)
        '        End If
        '        'DFPRBH.Text = CStr(xHour)
        '        'DFPRBM.Text = CStr(xMinute)
        '        ' MOD-End

        '    Case "假日"         '(例假日,國假日)
        '        DFPRBH.Text = CStr(xHour)
        '        DFPRBM.Text = CStr(xMinute)
        '    Case "國定假日"     '(國假日ykk)
        '        DFPRCH.Text = CStr(xHour)
        '        DFPRCM.Text = CStr(xMinute)
        '    Case Else   '平日
        '        DFPRAH1.Text = DFAH1.Text
        '        DFPRAM1.Text = DFAM1.Text
        '        DFPRAH2.Text = DFAH2.Text
        '        DFPRAM2.Text = DFAM2.Text
        'End Select

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     計算加班起迄時間 
    '**
    '*****************************************************************
    Function CalOverTime(ByVal sHH As String, ByVal sMM As String, ByVal eHH As String, ByVal eMM As String) As Integer
        Dim StartTime As Integer = CInt(sHH) * 60 + CInt(sMM)
        Dim EndTime As Integer = CInt(eHH) * 60 + CInt(eMM)
        CalOverTime = EndTime - StartTime
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編製委託No 
    '**
    '*****************************************************************
    Function SetNo(ByVal Seq As Integer) As String
        Dim Str As String = ""
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim i As Integer

        'Set當日日期
        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 10 - 1
            Str1 = "0" + Str1
        Next

        SetNo = Str + Str1
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案xx
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 2000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If

        'If UPFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
        'Check上傳檔案格式
        'Else
        'UPFileIsNormal = 9030
        'End If
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Disabled = True
        BNG1.Disabled = True
        BNG2.Disabled = True
        BSAVE.Disabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Disabled = False
        BNG1.Disabled = False
        BNG2.Disabled = False
        BSAVE.Disabled = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     輸入檢查
    '**
    '*****************************************************************
    Function InputCheck() As Integer
        InputCheck = 0
        'No
        If InputCheck = 0 Then
            If FindFieldInf("No") = 1 Then
                If DNo.Text = "" Then InputCheck = 1
            End If
        End If
        '日期
        If InputCheck = 0 Then
            If FindFieldInf("Date") = 1 Then
                If DDate.Text = "" Then InputCheck = 1
            End If
        End If
        '姓名
        If InputCheck = 0 Then
            If FindFieldInf("Name") = 1 Then
                If DName.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'EmpID
        If InputCheck = 0 Then
            If FindFieldInf("EmpID") = 1 Then
                If DEmpID.Text = "" Then InputCheck = 1
            End If
        End If
        '職稱
        If InputCheck = 0 Then
            If FindFieldInf("JobTitle") = 1 Then
                If DJobTitle.Text = "" Then InputCheck = 1
            End If
        End If
        '職稱代碼
        If InputCheck = 0 Then
            If FindFieldInf("JobCode") = 1 Then
                If DJobCode.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Name
        If InputCheck = 0 Then
            If FindFieldInf("DepoName") = 1 Then
                If DDepoName.Text = "" Then InputCheck = 1
            End If
        End If
        'Depo Code
        If InputCheck = 0 Then
            If FindFieldInf("DepoCode") = 1 Then
                If DDepoCode.Text = "" Then InputCheck = 1
            End If
        End If
        '部門
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.Text = "" Then InputCheck = 1
            End If
        End If
        '部門代碼
        If InputCheck = 0 Then
            If FindFieldInf("DivisionCode") = 1 Then
                If DDivisionCode.Text = "" Then InputCheck = 1
            End If
        End If
        '加班日期
        If InputCheck = 0 Then
            If FindFieldInf("OverTimeDate") = 1 Then
                If DOverTimeDate.Text = "" Then InputCheck = 1
            End If
        End If
        '加班別
        If InputCheck = 0 Then
            If FindFieldInf("DateType") = 1 Then
                If DDateType.Text = "" Then InputCheck = 1
            End If
        End If
        '所屬年月
        If InputCheck = 0 Then
            If FindFieldInf("SalaryYM") = 1 Then
                If DSalaryYM.Text = "" Then InputCheck = 1
            End If
        End If
        '調休
        If InputCheck = 0 Then
            If FindFieldInf("CVacation") = 1 Then
                If DCVacation.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '伙食
        If InputCheck = 0 Then
            If FindFieldInf("Food") = 1 Then
                If DFood.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '交通
        If InputCheck = 0 Then
            If FindFieldInf("Traffic") = 1 Then
                If DTraffic.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定開始-時
        If InputCheck = 0 Then
            If FindFieldInf("BStartH") = 1 Then
                If DBStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定開始-分
        If InputCheck = 0 Then
            If FindFieldInf("BStartM") = 1 Then
                If DBStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定終止-時
        If InputCheck = 0 Then
            If FindFieldInf("BEndH") = 1 Then
                If DBEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定終止-分
        If InputCheck = 0 Then
            If FindFieldInf("BEndM") = 1 Then
                If DBEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '預定計算-時
        If InputCheck = 0 Then
            If FindFieldInf("BH") = 1 Then
                If DBH.Text = "" Then InputCheck = 1
            End If
        End If
        '預定計算-分
        If InputCheck = 0 Then
            If FindFieldInf("BM") = 1 Then
                If DBM.Text = "" Then InputCheck = 1
            End If
        End If
        '實際開始-時
        If InputCheck = 0 Then
            If FindFieldInf("AStartH") = 1 Then
                If DAStartH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際開始-分
        If InputCheck = 0 Then
            If FindFieldInf("AStartM") = 1 Then
                If DAStartM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際終止-時
        If InputCheck = 0 Then
            If FindFieldInf("AEndH") = 1 Then
                If DAEndH.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際終止-分
        If InputCheck = 0 Then
            If FindFieldInf("AEndM") = 1 Then
                If DAEndM.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '實際計算-時
        If InputCheck = 0 Then
            If FindFieldInf("AH") = 1 Then
                If DAH.Text = "" Then InputCheck = 1
            End If
        End If
        '實際計算-分
        If InputCheck = 0 Then
            If FindFieldInf("AM") = 1 Then
                If DAM.Text = "" Then InputCheck = 1
            End If
        End If
        '核定平日2內-時
        If InputCheck = 0 Then
            If FindFieldInf("FAH1") = 1 Then
                If DFAH1.Text = "" Then InputCheck = 1
            End If
        End If
        '核定平日2內-分
        If InputCheck = 0 Then
            If FindFieldInf("FAM1") = 1 Then
                If DFAM1.Text = "" Then InputCheck = 1
            End If
        End If
        '核定平日2外-時
        If InputCheck = 0 Then
            If FindFieldInf("FAH2") = 1 Then
                If DFAH2.Text = "" Then InputCheck = 1
            End If
        End If
        '核定平日2外-分
        If InputCheck = 0 Then
            If FindFieldInf("FAM2") = 1 Then
                If DFAM2.Text = "" Then InputCheck = 1
            End If
        End If
        '核定假日-時
        If InputCheck = 0 Then
            If FindFieldInf("FBH") = 1 Then
                If DFBH.Text = "" Then InputCheck = 1
            End If
        End If
        '核定假日-分
        If InputCheck = 0 Then
            If FindFieldInf("FBM") = 1 Then
                If DFBM.Text = "" Then InputCheck = 1
            End If
        End If
        '核定國定假日-時
        If InputCheck = 0 Then
            If FindFieldInf("FCH") = 1 Then
                If DFCH.Text = "" Then InputCheck = 1
            End If
        End If
        '核定國定假日-分
        If InputCheck = 0 Then
            If FindFieldInf("FCM") = 1 Then
                If DFCM.Text = "" Then InputCheck = 1
            End If
        End If
        '加班理由
        If InputCheck = 0 Then
            If FindFieldInf("FReason") = 1 Then
                If DFReason.Text = "" Then InputCheck = 1
            End If
        End If

    End Function

End Class
