Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ManufInCTSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DOPContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BModify As System.Web.UI.WebControls.Button
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAttachFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.DropDownList

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
    'Dim wDepo As String = "CL"      '中壢行事曆(CL->中壢, TP->台北)
    '
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    'Modify-End

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
            If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
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

        Response.Cookies("DevNo").Value = ""        '開發No, DevNoPicker使用
        Response.Cookies("MapNo").Value = ""        '圖號, MapPicker使用
        Response.Cookies("Step").Value = Request.QueryString("pStep")  '隱藏用工程代碼

        Response.Cookies("PGM").Value = "ManufInCTSheet_01.aspx"
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
        Dim Message As String = ""
        'Check附件
        If DAttachFile.Visible Then
            If DAttachFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "附件"
                Else
                    Message = Message & ", " & "附件"
                End If
            End If
        End If

        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
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
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ManufInCTFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ManufInCTSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ManufInCTSheet")
        If DBDataSet1.Tables("F_ManufInCTSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Person"))      '擔當
            DSliderCode.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("SliderCode")       'Slider Code
            DLevel.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Level")                 '難易度
            DMapNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("MapNo")             '圖號
            If DMapNo.Text <> "" Then
                SQL = "Select FormNo, FormSno From F_MapSheet "
                SQL = SQL & " Where Sts = 1 "
                SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "MapSheet")
                If DBDataSet1.Tables("MapSheet").Rows.Count > 0 Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormNo") & _
                                                         "&pFormSno=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormSno")
                Else
                    SQL = "Select FormNo, FormSno From F_MapModSheet "
                    SQL = SQL & " Where Sts = 1 "
                    SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                    Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter4.Fill(DBDataSet1, "MapModSheet")
                    If DBDataSet1.Tables("MapModSheet").Rows.Count > 0 Then
                        LMapNo.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapModSheet").Rows(0).Item("FormNo") & _
                                                                "&pFormSno=" & DBDataSet1.Tables("MapModSheet").Rows(0).Item("FormSno")
                    End If
                End If
            Else
                LMapNo.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("OFormNo")             '圖號
            DOFormSno.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("OFormSno")             '圖號

            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000003" Then
                    LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                End If
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormNo")             '圖號
            If DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("NFormSno")             '圖號
            End If

            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ManufInModSheet_01.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
                If BModify.Visible = True Then
                    LNFormNo.Visible = False
                End If
                DReady.Text = "未閱讀"
            Else
                LNFormNo.Visible = False
                DReady.Text = ""
            End If

            SetFieldData("Target", DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Target"))
            DContent.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Content")             '圖號
            DDReason.Text = DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("Reason")             '圖號
            If DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("AttachFile") <> "" Then          '圖檔1
                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ManufInCTSheet").Rows(0).Item("AttachFile")
            Else
                LAttachFile.Visible = False
            End If

            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter2.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                DDecideDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("DecideDesc")       '說明
                DBStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BStartTime")       '預定開始
                DBEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime")           '預定完成
                DAStartTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AStartTime")       '實際開始
                If Not IsDBNull(DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")) Then
                    DAEndTime.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("AEndTime")           '實際完成
                Else
                    DAEndTime.Text = ""
                End If
                LBefOP.NavigateUrl = "BefOPList.aspx?pFormNo=" & wFormNo & "&pFormSno=" & wFormSno & "&pStep=" & wStep

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
                DFormSno.Text = "單號：" & DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FormSno")       '單號
            End If
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

        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DOPContactSheet.Visible = True   '表單Sheet-1
                DDescSheet.Visible = True       '說明Sheet
                DDelivery.Visible = True        '交期Sheet

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
                    Top = 752
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 648
                End If

                '交期
                DBStartTime.Visible = True      '預定開始
                DBEndTime.Visible = True        '預定完成
                DAStartTime.Visible = True      '實際開始
                DAEndTime.Visible = True        '實際完成
                DOPReady.Visible = True     '工程履歷閱讀
                DOPReadyDesc.Visible = True
                '需閱讀
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("FlowType") = 0 Then
                    DOPReady.BackColor = Color.LightGray
                Else
                    DOPReady.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "異常：需閱讀工程履歷")
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

                '連結顯示---需再修改
                LMapNo.Visible = True          '圖檔
                LOFormNo.Visible = True        '原委託
                LNFormNo.Visible = True        '新委託
                LAttachFile.Visible = True     '附件
                LBefOP.Visible = True          '工程履歷
                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
            Top = 464
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelivery.Visible = False   '交期Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False     '說明
            DBStartTime.Visible = False     '預定開始
            DBEndTime.Visible = False       '預定完成
            DAStartTime.Visible = False     '實際開始
            DAEndTime.Visible = False       '實際完成
            DReasonCode.Visible = False     '延遲理由代碼
            DReason.Visible = False         '延遲理由
            DReasonDesc.Visible = False     '延遲其他說明

            DOPReady.Visible = False    '工程履歷閱讀
            DOPReadyDesc.Visible = False

            '連結隱藏
            LMapNo.Visible = False          '圖檔
            LOFormNo.Visible = False        '原委託
            LNFormNo.Visible = False        '新委託
            LAttachFile.Visible = False     '附件
            LBefOP.Visible = False          '工程履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)      'NG1按鈕
            BNG2.Style.Add("Top", Top)      'NG2按鈕
            BSAVE.Style.Add("Top", Top)     '儲存按鈕
            BOK.Style.Add("Top", Top)       'OK按鈕
            DFormSno.Style.Add("Top", Top)  '單號
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
        BDate.Attributes("onclick") = "CalendarPicker('Form1.DDate');"  '日期選擇
        BIn.Attributes("onclick") = "DevNoPicker('In','000005');"       '內製
        BOMapNo.Attributes("onclick") = "MapPicker('Ori','000005');"    '原始圖號
        BMMapNo.Attributes("onclick") = "MapPicker('Mod','000005');"    '修改圖號
        BModify.Attributes("onclick") = "ModifySheet();"                '原委託
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
        If wFormSno > 0 And wStep > 2 Then      '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") >= NowDateTime Then
                    Top = 648
                Else
                    If DDelay.Visible = True Then
                        Top = 752
                    Else
                        Top = 648
                    End If
                End If
            End If
        Else
            Top = 464
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
    '**(ShowSheetField)
    '**     表單各欄位屬性及欄位輸入檢查等設定
    '**
    '*****************************************************************
    Sub SetFieldAttribute(ByVal pPost As String)
        '表單各欄位屬性及欄位輸入檢查等設定
        'No
        Select Case FindFieldInf("No")
            Case 0  '顯示
                DNo.BackColor = Color.LightGray
                DNo.ReadOnly = True
                DNo.Visible = True
                BIn.Visible = False
            Case 1  '修改+檢查
                DNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DnoRqd", "Dno", "異常：需輸入Ｎｏ")
                DNo.Visible = True
                BIn.Visible = True
            Case 2  '修改
                DNo.BackColor = Color.Yellow
                DNo.Visible = True
                BIn.Visible = True
            Case Else   '隱藏
                DNo.Visible = False
                BIn.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = ""
        '日期
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.ReadOnly = True
                DDate.Visible = True
                BDate.Visible = False
            Case 1  '修改+檢查
                DDate.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入日期")
                DDate.Visible = True
                BDate.Visible = True
            Case 2  '修改
                DDate.BackColor = Color.Yellow
                DDate.Visible = True
                BDate.Visible = True
            Case Else   '隱藏
                DDate.Visible = False
                BDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = CStr(DateTime.Now.Today)
        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
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
        If pPost = "New" Then SetFieldData("Division", "ZZZZZZ")
        '擔當
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.LightGray
                DPerson.Visible = True
            Case 1  '修改+檢查
                DPerson.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DPersonRqd", "DPerson", "異常：需輸入擔當")
                DPerson.Visible = True
            Case 2  '修改
                DPerson.BackColor = Color.Yellow
                DPerson.Visible = True
            Case Else   '隱藏
                DPerson.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Person", "ZZZZZZ")

        'Target
        Select Case FindFieldInf("Target")
            Case 0  '顯示
                DTarget.BackColor = Color.LightGray
                DTarget.Visible = True
            Case 1  '修改+檢查
                DTarget.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTargetRqd", "DTarget", "異常：需輸入目的")
                DTarget.Visible = True
            Case 2  '修改
                DTarget.BackColor = Color.Yellow
                DTarget.Visible = True
            Case Else   '隱藏
                DTarget.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Target", "ZZZZZZ")

        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '顯示
                DSliderCode.BackColor = Color.LightGray
                DSliderCode.ReadOnly = True
                DSliderCode.Visible = True
            Case 1  '修改+檢查
                DSliderCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSliderCodeRqd", "DSliderCode", "異常：需輸入 Slider Code")
                DSliderCode.Visible = True
            Case 2  '修改
                DSliderCode.BackColor = Color.Yellow
                DSliderCode.Visible = True
            Case Else   '隱藏
                DSliderCode.Visible = False
        End Select
        If pPost = "New" Then DSliderCode.Text = ""
        '難易度
        Select Case FindFieldInf("Level")
            Case 0  '顯示
                DLevel.BackColor = Color.LightGray
                DLevel.ReadOnly = True
                DLevel.Visible = True
            Case 1  '修改+檢查
                DLevel.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DLevelRqd", "DLevel", "異常：需輸入難易度")
                DLevel.Visible = True
            Case 2  '修改
                DLevel.BackColor = Color.Yellow
                DLevel.Visible = True
            Case Else   '隱藏
                DLevel.Visible = False
        End Select
        If pPost = "New" Then DLevel.Text = ""
        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.LightGray
                DMapNo.ReadOnly = True
                DMapNo.Visible = True
                BOMapNo.Visible = False
                BMMapNo.Visible = False
            Case 1  '修改+檢查
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入圖號")
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case 2  '修改
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
                BOMapNo.Visible = True
                BMMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
                BOMapNo.Visible = False
                BMMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""

        'OFormNo
        Select Case FindFieldInf("OFormNo")
            Case 0  '顯示
                DOFormNo.BackColor = Color.LightGray
                DOFormNo.ReadOnly = True
                DOFormNo.Visible = True
            Case 1  '修改+檢查
                DOFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormNoRqd", "DOFormNo", "異常：需輸入表單號碼")
                DOFormNo.Visible = True
            Case 2  '修改
                DOFormNo.BackColor = Color.Yellow
                DOFormNo.Visible = True
            Case Else   '隱藏
                DOFormNo.Visible = False
        End Select
        If pPost = "New" Then DOFormNo.Text = ""
        'OFormSno
        Select Case FindFieldInf("OFormSno")
            Case 0  '顯示
                DOFormSno.BackColor = Color.LightGray
                DOFormSno.ReadOnly = True
                DOFormSno.Visible = True
            Case 1  '修改+檢查
                DOFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOFormSnoRqd", "DOFormSno", "異常：需輸入單號")
                DOFormSno.Visible = True
            Case 2  '修改
                DOFormSno.BackColor = Color.Yellow
                DOFormSno.Visible = True
            Case Else   '隱藏
                DOFormSno.Visible = False
        End Select
        If pPost = "New" Then DOFormSno.Text = ""
        'NFormNo
        Select Case FindFieldInf("NFormNo")
            Case 0  '顯示
                DNFormNo.BackColor = Color.LightGray
                DNFormNo.ReadOnly = True
                DNFormNo.Visible = True
                BModify.Visible = False
            Case 1  '修改+檢查
                DNFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNFormNoRqd", "DNFormNo", "異常：需輸入表單號碼")
                DNFormNo.Visible = True
                BModify.Visible = True
            Case 2  '修改
                DNFormNo.BackColor = Color.Yellow
                DNFormNo.Visible = True
                BModify.Visible = True
            Case Else   '隱藏
                DNFormNo.Visible = False
                BModify.Visible = False
        End Select
        If pPost = "New" Then DNFormNo.Text = ""
        'NFormSno
        Select Case FindFieldInf("NFormSno")
            Case 0  '顯示
                DNFormSno.BackColor = Color.LightGray
                DNFormSno.ReadOnly = True
                DNFormSno.Visible = True
                BModify.Visible = False
            Case 1  '修改+檢查
                DNFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DNFormSnoRqd", "DNFormSno", "異常：需輸入單號")
                DNFormSno.Visible = True
                BModify.Visible = True
            Case 2  '修改
                DNFormSno.BackColor = Color.Yellow
                DNFormSno.Visible = True
                BModify.Visible = True
            Case Else   '隱藏
                DNFormSno.Visible = False
                BModify.Visible = False
        End Select
        If pPost = "New" Then DNFormSno.Text = ""

        'Content
        Select Case FindFieldInf("Content")
            Case 0  '顯示
                DContent.BackColor = Color.LightGray
                DContent.ReadOnly = True
                DContent.Visible = True
            Case 1  '修改+檢查
                DContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DContentRqd", "DContent", "異常：需輸入內容")
                DContent.Visible = True
            Case 2  '修改
                DContent.BackColor = Color.Yellow
                DContent.Visible = True
            Case Else   '隱藏
                DContent.Visible = False
        End Select
        If pPost = "New" Then DContent.Text = ""
        'Reason
        Select Case FindFieldInf("Reason")
            Case 0  '顯示
                DDReason.BackColor = Color.LightGray
                DDReason.ReadOnly = True
                DDReason.Visible = True
            Case 1  '修改+檢查
                DDReason.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DDReasonRqd", "DDReason", "異常：需輸入原因")
                DDReason.Visible = True
            Case 2  '修改
                DDReason.BackColor = Color.Yellow
                DDReason.Visible = True
            Case Else   '隱藏
                DDReason.Visible = False
        End Select
        If pPost = "New" Then DDReason.Text = ""
        '附件
        Select Case FindFieldInf("AttachFile")
            Case 0  '顯示
                DAttachFile.Visible = False
                DAttachFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAttachFileRqd", "DAttachFile", "異常：需輸入附件")
                DAttachFile.Visible = True
                DAttachFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DAttachFile.Visible = True
                DAttachFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DAttachFile.Visible = False
        End Select
        '單號
        If pPost = "New" Then DFormSno.Text = ""

    End Sub
    '*****************************************************************
    '**(ShowSheetField)
    '**     建立表單需輸入欄位
    '**
    '*****************************************************************
    Sub ShowRequiredFieldValidator(ByVal pID As String, ByVal pField As String, ByVal pMessage As String)
        Dim rqdVal As RequiredFieldValidator = New RequiredFieldValidator
        rqdVal.ID = pID
        rqdVal.ControlToValidate = pField
        rqdVal.ErrorMessage = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", CStr(Top + 25))
        rqdVal.Style.Add("Left", "8")
        rqdVal.Style.Add("Height", "20")
        rqdVal.Style.Add("Width", "250")
        rqdVal.Style.Add("POSITION", "absolute")
        Page.Controls(1).Controls.Add(rqdVal)
    End Sub

    '*****************************************************************
    '**(ShowSheetField)
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

        '部門
        If pFieldName = "Division" Then
            DDivision.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DDivision.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select DivName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("DivName")
                    ListItem1.Value = DBTable1.Rows(i).Item("DivName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DDivision.Items.Add(ListItem1)
                Next
            End If
        End If
        '擔當
        If pFieldName = "Person" Then
            DPerson.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DPerson.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select UserName From M_Users "
                SQL = SQL & " Where UserID = '" & Request.QueryString("pUserID") & "'"
                SQL = SQL & "   And Active = '1' "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Users")
                DBTable1 = DBDataSet1.Tables("M_Users")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("UserName")
                    ListItem1.Value = DBTable1.Rows(i).Item("UserName")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DPerson.Items.Add(ListItem1)
                Next
            End If
        End If
        '目的
        If pFieldName = "Target" Then
            DTarget.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DTarget.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='200' and DKey='TARGET' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DTarget.Items.Add(ListItem1)
                Next
            End If
        End If

        '難易度
        'If pFieldName = "Level" Then
        'DLevel.Items.Clear()
        'If idx = 0 Then
        'If pName <> "ZZZZZZ" Then
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = pName
        'ListItem1.Value = pName
        'DLevel.Items.Add(ListItem1)
        'End If
        'Else
        'SQL = "Select * From M_Referp Where Cat='007' and DKey='0' Order by DKey, Data "
        'Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "M_Referp")
        'DBTable1 = DBDataSet1.Tables("M_Referp")
        'For i = 0 To DBTable1.Rows.Count - 1
        'Dim ListItem1 As New ListItem
        'ListItem1.Text = DBTable1.Rows(i).Item("Data")
        'ListItem1.Value = DBTable1.Rows(i).Item("Data")
        'If ListItem1.Value = pName Then ListItem1.Selected = True
        'DLevel.Items.Add(ListItem1)
        'Next
        'End If
        'End If

        '延遲理由代碼
        If pFieldName = "ReasonCode" Then
            SQL = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
            DReasonCode.Items.Clear()
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
    '**(ShowSheetField)
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

                'Check附件Size及格式
                If ErrCode = 0 Then
                    If DAttachFile.Visible Then
                        If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DAttachFile)
                        End If
                    End If
                End If

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
                    If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
                    If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
                    If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
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
        Dim Message As String = ""
        Dim wQCLT As Integer = 0 'QC-L/T

        GetDataStatus()  '取得此關卡各按鈕的Data Status

        'Check是否已閱讀
        If ErrCode = 0 Then
            If DReady.Visible Then
                If DReady.Text = "未閱讀" Then
                    ErrCode = 9050
                End If
            End If
        End If
        'Check附件Size及格式
        If ErrCode = 0 Then
            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAttachFile)
                End If
            End If
        End If
        'Check延遲理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9040
                End If
            End If
        End If

        '--檢查委託書No---------
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                ErrCode = oCommon.CommissionNo("000005", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No

                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = DLevel.Text  '難易度

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
                Dim SQL As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
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

                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
                    '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                    '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 

                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
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
                                wNextGate = wNextGate & "," & pNextGate(i)
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
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)    '新增表單資料
                            AddCommissionNo(wFormNo, NewFormSno)
                            ModifyManufInData(pFun, 1, 1)   '更新原表單狀態
                        End If
                    Else
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            ModifyManufInData(pFun, 0, 1)   '更新原表單狀態

                            'If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                            'ReplaceManufInData()            '覆蓋原表單及更新狀態
                            'Else
                            'ModifyManufInData(pFun, 0, 1)   '更新原表單狀態
                            'End If
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
            If ErrCode = 9050 Then Message = "未閱讀新委託單內容,請點選新委託!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInCTFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_ManufInCTSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, "          '1~4
        SQl = SQl + "No, Date, Division, Person, SliderCode, "       '5~9
        SQl = SQl + "Level, MapNo, OFormNo, OFormSno, NFormNo, NFormSno, Suppiler, "  '10~14
        SQl = SQl + "Target, Content, Reason, AttachFile, "          '15~18
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "  '19~22
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "              '結案日
        SQl = SQl + " '000005', "                           '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "         '表單流水號
        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "                 'NO
        SQl = SQl + " '" + DDate.Text + "', "               '日期
        SQl = SQl + " '" + DDivision.SelectedValue + "', "  '部門
        SQl = SQl + " '" + DPerson.SelectedValue + "', "    '擔當
        SQl = SQl + " '" + YKK.ReplaceString(DSliderCode.Text) + "', "         'Slider Code
        SQl = SQl + " '" + YKK.ReplaceString(DLevel.Text) + "', "     '難易度
        SQl = SQl + " '" + DMapNo.Text + "', "              '圖號
        SQl = SQl + " '" + DOFormNo.Text + "', "            '表單No
        If DOFormSno.Text = "" Then                         '單號
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DOFormSno.Text + "', "
        End If
        SQl = SQl + " '" + DNFormNo.Text + "', "            '表單No
        If DNFormSno.Text = "" Then                         '單號
            SQl = SQl + " '0', "
        Else
            SQl = SQl + " '" + DNFormSno.Text + "', "
        End If
        SQl = SQl + " '" + "" + "', "          '外注商
        SQl = SQl + " N'" + YKK.ReplaceString(DTarget.SelectedValue) + "', "             '目的
        SQl = SQl + " N'" + YKK.ReplaceString(DContent.Text) + "', "            '內容
        SQl = SQl + " N'" + YKK.ReplaceString(DDReason.Text) + "', "             '原因

        FileName = ""
        If DAttachFile.Visible Then                         '附件
            If DAttachFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ManufInCTFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_ManufInCTSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " SliderCode = '" & YKK.ReplaceString(DSliderCode.Text) & "',"
        SQl = SQl + " Level = '" & YKK.ReplaceString(DLevel.Text) & "',"
        SQl = SQl + " MapNo = '" & DMapNo.Text & "',"
        SQl = SQl + " OFormNo = '" & DOFormNo.Text & "',"
        SQl = SQl + " OFormSno = '" & DOFormSno.Text & "',"
        SQl = SQl + " NFormNo = '" & DNFormNo.Text & "',"
        SQl = SQl + " NFormSno = '" & DNFormSno.Text & "',"
        SQl = SQl + " Suppiler = '" & "" & "',"
        SQl = SQl + " Target = N'" & YKK.ReplaceString(DTarget.SelectedValue) & "',"
        SQl = SQl + " Content = N'" & YKK.ReplaceString(DContent.Text) & "',"
        SQl = SQl + " Reason = N'" & YKK.ReplaceString(DDReason.Text) & "',"

        If DAttachFile.Visible Then
            If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DAttachFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " AttachFile = N'" & FileName & "',"
            End If
        End If

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
    '**(ReplaceManufInData)
    '**     覆蓋原資料(停用)
    '**
    '*****************************************************************
    Sub ReplaceManufInData()
        'Dim OleDbConnection1 As New OleDbConnection
        'OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        'Dim OleDBCommand1 As New OleDbCommand

        'Dim DBDataSet1 As New DataSet
        'Dim SQl As String
        'Dim RtnCode As Integer

        'SQl = "Select * From F_ManufInModSheet "
        'SQl = SQl & " Where FormNo =  '" & DNFormNo.Text & "'"
        'SQl = SQl & "   And FormSno =  '" & DNFormSno.Text & "'"
        'Dim DBAdapter1 As New OleDbDataAdapter(SQl, OleDbConnection1)
        'DBAdapter1.Fill(DBDataSet1, "F_ManufInModSheet")
        'If DBDataSet1.Tables("F_ManufInModSheet").Rows.Count > 0 Then
        '    SQl = "Update F_ManufInSheet Set "
        '    SQl = SQl + " No = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("No") & "',"
        '    SQl = SQl + " Date = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Date") & "',"
        '    SQl = SQl + " Division = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Division") & "',"
        '    SQl = SQl + " Person = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Person") & "',"
        '    SQl = SQl + " SliderCode = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderCode") & "',"
        '    SQl = SQl + " SliderGRCode = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderGRCode") & "',"
        '    SQl = SQl + " Spec = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Spec") & "',"
        '    SQl = SQl + " MapNo = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapNo") & "',"
        '    SQl = SQl + " OFormNo = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormNo") & "',"
        '    SQl = SQl + " OFormSno = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("OFormSno") & "',"
        '    SQl = SQl + " MapFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MapFile") & "',"
        '    SQl = SQl + " RefFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("RefFile") & "',"
        '    SQl = SQl + " Level = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Level") & "',"
        '    SQl = SQl + " Assembler = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Assembler") & "',"
        '    SQl = SQl + " SliderType1 = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType1") & "',"
        '    SQl = SQl + " SliderType2 = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SliderType2") & "',"
        '    SQl = SQl + " ManufPlace = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ManufPlace") & "',"
        '    SQl = SQl + " Material = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Material") & "',"
        '    SQl = SQl + " MaterialOther = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MaterialOther") & "',"
        '    SQl = SQl + " SellVendor = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SellVendor") & "',"
        '    SQl = SQl + " Buyer = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Buyer") & "',"
        '    SQl = SQl + " ConfirmFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ConfirmFile") & "',"
        '    SQl = SQl + " AuthorizeFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("AuthorizeFile") & "',"
        '    SQl = SQl + " DevReason = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("DevReason") & "',"
        '    SQl = SQl + " Sample = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Sample") & "',"
        '    SQl = SQl + " Price = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Price") & "',"
        '    SQl = SQl + " ArMoldFee = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ArMoldFee") & "',"
        '    SQl = SQl + " PurMold = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PurMold") & "',"
        '    SQl = SQl + " PullerPrice = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("PullerPrice") & "',"
        '    SQl = SQl + " Suppiler = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Suppiler") & "',"
        '    SQl = SQl + " MoldQty = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldQty") & "',"
        '    SQl = SQl + " MoldPoint = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("MoldQty") & "',"
        '    SQl = SQl + " Quality1 = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality1") & "',"
        '    SQl = SQl + " Quality2 = '" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("Quality2") & "',"
        '    SQl = SQl + " QAFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAFile") & "',"
        '    SQl = SQl + " SampleFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SampleFile") & "',"
        '    SQl = SQl + " ContactFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("ContactFile") & "',"

        '    SQl = SQl + " SAttachFile = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("SAttachFile") & "',"
        '    SQl = SQl + " QAttachFile1 = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile1") & "',"
        '    SQl = SQl + " QAttachFile2 = N'" & DBDataSet1.Tables("F_ManufInModSheet").Rows(0).Item("QAttachFile2") & "',"

        '    SQl = SQl + " CTSts = '" & "0" & "',"
        '    SQl = SQl + " Contact = '" & "1" & "',"
        '    SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        '    SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        '    SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
        '    SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
        '    OleDBCommand1.Connection = OleDbConnection1
        '    OleDBCommand1.CommandText = SQl
        '    OleDbConnection1.Open()
        '    OleDBCommand1.ExecuteNonQuery()
        '    OleDbConnection1.Close()
        'End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyManufInData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyManufInData(ByVal pFun As String, ByVal pSts As Integer, ByVal pContact As Integer)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_ManufInSheet Set "
        SQl = SQl + " CTSts = '" & CStr(pSts) & "',"
        SQl = SQl + " Contact = '" & CStr(pContact) & "',"
        SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & DOFormNo.Text & "'"
        SQl = SQl + "   And FormSno =  '" & DOFormSno.Text & "'"
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
                SQl = SQl + " '" + DMapNo.Text + "', "
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
                    SQl = SQl + " MapNo = '" & DMapNo.Text & "',"
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
    '**     檢查上傳檔案
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
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
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
    '**     列印表單
    '**
    '*****************************************************************
    Private Sub BPrint_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPrint.Click
        Dim URL As String = ""
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        SQL = "Select ViewURL From V_WaitHandle_01 "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "URL")
        If DBDataSet1.Tables("URL").Rows.Count > 0 Then
            URL = DBDataSet1.Tables("URL").Rows(0).Item("ViewURL")
        End If
        'DB連結關閉
        OleDbConnection1.Close()

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenPrintSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     流程模擬表單
    '**
    '*****************************************************************
    Private Sub BFlow_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFlow.Click
        Dim URL As String = ""
        Dim wLevel As String = DLevel.Text '難易度

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel

        'Call JavaScript
        Dim scriptString As String = "<script language=JavaScript> "
        scriptString = scriptString & "OpenSimulationSheet('" & URL & "'); "
        scriptString = scriptString & "</script>"
        If (Not Me.IsStartupScriptRegistered("Startup")) Then
            Me.RegisterStartupScript("Startup", scriptString)
        End If

    End Sub

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
        '部門
        If InputCheck = 0 Then
            If FindFieldInf("Division") = 1 Then
                If DDivision.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '擔當
        If InputCheck = 0 Then
            If FindFieldInf("Person") = 1 Then
                If DPerson.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Target
        If InputCheck = 0 Then
            If FindFieldInf("Target") = 1 Then
                If DTarget.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'Slider Code
        If InputCheck = 0 Then
            If FindFieldInf("SliderCode") = 1 Then
                If DSliderCode.Text = "" Then InputCheck = 1
            End If
        End If
        '難易度
        If InputCheck = 0 Then
            If FindFieldInf("Level") = 1 Then
                If DLevel.Text = "" Then InputCheck = 1
            End If
        End If
        '圖號
        If InputCheck = 0 Then
            If FindFieldInf("MapNo") = 1 Then
                If DMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        'OFormNo
        If InputCheck = 0 Then
            If FindFieldInf("OFormNo") = 1 Then
                If DOFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        'OFormSno
        If InputCheck = 0 Then
            If FindFieldInf("OFormSno") = 1 Then
                If DOFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        'NFormNo
        If InputCheck = 0 Then
            If FindFieldInf("NFormNo") = 1 Then
                If DNFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        'NFormSno
        If InputCheck = 0 Then
            If FindFieldInf("NFormSno") = 1 Then
                If DNFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        'Content
        If InputCheck = 0 Then
            If FindFieldInf("Content") = 1 Then
                If DContent.Text = "" Then InputCheck = 1
            End If
        End If
        'Reason
        If InputCheck = 0 Then
            If FindFieldInf("Reason") = 1 Then
                If DReason.Text = "" Then InputCheck = 1
            End If
        End If
        '附件
        If InputCheck = 0 Then
            If FindFieldInf("AttachFile") = 1 Then
                If DAttachFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

    End Function
End Class
