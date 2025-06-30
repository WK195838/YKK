Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class MapSheetMod_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DMapSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DMapSheet3 As System.Web.UI.WebControls.Image
    Protected WithEvents DMapSheet4 As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DMakeMap As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMapFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSellVendor As System.Web.UI.WebControls.TextBox
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DRefAttach As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BOriMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DOriMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBefMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DBefFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBefFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DModReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DModContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOriFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSample As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BPrint As System.Web.UI.WebControls.ImageButton
    Protected WithEvents LRefAttach As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LOriMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LBefMap As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BFlow As System.Web.UI.WebControls.ImageButton
    Protected WithEvents DManufType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBuyer As System.Web.UI.WebControls.TextBox
    Protected WithEvents DCpsc As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DOPReady As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOPReadyDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSuppiler As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LPdfFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DPdfFile As System.Web.UI.HtmlControls.HtmlInputFile

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

        Response.Cookies("OriMapNo").Value = ""     '原始圖號, MapPicker使用
        Response.Cookies("BefMapNo").Value = ""     '原始圖號, MapPicker使用

        Response.Cookies("PGM").Value = "MapModSheet_01.aspx"
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
        'Check參考附件
        If DRefAttach.Visible Then
            If DRefAttach.PostedFile.FileName <> "" Then
                Message = "參考附件"
            End If
        End If
        'Check圖檔
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "圖檔"
                Else
                    Message = Message & ", " & "圖檔"
                End If
            End If
        End If
        'Check圖檔
        If DPdfFile.Visible Then
            If DPdfFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "圖檔"
                Else
                    Message = Message & ", " & "圖檔"
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("MapModFilePath")
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_MapModSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_MapModSheet")
        If DBDataSet1.Tables("F_MapModSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Date")                   '日期

            DBuyer.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Buyer")       'Buyer
            DSellVendor.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("SellVendor")       '委託廠商
            SetFieldData("Division", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Sample", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Sample"))      'Sample
            SetFieldData("Person", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Person"))      '擔當
            SetFieldData("Cpsc", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Cpsc"))          'Cpsc
            DOriMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo")           '原始圖號
            DOriFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo")         '原始表單號碼
            DOriFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno")       '原始單號
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno") > 0 Then          '原圖號
                LOriMap.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormNo") & _
                                                      "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("OriFormSno"))
            Else
                LOriMap.Visible = False
            End If

            SetFieldData("Level", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Level"))        '難易度
            DBefMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo")           '前圖號
            DBefFormNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo")         '前表單號碼
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then
                DBefFormSno.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno")       '前單號
            Else
                DBefFormSno.Text = ""
            End If
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefMapNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") <> "" And _
               DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno") > 0 Then            '原圖號
                LBefMap.NavigateUrl = "MapModSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormNo") & _
                                                         "&pFormSno=" & CStr(DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("BefFormSno"))
            Else
                LBefMap.Visible = False
            End If

            SetFieldData("ModReasonCode", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonCode"))  '原因類別代碼
            DModReasonDesc.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModReasonDesc")           '原因說明
            DModContent.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ModContent")                 '修改細節說明
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach") <> "" Then              '參考附件
                LRefAttach.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("RefAttach")
            Else
                LRefAttach.Visible = False
            End If
            DMapNo.Text = DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapNo")                     '圖號
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile") <> "" Then                   '圖檔
                LMapFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MapFile")
            Else
                LMapFile.Visible = False
            End If
            If DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile") <> "" Then                   '圖檔
                LPdfFile.NavigateUrl = Path & DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("PdfFile")
            Else
                LPdfFile.Visible = False
            End If

            SetFieldData("MakeMap", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("MakeMap"))        '製圖者
            SetFieldData("ManufType", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("ManufType"))    '內外注
            SetFieldData("Suppiler", DBDataSet1.Tables("F_MapModSheet").Rows(0).Item("Suppiler"))    '外注商

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
                DMapSheet3.Visible = True   '圖面Sheet
                DMapSheet4.Visible = True   '說明Sheet
                DDelivery.Visible = True    '交期Sheet

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
                    Top = 936
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 824
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
                    '20190430 JESSICA 本田
                    DOPReady.BackColor = Color.LightGray

                    'DOPReady.BackColor = Color.GreenYellow
                    'ShowRequiredFieldValidator("DOPReadyRqd", "DOPReady", "異常：需閱讀工程履歷")
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
                LOriMap.Visible = True         '原圖號
                LBefMap.Visible = True         '前圖號
                LRefAttach.Visible = True      '附件
                LMapFile.Visible = True        '圖檔
                LPdfFile.Visible = True        '圖檔
                LBefOP.Visible = True          '工程履歷
                '按鈕位置
                BNG1.Style.Add("Top", Top)     'NG1按鈕
                BNG2.Style.Add("Top", Top)     'NG2按鈕
                BSAVE.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
            Top = 640
            'Sheet隱藏
            'DMapSheet3.Visible = False  '圖面Sheet
            DMapSheet4.Visible = False  '說明Sheet
            DDelivery.Visible = False   '交期Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DMapNo.Visible = False          '圖號
            DMapFile.Visible = False        '圖檔
            DPdfFile.Visible = False        '圖檔
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
            LOriMap.Visible = False         '原圖號
            LBefMap.Visible = False         '前圖號
            LRefAttach.Visible = False      '附件
            LMapFile.Visible = False        '圖檔
            LPdfFile.Visible = False        '圖檔
            LBefOP.Visible = False          '工程履歷
            '按鈕位置
            BNG1.Style.Add("Top", Top)     'NG1按鈕
            BNG2.Style.Add("Top", Top)     'NG2按鈕
            BSAVE.Style.Add("Top", Top)    '儲存按鈕
            BOK.Style.Add("Top", Top)      'OK按鈕
            DFormSno.Style.Add("Top", Top) '單號
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
        '日期選擇
        BDate.Attributes("onclick") = "calendarPicker('Form1.DDate');"
        '原始圖號選擇
        BOriMapNo.Attributes("onclick") = "MapModPicker('Form1.DOriMapNo');"
        '前圖號選擇
        BBefMapNo.Attributes("onclick") = "MapModPicker('Form1.DBefMapNo');"
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
                    Top = 824
                Else
                    If DDelay.Visible = True Then
                        Top = 936
                    Else
                        Top = 824
                    End If
                End If
            End If
        Else
            Top = 640
        End If
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

        'Buyer
        Select Case FindFieldInf("Buyer")
            Case 0  '顯示
                DBuyer.BackColor = Color.LightGray
                DBuyer.ReadOnly = True
                DBuyer.Visible = True
            Case 1  '修改+檢查
                DBuyer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBuyerRqd", "DBuyer", "異常：需輸入Buyer")
                DBuyer.Visible = True
            Case 2  '修改
                DBuyer.BackColor = Color.Yellow
                DBuyer.Visible = True
            Case Else   '隱藏
                DBuyer.Visible = False
        End Select
        If pPost = "New" Then DBuyer.Text = ""

        '委託廠商
        Select Case FindFieldInf("SellVendor")
            Case 0  '顯示
                DSellVendor.BackColor = Color.LightGray
                DSellVendor.ReadOnly = True
                DSellVendor.Visible = True
            Case 1  '修改+檢查
                DSellVendor.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSellVendorRqd", "DSellVendor", "異常：需輸入委託廠商")
                DSellVendor.Visible = True
            Case 2  '修改
                DSellVendor.BackColor = Color.Yellow
                DSellVendor.Visible = True
            Case Else   '隱藏
                DSellVendor.Visible = False
        End Select
        If pPost = "New" Then DSellVendor.Text = ""

        '部門
        Select Case FindFieldInf("Division")
            Case 0  '顯示
                DDivision.BackColor = Color.LightGray
                'DDivision.Enabled = False
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

        'sample
        Select Case FindFieldInf("Sample")
            Case 0  '顯示
                DSample.BackColor = Color.LightGray
                'DSample.Enabled = False
                DSample.Visible = True
            Case 1  '修改+檢查
                DSample.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSampleRqd", "DSample", "異常：需輸入是否附樣品")
                DSample.Visible = True
            Case 2  '修改
                DSample.BackColor = Color.Yellow
                DSample.Visible = True
            Case Else   '隱藏
                DSample.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Sample", "ZZZZZZ")

        '擔當
        Select Case FindFieldInf("Person")
            Case 0  '顯示
                DPerson.BackColor = Color.LightGray
                'DPerson.Enabled = False
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

        'CPSC
        Select Case FindFieldInf("Cpsc")
            Case 0  '顯示
                DCpsc.BackColor = Color.LightGray
                'DCPSC.Enabled = False
                DCpsc.Visible = True
            Case 1  '修改+檢查
                DCpsc.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DCpscRqd", "DCpsc", "異常：需輸入CPSC")
                DCpsc.Visible = True
            Case 2  '修改
                DCpsc.BackColor = Color.Yellow
                DCpsc.Visible = True
            Case Else   '隱藏
                DCpsc.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Cpsc", "ZZZZZZ")

        '原始圖號
        Select Case FindFieldInf("OriMapNo")
            Case 0  '顯示
                DOriMapNo.BackColor = Color.LightGray
                DOriMapNo.ReadOnly = True
                DOriMapNo.Visible = True
                BOriMapNo.Visible = False
            Case 1  '修改+檢查
                DOriMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOriMapNoRqd", "DOriMapNo", "異常：需輸入原始圖號")
                DOriMapNo.Visible = True
                BOriMapNo.Visible = True
            Case 2  '修改
                DOriMapNo.BackColor = Color.Yellow
                DOriMapNo.Visible = True
                BOriMapNo.Visible = True
            Case Else   '隱藏
                DOriMapNo.Visible = False
                BOriMapNo.Visible = False
        End Select
        If pPost = "New" Then DOriMapNo.Text = ""

        '原始表單號碼
        Select Case FindFieldInf("OriFormNo")
            Case 0  '顯示
                DOriFormNo.BackColor = Color.LightGray
                DOriFormNo.ReadOnly = True
                DOriFormNo.Visible = True
            Case 1  '修改+檢查
                DOriFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOriFormNoRqd", "DOriFormNo", "異常：需輸入原始表單號碼")
                DOriFormNo.Visible = True
            Case 2  '修改
                DOriFormNo.BackColor = Color.Yellow
                DOriFormNo.Visible = True
            Case Else   '隱藏
                DOriFormNo.Visible = False
        End Select
        If pPost = "New" Then DOriFormNo.Text = ""

        '原始單號
        Select Case FindFieldInf("OriFormSno")
            Case 0  '顯示
                DOriFormSno.BackColor = Color.LightGray
                DOriFormSno.ReadOnly = True
                DOriFormSno.Visible = True
            Case 1  '修改+檢查
                DOriFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DOriFormSnoRqd", "DOriFormSno", "異常：需輸入原始單號")
                DOriFormSno.Visible = True
            Case 2  '修改
                DOriFormSno.BackColor = Color.Yellow
                DOriFormSno.Visible = True
            Case Else   '隱藏
                DOriFormSno.Visible = False
        End Select
        If pPost = "New" Then DOriFormSno.Text = ""

        '難易度
        Select Case FindFieldInf("Level")
            Case 0  '顯示
                DLevel.BackColor = Color.LightGray
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
        If pPost = "New" Then SetFieldData("Level", "ZZZZZZ")

        '內外製
        Select Case FindFieldInf("ManufType")
            Case 0  '顯示
                DManufType.BackColor = Color.LightGray
                DManufType.Visible = True
            Case 1  '修改+檢查
                DManufType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DManufTypeRqd", "DManufType", "異常：需輸入內製或外注")
                DManufType.Visible = True
            Case 2  '修改
                DManufType.BackColor = Color.Yellow
                DManufType.Visible = True
            Case Else   '隱藏
                DManufType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ManufType", "ZZZZZZ")

        '外注商
        Select Case FindFieldInf("Suppiler")
            Case 0  '顯示
                DSuppiler.BackColor = Color.LightGray
                DSuppiler.Visible = True
            Case 1  '修改+檢查
                DSuppiler.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSuppilerRqd", "DSuppiler", "異常：需輸入外注商")
                DSuppiler.Visible = True
            Case 2  '修改
                DSuppiler.BackColor = Color.Yellow
                DSuppiler.Visible = True
            Case Else   '隱藏
                DSuppiler.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Suppiler", "ZZZZZZ")

        '前圖號
        Select Case FindFieldInf("BefMapNo")
            Case 0  '顯示
                DBefMapNo.BackColor = Color.LightGray
                DBefMapNo.ReadOnly = True
                DBefMapNo.Visible = True
                BBefMapNo.Visible = False
            Case 1  '修改+檢查
                DBefMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBefMapNoRqd", "DBefMapNo", "異常：需輸入前圖號")
                DBefMapNo.Visible = True
                BBefMapNo.Visible = True
            Case 2  '修改
                DBefMapNo.BackColor = Color.Yellow
                DBefMapNo.Visible = True
                BBefMapNo.Visible = True
            Case Else   '隱藏
                DBefMapNo.Visible = False
                BBefMapNo.Visible = False
        End Select
        If pPost = "New" Then DBefMapNo.Text = ""

        '前圖表單號碼
        Select Case FindFieldInf("BefFormNo")
            Case 0  '顯示
                DBefFormNo.BackColor = Color.LightGray
                DBefFormNo.ReadOnly = True
                DBefFormNo.Visible = True
            Case 1  '修改+檢查
                DBefFormNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBefFormNoRqd", "DBefFormNo", "異常：需輸入前圖表單號碼")
                DBefFormNo.Visible = True
            Case 2  '修改
                DBefFormNo.BackColor = Color.Yellow
                DBefFormNo.Visible = True
            Case Else   '隱藏
                DBefFormNo.Visible = False
        End Select
        If pPost = "New" Then DBefFormNo.Text = ""

        '前圖單號
        Select Case FindFieldInf("BefFormSno")
            Case 0  '顯示
                DBefFormSno.BackColor = Color.LightGray
                DBefFormSno.ReadOnly = True
                DBefFormSno.Visible = True
            Case 1  '修改+檢查
                DBefFormSno.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBefFormSnoRqd", "DBefFormSno", "異常：需輸入前圖單號")
                DBefFormSno.Visible = True
            Case 2  '修改
                DBefFormSno.BackColor = Color.Yellow
                DBefFormSno.Visible = True
            Case Else   '隱藏
                DBefFormSno.Visible = False
        End Select
        If pPost = "New" Then DBefFormSno.Text = ""

        '原因類別
        Select Case FindFieldInf("ModReasonCode")
            Case 0  '顯示
                DModReasonCode.BackColor = Color.LightGray
                'DModReasonCode.Enabled = False
                DModReasonCode.Visible = True
            Case 1  '修改+檢查
                DModReasonCode.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DModReasonCodeRqd", "DModReasonCode", "異常：需輸入原因類別")
                DModReasonCode.Visible = True
            Case 2  '修改
                DModReasonCode.BackColor = Color.Yellow
                DModReasonCode.Visible = True
            Case Else   '隱藏
                DModReasonCode.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("ModReasonCode", "ZZZZZZ")

        '修改說明
        Select Case FindFieldInf("ModReasonDesc")
            Case 0  '顯示
                DModReasonDesc.BackColor = Color.LightGray
                DModReasonDesc.ReadOnly = True
                DModReasonDesc.Visible = True
            Case 1  '修改+檢查
                DModReasonDesc.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DModReasonDescRqd", "DModReasonDesc", "異常：需輸入修改說明")
                DModReasonDesc.Visible = True
            Case 2  '修改
                DModReasonDesc.BackColor = Color.Yellow
                DModReasonDesc.Visible = True
            Case Else   '隱藏
                DModReasonDesc.Visible = False
        End Select
        If pPost = "New" Then DModReasonDesc.Text = ""

        '修改細節內容
        Select Case FindFieldInf("ModContent")
            Case 0  '顯示
                DModContent.BackColor = Color.LightGray
                DModContent.ReadOnly = True
                DModContent.Visible = True
            Case 1  '修改+檢查
                DModContent.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DModContentRqd", "DModContent", "異常：需輸入修改細節內容")
                DModContent.Visible = True
            Case 2  '修改
                DModContent.BackColor = Color.Yellow
                DModContent.Visible = True
            Case Else   '隱藏
                DModContent.Visible = False
        End Select
        If pPost = "New" Then DModContent.Text = ""

        '附件
        Select Case FindFieldInf("RefAttach")
            Case 0  '顯示
                DRefAttach.Visible = False
                DRefAttach.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DRefAttachRqd", "DRefAttach", "異常：需輸入附件")
                DRefAttach.Visible = True
                DRefAttach.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DRefAttach.Visible = True
                DRefAttach.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DRefAttach.Visible = False
        End Select

        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.LightGray
                DMapNo.ReadOnly = True
                DMapNo.Visible = True
            Case 1  '修改+檢查
                DMapNo.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMapNoRqd", "DMapNo", "異常：需輸入圖號")
                DMapNo.Visible = True
            Case 2  '修改
                DMapNo.BackColor = Color.Yellow
                DMapNo.Visible = True
            Case Else   '隱藏
                DMapNo.Visible = False
        End Select
        If pPost = "New" Then DMapNo.Text = ""

        '圖檔
        Select Case FindFieldInf("MapFile")
            Case 0  '顯示
                DMapFile.Visible = False
                DMapFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
                DPdfFile.Visible = False
                DPdfFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DMapFileRqd", "DMapFile", "異常：需輸入圖檔")
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
                ShowRequiredFieldValidator("DPdfFileRqd", "DPdfFile", "異常：需輸入圖檔")
                DPdfFile.Visible = True
                DPdfFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DMapFile.Visible = True
                DMapFile.Style.Add("BACKGROUND-COLOR", "Yellow")
                DPdfFile.Visible = True
                DPdfFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DMapFile.Visible = False
                DPdfFile.Visible = False
        End Select

        '製圖者
        Select Case FindFieldInf("MakeMap")
            Case 0  '顯示
                DMakeMap.BackColor = Color.LightGray
                'DMakeMap.Enabled = False
                DMakeMap.Visible = True
            Case 1  '修改+檢查
                DMakeMap.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DMakeMapRqd", "DMakeMap", "異常：需輸入製圖者")
                DMakeMap.Visible = True
            Case 2  '修改
                DMakeMap.BackColor = Color.Yellow
                DMakeMap.Visible = True
            Case Else   '隱藏
                DMakeMap.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("MakeMap", "ZZZZZZ")

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

        'Sample
        If pFieldName = "Sample" Then
            DSample.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSample.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and DKey='SAMPLE' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSample.Items.Add(ListItem1)
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
        'CPSC
        If pFieldName = "Cpsc" Then
            DCpsc.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DCpsc.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and DKey='CPSC' Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DCpsc.Items.Add(ListItem1)
                Next
            End If
        End If


        '難易度
        If pFieldName = "Level" Then
            DLevel.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DLevel.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='007' and (DKey like 'In%' or DKey like 'Out%' or DKey = '')  Order by DKey, Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DLevel.Items.Add(ListItem1)
                Next
            End If
        End If

        '內外製
        If pFieldName = "ManufType" Then
            DManufType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DManufType.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='150' and DKey='MANUFTYPE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DManufType.Items.Add(ListItem1)
                Next
            End If
        End If

        'Suppiler
        If pFieldName = "Suppiler" Then
            DSuppiler.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSuppiler.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='700' and DKey='SUPPLIER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSuppiler.Items.Add(ListItem1)
                Next
            End If
        End If

        '製圖者
        If pFieldName = "MakeMap" Then
            DMakeMap.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DMakeMap.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='100' and DKey='MAKEMAP' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DMakeMap.Items.Add(ListItem1)
                Next
            End If
        End If

        '修改原因代碼
        If pFieldName = "ModReasonCode" Then
            DModReasonCode.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DModReasonCode.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='150' and DKey='MODREASONCODE' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DModReasonCode.Items.Add(ListItem1)
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
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    DReason.Text = pName
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    DReason.Text = DBTable1.Rows(i).Item("Data")
                Next
            End If
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

                'Check上傳參考附件及格式
                If ErrCode = 0 Then
                    If DRefAttach.Visible Then
                        If DRefAttach.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DRefAttach)
                        End If
                    End If
                End If
                'Check開發圖檔Size及格式
                If ErrCode = 0 Then
                    If DMapFile.Visible Then
                        If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DMapFile)
                        End If
                    End If
                End If
                'Check開發圖檔Size及格式
                If ErrCode = 0 Then
                    If DPdfFile.Visible Then
                        If DPdfFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                            ErrCode = UPFileIsNormal(DPdfFile)
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

        'Check上傳參考附件及格式
        If ErrCode = 0 Then
            If DRefAttach.Visible Then
                If DRefAttach.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DRefAttach)
                End If
            End If
        End If

        'Check開發圖檔Size及格式
        If ErrCode = 0 Then
            If DMapFile.Visible Then
                If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DMapFile)
                End If
            End If
        End If
        'Check開發圖檔Size及格式
        If ErrCode = 0 Then
            If DPdfFile.Visible Then
                If DPdfFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DPdfFile)
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

                ErrCode = oCommon.CommissionNo("000002", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No

                If ErrCode <> 0 Then
                    ErrCode = 9060
                End If
            End If
        End If

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim MultiJob As Integer = 0         '多人核定
            Dim wLevel As String = ""           '樣品難易度
            If DSample.SelectedValue = "YES" Then wLevel = "Z"

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
                    Dim DBDataSet1 As New DataSet
                    Dim wAllocateID As String = ""
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

                    OleDbConnection1.Open()
                    '製圖者
                    SQL = "Select UserID From M_Users "
                    SQL = SQL & " Where Active = 1 "
                    SQL = SQL & "   And UserName =  '" & DMakeMap.SelectedValue & "'"
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "M_Users")
                    If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '  日期:2008/10/24, 2010/3(再修改)
                    '  對應:企劃問題-->企劃起單,  EA技術問題-->EA起單
                    If wStep = 130 Then     '圖面修改完成EA管理窗口
                        DBDataSet1.Clear()
                        SQL = "Select CreateUser From F_MapSheet "
                        SQL = SQL + " Where FormNo  =  '" & DOriFormNo.Text & "'"
                        SQL = SQL + "   And FormSno =  '" & DOriFormSno.Text & "'"
                        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                        DBAdapter2.Fill(DBDataSet1, "F_MapSheet")
                        If DBDataSet1.Tables("F_MapSheet").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("CreateUser")
                    End If
                    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    OleDbConnection1.Close()


                    RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                                  pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)
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
                    If wFormSno = 0 And wStep < 3 Then      '判斷是否起單
                        If NewFormSno <> 0 Then
                            AppendData(pFun, NewFormSno)  'Insert
                            AddCommissionNo(wFormNo, NewFormSno)
                            UpdateOriMapData("START")  'Update Ori Map Data
                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
                        If pNextStep = 999 Then     '工程結束嗎?
                            If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                            If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                            If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                            UpdateOriMapData("END")  'Update Ori Map Data
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
    '**(UpdateOriMapData)
    '**     修改原表單資料狀態
    '**
    '*****************************************************************
    Sub UpdateOriMapData(ByVal pFun As String)
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        If pFun = "START" Then
            SQl = "Update F_MapSheet Set "
            SQl = SQl + " UPDSts = '1', ModMap = '1', "
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "', "
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & DOriFormNo.Text & "'"
            SQl = SQl + "   And FormSno =  '" & DOriFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            If DBefFormNo.Text <> "" And DBefFormSno.Text <> "" Then
                SQl = "Update F_MapModSheet Set "
                SQl = SQl + " UPDSts = '1', ModMap = '1', "
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "', "
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & DBefFormNo.Text & "'"
                SQl = SQl + "   And FormSno =  '" & DBefFormSno.Text & "'"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDbConnection1.Open()
                OleDBCommand1.ExecuteNonQuery()
                OleDbConnection1.Close()
            End If
        Else
            SQl = "Update F_MapSheet Set "
            SQl = SQl + " UPDSts = '0', ModMap = '1', "
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "', "
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & DOriFormNo.Text & "'"
            SQl = SQl + "   And FormSno =  '" & DOriFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQl
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            If DBefFormNo.Text <> "" And DBefFormSno.Text <> "" Then
                SQl = "Update F_MapModSheet Set "
                SQl = SQl + " UPDSts = '0', ModMap = '1', "
                SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "', "
                SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
                SQl = SQl + " Where FormNo  =  '" & DBefFormNo.Text & "'"
                SQl = SQl + "   And FormSno =  '" & DBefFormSno.Text & "'"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = SQl
                OleDbConnection1.Open()
                OleDBCommand1.ExecuteNonQuery()
                OleDbConnection1.Close()
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapModFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_MapModSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSNo, No, "         '1~5
        SQl = SQl + "Date, Buyer, SellVendor, Division, Sample, Person, Cpsc, "       '6~10
        SQl = SQl + "OriMapNo, OriFormNo, OriFormSNo, Level, "          '11~14
        SQl = SQl + "BefMapNo, BefFormNo, BefFormSNo, "                 '15~17
        SQl = SQl + "ModReasonCode, ModReason, ModReasonDesc, ModContent, RefAttach, "   '18~23
        SQl = SQl + "MapNo, MapFile,PdfFile, MakeMap, UPDSts, ModMap, ManufType, Suppiler, "                         '24~26
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "   '27~30
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "
        SQl = SQl + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        SQl = SQl + " '" + NowDateTime + "', "        '結案日
        SQl = SQl + " '000002', "                     '表單代號
        SQl = SQl + " '" + CStr(NewFormSno) + "', "   '表單流水號
        SQl = SQl + " '" + YKK.ReplaceString(DNo.Text) + "', "           'NO
        SQl = SQl + " '" + DDate.Text + "', "         '日期
        SQl = SQl + " N'" + DBuyer.Text + "', "        'Buyer
        SQl = SQl + " N'" + YKK.ReplaceString(DSellVendor.Text) + "', "         '委託廠商
        SQl = SQl + " '" + DDivision.SelectedValue + "', "  '部門
        SQl = SQl + " '" + DSample.SelectedValue + "', "  'Sample
        SQl = SQl + " '" + DPerson.SelectedValue + "', "    '擔當
        SQl = SQl + " '" + DCpsc.SelectedValue + "', "    'Cpsc
        SQl = SQl + " '" + DOriMapNo.Text + "', "     '原始圖號
        SQl = SQl + " '" + DOriFormNo.Text + "', "    '原始表單號碼
        SQl = SQl + " '" + DOriFormSno.Text + "', "   '原始單號
        SQl = SQl + " '" + DLevel.SelectedValue + "', "        '難易度
        SQl = SQl + " '" + DBefMapNo.Text + "', "     '前圖號
        SQl = SQl + " '" + DBefFormNo.Text + "', "    '前表單號碼
        SQl = SQl + " '" + DBefFormSno.Text + "', "   '前單號
        SQl = SQl + " '" + DModReasonCode.SelectedValue + "', "   '原因類別代碼
        SQl = SQl + " '" + "" + "', "      '類別內容
        SQl = SQl + " N'" + YKK.ReplaceString(DModReasonDesc.Text) + "', "  '說明
        SQl = SQl + " N'" + YKK.ReplaceString(DModContent.Text) + "', "     '修改細節說明
        FileName = ""
        If DRefAttach.Visible Then
            If DRefAttach.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "RAttach" & "-" & UploadDateTime & "-" & Right(DRefAttach.PostedFile.FileName, InStr(StrReverse(DRefAttach.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "RAttach" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefAttach.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳參考附件
                    DRefAttach.PostedFile.SaveAs(Path & FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "                 '參考附件
        SQl = SQl + " '" + YKK.ReplaceString(DMapNo.Text) + "', "              '圖號
        FileName = ""
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "                 '圖檔

        FileName = ""
        If DPdfFile.Visible Then
            If DPdfFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DPdfFile.PostedFile.FileName, InStr(StrReverse(DPdfFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(NewFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPdfFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DPdfFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "                 '圖檔

        SQl = SQl + " '" + DMakeMap.SelectedValue + "', "         '製圖者
        SQl = SQl + " '" + "0" + "', "          'UPDSts
        SQl = SQl + " '" + "0" + "', "          'ModMap
        SQl = SQl + " '" + DManufType.SelectedValue + "', "       '內外注
        SQl = SQl + " N'" + DSuppiler.SelectedValue + "', "           '外注商

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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("MapModFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_MapModSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & YKK.ReplaceString(DNo.Text) & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"
        SQl = SQl + " Buyer = N'" & DBuyer.Text & "',"
        SQl = SQl + " SellVendor = N'" & YKK.ReplaceString(DSellVendor.Text) & "',"
        SQl = SQl + " Division = '" & DDivision.SelectedValue & "',"
        SQl = SQl + " Sample = '" & DSample.SelectedValue & "',"
        SQl = SQl + " Person = '" & DPerson.SelectedValue & "',"
        SQl = SQl + " Cpsc = '" & DCpsc.SelectedValue & "',"
        SQl = SQl + " OriMapNo = '" & DOriMapNo.Text & "',"
        SQl = SQl + " OriFormNo = '" & DOriFormNo.Text & "',"
        SQl = SQl + " OriFormSno = '" & DOriFormSno.Text & "',"
        SQl = SQl + " Level = '" & DLevel.SelectedValue & "',"
        SQl = SQl + " BefMapNo = '" & DBefMapNo.Text & "',"
        SQl = SQl + " BefFormNo = '" & DBefFormNo.Text & "',"
        SQl = SQl + " BefFormSno = '" & DBefFormSno.Text & "',"
        SQl = SQl + " ModReasonCode = '" & DModReasonCode.SelectedValue & "',"
        SQl = SQl + " ModReason = '" & "" & "',"
        SQl = SQl + " ModReasonDesc = N'" & YKK.ReplaceString(DModReasonDesc.Text) & "',"
        SQl = SQl + " ModContent = N'" & YKK.ReplaceString(DModContent.Text) & "',"
        If DRefAttach.Visible Then
            If DRefAttach.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "RMap" & "-" & UploadDateTime & "-" & Right(DRefAttach.PostedFile.FileName, InStr(StrReverse(DRefAttach.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "RMap" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DRefAttach.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳參考附件
                    DRefAttach.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " RefAttach = N'" & FileName & "',"
            End If
        End If
        SQl = SQl + " MapNo = '" & YKK.ReplaceString(DMapNo.Text) & "',"
        If DMapFile.Visible Then
            If DMapFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DMapFile.PostedFile.FileName, InStr(StrReverse(DMapFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DMapFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DMapFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " MapFile = N'" & FileName & "',"
            End If
        End If

        If DPdfFile.Visible Then
            If DPdfFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                '*** IE8對應-Start 2011/1/4
                'FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & Right(DPdfFile.PostedFile.FileName, InStr(StrReverse(DPdfFile.PostedFile.FileName), "\") - 1)
                FileName = CStr(wFormSno) & "-" & "Map" & "-" & UploadDateTime & "-" & IO.Path.GetFileName(DPdfFile.PostedFile.FileName)
                '*** IE8對應-End
                Try    '上傳圖檔
                    DPdfFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " PdfFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " MakeMap = '" & DMakeMap.SelectedValue & "',"
        SQl = SQl + " ManufType = '" & DManufType.SelectedValue & "',"  '內外注
        SQl = SQl + " Suppiler = N'" & DSuppiler.SelectedValue & "',"

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
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".ai", ".xls", ".doc", ".ppt", ".xlsx", ".docx", ".pptx"}   '定義允許的檔案格式
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
        Dim SQL As String = ""
        Dim wLevel As String = ""           '樣品難易度
        Dim wMakeMap As String = ""         '製圖者
        If DSample.SelectedValue = "YES" Then wLevel = "Z"

        If DMakeMap.SelectedValue <> "" Then
            Dim DBDataSet1 As New DataSet
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            SQL = "Select UserID From M_Users "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And UserName =  '" & DMakeMap.SelectedValue & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Users")
            If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then wMakeMap = DBDataSet1.Tables("M_Users").Rows(0).Item("UserID")

            OleDbConnection1.Close()
        End If

        URL = "Simulation.aspx?pFormNo=" & wFormNo & _
                              "&pFormSno=" & wFormSno & _
                              "&pStep=" & wStep & _
                              "&pLevel=" & wLevel & _
                              "&pAllocateID=" & wMakeMap

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
        'Buyer
        If InputCheck = 0 Then
            If FindFieldInf("Buyer") = 1 Then
                If DBuyer.Text = "" Then InputCheck = 1
            End If
        End If
        '委託廠商
        If InputCheck = 0 Then
            If FindFieldInf("SellVendor") = 1 Then
                If DSellVendor.Text = "" Then InputCheck = 1
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
        'Sample
        If InputCheck = 0 Then
            If FindFieldInf("Sample") = 1 Then
                If DSample.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        'CPSC
        If InputCheck = 0 Then
            If FindFieldInf("Cpsc") = 1 Then
                If DCpsc.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '原始圖號
        If InputCheck = 0 Then
            If FindFieldInf("OriMapNo") = 1 Then
                If DOriMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        '原始表單號碼
        If InputCheck = 0 Then
            If FindFieldInf("OriFormNo") = 1 Then
                If DOriFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        '原始單號
        If InputCheck = 0 Then
            If FindFieldInf("OriFormSno") = 1 Then
                If DOriFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        '難易度
        If InputCheck = 0 Then
            If FindFieldInf("Level") = 1 Then
                If DLevel.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '內外製
        If InputCheck = 0 Then
            If FindFieldInf("ManufType") = 1 Then
                If DManufType.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '外注商
        If InputCheck = 0 Then
            If FindFieldInf("Suppiler") = 1 Then
                If DSuppiler.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '前圖號
        If InputCheck = 0 Then
            If FindFieldInf("BefMapNo") = 1 Then
                If DBefMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        '前圖表單號碼
        If InputCheck = 0 Then
            If FindFieldInf("BefFormNo") = 1 Then
                If DBefFormNo.Text = "" Then InputCheck = 1
            End If
        End If
        '前圖單號
        If InputCheck = 0 Then
            If FindFieldInf("BefFormSno") = 1 Then
                If DBefFormSno.Text = "" Then InputCheck = 1
            End If
        End If
        '原因類別
        If InputCheck = 0 Then
            If FindFieldInf("ModReasonCode") = 1 Then
                If DModReasonCode.SelectedValue = "" Then InputCheck = 1
            End If
        End If
        '修改說明
        If InputCheck = 0 Then
            If FindFieldInf("ModReasonDesc") = 1 Then
                If DModReasonDesc.Text = "" Then InputCheck = 1
            End If
        End If
        '修改細節內容
        If InputCheck = 0 Then
            If FindFieldInf("ModContent") = 1 Then
                If DModContent.Text = "" Then InputCheck = 1
            End If
        End If
        '附件
        If InputCheck = 0 Then
            If FindFieldInf("RefAttach") = 1 Then
                If DRefAttach.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '圖號
        If InputCheck = 0 Then
            If FindFieldInf("MapNo") = 1 Then
                If DMapNo.Text = "" Then InputCheck = 1
            End If
        End If
        '圖檔
        If InputCheck = 0 Then
            If FindFieldInf("MapFile") = 1 Then
                If DMapFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If
        '圖檔
        If InputCheck = 0 Then
            If FindFieldInf("PdfFile") = 1 Then
                If DPdfFile.PostedFile.FileName = "" Then InputCheck = 1
            End If
        End If

        '製圖者
        If InputCheck = 0 Then
            If FindFieldInf("MakeMap") = 1 Then
                If DMakeMap.SelectedValue = "" Then InputCheck = 1
            End If
        End If
    End Function

End Class
