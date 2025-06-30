Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class ContactOutSheet
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DDivision As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DPerson As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BDate As System.Web.UI.WebControls.Button
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSliderCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents BMMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents BOMapNo As System.Web.UI.WebControls.Button
    Protected WithEvents DMapNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DContent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents BOK As System.Web.UI.WebControls.Button
    Protected WithEvents BNG As System.Web.UI.WebControls.Button
    Protected WithEvents BSave As System.Web.UI.WebControls.Button
    Protected WithEvents DFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DDelivery As System.Web.UI.WebControls.Image
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndTime As System.Web.UI.WebControls.TextBox
    Protected WithEvents LBefOP As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents BIn As System.Web.UI.WebControls.Button
    Protected WithEvents DOFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents LAttachFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DNFormNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNFormSno As System.Web.UI.WebControls.TextBox
    Protected WithEvents LMapNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents BModify As System.Web.UI.WebControls.Button
    Protected WithEvents LNFormNo As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DContactSheet As System.Web.UI.WebControls.Image
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents DAttachFile As System.Web.UI.HtmlControls.HtmlInputFile

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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(44) As String     '各欄位
    Dim Attribute(44) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim NowDateTime As String       '現在日期時間

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
            If wFormSno > 0 And wStep > 1 Then    '判斷是否[簽核]
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
        Response.Cookies("DevNo").Value = ""        '開發No, DevNoPicker使用
        Response.Cookies("MapNo").Value = ""        '圖號, MapPicker使用
        Response.Cookies("Step").Value = Request.QueryString("pStep")  '隱藏用工程代碼
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
        Dim oUpdateFlow As Object
        oUpdateFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
        oUpdateFlow.pFormNo = wFormNo      '表單號碼
        oUpdateFlow.pFormSno = wFormSno    '表單流水號
        oUpdateFlow.pStep = wStep          '工程關卡號碼
        oUpdateFlow.pSeqNo = wSeqNo        '序號
        oUpdateFlow.UpdateFlow(Request.Cookies("UserID").Value)
    End Sub

    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ContactOutFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_ContactOutSheet "
        SQL = SQL & " Where Sts = 0 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ContactOutSheet")
        If DBDataSet1.Tables("F_ContactOutSheet").Rows.Count > 0 Then
            DNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Date")                   '日期
            SetFieldData("Division", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Division"))  '部門
            SetFieldData("Person", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Person"))      '擔當
            DSliderCode.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("SliderCode")       'Slider Code
            SetFieldData("Level", DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Level"))                '難易度
            DMapNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("MapNo")             '圖號
            If DMapNo.Text <> "" Then
                SQL = "Select FormNo, FormSno From F_MapSheet "
                SQL = SQL & " Where Sts = 2 "
                SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter3.Fill(DBDataSet1, "MapSheet")
                If DBDataSet1.Tables("MapSheet").Rows.Count > 0 Then
                    LMapNo.NavigateUrl = "MapSheet_02.aspx?pFormNo=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormNo") & _
                                                         "&pFormSno=" & DBDataSet1.Tables("MapSheet").Rows(0).Item("FormSno")
                Else
                    SQL = "Select FormNo, FormSno From F_ModMapSheet "
                    SQL = SQL & " Where Sts = 2 "
                    SQL = SQL & "   And MapNo =  '" & DMapNo.Text & "'"
                    Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter4.Fill(DBDataSet1, "ModMapSheet")
                    If DBDataSet1.Tables("ModMapSheet").Rows.Count > 0 Then
                        LMapNo.NavigateUrl = "MapSheetMod_02.aspx?pFormNo=" & DBDataSet1.Tables("ModMapSheet").Rows(0).Item("FormNo") & _
                                                                "&pFormSno=" & DBDataSet1.Tables("ModMapSheet").Rows(0).Item("FormSno")
                    End If
                End If
            Else
                LMapNo.Visible = False
            End If

            DOFormNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("OFormNo")             '圖號
            DOFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("OFormSno")             '圖號
            If DOFormNo.Text <> "" And DOFormSno.Text <> "" Then
                If DOFormNo.Text = "000003" Then
                    LOFormNo.NavigateUrl = "ManufInSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                Else
                    LOFormNo.NavigateUrl = "ManufOutSheet_02.aspx?pFormNo=" & DOFormNo.Text & "&pFormSno=" & CInt(DOFormSno.Text)
                End If
            Else
                LOFormNo.Visible = False
            End If

            DNFormNo.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormNo")             '圖號
            DNFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno")             '圖號
            If DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno") = 0 Then
                DNFormSno.Text = ""
            Else
                DNFormSno.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("NFormSno")             '圖號
            End If
            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                LNFormNo.NavigateUrl = "ModManufOutSheet_01.aspx?pFormNo=" & DNFormNo.Text & "&pFormSno=" & CInt(DNFormSno.Text) & "&pOFormNo=" & DOFormNo.Text & "&pOFormSno=" & CInt(DOFormSno.Text) & "&pStep=" & wStep
                If BModify.Visible = True Then
                    LNFormNo.Visible = False
                End If
            Else
                LNFormNo.Visible = False
            End If

            DTarget.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Target")             '圖號
            DContent.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Content")             '圖號
            DDReason.Text = DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("Reason")             '圖號
            If DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("AttachFile") <> "" Then          '圖檔1
                LAttachFile.NavigateUrl = Path & DBDataSet1.Tables("F_ContactOutSheet").Rows(0).Item("AttachFile")
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
                    SetFieldData("ReasonCode", DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode"))    '延遲理由代碼
                    If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonCode") = "" Then
                        SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                        DReasonDesc.Text = ""   '延遲其他說明
                    Else
                        DReason.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("Reason")  '延遲理由
                        DReasonDesc.Text = DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("ReasonDesc")     '延遲其他說明
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
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        If wFormSno > 0 And wStep > 1 Then    '判斷是否[簽核]
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "T_WaitHandle")
            If DBDataSet1.Tables("T_WaitHandle").Rows.Count > 0 Then
                'Sheet顯示
                DContactSheet.Visible = True   '表單Sheet-1
                DDescSheet.Visible = True       '說明Sheet
                DDelivery.Visible = True        '交期Sheet
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    DDelay.Visible = True   '延遲Sheet
                Else
                    DDelay.Visible = False  '延遲Sheet
                End If
                '欄位顯示
                DDecideDesc.Visible = True      '說明
                DBStartTime.Visible = True      '預定開始
                DBEndTime.Visible = True        '預定完成
                DAStartTime.Visible = True      '實際開始
                DAEndTime.Visible = True        '實際完成
                If DBDataSet1.Tables("T_WaitHandle").Rows(0).Item("BEndTime") < NowDateTime Then
                    DReasonCode.Visible = True     '延遲理由代碼
                    DReasonCode.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                    DReason.Visible = True         '延遲理由
                    DReason.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                    DReasonDesc.Visible = True     '延遲其他說明
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                End If
                '連結顯示---需再修改
                LMapNo.Visible = True          '圖檔
                LOFormNo.Visible = True        '原委託
                LNFormNo.Visible = True        '新委託
                LAttachFile.Visible = True     '附件
                LBefOP.Visible = True          '工程履歷
                '按鈕位置
                BNG.Style.Add("Top", Top)      'NG按鈕
                BSave.Style.Add("Top", Top)    '儲存按鈕
                BOK.Style.Add("Top", Top)      'OK按鈕
                DFormSno.Style.Add("Top", Top) '單號
            End If
        Else
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
            '連結隱藏
            LMapNo.Visible = False          '圖檔
            LOFormNo.Visible = False        '原委託
            LNFormNo.Visible = False        '新委託
            LAttachFile.Visible = False     '附件
            LBefOP.Visible = False          '工程履歷
            '按鈕位置
            BNG.Style.Add("Top", Top)       'NG按鈕
            BSave.Style.Add("Top", Top)     '儲存按鈕
            BOK.Style.Add("Top", Top)       'OK按鈕
            DFormSno.Style.Add("Top", Top)  '單號
        End If

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
                BSave.Visible = True
            Else
                BSave.Visible = False
            End If
            'NG按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun") = 1 Then
                BNG.Visible = True
            Else
                BNG.Visible = False
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
            Else
                BOK.Visible = False
            End If
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
        BIn.Attributes("onclick") = "DevNoPicker('Out','000009');"       '內製
        BOMapNo.Attributes("onclick") = "MapPicker('Ori');"             '原始圖號
        BMMapNo.Attributes("onclick") = "MapPicker('Mod');"             '修改圖號
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
        If wFormSno > 0 And wStep > 1 Then    '判斷是否[簽核]
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
                    Top = 752
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
        Dim oFieldAttribute As Object
        oFieldAttribute = Server.CreateObject("GetFieldAttb.WFField")
        oFieldAttribute.pFormNo = wFormNo      '表單號碼
        oFieldAttribute.pStep = wStep          '工程關卡號碼
        oFieldAttribute.GetFieldAttribute(FieldName, Attribute)       '欄位名,欄位屬性

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
                DNo.BackColor = Color.PeachPuff
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
                DDate.BackColor = Color.PeachPuff
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
                DDivision.BackColor = Color.PeachPuff
                DDivision.Enabled = False
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
                DPerson.BackColor = Color.PeachPuff
                DPerson.Enabled = False
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
        'Slider Code
        Select Case FindFieldInf("SliderCode")
            Case 0  '顯示
                DSliderCode.BackColor = Color.PeachPuff
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
                DLevel.BackColor = Color.PeachPuff
                DLevel.Enabled = False
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
        '圖號
        Select Case FindFieldInf("MapNo")
            Case 0  '顯示
                DMapNo.BackColor = Color.PeachPuff
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
                DOFormNo.BackColor = Color.PeachPuff
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
                DOFormSno.BackColor = Color.PeachPuff
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
                DNFormNo.BackColor = Color.PeachPuff
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
                DNFormSno.BackColor = Color.PeachPuff
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
        'Target
        Select Case FindFieldInf("Target")
            Case 0  '顯示
                DTarget.BackColor = Color.PeachPuff
                DTarget.ReadOnly = True
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
        If pPost = "New" Then DTarget.Text = ""
        'Content
        Select Case FindFieldInf("Content")
            Case 0  '顯示
                DContent.BackColor = Color.PeachPuff
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
                DDReason.BackColor = Color.PeachPuff
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
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DAttachFileRqd", "DAttachFile", "異常：需輸入附件")
                DAttachFile.Visible = True
            Case 2  '修改
                DAttachFile.Visible = True
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
        Dim i As Integer
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '部門
        If pFieldName = "Division" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='DIVISION' Order by Data "
            DDivision.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DDivision.Items.Add(ListItem1)
            Next
        End If
        '擔當
        If pFieldName = "Person" Then
            SQL = "Select * From M_Referp Where Cat='100' and DKey='PERSON' Order by Data "
            DPerson.Items.Clear()
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "M_Referp")
            DBTable1 = DBDataSet1.Tables("M_Referp")
            For i = 0 To DBTable1.Rows.Count - 1
                Dim ListItem1 As New ListItem
                ListItem1.Text = DBTable1.Rows(i).Item("Data")
                ListItem1.Value = DBTable1.Rows(i).Item("Data")
                If ListItem1.Value = pName Then ListItem1.Selected = True
                DPerson.Items.Add(ListItem1)
            Next
        End If
        '難易度
        If pFieldName = "Level" Then
            SQL = "Select * From M_Referp Where Cat='007' and DKey<>'Z' Order by DKey, Data "
            DLevel.Items.Clear()
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
        While i < 44 And Run
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
    Private Sub BSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        '上傳日期
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)

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
            Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
            Dim FileName As String
            Dim SQL As String

            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            Dim OleDBCommand1 As New OleDbCommand

            '儲存表單資料
            SQL = "Update F_ContactOutSheet Set "
            SQL = SQL + " No = '" & DNo.Text & "',"
            SQL = SQL + " Date = '" & DDate.Text & "',"
            SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
            SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
            SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
            SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
            SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
            SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
            SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
            SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
            SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
            SQL = SQL + " Target = '" & DTarget.Text & "',"
            SQL = SQL + " Content = '" & DContent.Text & "',"
            SQL = SQL + " Reason = '" & DDReason.Text & "',"

            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                    Try    '上傳圖檔
                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    SQL = SQL + " AttachFile = '" & FileName & "',"
                End If
            End If

            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '儲存交易資料
            SQL = "Update T_WaitHandle Set "
            If DReasonCode.Visible = True Then
                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQL = SQL + " Reason = '" & DReason.Text & "',"
                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
            End If
            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL + "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()
        Else
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9210 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0

        If ErrCode = 0 Then
            Dim URL As String = "MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim Message As String = ""

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

        If ErrCode = 0 Then
            Dim Run As Boolean = True           '是否執行
            Dim RepeatRun As Boolean = False    '是否重覆執行
            Dim wLevel As String = DLevel.SelectedValue     '難易度

            While Run = True
                Run = False     '執行Flag=不執行
                '--取得下一關參數設定---------
                Dim oNextGate As Object
                Dim pNextGate(10) As String
                Dim pNextStep As Integer = 0
                Dim pFlowType As Integer = 0    '0=通知
                Dim pCount As Integer
                '--取得LeadTime參數設定---------
                Dim oGetLeadTime As Object
                Dim pCTime, pStartTime, pEndTime As DateTime
                '--取得工程負荷參數設定---------
                Dim oGetLoading As Object
                Dim pLastTime As DateTime
                Dim pCount1 As Integer
                '--取得表單流水號設定---------
                Dim oGetSeqNo As Object
                '--流程資料設定---------
                Dim oNextFlow As Object
                '--郵件傳送---------
                Dim oMail As Object
                '--日程調整---------
                Dim oSchedule As Object
                '--其他SubFile---------
                Dim oPutSubFile As Object

                Dim RtnCode, i As Integer
                Dim NewFormSno As Integer = wFormSno    '表單流水號
                Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
                Dim SQL As String

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep = 1 Then    '判斷是否起單
                        '取得表單流水號
                        oGetSeqNo = Server.CreateObject("GetSeqno.WFFormInf")
                        RtnCode = oGetSeqNo.Seqno(wFormNo, NewFormSno)   '表單號碼, 表單流水號
                        If RtnCode <> 0 Then
                            ErrCode = 9110
                        Else
                            '申請流程資料建置
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo       '表單號碼
                            oNextFlow.pFormSno = NewFormSno   '表單流水號
                            oNextFlow.pStep = 1               '工程關卡號碼
                            oNextFlow.pSeqNo = 1              '序號
                            RtnCode = oNextFlow.NewFlow(Request.Cookies("UserID").Value, wApplyID)   '建置者, 申請者
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            Dim OleDbConnection1 As New OleDbConnection
                            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                            Dim OleDBCommand1 As New OleDbCommand

                            SQL = "Update T_WaitHandle Set "
                            SQL = SQL + " Active = '" & "0" & "',"
                            SQL = SQL + " Sts = '" & "1" & "',"
                            SQL = SQL + " AEndTime = '" & NowDateTime & "',"
                            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
                            If DDelay.Visible = True Then
                                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                                SQL = SQL + " Reason = '" & DReason.Text & "',"
                                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
                            End If
                            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
                            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                            SQL = SQL + "   And Step    =  '" & CStr(wStep) & "'"
                            SQL = SQL + "   And EveryOne =  '1' "
                            'SQL = SQL + "   And SeqNo   =  '" & CStr(wSeqNo) & "'"
                            OleDBCommand1.Connection = OleDbConnection1
                            OleDBCommand1.CommandText = SQL
                            OleDbConnection1.Open()
                            OleDBCommand1.ExecuteNonQuery()
                            OleDbConnection1.Close()

                            '流程資料結束
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo   '表單號碼
                            oNextFlow.pFormSno = wFormSno                        '表單流水號
                            oNextFlow.pStep = wStep       '工程關卡號碼
                            RtnCode = oNextFlow.CheckFlow(Request.Cookies("UserID").Value, pRunNextStep)   '建置者, 流程結束否(會簽)
                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    oNextGate = Server.CreateObject("NextGate.WFNextGate")
                    oNextGate.pFormNo = wFormNo     '表單號碼
                    oNextGate.pStep = wStep         '工程關卡號碼
                    oNextGate.pUserID = Request.Cookies("UserID").Value       '簽核者ID
                    oNextGate.pApplyID = wApplyID                                     '申請者ID
                    RtnCode = oNextGate.NextGate(pNextStep, pNextGate, pCount, pFlowType)  '下一工程的, 號碼, 擔當者, 人數, 處理方法 
                    If RtnCode <> 0 Then ErrCode = 9130
                    If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                End If

                '建置流程資料
                If ErrCode = 0 And pRunNextStep = 1 Then
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            '取得工程負荷最後日時
                            oGetLoading = Server.CreateObject("GetLoading.WFOPINF")
                            RtnCode = oGetLoading.GetLastTime(pNextGate(i), wFormNo, pNextStep, NowDateTime, pLastTime, pCount1)  '表單號碼, 工程號碼, 開始日時, 最後日時, 件數
                            '取得預定開始,完成日程計算
                            oGetLeadTime = Server.CreateObject("GetLeadTime.WFOPInf")
                            oGetLeadTime.pFormNo = wFormNo      '表單號碼
                            oGetLeadTime.pStep = pNextStep      '工程號碼
                            oGetLeadTime.pLevel = wLevel        '難易度
                            RtnCode = oGetLeadTime.LeadTime(pLastTime, pStartTime, pEndTime)  '現在時間, 預定開始日時, 預定完成日時
                            '建置流程資料
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo       '表單號碼
                            oNextFlow.pFormSno = NewFormSno   '表單流水號
                            oNextFlow.pStep = pNextStep       '工程關卡號碼
                            oNextFlow.pSeqNo = i              '序號
                            oNextGate.pApplyID = wApplyID     '申請者ID
                            RtnCode = oNextFlow.NextFlow(Request.Cookies("UserID").Value, pNextGate(i), wApplyID, pStartTime, pEndTime, 0)   '建置者, 簽核者, 申請者, 預定開始日, 預定完成日, 重要性
                            If RtnCode <> 0 Then
                                ErrCode = 9150
                                Exit For
                            End If
                        Next i
                    Else
                        oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                        oNextFlow.pFormNo = wFormNo       '表單號碼
                        oNextFlow.pFormSno = wFormSno     '表單流水號
                        oNextFlow.pStep = pNextStep       '工程關卡號碼
                        oNextFlow.pSeqNo = 1              '序號
                        RtnCode = oNextFlow.EndFlow(Request.Cookies("UserID").Value, wApplyID)   '建置者, 申請者
                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        oSchedule = Server.CreateObject("Schedule.WFSchedule")
                        RtnCode = oSchedule.AdjustSchedule(Request.Cookies("UserID").Value, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel)
                    End If
                End If
                '儲存表單資料
                If ErrCode = 0 Then
                    'If RepeatRun = False Then
                    Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
                    Dim FileName As String

                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
                    Dim OleDBCommand1 As New OleDbCommand

                    If wFormSno = 0 And wStep = 1 Then    '判斷是否起單
                        If NewFormSno <> 0 Then
                            '表單Table處理
                            SQL = "Insert into F_ContactOutSheet "
                            SQL = SQL + "( "
                            SQL = SQL + "Sts, CompletedTime, FormNo, FormSNo, "          '1~4
                            SQL = SQL + "No, Date, Division, Person, SliderCode, "       '5~9
                            SQL = SQL + "Level, MapNo, OFormNo, OFormSno, NFormNo, NFormSno, "  '10~14
                            SQL = SQL + "Target, Content, Reason, AttachFile, "          '15~18
                            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "  '19~22
                            SQL = SQL + ")  "
                            SQL = SQL + "VALUES( "
                            SQL = SQL + " '0', "                                '狀態(0:未結,1:已結NG,2:已結OK)
                            SQL = SQL + " '" + NowDateTime + "', "              '結案日
                            SQL = SQL + " '000009', "                           '表單代號
                            SQL = SQL + " '" + CStr(NewFormSno) + "', "         '表單流水號
                            SQL = SQL + " '" + DNo.Text + "', "                 'NO
                            SQL = SQL + " '" + DDate.Text + "', "               '日期
                            SQL = SQL + " '" + DDivision.SelectedValue + "', "  '部門
                            SQL = SQL + " '" + DPerson.SelectedValue + "', "    '擔當
                            SQL = SQL + " '" + DSliderCode.Text + "', "         'Slider Code
                            SQL = SQL + " '" + DLevel.SelectedValue + "', "     '難易度
                            SQL = SQL + " '" + DMapNo.Text + "', "              '圖號
                            SQL = SQL + " '" + DOFormNo.Text + "', "            '表單No
                            If DOFormSno.Text = "" Then                         '單號
                                SQL = SQL + " '0', "
                            Else
                                SQL = SQL + " '" + DOFormSno.Text + "', "
                            End If
                            SQL = SQL + " '" + DNFormNo.Text + "', "            '表單No
                            If DNFormSno.Text = "" Then                         '單號
                                SQL = SQL + " '0', "
                            Else
                                SQL = SQL + " '" + DNFormSno.Text + "', "
                            End If
                            SQL = SQL + " '" + DTarget.Text + "', "             '目的
                            SQL = SQL + " '" + DContent.Text + "', "            '內容
                            SQL = SQL + " '" + DDReason.Text + "', "             '原因

                            FileName = ""
                            If DAttachFile.Visible Then                         '附件
                                If DAttachFile.PostedFile.FileName <> "" Then     '判斷有檔案上傳
                                    FileName = CStr(NewFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                                    Try    '上傳圖檔
                                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                                    Catch ex As Exception
                                    End Try
                                End If
                            Else
                                FileName = ""
                            End If
                            SQL = SQL + " '" + FileName + "', "

                            SQL = SQL + " '" + Request.Cookies("UserID").Value + "', "     '作成者
                            SQL = SQL + " '" + NowDateTime + "', "       '作成時間
                            SQL = SQL + " '" + "" + "', "                       '修改者
                            SQL = SQL + " '" + NowDateTime + "' "       '修改時間
                            SQL = SQL + " ) "
                            OleDBCommand1.Connection = OleDbConnection1
                            OleDBCommand1.CommandText = SQL
                            OleDbConnection1.Open()
                            OleDBCommand1.ExecuteNonQuery()
                            OleDbConnection1.Close()

                            'Update 原委託單狀態
                            If DOFormNo.Text = "000003" Then
                                SQL = "Update F_ManufInSheet Set "
                            Else
                                SQL = "Update F_ManufOutSheet Set "
                            End If
                            SQL = SQL + " Status = '" & "1" & "',"
                            SQL = SQL + " StatusDesc = '" & "" & "',"
                            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                            SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                            SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                            OleDBCommand1.Connection = OleDbConnection1
                            OleDBCommand1.CommandText = SQL
                            OleDbConnection1.Open()
                            OleDBCommand1.ExecuteNonQuery()
                            OleDbConnection1.Close()

                        End If  'pSeqno <> 0
                    Else    '判斷是否起單
                        SQL = "Update F_ContactOutSheet Set "
                        If pNextStep = 999 Then     '工程結束嗎?
                            SQL = SQL + " Sts = '" & "2" & "',"
                        Else
                            SQL = SQL + " Sts = '" & "0" & "',"
                        End If
                        SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
                        SQL = SQL + " No = '" & DNo.Text & "',"
                        SQL = SQL + " Date = '" & DDate.Text & "',"
                        SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
                        SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
                        SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
                        SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
                        SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
                        SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
                        SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
                        SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
                        SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
                        SQL = SQL + " Target = '" & DTarget.Text & "',"
                        SQL = SQL + " Content = '" & DContent.Text & "',"
                        SQL = SQL + " Reason = '" & DDReason.Text & "',"

                        If DAttachFile.Visible Then
                            If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                                FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                                Try    '上傳圖檔
                                    DAttachFile.PostedFile.SaveAs(Path + FileName)
                                Catch ex As Exception
                                End Try
                                SQL = SQL + " AttachFile = '" & FileName & "',"
                            End If
                        End If

                        SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                        SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                        SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
                        SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
                        OleDBCommand1.Connection = OleDbConnection1
                        OleDBCommand1.CommandText = SQL
                        OleDbConnection1.Open()
                        OleDBCommand1.ExecuteNonQuery()
                        OleDbConnection1.Close()

                        If pNextStep = 999 Then     '工程結束嗎? 結束時將新版壓舊版
                            If DNFormNo.Text <> "" And DNFormSno.Text <> "" Then
                                Dim DBDataSet1 As New DataSet
                                SQL = "Select * From F_ModManufSheet "
                                SQL = SQL & " Where FormNo =  '" & DNFormNo.Text & "'"
                                SQL = SQL & "   And FormSno =  '" & DNFormSno.Text & "'"
                                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                                DBAdapter1.Fill(DBDataSet1, "F_ModManufSheet")
                                If DBDataSet1.Tables("F_ModManufSheet").Rows.Count > 0 Then
                                    If DOFormNo.Text = "000003" Then
                                        SQL = "Update F_ManufInSheet Set "
                                    Else
                                        SQL = "Update F_ManufOutSheet Set "
                                    End If
                                    SQL = SQL + " No = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("No") & "',"
                                    SQL = SQL + " Date = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Date") & "',"
                                    SQL = SQL + " Division = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Division") & "',"
                                    SQL = SQL + " Person = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Person") & "',"
                                    SQL = SQL + " SliderCode = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderCode") & "',"
                                    SQL = SQL + " SliderGRCode = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderGRCode") & "',"
                                    SQL = SQL + " Spec = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Spec") & "',"
                                    SQL = SQL + " MapNo = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapNo") & "',"
                                    SQL = SQL + " MapFile1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile1") & "',"
                                    SQL = SQL + " MapFile2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MapFile2") & "',"
                                    SQL = SQL + " Level = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Level") & "',"
                                    SQL = SQL + " Assembler = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Assembler") & "',"
                                    SQL = SQL + " SliderType1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType1") & "',"
                                    SQL = SQL + " SliderType2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SliderType2") & "',"
                                    SQL = SQL + " ManufPlace = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ManufPlace") & "',"
                                    SQL = SQL + " Material = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Material") & "',"
                                    SQL = SQL + " MaterialDetail = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialDetail") & "',"
                                    SQL = SQL + " MaterialOther = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MaterialOther") & "',"
                                    SQL = SQL + " SellVendor = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SellVendor") & "',"
                                    SQL = SQL + " Buyer = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Buyer") & "',"
                                    SQL = SQL + " ConfirmFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ConfirmFile") & "',"
                                    SQL = SQL + " AuthorizeFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("AuthorizeFile") & "',"
                                    SQL = SQL + " DevReason = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("DevReason") & "',"
                                    SQL = SQL + " Sample = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Sample") & "',"
                                    SQL = SQL + " Price = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Price") & "',"
                                    SQL = SQL + " ArMoldFee = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("ArMoldFee") & "',"
                                    SQL = SQL + " PurMold = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PurMold") & "',"
                                    SQL = SQL + " PullerPrice = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("PullerPrice") & "',"
                                    SQL = SQL + " MoldName = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldName") & "',"
                                    SQL = SQL + " MoldQty = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty") & "',"
                                    SQL = SQL + " MoldPoint = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("MoldQty") & "',"
                                    SQL = SQL + " Quality1 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality1") & "',"
                                    SQL = SQL + " Quality2 = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("Quality2") & "',"
                                    SQL = SQL + " QAAttachFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("QAAttachFile") & "',"
                                    SQL = SQL + " SampleFile = '" & DBDataSet1.Tables("F_ModManufSheet").Rows(0).Item("SampleFile") & "',"
                                    SQL = SQL + " Status = '" & "0" & "',"
                                    SQL = SQL + " StatusDesc = '" & "" & "',"
                                    SQL = SQL + " Contact = '" & "1" & "',"
                                    SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                                    SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                    SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                                    SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                                    OleDBCommand1.Connection = OleDbConnection1
                                    OleDBCommand1.CommandText = SQL
                                    OleDbConnection1.Open()
                                    OleDBCommand1.ExecuteNonQuery()
                                    OleDbConnection1.Close()
                                    'Work Sub File Transfer
                                    oPutSubFile = Server.CreateObject("MSubFile.SubFile")
                                    RtnCode = oPutSubFile.Transfer(DNFormNo.Text, DNFormSno.Text, DOFormNo.Text, DOFormSno.Text, Request.Cookies("UserID").Value)
                                End If
                            Else
                                If DOFormNo.Text = "000003" Then
                                    SQL = "Update F_ManufInSheet Set "
                                Else
                                    SQL = "Update F_ManufOutSheet Set "
                                End If
                                SQL = SQL + " Status = '" & "0" & "',"
                                SQL = SQL + " StatusDesc = '" & "" & "',"
                                SQL = SQL + " Contact = '" & "1" & "',"
                                SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
                                SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
                                SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
                                SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
                                OleDBCommand1.Connection = OleDbConnection1
                                OleDBCommand1.CommandText = SQL
                                OleDbConnection1.Open()
                                OleDBCommand1.ExecuteNonQuery()
                                OleDbConnection1.Close()
                            End If  'NFormNo and NFormSno = ""
                        End If  '=999
                    End If  'Insert / Update
                    'End If
                    '傳送郵件
                    If pNextStep <> 999 Then
                        For i = 1 To pCount
                            oMail = Server.CreateObject("SendMail.WFMail")
                            oMail.Send(Request.Cookies("UserID").Value, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "OK")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼
                        Next i
                    Else
                        oMail = Server.CreateObject("SendMail.WFMail")
                        oMail.Send(Request.Cookies("UserID").Value, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "OK")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼
                    End If

                    If pFlowType = 0 Then
                        Run = True
                        RepeatRun = True
                    End If
                    wStep = pNextStep     '下一工程關卡號碼
                Else
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
                Dim URL As String = "MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                    "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            Response.Write(YKK.ShowMessage(Message))
        End If      '上傳檔案ErrCode=0
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG按鈕點選後事件
    '**
    '*****************************************************************
    Private Sub BNG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BNG.Click
        Dim SQL As String
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim RtnCode As Integer = 0
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("ContactOutFilePath"))
        Dim FileName As String
        Dim Message As String = ""
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        '--流程資料設定---------
        Dim oNextFlow As Object
        Dim pRunNextStep As Integer = 0         '是否執行計算下一關(會簽)
        '--郵件傳送---------
        Dim oMail As Object
        '--日程調整---------
        Dim oSchedule As Object

        'Check理由
        If ErrCode = 0 Then
            If DReasonCode.Visible = True Then
                If DReasonCode.SelectedValue = "99" Then
                    If DReasonDesc.Text = "" Then ErrCode = 9310
                End If
            End If
        End If

        '更新交易資料
        If ErrCode = 0 Then
            SQL = "Update T_WaitHandle Set "
            SQL = SQL + " Active = '" & "0" & "',"
            SQL = SQL + " Sts = '" & "2" & "',"
            SQL = SQL + " AEndTime = '" & NowDateTime & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQL = SQL + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQL = SQL + " Reason = '" & DReason.Text & "',"
                SQL = SQL + " ReasonDesc = '" & DReasonDesc.Text & "',"
            End If
            SQL = SQL + " DecideDesc = '" & DDecideDesc.Text & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL + "   And Step    =  '" & CStr(wStep) & "'"
            SQL = SQL + "   And EveryOne =  '1' "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '流程資料結束
            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
            oNextFlow.pFormNo = wFormNo   '表單號碼
            oNextFlow.pFormSno = wFormSno                        '表單流水號
            oNextFlow.pStep = wStep       '工程關卡號碼
            RtnCode = oNextFlow.CheckFlow(Request.Cookies("UserID").Value, pRunNextStep)  '建置者, 流程結束否(會簽)
            If RtnCode <> 0 Then ErrCode = 9320
        End If

        '當工程日程調整
        If ErrCode = 0 Then
            Dim wLevel As String = DLevel.SelectedValue     '難易度
            oSchedule = Server.CreateObject("Schedule.WFSchedule")
            RtnCode = oSchedule.AdjustSchedule(Request.Cookies("UserID").Value, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel)
        End If

        '更新表單資料
        If ErrCode = 0 Then
            SQL = "Update F_ContactOutSheet Set "
            SQL = SQL + " Sts = '" & "1" & "',"
            SQL = SQL + " CompletedTime = '" & NowDateTime & "',"
            SQL = SQL + " No = '" & DNo.Text & "',"
            SQL = SQL + " Date = '" & DDate.Text & "',"
            SQL = SQL + " Division = '" & DDivision.SelectedValue & "',"
            SQL = SQL + " Person = '" & DPerson.SelectedValue & "',"
            SQL = SQL + " SliderCode = '" & DSliderCode.Text & "',"
            SQL = SQL + " Level = '" & DLevel.SelectedValue & "',"
            SQL = SQL + " MapNo = '" & DMapNo.Text & "',"
            SQL = SQL + " OFormNo = '" & DOFormNo.Text & "',"
            SQL = SQL + " OFormSno = '" & DOFormSno.Text & "',"
            SQL = SQL + " NFormNo = '" & DNFormNo.Text & "',"
            SQL = SQL + " NFormSno = '" & DNFormSno.Text & "',"
            SQL = SQL + " Target = '" & DTarget.Text & "',"
            SQL = SQL + " Content = '" & DContent.Text & "',"
            SQL = SQL + " Reason = '" & DDReason.Text & "',"

            If DAttachFile.Visible Then
                If DAttachFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    FileName = CStr(wFormSno) & "-" & "Attach" & "-" & UploadDateTime & "-" & Right(DAttachFile.PostedFile.FileName, InStr(StrReverse(DAttachFile.PostedFile.FileName), "\") - 1)
                    Try    '上傳圖檔
                        DAttachFile.PostedFile.SaveAs(Path + FileName)
                    Catch ex As Exception
                    End Try
                    SQL = SQL + " AttachFile = '" & FileName & "',"
                End If
            End If

            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & wFormNo & "'"
            SQL = SQL + "   And FormSno =  '" & CStr(wFormSno) & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '原表單處理
            If DOFormNo.Text = "000003" Then
                SQL = "Update F_ManufInSheet Set "
            Else
                SQL = "Update F_ManufOutSheet Set "
            End If
            SQL = SQL + " Status = '" & "0" & "',"
            SQL = SQL + " StatusDesc = '" & "" & "',"
            SQL = SQL + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where FormNo  =  '" & DOFormNo.Text & "'"
            SQL = SQL + "   And FormSno =  '" & DOFormSno.Text & "'"
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDbConnection1.Open()
            OleDBCommand1.ExecuteNonQuery()
            OleDbConnection1.Close()

            '傳送郵件
            oMail = Server.CreateObject("SendMail.WFMail")
            oMail.Send(Request.Cookies("UserID").Value, wApplyID, wApplyID, wFormNo, wFormSno, wStep, "NG")
            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼
        End If

        If ErrCode = 0 Then
            Dim URL As String = "MessagePage.aspx?pMSGID=N&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
            Response.Redirect(URL)
        Else
            If ErrCode = 9310 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9320 Then Message = "流程資料更新異常,請確認或連絡系統人員!"
            Response.Write(YKK.ShowMessage(Message))
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".jpg", ".jpeg", ".gif", ".xls", ".doc"}   '定義允許的檔案格式
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

End Class
