Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class IT_SelectDataSheet_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BFinishDate As System.Web.UI.WebControls.Button
    Protected WithEvents DSystem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBEndDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents BBStartDate As System.Web.UI.WebControls.Button
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DRemark As System.Web.UI.WebControls.TextBox
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFinishDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label
    Protected WithEvents DReasonDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DReasonCode As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDelay As System.Web.UI.WebControls.Image
    Protected WithEvents DDecideDesc As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDescSheet As System.Web.UI.WebControls.Image
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DEngineer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DRefFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents BSAVE As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BNG1 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents BOK As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents DSelectDataSheet1 As System.Web.UI.WebControls.Image

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
    Dim wDepo As String = "TP"      '台北行事曆(CL->中壢, TP->台北)
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

        Response.Cookies("PGM").Value = "IT_SelectDataSheet_01.aspx"
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
        '參考附件
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then
                If Message = "" Then
                    Message = "參考附件"
                Else
                    Message = Message & ", " & "參考附件"
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
        oUpdateFlow.pDepo = wDepo          '台北
        oUpdateFlow.UpdateFlow(Request.Cookies("UserID").Value)
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("SelectDataFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_SelectDataSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SelectDataSheet")
        If DBDataSet1.Tables("F_SelectDataSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("No")
            DDate.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Date")
            DName.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Name")
            DEmpID.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("JobTitle")
            DJobCode.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("JobCode")
            DDepoName.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("DepoName")
            DDepoCode.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("DepoCode")
            DDivision.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Division")
            DDivisionCode.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("DivisionCode")
            DFinishDate.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("FinishDate")
            DTarget.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Target")

            If DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("RefFile") <> "" Then
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If
            If DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("BStartDate") = "1900/1/1" Then
                DBStartDate.Text = ""
            Else
                DBStartDate.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("BStartDate")
            End If

            If DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("BEndDate") = "1900/1/1" Then
                DBEndDate.Text = ""
            Else
                DBEndDate.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("BEndDate")
            End If

            DBDays.Text = CStr(DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("BDays"))
            SetFieldData("Engineer", DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Engineer"))
            SetFieldData("System", DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("System"))
            DRemark.Text = DBDataSet1.Tables("F_SelectDataSheet").Rows(0).Item("Remark")
            '交易資料
            DBDataSet1.Clear()
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet1, "T_WaitHandle")
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
            SQL = SQL + "Order by Unique_ID Desc "
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
                BSAVE.Attributes("onclick") = "Button('SAVE', '" + BSAVE.Value + "');"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc1")
                BNG1.Attributes("onclick") = "Button('NG1', '" + BNG1.Value + "');"
                wNGSts1 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGDesc2")
                BNG2.Attributes("onclick") = "Button('NG2', '" + BNG2.Value + "');"
                wNGSts2 = DBDataSet1.Tables("M_Flow").Rows(0).Item("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If DBDataSet1.Tables("M_Flow").Rows(0).Item("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Value = DBDataSet1.Tables("M_Flow").Rows(0).Item("OKDesc")
                BOK.Attributes("onclick") = "Button('OK', '" + BOK.Value + "');"
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
                DSelectDataSheet1.Visible = True   '表單Sheet-1
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
                    Top = 666
                Else
                    DReasonCode.Visible = False     '延遲理由代碼
                    DReason.Visible = False         '延遲理由
                    DReasonDesc.Visible = False     '延遲其他說明
                    Top = 554
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
                LRefFile.Visible = True     '附件
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
            Top = 482
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
            '連結隱藏
            LRefFile.Visible = False       '附件
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
        BFinishDate.Attributes("onclick") = "CalendarPicker('Form1.DFinishDate');"  '希望完成日
        BBStartDate.Attributes("onclick") = "CalendarPicker('Form1.DBStartDate');"  '預定開始日
        BBEndDate.Attributes("onclick") = "CalendarPicker('Form1.DBEndDate');"      '預定完成日
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
                    Top = 554
                Else
                    If DDelay.Visible = True Then
                        Top = 666
                    Else
                        Top = 554
                    End If
                End If
            End If
        Else
            Top = 482
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
        Dim wEmpID, wJobTitle, wJobCode, wDivision, wDivisionCode, wDepoName, wDepoCode As String

        OleDbConnection1.Open()
        '取得申請者資訊
        SQL = "Select UserName, EmpID, JobName, JobID, DivName, DivID, DepoID, DepoName From M_Users "
        SQL = SQL & " Where UserID = '" & Request.Cookies("UserID").Value & "'"
        SQL = SQL & "   And Active = '1' "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Users")
        If DBDataSet1.Tables("M_Users").Rows.Count > 0 Then
            wDepoCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID")
            wDepoName = DBDataSet1.Tables("M_Users").Rows(0).Item("DepoName")
            If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "01" Then wDepo = "CL"
            If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "10" Then wDepo = "TP"
            If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "11" Then wDepo = "TP"
            If DBDataSet1.Tables("M_Users").Rows(0).Item("DepoID") = "51" Then wDepo = "YA"
            wUserName = DBDataSet1.Tables("M_Users").Rows(0).Item("UserName")
            wEmpID = DBDataSet1.Tables("M_Users").Rows(0).Item("EmpID")
            wJobTitle = DBDataSet1.Tables("M_Users").Rows(0).Item("JobName")
            wJobCode = DBDataSet1.Tables("M_Users").Rows(0).Item("JobID")
            wDivision = DBDataSet1.Tables("M_Users").Rows(0).Item("DivName")
            wDivisionCode = DBDataSet1.Tables("M_Users").Rows(0).Item("DivID")
        End If
        OleDbConnection1.Close()
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
        If pPost = "New" Then DName.Text = wUserName
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
        '工作日
        Select Case FindFieldInf("BDays")
            Case 0  '顯示
                DBDays.BackColor = Color.LightGray
                DBDays.ReadOnly = True
                DBDays.Visible = True
            Case 1  '修改+檢查
                DBDays.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DBDaysRqd", "DBDays", "異常：需輸入工作日")
                DBDays.Visible = True
            Case 2  '修改
                DBDays.BackColor = Color.Yellow
                DBDays.Visible = True
            Case Else   '隱藏
                DBDays.Visible = False
        End Select
        If pPost = "New" Then DBDays.Text = ""
        '希望完成日
        Select Case FindFieldInf("FinishDate")
            Case 0  '顯示
                DFinishDate.BackColor = Color.LightGray
                DFinishDate.ReadOnly = True
                DFinishDate.Visible = True
            Case 1  '修改+檢查
                DFinishDate.BackColor = Color.GreenYellow
                DFinishDate.ReadOnly = True
                ShowRequiredFieldValidator("DFinishDateRqd", "DFinishDate", "異常：需輸入希望完成日")
                DFinishDate.Visible = True
            Case 2  '修改
                DFinishDate.BackColor = Color.Yellow
                DFinishDate.ReadOnly = True
                DFinishDate.Visible = True
            Case Else   '隱藏
                DFinishDate.Visible = False
        End Select
        If pPost = "New" Then DFinishDate.Text = ""
        '目的
        Select Case FindFieldInf("Target")
            Case 0  '顯示
                DTarget.BackColor = Color.LightGray
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
        '參考附件
        Select Case FindFieldInf("RefFile")
            Case 0  '顯示
                DRefFile.Visible = False
                DRefFile.Style.Add("BACKGROUND-COLOR", "LightGrey")
            Case 1  '修改+檢查
                ShowRequiredFieldValidator("DRefFileRqd", "DRefFile", "異常：需輸入參考附件")
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "GreenYellow")
            Case 2  '修改
                DRefFile.Visible = True
                DRefFile.Style.Add("BACKGROUND-COLOR", "Yellow")
            Case Else   '隱藏
                DRefFile.Visible = False
        End Select
        '預定開始日
        Select Case FindFieldInf("BStartDate")
            Case 0  '顯示
                DBStartDate.BackColor = Color.LightGray
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
            Case 1  '修改+檢查
                DBStartDate.BackColor = Color.GreenYellow
                DBStartDate.ReadOnly = True
                ShowRequiredFieldValidator("DBStartDateRqd", "DBStartDate", "異常：需輸入預定開始日")
                DBStartDate.Visible = True
            Case 2  '修改
                DBStartDate.BackColor = Color.Yellow
                DBStartDate.ReadOnly = True
                DBStartDate.Visible = True
            Case Else   '隱藏
                DBStartDate.Visible = False
        End Select
        If pPost = "New" Then DBStartDate.Text = ""
        '預定完成日
        Select Case FindFieldInf("BEndDate")
            Case 0  '顯示
                DBEndDate.BackColor = Color.LightGray
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
            Case 1  '修改+檢查
                DBEndDate.BackColor = Color.GreenYellow
                DBEndDate.ReadOnly = True
                ShowRequiredFieldValidator("DBEndDateRqd", "DBEndDate", "異常：需輸入預定完成日")
                DBEndDate.Visible = True
            Case 2  '修改
                DBEndDate.BackColor = Color.Yellow
                DBEndDate.ReadOnly = True
                DBEndDate.Visible = True
            Case Else   '隱藏
                DBEndDate.Visible = False
        End Select
        If pPost = "New" Then DBEndDate.Text = ""
        '開發擔當
        Select Case FindFieldInf("Engineer")
            Case 0  '顯示
                DEngineer.BackColor = Color.LightGray
                DEngineer.Visible = True
            Case 1  '修改+檢查
                DEngineer.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DEngineerRqd", "DEngineer", "異常：需輸入開發擔當")
                DEngineer.Visible = True
            Case 2  '修改
                DEngineer.BackColor = Color.Yellow
                DEngineer.Visible = True
            Case Else   '隱藏
                DEngineer.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Engineer", "ZZZZZZ")
        '系統
        Select Case FindFieldInf("System")
            Case 0  '顯示
                DSystem.BackColor = Color.LightGray
                DSystem.Visible = True
            Case 1  '修改+檢查
                DSystem.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DSystemRqd", "DSystem", "異常：需輸入系統")
                DSystem.Visible = True
            Case 2  '修改
                DSystem.BackColor = Color.Yellow
                DSystem.Visible = True
            Case Else   '隱藏
                DSystem.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("System", "ZZZZZZ")
        '備註
        Select Case FindFieldInf("Remark")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.ReadOnly = True
                DRemark.Visible = True
            Case 1  '修改+檢查
                DRemark.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入備註")
                DRemark.Visible = True
            Case 2  '修改
                DRemark.BackColor = Color.Yellow
                DRemark.Visible = True
            Case Else   '隱藏
                DRemark.Visible = False
        End Select
        If pPost = "New" Then DRemark.Text = ""
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

        '開發擔當
        If pFieldName = "Engineer" Then
            DEngineer.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DEngineer.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1104' and DKey='ENGINEER' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DEngineer.Items.Add(ListItem1)
                Next
            End If
        End If
        '系統
        If pFieldName = "System" Then
            DSystem.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DSystem.Items.Add(ListItem1)
                End If
            Else
                SQL = "Select * From M_Referp Where Cat='1104' and DKey='SYSTEM' Order by Data "
                Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1.Fill(DBDataSet1, "M_Referp")
                DBTable1 = DBDataSet1.Tables("M_Referp")
                For i = 0 To DBTable1.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = DBTable1.Rows(i).Item("Data")
                    ListItem1.Value = DBTable1.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DSystem.Items.Add(ListItem1)
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
            DisabledButton()   '停止Button運作
            Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
            Dim Message As String = ""

            'Check上傳附件Size及格式
            If ErrCode = 0 Then
                If DRefFile.Visible Then
                    If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                        ErrCode = UPFileIsNormal(DRefFile)
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
                Dim URL As String = "MessagePage.aspx?pMSGID=S&pFormNo=" & wFormNo & "&pStep=" & wStep & _
                                                    "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
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
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
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
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")
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
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
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

        'Check上傳附件Size及格式
        If ErrCode = 0 Then
            If DRefFile.Visible Then
                If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DRefFile)
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
        '檢查委託書No
        If ErrCode = 0 Then
            If DNo.Text <> "" Then
                Dim oCheckNo As Object
                oCheckNo = Server.CreateObject("CheckNo.CommissionNo")
                ErrCode = oCheckNo.CommissionNo("001104", wFormSno, wStep, DNo.Text)    '表單號碼, 表單流水號, 工程, 委託書No
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
                Dim oNextGate As Object
                Dim pNextGate(10) As String
                Dim pAgentGate(10) As String
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

                '取得表單流水號或更新交易資料
                If ErrCode = 0 Then
                    If wFormSno = 0 And wStep < 3 Then    '判斷是否起單
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
                            oNextFlow.pStep = wStep           '工程關卡號碼
                            oNextFlow.pSeqNo = 1              '序號
                            oNextFlow.pDepo = wDepo           '台北
                            RtnCode = oNextFlow.NewFlow(Request.Cookies("UserID").Value, wApplyID)   '建置者, 申請者
                            '設定委託No
                            DNo.Text = SetNo(NewFormSno)
                        End If
                        pRunNextStep = 1
                    Else
                        If RepeatRun = False Then   '不是通知的重覆執行
                            '更新交易資料
                            ModifyTranData(pFun, pSts)
                            '流程資料結束
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo   '表單號碼
                            oNextFlow.pFormSno = wFormSno                        '表單流水號
                            oNextFlow.pStep = wStep       '工程關卡號碼
                            oNextFlow.pDepo = wDepo           '台北
                            RtnCode = oNextFlow.CheckFlow(Request.Cookies("UserID").Value, pRunNextStep)   '建置者, 流程結束否(會簽)
                            If RtnCode <> 0 Then ErrCode = 9120
                        Else
                            pRunNextStep = 1    '是通知的重覆執行
                        End If
                    End If
                End If

                '取得下一關
                If ErrCode = 0 And pRunNextStep = 1 Then
                    Dim SQL As String = ""
                    Dim wAllocateID As String = ""

                    Dim DBDataSet1 As New DataSet
                    Dim OleDbConnection1 As New OleDbConnection
                    OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

                    OleDbConnection1.Open()
                    SQL = "Select Data From M_Referp "
                    SQL = SQL & " Where Cat = '1104' "
                    SQL = SQL & "   And DKey =  '" & DSystem.SelectedValue & "'"
                    Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter1.Fill(DBDataSet1, "M_Refrp")
                    If DBDataSet1.Tables("M_Refrp").Rows.Count > 0 Then wAllocateID = DBDataSet1.Tables("M_Refrp").Rows(0).Item("Data")
                    OleDbConnection1.Close()

                    oNextGate = Server.CreateObject("NextGate.WFNextGate")
                    oNextGate.pFormNo = wFormNo     '表單號碼
                    oNextGate.pStep = wStep         '工程關卡號碼
                    oNextGate.pUserID = Request.Cookies("UserID").Value       '簽核者ID
                    oNextGate.pApplyID = wApplyID                             '申請者ID
                    oNextGate.pAgentID = wAgentID                             '被代理人ID
                    oNextGate.pAllocateID = wAllocateID                       '指定ID
                    oNextGate.pMultiJob = MultiJob                            '多人核定
                    RtnCode = oNextGate.NextGate(pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)  '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save) 
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
                            '取得工程負荷最後日時
                            oGetLoading = Server.CreateObject("GetLoading.WFOPLoad")
                            RtnCode = oGetLoading.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)  '表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數

                            '取得預定開始,完成日程計算
                            oGetLeadTime = Server.CreateObject("GetLeadTime.WFOPInf")
                            oGetLeadTime.pFormNo = wFormNo      '表單號碼
                            oGetLeadTime.pStep = pNextStep      '工程號碼
                            oGetLeadTime.pLevel = wLevel        '難易度
                            oGetLeadTime.pQCLeadTime = wQCLT    'QC-L/T
                            RtnCode = oGetLeadTime.LeadTime(pLastTime, pStartTime, pEndTime, wDepo)  '現在時間, 預定開始日時, 預定完成日時, 行事曆
                            '建置流程資料
                            oNextFlow = Server.CreateObject("FlowEngine.WFFlowEngine")
                            oNextFlow.pFormNo = wFormNo       '表單號碼
                            oNextFlow.pFormSno = NewFormSno   '表單流水號
                            oNextFlow.pStep = pNextStep       '工程關卡號碼
                            oNextFlow.pSeqNo = i              '序號
                            oNextGate.pApplyID = wApplyID     '申請者ID
                            oNextFlow.pDepo = wDepo           '台北
                            RtnCode = oNextFlow.NextFlow(Request.Cookies("UserID").Value, pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)   '建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性
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
                        oNextFlow.pDepo = wDepo           '台北
                        RtnCode = oNextFlow.EndFlow(Request.Cookies("UserID").Value, wApplyID)   '建置者, 申請者
                        If RtnCode <> 0 Then ErrCode = 9160
                    End If
                End If
                '當工程日程調整
                If ErrCode = 0 Then
                    If RepeatRun = False Then
                        oSchedule = Server.CreateObject("Schedule.WFSchedule")
                        RtnCode = oSchedule.AdjustSchedule(Request.Cookies("UserID").Value, wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDepo)
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
                            oMail = Server.CreateObject("SendMail.WFMail")
                            oMail.Send(Request.Cookies("UserID").Value, pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼
                        Next i
                    Else
                        oMail = Server.CreateObject("SendMail.WFMail")
                        oMail.Send(Request.Cookies("UserID").Value, wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼
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
                Dim oMail As Object
                oMail = Server.CreateObject("SendMail.WFMail")
                oMail.SendMail()

                Dim URL As String = "MessagePage.aspx?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                    "&pUserID=" & Request.Cookies("UserID").Value & "&pApplyID=" & wApplyID
                Response.Redirect(URL)
            End If
        Else
            EnabledButton()   '起動Button運作
            If ErrCode = 9010 Then Message = "上傳檔案格式錯誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案Size超過1024KB,請確認!"
            If ErrCode = 9030 Then Message = "上傳檔案沒指定,請確認!"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "預定或實際加班時間需填寫,請確認!"
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SelectDataFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Insert into F_SelectDataSheet "
        SQl = SQl + "( "
        SQl = SQl + "Sts, CompletedTime, FormNo, FormSno, "
        SQl = SQl + "No, Date, Name, EmpID, JobTitle, "
        SQl = SQl + "JobCode, DepoName, DepoCode, Division, DivisionCode, "
        SQl = SQl + "BDays, FinishDate, Target, RefFile, "
        SQl = SQl + "BStartDate, BEndDate, Engineer, Remark, System, "
        SQl = SQl + "CreateUser, CreateTime, ModifyUser, ModifyTime "
        SQl = SQl + ")  "
        SQl = SQl + "VALUES( "

        SQl = SQl + " '0', "
        SQl = SQl + " '" + NowDateTime + "', "
        SQl = SQl + " '001104', "
        SQl = SQl + " '" + CStr(NewFormSno) + "', "
        SQl = SQl + " '" + DNo.Text + "', "
        SQl = SQl + " '" + DDate.Text + "', "
        SQl = SQl + " N'" + DName.Text + "', "
        SQl = SQl + " '" + DEmpID.Text + "', "
        SQl = SQl + " N'" + DJobTitle.Text + "', "
        SQl = SQl + " '" + DJobCode.Text + "', "
        SQl = SQl + " N'" + DDepoName.Text + "', "
        SQl = SQl + " '" + DDepoCode.Text + "', "
        SQl = SQl + " N'" + DDivision.Text + "', "
        SQl = SQl + " '" + DDivisionCode.Text + "', "

        SQl = SQl + " '" + DBDays.Text + "', "
        SQl = SQl + " '" + DFinishDate.Text + "', "
        SQl = SQl + " N'" + DTarget.Text + "', "

        FileName = ""
        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(NewFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                Try    '上傳圖檔
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
            End If
        Else
            FileName = ""
        End If
        SQl = SQl + " N'" + FileName + "', "

        SQl = SQl + " '" + DBStartDate.Text + "', "
        SQl = SQl + " '" + DBEndDate.Text + "', "
        SQl = SQl + " '" + DEngineer.SelectedValue + "', "
        SQl = SQl + " N'" + DRemark.Text + "', "
        SQl = SQl + " '" + DSystem.SelectedValue + "', "
        '--------------------------------------------
        SQl = SQl + " '" + Request.Cookies("UserID").Value + "', "     '作成者
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
        Dim Path As String = Server.MapPath(ConfigurationSettings.AppSettings.Get("SelectDataFilePath"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)     '上傳日期
        Dim FileName As String
        Dim SQl As String

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        SQl = "Update F_SelectDataSheet Set "
        If pFun <> "SAVE" Then
            SQl = SQl + " Sts = '" & pSts & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
        End If
        SQl = SQl + " No = '" & DNo.Text & "',"
        SQl = SQl + " Date = '" & DDate.Text & "',"

        SQl = SQl + " Name = N'" & DName.Text & "',"
        SQl = SQl + " EmpID = '" & DEmpID.Text & "',"
        SQl = SQl + " JobTitle = N'" & DJobTitle.Text & "',"
        SQl = SQl + " JobCode = '" & DJobCode.Text & "',"
        SQl = SQl + " DepoName = N'" & DDepoName.Text & "',"
        SQl = SQl + " DepoCode = '" & DDepoCode.Text & "',"
        SQl = SQl + " Division = N'" & DDivision.Text & "',"
        SQl = SQl + " DivisionCode = '" & DDivisionCode.Text & "',"

        SQl = SQl + " BDays = '" & DBDays.Text & "',"
        SQl = SQl + " FinishDate = '" & DFinishDate.Text & "',"
        SQl = SQl + " Target = N'" & DTarget.Text & "',"

        If DRefFile.Visible Then
            If DRefFile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                FileName = CStr(wFormSno) & "-" & "Ref" & "-" & UploadDateTime & "-" & Right(DRefFile.PostedFile.FileName, InStr(StrReverse(DRefFile.PostedFile.FileName), "\") - 1)
                Try    '上傳草圖
                    DRefFile.PostedFile.SaveAs(Path + FileName)
                Catch ex As Exception
                End Try
                SQl = SQl + " RefFile = N'" & FileName & "',"
            End If
        End If

        SQl = SQl + " BStartDate = '" & DBStartDate.Text & "',"
        SQl = SQl + " BEndDate = '" & DBEndDate.Text & "',"
        SQl = SQl + " Engineer = '" & DEngineer.SelectedItem.Text & "',"
        SQl = SQl + " Remark = N'" & DRemark.Text & "',"
        SQl = SQl + " System = '" & DSystem.SelectedItem.Text & "',"
        '--------------------------------------------
        SQl = SQl + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
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
            SQl = SQl + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
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
            SQl = SQl + " ModifyUser = '" & Request.Cookies("UserID").Value & "',"
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
                SQl = SQl + " '" + Request.Cookies("UserID").Value + "', "
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
                    SQl = SQl + " CreateUser = '" & Request.Cookies("UserID").Value & "',"
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

End Class
