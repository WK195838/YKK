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
Imports System.Web.UI.WebControls




Partial Class StockInSheet_01
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
    Dim oWaves As New WAVES.CommonService
    Dim UploadName As String
    Dim griderror, inserterror, fielderror As Boolean
    Dim Message As String = ""
    Dim str1 As String
    Dim AddFileName As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BASDate.Attributes("onclick") = "calendarPicker('DCheckDate');"

        DAttachfile.Attributes("onchange") = "UploadFile(this)"

        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        SetControlPosition()    ' 設定控制項位置

        If Not Me.IsPostBack Then   '不是PostBack
            ShowSheetField("New")   '表單欄位顯示及欄位輸入檢查
            ShowSheetFunction()     '表單功能按鈕顯示

            If wFormSno > 0 And wStep > 2 Then    '判斷是否[簽核]
                ShowFormData()      '顯示表單資料
                UpdateTranFile()    '更新交易資料
                TopPosition()
                SetControlPosition()    ' 設定控制項位置

            End If
            SetPopupFunction()      '設定彈出視窗事件


        Else
            ShowSheetFunction()     '表單功能按鈕顯示
            ' fpObj.GetDeafultNextGate(DNextGate, DDivision.Text,Request.QueryString("pUserID")) ' 設定預設的簽核者
            ShowSheetField("Post")  '表單欄位顯示及欄位輸入檢查

            ShowMessage()           '上傳資料檢查及顯示訊息

            '上傳資料檢查及顯示訊息

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
        If wFormSno > 0 And wStep > 3 Then      '判斷是否[簽核]
            Top = 670
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 670
                    End If
                End If
            End If
        Else
            Top = 670
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     '上傳資料檢查及顯示訊息
    '**
    '*****************************************************************
    Sub ShowMessage()
        Dim Message As String = ""
        If Message <> "" Then
            Message = "下列所設定的附加檔案將會遺失 (" & Message & ") " & ",請重新設定!"
            Response.Write(YKK.ShowMessage(Message))
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetControlPosition)
    '**    設定按鈕,說明,延遲,核定履歷的位置
    '**
    '*****************************************************************
    Sub SetControlPosition()



        If DDescSheet.Visible Then                                      ' 說明
            DDescSheet.Style("top") = Top & "px"
            DDecideDesc.Style("top") = Top + 6 & "px"
            Top = Top + 50
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top + 20 & "px"
        BNG1.Style("top") = Top + 20 & "px"
        BNG2.Style("top") = Top + 20 & "px"
        BOK.Style("top") = Top + 20 & "px"


        Top = Top + 50
        If GridView2.Rows.Count > 0 Then                                ' 核定履歷
            DHistoryLabel.Style("top") = Top & "px"
            Top += 20
            GridView2.Style("top") = Top & "px"
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
        wStep = Request.QueryString("pStep")        '工程代碼
        wbFormSno = Request.QueryString("pFormSno")  '連續起單用流水號
        wbStep = Request.QueryString("pStep")        '連續起單用工程代碼
        wSeqNo = Request.QueryString("pSeqNo")      '序號
        wApplyID = Request.QueryString("pApplyID")  '申請者ID
        wAgentID = Request.QueryString("pAgentID")  '被代理人ID
        wApplyCalendar = oCommon.GetCalendarGroup(wApplyID)
        wDecideCalendar = oCommon.GetCalendarGroup(Request.QueryString("pUserID"))
        '
        Response.Cookies("PGM").Value = "ExpenseSheet_01.aspx"
        Response.Cookies("PGMFORMSNO").Value = Request.QueryString("pFormSno")  '表單流水號
        Response.Cookies("PGMSTEP").Value = Request.QueryString("pStep")        '工程代碼


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub UpdateTranFile()
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDecideCalendar, Request.QueryString("pUserID"))
        '表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者


       


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowSheetFunction)
    '**     表單功能按鈕顯示
    '**
    '*****************************************************************
    Sub ShowSheetFunction()
        Dim wDelay As Integer = 0
        Dim sql As String = ""
        sql = "Select * From M_Flow "
        sql &= " Where Active = 1 "
        sql &= "   And FormNo =  '" & wFormNo & "'"
        sql &= "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(sql)
        If dtFlow.Rows.Count > 0 Then
            '電子簽章未使用
            If dtFlow.Rows(0)("SignImage") = 1 Then
            Else
            End If
            '附加檔案未使用(由FormField中設定)
            If dtFlow.Rows(0)("Attach") = 1 Then
            Else
            End If
            '儲存按鈕
            If dtFlow.Rows(0)("SaveFun") = 1 Then
                BSAVE.Visible = True
                BSAVE.Text = dtFlow.Rows(0)("SaveDesc")
                'BSAVE.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BSAVE.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
            Else
                BSAVE.Visible = False
            End If
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                BNG1.Visible = True
                BNG1.Text = dtFlow.Rows(0)("NGDesc1")
                'BNG1.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG1.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            Else
                BNG1.Visible = False
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                BNG2.Visible = True
                BNG2.Text = dtFlow.Rows(0)("NGDesc2")
                'BNG2.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BNG2.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            Else
                BNG2.Visible = False
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                BOK.Visible = True
                BOK.Text = dtFlow.Rows(0)("OKDesc")
                'BOK.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[" + BOK.Text + "]作業嗎？ 請確認...." + "');if(!ok){return false;}"
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            Else
                BOK.Visible = False
            End If
            '遲納管理
            If dtFlow.Rows(0)("Delay") = 1 Then
                wDelay = 1
            End If
        End If
        '
        If wFormSno > 0 And wStep > 3 Then    '判斷是否[簽核]
            sql = "Select * From T_WaitHandle "
            sql = sql & " Where Active = 1 "
            sql = sql & "   And FormNo =  '" & wFormNo & "'"
            sql = sql & "   And FormSno =  '" & CStr(wFormSno) & "'"
            sql = sql & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(sql)
            If dtWaitHandle.Rows.Count > 0 Then
                '說明設定
                Top = 670
                DDescSheet.Visible = True
                DDecideDesc.Visible = True
                If dtWaitHandle.Rows(0)("FlowType") = 0 Then
                    DDecideDesc.BackColor = Color.LightGray
                Else
                    DDecideDesc.BackColor = Color.GreenYellow
                    ShowRequiredFieldValidator("DDecideDescRqd", "DDecideDesc", "異常：需輸入說明")
                End If
                '遲納管理
                If wDelay = 1 Then
                    If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                        DDelay.Visible = True   '延遲Sheet
                        DReasonCode.Visible = True     '延遲理由代碼
                        DReasonCode.BackColor = Color.GreenYellow
                        DReason.Visible = True         '延遲理由
                        DReason.BackColor = Color.GreenYellow
                        DReasonDesc.Visible = True     '延遲其他說明
                        ShowRequiredFieldValidator("DReasonCodeRqd", "DReasonCode", "異常：需輸入延遲理由")
                        ShowRequiredFieldValidator("DReasonRqd", "DReason", "異常：需輸入延遲理由")
                        Top = Top + 110
                    Else
                        DDelay.Visible = False  '延遲Sheet
                        DReasonCode.Visible = False     '延遲理由代碼
                        DReason.Visible = False         '延遲理由
                        DReasonDesc.Visible = False     '延遲其他說明
                    End If
                End If

            End If
        Else
            Top = 670
            'Sheet隱藏
            DDescSheet.Visible = False  '說明Sheet
            DDelay.Visible = False      '延遲Sheet
            '欄位隱藏
            DDecideDesc.Visible = False    '說明
            DReasonCode.Visible = False    '延遲理由代碼
            DReason.Visible = False        '延遲理由
            DReasonDesc.Visible = False    '延遲其他說明
            DHistoryLabel.Visible = False  '核定履歷
        End If
        '按鈕及超連結設值


        BSAVE.Style("top") = Top & "px"
        BNG1.Style("top") = Top & "px"
        BNG2.Style("top") = Top & "px"
        BOK.Style("top") = Top & "px"
        DHistoryLabel.Style("top") = Top + 32 & "px"
        GridView2.Style("top") = Top + 32 + 16 & "px"
        '簽核,通知 移除控制設定

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
        Dim sql As String
        Sql = "Select Divname,Username From M_Users "
        Sql = Sql & " Where UserID = '" & wApplyID & "'"
        Sql = Sql & "   And Active = '1' "
        Dim DBUser As DataTable = uDataBase.GetDataTable(Sql)



        DName.Text = DBUser.Rows(0).Item("Username")
        DDepName.Text = DBUser.Rows(0).Item("Divname")

        If DTNo.Text = "" Then
            DTNo.Text = Now.ToString("yyyyMMddHHmmss") '虛擬單號
        End If


        'No
        Select Case FindFieldInf("NO")
            Case 0  '顯示
                DNo.BackColor = Color.White
                DNo.Visible = True
                DNo.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DNo.Visible = True
                DNo.BackColor = Color.GreenYellow
                DNo.ReadOnly = False
                ShowRequiredFieldValidator("DNoRqd", "DNo", "異常：需輸入Ｎｏ")
            Case 2  '修改
                DNo.Visible = True
                DNo.BackColor = Color.Yellow
                DNo.ReadOnly = False
            Case Else   '隱藏
                DNo.Visible = False
        End Select

        'Date
        Select Case FindFieldInf("Date")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.Visible = True
                DDate.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDate.Visible = True
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = False
                ShowRequiredFieldValidator("DDateRqd", "DDate", "異常：需輸入申請日期")
            Case 2  '修改
                DDate.Visible = True
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = False
            Case Else   '隱藏
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


        'Name
        Select Case FindFieldInf("Name")
            Case 0  '顯示
                DName.BackColor = Color.LightGray
                DName.Visible = True
                DName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DName.Visible = True
                DName.BackColor = Color.GreenYellow
                DName.ReadOnly = False
                ShowRequiredFieldValidator("DNameRqd", "DName", "異常：需輸入姓名")
            Case 2  '修改
                DName.Visible = True
                DName.BackColor = Color.Yellow
                DName.ReadOnly = False
            Case Else   '隱藏
                DName.Visible = False
        End Select





        'Select Case FindFieldInf("Division")
        Select Case FindFieldInf("DepName")
            Case 0  '顯示
                DDepName.BackColor = Color.LightGray
                DDepName.Visible = True
                DDepName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDepName.Visible = True
                DDepName.BackColor = Color.GreenYellow
                DDepName.ReadOnly = False
                ShowRequiredFieldValidator("DDepNameRqd", "DDepName", "異常：需輸入部門")
            Case 2  '修改
                DDepName.Visible = True
                DDepName.BackColor = Color.Yellow
                DDepName.ReadOnly = False
            Case Else   '隱藏
                DDepName.Visible = False
        End Select




        '類別
        Select Case FindFieldInf("Type")
            Case 0  '顯示
                DType.BackColor = Color.LightGray
                DType.Visible = True

            Case 1  '修改+檢查
                DType.BackColor = Color.GreenYellow
                ShowRequiredFieldValidator("DTypeRqd", "DType", "異常：需輸入類別")
                DType.Visible = True
            Case 2  '修改
                DType.BackColor = Color.Yellow
                DType.Visible = True
            Case Else   '隱藏
                DType.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("Type", "ZZZZZZ")



        '類別
        Select Case FindFieldInf("StockNo")
            Case 0  '顯示
                DStockNo.BackColor = Color.LightGray
                DStockNo.Visible = True

            Case 1  '修改+檢查
                DStockNo.BackColor = Color.GreenYellow
                ' ShowRequiredFieldValidator("DStockNoRqd", "DStockNo", "異常：需輸入輸入PO/PR/棧板號")
                DStockNo.Visible = True
            Case 2  '修改
                DStockNo.BackColor = Color.Yellow
                DStockNo.Visible = True
            Case Else   '隱藏
                DStockNo.Visible = False
        End Select
        If DType.SelectedValue = "EXCEL上傳" Then
            DStockNo.Visible = False
        End If
        If pPost = "New" Then SetFieldData("StockNo", "ZZZZZZ")

        If wStep = 1 Then
            LAttachfile.Visible = False
        Else
            LAttachfile.Visible = True
        End If

      
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
                    System.Configuration.ConfigurationManager.AppSettings("ExpensePath")
        Dim FileName As String = ""

        Dim SQL As String
        SQL = "Select * From F_StockInSheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtData.Rows(0).Item("Date")                         'No 
            DDepName.Text = dtData.Rows(0).Item("DepName")

            DName.Text = dtData.Rows(0).Item("Name")                         'No

            SetFieldData("Type", dtData.Rows(0).Item("Type"))    '延遲理由

            DStockNo.Text = dtData.Rows(0).Item("StockNo")


            If dtData.Rows(0).Item("AttachFile") <> "" Then
                LAttachfile.NavigateUrl = Path & dtData.Rows(0).Item("AttachFile") '折扣
                LAttachfile.Visible = True
            Else
                LAttachfile.Visible = False
            End If



            '交易資料
            SQL = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step =  '" & CStr(wStep) & "'"
            SQL = SQL & "   And SeqNo =  '" & CStr(wSeqNo) & "'"
            Dim dtWaitHandle As DataTable = uDataBase.GetDataTable(SQL)
            If dtWaitHandle.Rows.Count > 0 Then
                DDecideDesc.Text = dtWaitHandle.Rows(0)("DecideDesc")           '說明

                If dtWaitHandle.Rows(0)("BEndTime") < NowDateTime Then
                    If DDelay.Visible = True Then
                        SetFieldData("ReasonCode", dtWaitHandle.Rows(0)("ReasonCode"))    '延遲理由代碼
                        If dtWaitHandle.Rows(0)("ReasonCode") = "" Then
                            SetFieldData("Reason", DReasonCode.SelectedValue)    '延遲理由
                            DReasonDesc.Text = ""   '延遲其他說明
                        Else
                            DReason.Text = dtWaitHandle.Rows(0)("Reason")  '延遲理由
                            DReasonDesc.Text = dtWaitHandle.Rows(0)("ReasonDesc")     '延遲其他說明
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
            SQL = SQL + "Order by Order by CreateTime Desc, Step Desc, SeqNo Desc "
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
        Dim idx As Integer = FindFieldInf(pFieldName)
        Dim i As Integer
        '擔當者及部門 

      


        '類別
        If pFieldName = "Type" Then
            DType.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DType.Items.Add(ListItem1)
                End If
            Else
                sql = "  Select  * from M_referp"
                sql = sql & " where  cat = '3112'"
                sql = sql & " and dkey = 'Type'"
                sql = sql & " order by unique_id"

                Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
                DType.Items.Add("")
                For i = 0 To dtReferp.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReferp.Rows(i).Item("Data")
                    ListItem1.Value = dtReferp.Rows(i).Item("Data")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DType.Items.Add(ListItem1)
                Next
                dtReferp.Clear()
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
                sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dtReasonCode.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtReasonCode.Rows(i)("DKey")
                    ListItem1.Value = dtReasonCode.Rows(i)("DKey")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DReasonCode.Items.Add(ListItem1)
                Next
            End If
        End If
        '延遲理由
        If pFieldName = "Reason" Then
            sql = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim dtReason As DataTable = uDataBase.GetDataTable(sql)
            For i = 0 To dtReason.Rows.Count - 1
                DReason.Text = dtReason.Rows(i)("Data")
            Next
        End If
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
        rqdVal.Text = pMessage
        rqdVal.Display = ValidatorDisplay.Dynamic
        rqdVal.Style.Add("Top", Top & "px")
        rqdVal.Style.Add("Left", "8px")
        rqdVal.Style.Add("Height", "20px")
        rqdVal.Style.Add("Width", "250px")
        rqdVal.Style.Add("POSITION", "absolute")
        Me.Form.Controls.Add(rqdVal)
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
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()


    End Sub
    '##AgentApprovProc-End
    '##
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
        Dim Run As Boolean = True           '是否執行
        Dim RepeatRun As Boolean = False    '是否重覆執行
        Dim MultiJob As Integer = 0         '多人核定
        Dim wLevel As String = ""           '難易度
        '
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
                    DNo.Text = SetNo(NewFormSno)
                    '儲存表單資料
                    If NewFormSno <> 0 Then
                        AppendData(pFun, NewFormSno)  'Insert
                        AddCommissionNo(wFormNo, NewFormSno)
                    End If
                    If RtnCode <> 0 Then
                        ErrCode = 9110
                    Else
                        '申請流程資料建置(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                       
                        AddCommissionNo(wFormNo, wFormSno)
                    End If
                    pRunNextStep = 1
                Else
                    If RepeatRun = False Then   '不是通知的重覆執行
                        '更新交易資料
                        ModifyTranData(pFun, pSts)
                        '流程資料結束(表單號碼,表單流水號,工程關卡號碼,行事曆,簽核者, 流程結束否(會簽))
                        RtnCode = oFlow.CheckFlow(wFormNo, wFormSno, wStep, wDecideCalendar, Request.QueryString("pUserID"), pRunNextStep)

                        If RtnCode <> 0 Then
                            ErrCode = 9120
                        End If
                    Else
                        pRunNextStep = 1    '是通知的重覆執行
                    End If
                End If
            End If


            '取得下一關
            If ErrCode = 0 And pRunNextStep = 1 Then
                '指定簽核者User ID
                Dim wAllocateID As String = ""

                '取得簽核者
                '表單號碼,工程關卡號碼,簽核者,申請者,被代理者,被指定者,多人核定工程No,
                '下一工程的, 號碼, 擔當者, 被代理者, 人數, 處理方法, 動作(0:OK,1:NG,2:Save)

                RtnCode = oCommon.GetNextGate(wFormNo, wStep, Request.QueryString("pUserID"), wApplyID, wAgentID, wAllocateID, MultiJob, _
                                              pNextStep, pNextGate, pAgentGate, pCount, pFlowType, pAction)

                If RtnCode <> 0 Then ErrCode = 9130
                If pCount = 0 And pNextStep <> 999 Then ErrCode = 9131
                If ErrCode = 0 Then pAction = 0
            End If

            '建置流程資料
            If ErrCode = 0 And pRunNextStep = 1 Then
                '設定委託No
                If pNextStep = 999 Then     '工程結束嗎?
                    If pFun = "OK" Then ModifyData(pFun, CStr(wOKSts)) '更新表單資料
                    If pFun = "NG1" Then ModifyData(pFun, CStr(wNGSts1))
                    If pFun = "NG2" Then ModifyData(pFun, CStr(wNGSts2))
                Else
                    ModifyData(pFun, "0")         '更新表單資料 Sts=0(未結)
                End If

                If pNextStep <> 999 Then
                    wNextGate = ""
                    For i = 1 To pCount
                        '取得下一關人員(訊息時使用)
                        If wNextGate = "" Then
                            wNextGate = pNextGate(i)
                        Else
                            wNextGate = "," & pNextGate(i)
                        End If
                        '取得核定者-群組行是曆
                        wNextGateCalendar = oCommon.GetCalendarGroup(pNextGate(i))

                        '取得工程負荷最後日時(核定者, 表單號碼, 工程號碼, 類別(0:通知,1:核定), 開始日時, 最後日時, 件數)
                        oSchedule.GetLastTime(pNextGate(i), wFormNo, pNextStep, pFlowType, NowDateTime, pLastTime, pCount1)

                        '取得預定開始,完成日程計算(表單號碼,工程號碼,難易度,QC-L/T,現在時間, 預定開始日時, 預定完成日時, 行事曆)
                        oSchedule.GetLeadTime(wFormNo, pNextStep, wLevel, wQCLT, pLastTime, pStartTime, pEndTime, wNextGateCalendar)

                        '建置流程資料(表單號碼,表單流水號,工程關卡號碼,序號,申請者ID,行事曆,建置者, 簽核者, 被代理者, 申請者, 預定開始日, 預定完成日, 重要性)
                        RtnCode = oFlow.NextFlow(wFormNo, NewFormSno, pNextStep, i, wNextGateCalendar, Request.QueryString("pUserID"), pNextGate(i), pAgentGate(i), wApplyID, pStartTime, pEndTime, 0)

                        If RtnCode <> 0 Then
                            ErrCode = 9150
                            Exit For
                        End If
                    Next i
                Else
                    '流程結束(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                    RtnCode = oFlow.EndFlow(wFormNo, wFormSno, pNextStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)

                    If RtnCode <> 0 Then ErrCode = 9160
                End If
            End If
            '當工程日程調整
            If ErrCode = 0 Then
                If RepeatRun = False Then
                    '簽核者,表單號碼,表單流水號,工程關卡號碼,序號,現在日時,難易度,行事曆
                    RtnCode = oSchedule.AdjustSchedule(Request.QueryString("pUserID"), wFormNo, wFormSno, wStep, wSeqNo, NowDateTime, wLevel, wDecideCalendar)
                End If
            End If

            If ErrCode = 0 Then
                '傳送郵件
                If pNextStep <> 999 Then
                    For i = 1 To pCount
                        'oCommon.Send(Request.QueryString("pUserID"), pNextGate(i), wApplyID, wFormNo, NewFormSno, pNextStep, "FLOW")
                        '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼,訊息類別
                    Next i
                Else
                    'oCommon.Send(Request.QueryString("pUserID"), wApplyID, wApplyID, wFormNo, NewFormSno, pNextStep, "END")
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
                uJavaScript.PopMsg(Me, Message)
            End If      '儲存表單ErrCode=0
        End While       '重覆執行

        If ErrCode = 0 Then
            '--郵件傳送---------
            'oCommon.SendMail()
            '
            Dim URL As String
            If DType.SelectedValue = "EXCEL上傳" Then
                URL = "StockInSheet_04.aspx?&pFormSno=" & wFormSno & "&pUserID=" & Request.QueryString("pUserID")
            Else
                URL = "StockInSheet_03.aspx?&pFormSno=" & wFormSno
            End If

            Response.Redirect(URL)

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AppendData)
    '**     新增表單資料
    '**
    '*****************************************************************
    Sub AppendData(ByVal pFun As String, ByVal NewFormSno As Integer)
        Dim Code As String

        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                           CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                    System.Configuration.ConfigurationManager.AppSettings("StockInPath")
        Dim sql As String


        sql = " Insert into F_StockInSheet (Sts, CompletedTime, FormNo, FormSno,"
        sql &= " NO,Date,Name,DepName,Type,StockNo,StockIn,Attachfile,"
        sql = sql + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + "VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日PO00000283
        sql = sql + " '003112', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號
        sql = sql + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDate.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DName.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDepName.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DType.Text) + "', "   'NO  1    
        sql = sql + " N'" + YKK.ReplaceString(UCase(DStockNo.Text)) + "', "   'NO  1
        sql = sql + " 0, "   'NO  1


        If AddFileName <> "" Then
            sql = sql + "'" + AddFileName + "'," '附檔1
        Else
            sql = sql + "''," '附檔1
        End If
        'If DAttachfile.Visible = True Then
        '    '   And LCertifcateFile.NavigateUrl = "" Then
        '    If (DAttachfile.PostedFile.FileName <> "") Then

        '        '20170912 將檔名修改成不含原始檔名
        '        Dim fileExtension As String  '副檔名
        '        fileExtension = IO.Path.GetExtension(DAttachfile.PostedFile.FileName).ToLower   '取得檔案格式

        '        FileName = CStr(NewFormSno) + "-Attachfile1-" + UploadDateTime + fileExtension

        '        If DAttachfile.PostedFile.FileName = "" Then
        '            '  FileName = Right(LDISPOSALFILE.NavigateUrl, InStr(StrReverse(LDISPOSALFILE.NavigateUrl), "\") - 1)
        '            DAttachfile.PostedFile.SaveAs(Path + FileName)
        '        Else
        '            ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime & "-" & Right(DAttachFile1.PostedFile.FileName, InStr(StrReverse(DAttachFile1.PostedFile.FileName), "\") - 1)
        '            ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension
        '            DAttachfile.PostedFile.SaveAs(Path + FileName)
        '        End If

        '    Else
        '        FileName = ""
        '    End If
        '    If FileName <> "" Then
        '        sql = sql + "'" + FileName + "'," '附檔1
        '    End If
        'Else
        '    sql = sql + "''," '附檔1
        'End If


        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)

        '明細新增
        '存入明細

      
        sql = " insert into F_StockInSheetdt (No,Type,CheckStockIn,StockInDate,ItemCode,Color,Size,Length,QTY,UNIT,BOXNO,STOCKNO,WINGSTS,CreateUser, CreateTime,ModifyUser, ModifyTime) "
        sql = sql + " select  '" + YKK.ReplaceString(DNo.Text) + "',Type,CheckStockIn,StockInDate,ItemCode,Color,Size,Length,QTY,UNIT,BOXNO,STOCKNO,0, "
        sql = sql + " '" + Request.QueryString("pUserID") + "', "        '作成者
        sql = sql + " '" + NowDateTime + "', "                            '作成時間
        sql = sql + " '" + "" + "', "                                     '修改者
        sql = sql + " '" + NowDateTime + "' "                             '修改時間
        sql = sql + "  from F_StockInSheetdttemp where NO = '" + DTNo.Text + "'"
        uDataBase.ExecuteNonQuery(sql)

        '存到wings 
        If DType.SelectedValue = "EXCEL上傳" Then
            sql = "  Select  * from F_StockInSheetdt"
            sql = sql & " where  NO ='" + YKK.ReplaceString(DNo.Text) + "'"


            Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)
            Dim I As Integer
            For i = 0 To dtReferp.Rows.Count - 1
                Code = oWaves.StockInNew2Wings(UCase(dtReferp.Rows(I).Item("stockno")))
            Next
        Else
            Code = oWaves.StockInNew2Wings(UCase(DStockNo.Text))
        End If

        If Code <> "0" Then
            uJavaScript.PopMsg(Me, "匯入WINS異常，請洽電腦室")

        End If
       


    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)
        Dim sql As String
        Dim FileName As String
        Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                        CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
        System.Configuration.ConfigurationManager.AppSettings("StockInPath")



        sql = "Update  F_StockInSheet  Set "
        If pFun <> "SAVE" Then
            sql &= " Sts = '" & pSts & "',"
            sql &= " CompletedTime = '" & NowDateTime & "',"
        End If
        sql &= " FormNo = '" & wFormNo & "',"
        sql &= " FormsNo = '" & wFormSno & "',"
        sql &= " No = '" & DNo.Text & "',"
        sql &= " Date = '" & DDate.Text & "',"
        sql &= " Name = '" & DName.Text & "',"
        sql &= " DepName = '" & DDepName.Text & "',"
        sql &= " Type = '" & DType.Text & "',"
        sql &= " StockNo = '" & DStockNo.Text & "',"
        FileName = ""
        If DAttachfile.Visible = True Then
            '   And LCertifcateFile.NavigateUrl = "" Then
            If DAttachfile.PostedFile.FileName <> "" Then

                '20170912 將檔名修改成不含原始檔名
                Dim fileExtension As String  '副檔名
                fileExtension = IO.Path.GetExtension(DAttachfile.PostedFile.FileName).ToLower   '取得檔案格式

                FileName = CStr(wFormSno) + "-Attachfile1-" + UploadDateTime + fileExtension

                If DAttachfile.PostedFile.FileName = "" Then
                    '  FileName = Right(LDISPOSALFILE.NavigateUrl, InStr(StrReverse(LDISPOSALFILE.NavigateUrl), "\") - 1)
                    DAttachfile.PostedFile.SaveAs(Path + FileName)
                Else
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime & "-" & Right(DAttachFile1.PostedFile.FileName, InStr(StrReverse(DAttachFile1.PostedFile.FileName), "\") - 1)
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension
                    DAttachfile.PostedFile.SaveAs(Path + FileName)
                End If

            Else
                FileName = ""
            End If
            If FileName <> "" Then
                sql = sql + " Attachfile='" + FileName + "'," '附檔1
            End If

        End If

      
        sql &= " ModifyUser = '" & Request.QueryString("pUserID") & "',"
        sql &= " ModifyTime = '" & NowDateTime & "' "
        sql &= " Where FormNo  =  '" & wFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        '
        uDataBase.ExecuteNonQuery(sql)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Add T_CommissionNo)
    '**     追加交易資料和委託單對照表
    '**
    '*****************************************************************
    Sub AddCommissionNo(ByVal pFormNo As String, ByVal pFormSno As Integer)
        Dim SQl As String
        SQl = "Select * From T_CommissionNo "
        SQl = SQl & " Where FormNo =  '" & pFormNo & "'"
        SQl = SQl & "   And FormSno =  '" & CStr(pFormSno) & "'"
        Dim dtCommissionNo As DataTable = uDataBase.GetDataTable(SQl)

        If dtCommissionNo.Rows.Count <= 0 Then
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
                uDataBase.ExecuteNonQuery(SQl)
            End If
        Else
            If DNo.Text <> "" Then
                If DNo.Text <> dtCommissionNo.Rows(0)("No") Then  'No
                    SQl = "Update T_CommissionNo Set "
                    SQl = SQl + " No = '" & DNo.Text & "',"
                    SQl = SQl + " MapNo = '" & "" & "',"
                    SQl = SQl + " CreateUser = '" & Request.QueryString("pUserID") & "',"
                    SQl = SQl + " CreateTime = '" & NowDateTime & "' "
                    SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
                    SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
                    uDataBase.ExecuteNonQuery(SQl)
                End If
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Sub ModifyTranData(ByVal pFun As String, ByVal pSts As String)
        Dim SQl As String
        If pFun <> "SAVE" Then      '<> Save
            SQl = "Update T_WaitHandle Set "
            SQl = SQl + " Active = '" & "0" & "',"
            SQl = SQl + " Sts = '" & pSts & "',"
            If pSts = "1" Then SQl = SQl + " StsDes = '" & BOK.Text & "',"
            If pSts = "2" Then SQl = SQl + " StsDes = '" & BNG1.Text & "',"
            If pSts = "3" Then SQl = SQl + " StsDes = '" & BNG2.Text & "',"

            SQl = SQl + " AEndTime = '" & NowDateTime & "',"
            SQl = SQl + " CompletedTime = '" & NowDateTime & "',"
            If DDelay.Visible = True Then
                SQl = SQl + " ReasonCode = '" & DReasonCode.SelectedValue & "',"
                SQl = SQl + " Reason = '" & DReason.Text & "',"
                SQl = SQl + " ReasonDesc = N'" & DReasonDesc.Text & "',"
            End If
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
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
            SQl = SQl + " DecideDesc = N'" & uCommon.ReplaceString(DDecideDesc.Text) & "',"
            SQl = SQl + " ModifyUser = '" & Request.QueryString("pUserID") & "',"
            SQl = SQl + " ModifyTime = '" & NowDateTime & "' "
            SQl = SQl + " Where FormNo  =  '" & wFormNo & "'"
            SQl = SQl + "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQl = SQl + "   And Step =  '" & CStr(wStep) & "'"
            SQl = SQl + "   And SeqNo =  '" & CStr(wSeqNo) & "'"
        End If
        uDataBase.ExecuteNonQuery(SQl)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BSAVE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSAVE.Click

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-2按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG2.Click
        If InputDataOK(2) Then
            DisabledButton()   '停止Button運作
            FlowControl("NG2", 2, "3")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     NG-1按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BNG1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNG1.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            FlowControl("NG1", 1, "2")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     OK按鈕點選後事件
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        If InputDataOK(1) Then
            DisabledButton()   '停止Button運作
            FlowControl("OK", 0, "1")
        Else
            EnabledButton()   '起動Button運作
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     停止Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub DisabledButton()
        BOK.Enabled = False
        BNG1.Enabled = False
        BNG2.Enabled = False
        BSAVE.Enabled = False
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     起動Button(OK, Ng1, NG2, Save)運作 
    '**
    '*****************************************************************
    Private Sub EnabledButton()
        BOK.Enabled = True
        BNG1.Enabled = True
        BNG2.Enabled = True
        BSAVE.Enabled = True
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
    '**(GetDataStatus)
    '**     取得表單進度狀態
    '**
    '*****************************************************************
    Sub GetDataStatus()
        Dim SQL As String
        SQL = "Select * From M_Flow "
        SQL = SQL & " Where Active = 1 "
        SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And Step   =  '" & wStep & "'"
        Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
        If dtFlow.Rows.Count > 0 Then
            'NG-1按鈕
            If dtFlow.Rows(0)("NGFun1") = 1 Then
                wNGSts1 = dtFlow.Rows(0)("NGSts1") + 1
            End If
            'NG-2按鈕
            If dtFlow.Rows(0)("NGFun2") = 1 Then
                wNGSts2 = dtFlow.Rows(0)("NGSts2") + 1
            End If
            'OK按鈕
            If dtFlow.Rows(0)("OKFun") = 1 Then
                wOKSts = dtFlow.Rows(0)("OKSts") + 1
            End If
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As FileUpload) As Integer
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
                UPFileIsNormal = 9020
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9030
            End If
        End If
    End Function

    '*****************************************************************
    '**
    '**     判斷是否可繼續執行(驗證資料)
    '**
    '*****************************************************************
    Function InputDataOK(ByVal pAction As Integer) As Boolean
        Dim isOK As Boolean = False
        Dim ErrCode As Integer = 0   '宣告一個Error用來判別是否資料正常，預設為False
        Dim Message As String = ""
        Dim SQL As String


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
                ErrCode = oCommon.CommissionNo("003112", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If

        End If

        'Check上傳附件Size及格式
        If ErrCode = 0 Then
            If DAttachfile.Visible Then
                If DAttachfile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
                    ErrCode = UPFileIsNormal(DAttachfile)
                End If
            End If
        End If


        '檢查是否有附檔
        If ErrCode = 0 Then
            If DType.SelectedValue = "EXCEL上傳" And DAttachfile.FileName = "" Then
                If ErrCode <> 0 Then
                    ErrCode = 9020
                End If
            End If

        End If


        If DType.SelectedValue = "EXCEL上傳" Then
            If DAttachfile.FileName <> "" Then
                'upload()
                'Insert()

            Else
                'ErrCode = 9020

            End If

        Else
            If Left(DType.SelectedValue, 2) <> "PO" And Left(DType.SelectedValue, 2) <> "PR" Then
                ErrCode = 9040
            Else
                '  Waveupload()
                Insert()
            End If

        End If


        '檢查PO/PR/棧板號重覆
        If wStep = 1 And ErrCode = 0 Then
            If DType.SelectedValue <> "EXCEL上傳" Then
                SQL = "Select * From F_StockInSheet "
                SQL &= " Where stockNo ='" & DStockNo.Text & "'"
                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
                If dtFlow.Rows.Count > 0 Then
                    ErrCode = 9010
                End If
            Else
                SQL = SQL + " select  * from  "
                SQL = SQL + " F_StockInSheetdttemp where NO = '" + DTNo.Text + "'"
                SQL = SQL + " and stockno in ("
                SQL = SQL + " select stockno from F_StockInSheetdt)"
                Dim dtFlow As DataTable = uDataBase.GetDataTable(SQL)
                If dtFlow.Rows.Count > 0 Then
                    ErrCode = 9010
                End If
            End If
        End If


      


        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "PO/PR/棧板號重覆,請確認!"
            If ErrCode = 9020 Then Message = "請上傳檔案!"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!"
            If ErrCode = 9040 Then Message = "請確認棧板號是否正確!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            If ErrCode = 9073 Then Message = "金額需大於0，請確認!"

            If ErrCode = 9075 Then Message = "非數字格式,請確認!"
            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '
        Return isOK
    End Function



    Sub GetData()


        Dim SQL As String
        SQL = " Select unique_id,type,stockno,color,boxno,itemcode,convert(decimal(9,0),qty)qty,case when convert(char(10),StockInDate,112) = '19000101' then '' else   convert(char(10),StockInDate,112)  end  as Etime from   F_StockInSheetdttemp where  "
        SQL = SQL + " NO = '" + DTNo.Text + "'"
        SQL = SQL + " order by unique_id,StockInDate "
        Dim dtdData As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = dtdData
        GridView1.DataBind()

        SQL = " Select ItemCode, Color, Size, Length, convert(decimal(9,0),qty)qty, Unit, BOXNO, STOCKNO,case when convert(char(10),StockInDate,112) = '19000101' then '' else   convert(char(10),StockInDate,112)  end  as Etime from   F_StockInSheetdttemp where  "
        SQL = SQL + " NO = '" + DTNo.Text + "'"
        SQL = SQL + " order by unique_id,StockInDate "
        Dim dtdData1 As DataTable = uDataBase.GetDataTable(SQL)
        GridView3.DataSource = dtdData1
        GridView3.DataBind()


    End Sub
 

    Protected Sub DType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DType.SelectedIndexChanged
        If DType.SelectedValue = "EXCEL上傳" Then
            DAttachfile.BackColor = Color.Yellow
            DAttachfile.Visible = True
            DStockNo.Visible = False
            BAdd.Visible = False
           
        Else
            DAttachfile.BackColor = Color.LightGray
            DAttachfile.Visible = False
            DStockNo.Visible = True
            BAdd.Visible = True

        End If
        GridView1.DataSource = Nothing
        GridView1.DataBind()
        GridView3.DataSource = Nothing
        GridView3.DataBind()
    End Sub
    Function upload() As Boolean
        'Try
        If DAttachfile.HasFile Then

            '上傳附檔
            Dim FileName1 As String
            UploadName = DAttachfile.PostedFile.FileName

            Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                          CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
            ' Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("DISPOSALData1")
            Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                            System.Configuration.ConfigurationManager.AppSettings("StockInData")
            'System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
            '               System.Configuration.ConfigurationManager.AppSettings("DISPOSALData")

            '20170912 將檔名修改成不含原始檔名
            Dim fileExtension As String  '副檔名
            fileExtension = IO.Path.GetExtension(UploadName).ToLower   '取得檔案格式
            FileName1 = Path1 + CStr(DNo.Text) + UploadDateTime + fileExtension
            DAttachfile.PostedFile.SaveAs(FileName1)

            '展開
            Dim FileName As String = Path.GetFileName(DAttachfile.PostedFile.FileName)
            AddFileName = FileName
            Dim Extension As String = Path.GetExtension(DAttachfile.PostedFile.FileName)
            '  Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

            ' FileName1 = "http://10.245.1.6/DASW/Document/006002/" + CStr(DNo.Text) + UploadDateTime + fileExtension
            'Dim FilePath As String = Server.MapPath(FolderPath + FileName)
            ' DFileUpload1.SaveAs(FilePath)
            'Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text)
            Import_To_Grid(FileName1, Extension, rbHDR.SelectedItem.Text)


        End If
        'Catch ex As Exception
        '    uJavaScript.PopMsg(Me, "上傳檔案格式有誤upload,請確認!")
        '    griderror = False
        'End Try


    End Function


    Private Sub Import_To_Grid(ByVal FilePath As String, ByVal Extension As String, ByVal isHDR As String)
        Dim conStr As String = ""
        Select Case Extension
            Case ".xls"
                'Excel 97-03
                conStr = ConfigurationManager.ConnectionStrings("Excel03ConString").ConnectionString
                Exit Select
            Case ".xlsx"
                'Excel 07
                conStr = ConfigurationManager.ConnectionStrings("Excel07ConString").ConnectionString
                Exit Select
        End Select
        conStr = String.Format(conStr, FilePath, isHDR)

        Dim connExcel As New OleDbConnection(conStr)
        Dim cmdExcel As New OleDbCommand()
        Dim oda As New OleDbDataAdapter()
        Dim dt As New DataTable()

        cmdExcel.Connection = connExcel

        'Get the name of First Sheet
        connExcel.Open()
        Dim dtExcelSchema As DataTable
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        Dim i As Integer
        Dim SheetName As String = ""

        For i = 0 To dtExcelSchema.Rows.Count - 1
            If dtExcelSchema.Rows.Count > 1 Then
                SheetName = dtExcelSchema.Rows(1)("TABLE_NAME").ToString()
                If SheetName = "StockIn$" Then
                Else
                    SheetName = "StockIn$"
                End If
            Else
                SheetName = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            End If

        Next



        connExcel.Close()

        'Read Data from First Sheet
        connExcel.Open()
        cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
        oda.SelectCommand = cmdExcel
        oda.Fill(dt)
        connExcel.Close()

        'Bind Data to GridView
        '顯示用
        GridView3.Caption = Path.GetFileName(FilePath)
        GridView3.DataSource = dt
        GridView3.DataBind()

        'GridView1.DataSource = dt
        'GridView1.DataBind()
    End Sub


    Sub Insert()
        '檢查是否有同月份同申請人的檔案，若有則將之前的全刪除
        Dim a As String = ""
        Dim uploadflag As Integer = 1
        Dim nullflag As Integer = 1
        Dim sno As String
        Dim NewFormSno As Integer = wFormSno    '表單流水號
        Dim SQL As String
       

        ' If uploadflag = 1 Then
        Try
            sno = DNo.Text
           
            wApplyID = Request.QueryString("pApplyID")  '申請者ID

            SQL = " delete from  F_StockInSheetdtTEMP where No='" + DTNo.Text + "'"
            uDataBase.ExecuteNonQuery(SQL)

            '上傳到資料庫
            Dim j As Integer
            Dim jSQL As String
            jSQL = ""

            Dim i As Integer

            For i = 0 To Me.GridView3.Rows.Count - 1 Step i + 1



                SQL = "Insert into F_StockInSheetdtTEMP (No,Type,CheckStockIn,ItemCode,Color,Size,Length,QTY,UNIT,BOXNO,STOCKNO,StockInDate,CreateTime,CreateUser)"
                SQL = SQL + " values('" + YKK.ReplaceString(DTNo.Text) + "','" + DType.SelectedValue + "','',"

                For j = 0 To 8

                    a = GridView3.Rows(i).Cells(0).Text
                    If j = 0 Then

                        If GridView3.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = "''"
                        Else
                            jSQL = "'" + YKK.ReplaceString(GridView3.Rows(i).Cells(j).Text) + "'"

                        End If
                    Else

                        If GridView3.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = jSQL + "," + "''"
                        Else
                            jSQL = jSQL + ",'" + YKK.ReplaceString(GridView3.Rows(i).Cells(j).Text) + "'"

                        End If

                    End If



                    If a = "&nbsp;" Then  '檢查第一欄是不是NULL
                        nullflag = 0
                    End If


                Next

                SQL = SQL + jSQL + ","
                SQL = SQL + "getdate(),'" + Request.QueryString("pUserID") + "')"

                'If CStr(i + 1) = "51" Then
                'uJavaScript.PopMsg(Me, a)
                'End If


                If nullflag = 1 Then '第一欄是空白就不要匯入

                    uDataBase.ExecuteNonQuery(SQL)

                End If


            Next


            ' uJavaScript.PopMsg(Me, "上傳成功")


         
        Catch ex As Exception

            uJavaScript.PopMsg(Me, "Insert 有誤,請確認!")

            inserterror = False
        End Try

        '  End If


        GridView3.DataSource = Nothing
        GridView3.DataBind()




    End Sub



    Sub Waveupload()
        Dim sql As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection

      

 
        If DType.Text = "PO NO" Then
            sql = " SELECT  ITMCEM as ITEMCODE,CLRCEM as COLOR,'' as SIZE,'' as LENGTH,PODQEM as QTY, PQUCEM as UNIT,'' as BOXNO,'" + DStockNo.Text + "' as STOCKNO,ETAUEM as ETime FROM WAVEDLIB.PEM00 "
            sql = sql + "  WHERE PODNEM ='" + DStockNo.Text + "'"
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("WAVESOLEDB")
        ElseIf DType.Text = "PR NO" Then
            sql = " SELECT  ITMC9G as ITEMCODE,CLRC9G as COLOR,'' as SIZE,'' AS LENGTH,PSHQ9G as QTY,QUNC9G as UNIT,'' as BOXNO,'" + DStockNo.Text + "' as STOCKNO,'' as ETIME FROM WAVEDLIB.F9G00 "
            sql = sql + "  WHERE PSHN9G ='" + DStockNo.Text + "'"
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("WAVESOLEDB")


        End If

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(sql, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Data")
        GridView3.DataSource = DBDataSet1
        GridView3.DataBind()
        ' GridView1.DataSource = DBDataSet1
        ' GridView1.DataBind()

        OleDbConnection1.Close()

    End Sub

    Protected Sub BAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BAdd.Click
        Waveupload()
        Insert()
        GetData()
    End Sub
 


   

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim l As Object
            ' e.Row.Controls(4) is Delete button.
            l = e.Row.Controls(0)
            l.Attributes.Add("onclick", "javascript:return confirm('確認要刪除嗎?')")
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String

        Dim id As String = GridView1.DataKeys(e.RowIndex).Value

        SQL = " delete  from   F_StockInSheetdtTEMP "
        SQL = SQL & " where unique_id = " & id & " "
        uDataBase.ExecuteNonQuery(SQL)
        '


        GetData()
    End Sub

 


    Protected Sub btn_upload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_upload.Click
        If DAttachfile.FileName <> "" Then
            upload()
            Insert()
            GetData()
        End If


    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(1).Visible = False
    End Sub
End Class
