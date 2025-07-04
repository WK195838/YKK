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


Partial Class RPAStockModifySheet_01
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
    Dim wUserIP As String = ""
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
    Dim UploadName As String
    Dim griderror, inserterror, fielderror, qtyerror As Boolean
    Dim AttachFileName As String
    Dim sourceDir, backupDir As String
    Dim Message As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
            Top = 350
            Dim SQL As String = "Select * From T_WaitHandle "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo =  '" & wFormNo & "'"
            SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
            SQL = SQL & "   And Step   =  '" & CStr(wStep) & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("BEndTime").ToString < NowDateTime Then
                    If DDelay.Visible = True Then
                        Top = 350
                    End If
                End If
            End If
        Else
            Top = 350
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
            Top = Top - 80
        End If

        If DDelay.Visible Then                                          ' 延遲說明
            DDelay.Style("top") = Top & "px"
            DReasonCode.Style("top") = Top + 9 & "px"
            DReason.Style("top") = Top + 6 & "px"
            DReasonDesc.Style("top") = Top + 39 & "px"
            Top += 161
        End If


        BSAVE.Style("top") = Top + 200 & "px"
        BNG1.Style("top") = Top + 200 & "px"
        BNG2.Style("top") = Top + 200 & "px"
        BOK.Style("top") = Top + 200 & "px"

        Top = Top + 350
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
        Response.Cookies("PGM").Value = "EApprovalSheet_01.aspx"
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
                Top = 350
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
            Top = 350
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


        BSAVE.Style("top") = Top + 32 & "px"
        BNG1.Style("top") = Top + 32 & "px"
        BNG2.Style("top") = Top + 32 & "px"
        BOK.Style("top") = Top + 32 & "px"
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
            Case 2  '修改
                DNo.Visible = True
                DNo.BackColor = Color.Yellow
                DNo.ReadOnly = False
            Case Else   '隱藏
                DNo.Visible = False
        End Select
        If pPost = "New" Then DNo.Text = SetNo(wFormSno)

        'Date
        Select Case FindFieldInf("DATE")
            Case 0  '顯示
                DDate.BackColor = Color.LightGray
                DDate.Visible = True
                DDate.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDate.Visible = True
                DDate.BackColor = Color.GreenYellow
                DDate.ReadOnly = False
            Case 2  '修改
                DDate.Visible = True
                DDate.BackColor = Color.Yellow
                DDate.ReadOnly = False
            Case Else   '隱藏
                DDate.Visible = False
        End Select
        If pPost = "New" Then DDate.Text = Now.ToString("yyyy/MM/dd") '現在日時


        'Name
        Select Case FindFieldInf("APPNAME")
            Case 0  '顯示
                DAppName.BackColor = Color.LightGray
                DAppName.Visible = True
                DAppName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAppName.Visible = True
                DAppName.BackColor = Color.GreenYellow
                DAppName.ReadOnly = False
            Case 2  '修改
                DAppName.Visible = True
                DAppName.BackColor = Color.Yellow
                DAppName.ReadOnly = False
            Case Else   '隱藏
                DAppName.Visible = False
        End Select


        'DepID
        Select Case FindFieldInf("DEPID")
            Case 0  '顯示
                DDepID.BackColor = Color.LightGray
                DDepID.Visible = True
                DDepID.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDepID.BackColor = Color.GreenYellow
                DDepID.Visible = True
            Case 2  '修改
                DDepID.BackColor = Color.Yellow
                DDepID.Visible = True
            Case Else   '隱藏
                DDepID.Visible = False
        End Select

        'DepName
        Select Case FindFieldInf("DEPNAME")
            Case 0  '顯示
                DDepName.BackColor = Color.LightGray
                DDepName.Visible = True
                DDepName.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DDepName.BackColor = Color.GreenYellow
                DDepName.Visible = True
            Case 2  '修改
                DDepName.BackColor = Color.Yellow
                DDepName.Visible = True
            Case Else   '隱藏
                DDepName.Visible = False
        End Select


        'Remark
        Select Case FindFieldInf("REMARK")
            Case 0  '顯示
                DRemark.BackColor = Color.LightGray
                DRemark.Visible = True
                DRemark.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRemark.Visible = True
                DRemark.BackColor = Color.GreenYellow
                DRemark.ReadOnly = False
                ShowRequiredFieldValidator("DRemarkRqd", "DRemark", "異常：需輸入備註")
            Case 2  '修改
                DRemark.Visible = True
                DRemark.BackColor = Color.Yellow
                DRemark.ReadOnly = False
            Case Else   '隱藏
                DRemark.Visible = False
        End Select

        '附檔
        Select Case FindFieldInf("FILEUPLOAD1")
            Case 0  '顯示
                DFileUpload1.Visible = False
                If LAttachfile.NavigateUrl = "" Then
                    LAttachfile.Visible = False
                End If
                DFileUpload1.BackColor = Color.LightGray
                DFileUpload1.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DFileUpload1.BackColor = Color.GreenYellow
                DFileUpload1.Visible = True
                LAttachfile.Visible = False
                ShowRequiredFieldValidator("DFILEUPLOAD1Rqd", "DFILEUPLOAD1", "異常：需上傳在庫修正表")
            Case 2  '修改
                DFileUpload1.Visible = True
                DFileUpload1.BackColor = Color.Yellow
                LAttachfile.Visible = False
            Case Else   '隱藏
                DFileUpload1.Visible = True
                LAttachfile.Visible = True
        End Select


        'RPAREASON
        Select Case FindFieldInf("RPAREASON")
            Case 0  '顯示
                DRPAREASON.BackColor = Color.LightGray
                DRPAREASON.Visible = True
                DRPAREASON.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRPAREASON.Visible = True
                DRPAREASON.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DRPAREASONRqd", "DRPAREASON", "異常：需輸入原因")
            Case 2  '修改
                DRPAREASON.Visible = True
                DRPAREASON.BackColor = Color.Yellow

            Case Else   '隱藏
                DRPAREASON.Visible = False
        End Select
        If pPost = "New" Then SetFieldData("RPAREASON", "ZZZZZZ")


        'AMOUNT
        Select Case FindFieldInf("AMOUNT")
            Case 0  '顯示
                DAMOUNT.BackColor = Color.LightGray
                DAMOUNT.Visible = True
                DAMOUNT.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DAMOUNT.Visible = True
                DAMOUNT.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DAMOUNTRqd", "DAMOUNT", "異常：需輸入金額")
            Case 2  '修改
                DAMOUNT.Visible = True
                DAMOUNT.BackColor = Color.Yellow

            Case Else   '隱藏
                DAMOUNT.Visible = False
        End Select



        'RPAComment
        Select Case FindFieldInf("RPAComment")
            Case 0  '顯示
                DRPAComment.BackColor = Color.LightGray
                DRPAComment.Visible = True
                DRPAComment.Attributes.Add("readonly", "true")
            Case 1  '修改+檢查
                DRPAComment.Visible = True
                DRPAComment.BackColor = Color.GreenYellow

                ShowRequiredFieldValidator("DRPACommentRqd", "DRPAComment", "異常：需輸入金額")
            Case 2  '修改
                DRPAComment.Visible = True
                DRPAComment.BackColor = Color.Yellow

            Case Else   '隱藏
                DRPAComment.Visible = False
        End Select



        'RPAFILE
        Select Case FindFieldInf("RPAFILE")
            Case 0  '顯示

                LRPAFILE.Visible = True

            Case 1  '修改+檢查
                LRPAFILE.Visible = False

            Case 2  '修改
                LRPAFILE.Visible = False

            Case Else   '隱藏
                LRPAFILE.Visible = False

        End Select





        '擔當者及部門 

        Try
            Dim sql As String = ""
            sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
            sql = sql & "  where a.empname =b.username"
            sql = sql & " and UserID = '" & wApplyID & "'"
            sql = sql & "   And Active = '1' "
            Dim DBUser As DataTable = uDataBase.GetDataTable(sql)

            If wApplyID = "gas004" And DBUser.Rows(0).Item("EmpName") = "陳怡君" Then
                sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
                sql = sql & "  where a.empname =b.username "
                sql = sql & " and UserID = '" & wApplyID & "'"
                sql = sql & " and a.EmpID = '260465' "
                sql = sql & "   And Active = '1' "
                Dim DBUser1 As DataTable = uDataBase.GetDataTable(sql)
                DDepID.Text = DBUser1.Rows(0).Item("DepID")
                DDepName.Text = DBUser1.Rows(0).Item("Depname")
                DAppName.Text = DBUser1.Rows(0).Item("Empname")
            ElseIf wApplyID = "sl026" And DBUser.Rows(0).Item("EmpName") = "陳怡君" Then
                sql = " Select a.* From (Select * From m_wfsemp union all Select  *  From m_wfsempCL) a ,m_users b"
                sql = sql & "  where a.empname =b.username "
                sql = sql & " and UserID = '" & wApplyID & "'"
                sql = sql & " and a.EmpID = '260195' "
                sql = sql & "   And Active = '1' "
                Dim DBUser1 As DataTable = uDataBase.GetDataTable(sql)
                DDepID.Text = DBUser1.Rows(0).Item("DepID")
                DDepName.Text = DBUser1.Rows(0).Item("Depname")
                DAppName.Text = DBUser1.Rows(0).Item("Empname")
            Else
                DDepID.Text = DBUser.Rows(0).Item("DepID")
                DDepName.Text = DBUser.Rows(0).Item("Depname")
                DAppName.Text = DBUser.Rows(0).Item("Empname")
            End If


        Catch
            uJavaScript.PopMsg(Me, wApplyID + "-無此員工資料，請連絡電腦室")
        End Try


    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()

        Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                            System.Configuration.ConfigurationManager.AppSettings("RPAStockPath")  'WIS-TempPath


        Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("http") & _
                            System.Configuration.ConfigurationManager.AppSettings("RPARefPath")  'WIS-TempPath


        Dim SQL As String
        SQL = "Select * From F_RPAStockModifySheet "
        SQL &= " Where FormNo =  '" & wFormNo & "'"
        SQL &= "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            '表單資料
            DNo.Text = dtData.Rows(0).Item("No")                         'No
            '   DBarCode.ImageUrl = "imga.aspx?mycode=" & DNo.Text  BARCODE
            DDate.Text = dtData.Rows(0).Item("Date")
            DAppName.Text = dtData.Rows(0).Item("AppName") 'No

            DRemark.Text = dtData.Rows(0).Item("Remark")                         'No

            SetFieldData("RPAREASON", dtData.Rows(0).Item("RPAREASON"))
            DAMOUNT.Text = dtData.Rows(0).Item("AMOUNT")

            If dtData.Rows(0).Item("FileUpload1") <> "" Then
                LAttachfile.NavigateUrl = Path & dtData.Rows(0).Item("FileUpload1") '折扣
                LAttachfile.Visible = True
            Else
                LAttachfile.Visible = False
            End If


           


            If dtData.Rows(0).Item("RPASts") <> 0 Then
                DRPAComment.Text = dtData.Rows(0).Item("RPAComment") 'No
                LRPAFILE.Visible = True
                LRPAFILE.NavigateUrl = Path1 & DNo.Text + "_.xlsm"

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
                'DDecideDesc.Text = dtWaitHandle.Rows(0)("DecideDesc")           '說明

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
        Dim Sql As String = ""
        Dim idx As Integer = FindFieldInf(pFieldName)
        Dim i As Integer



        '修正理由
        If pFieldName = "RPAREASON" Then
            DRPAREASON.Items.Clear()
            If idx = 0 Then
                If pName <> "ZZZZZZ" Then
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = pName
                    ListItem1.Value = pName
                    DRPAREASON.Items.Add(ListItem1)
                End If
            Else
                Sql = "Select DATA From M_Referp Where Cat='6003' AND DKEY = 'RPAREASON' Order by DKey "
                Dim dtRPAREASON As DataTable = uDataBase.GetDataTable(Sql)
                DRPAREASON.Items.Add("")
                For i = 0 To dtRPAREASON.Rows.Count - 1
                    Dim ListItem1 As New ListItem
                    ListItem1.Text = dtRPAREASON.Rows(i)("DATA")
                    ListItem1.Value = dtRPAREASON.Rows(i)("DATA")
                    If ListItem1.Value = pName Then ListItem1.Selected = True
                    DRPAREASON.Items.Add(ListItem1)
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
                Sql = "Select * From M_Referp Where Cat='014' Order by DKey, Data "
                Dim dtReasonCode As DataTable = uDataBase.GetDataTable(Sql)
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
            Sql = "Select * From M_Referp Where Cat='014' and DKey= '" & DReasonCode.SelectedValue & "' Order by Data "
            Dim dtReason As DataTable = uDataBase.GetDataTable(Sql)
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
        rqdVal.Style.Add("z-index", "137")
        rqdVal.Style.Add("Top", "350px")
        rqdVal.Style.Add("Left", "30px")
        rqdVal.Style.Add("Width", "450px")
        rqdVal.Style.Add("position", "absolute")
        rqdVal.Style.Add("display", "inline")
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
        'Button1.Attributes("onclick") = "calendarPicker('Form1.DAppDate');"
    End Sub
    '##-1
    '##AgentApprovProc-Start
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(AgentApprovProc)
    '**     流程控制
    '**        pFun=OK, NG1, NG2, SAVE  
    '**        pAction=0:OK, 1:NG1, 2:NG2, 3:Save   下一關卡時使用 
    '**        pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單  更新T_Waithandle狀態
    '**     
    '*****************************************************************
    Sub AgentApprovProc(ByVal pFun As String, ByVal pAction As Integer, ByVal pSts As String)
        Dim ErrCode As Integer = 0
        '------------------------------------------------
        '處理button說明
        'StsDes
        Dim wStsDes As String = ""
        If pSts = "1" Then wStsDes = BOK.Text
        If pSts = "2" Then wStsDes = BNG1.Text
        If pSts = "3" Then wStsDes = BNG2.Text
        '--
        '------------------------------------------------
        '指定下工程簽核者
        'AllocteID
        Dim wAllocateID As String = ""
        '--
        '------------------------------------------------
        '主表單
        'TabelName
        Dim wTableName As String = "F_RPAStockModifySheet"
        '--
        '------------------------------------------------
        '代理簽核
        'RunAgentApprov
        ErrCode = oCommon.RunAgentApprov(pFun, _
                                            pAction, _
                                            pSts, _
                                            wStsDes, _
                                            wFormNo, _
                                            wFormSno, _
                                            wStep, _
                                            wSeqNo, _
                                            wDecideCalendar, _
                                            Request.QueryString("pUserID"), _
                                            wApplyID, _
                                            wAgentID, _
                                            wAllocateID, _
                                            uCommon.ReplaceString(DDecideDesc.Text), _
                                            wTableName, _
                                            wUserIP)
        '--
        '------------------------------------------------
        '表單資料
        If ErrCode = 0 Then
            ModifyData("AGENTAPPROVE", "0")
            AddCommissionNo(wFormNo, wFormSno)
            '
            Dim URL As String = "http://10.245.1.10/WorkFlow/WaitHandle.aspx?pUserID=" + Request.QueryString("pUserID")
            Response.Redirect(URL)
        Else
            EnabledButton()   '起動Button運作
            uJavaScript.PopMsg(Me, "代理簽核異常,請確認或連絡系統人員!")
        End If
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
                    If RtnCode <> 0 Then
                        ErrCode = 9110
                    Else
                        '申請流程資料建置(表單號碼,表單流水號,工程關卡號碼,序號,行事曆,簽核者, 申請者)
                        oFlow.NewFlow(wFormNo, NewFormSno, wStep, 1, wDecideCalendar, Request.QueryString("pUserID"), wApplyID)
                        '設定委託No
                        DNo.Text = SetNo(NewFormSno)
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

                If (wStep = 1 Or wStep = 500) And pFun = "OK" Then
                    wAllocateID = GetRelated10(DAppName.Text)
                End If

                If (wStep = 10) And pFun = "OK" Then
                    wAllocateID = GetRelated20(DAppName.Text)
                End If

                If (wStep = 40) And pFun = "OK" Then
                    If CInt(DAMOUNT.Text) > 100000 Then
                        pAction = 2
                    End If
                End If


                If (wStep = 50) And pFun = "OK" Then
                    If CInt(DAMOUNT.Text) > 150000 Then
                        pAction = 2
                    End If
                End If



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
                uJavaScript.PopMsg(Me, Message)
            End If      '儲存表單ErrCode=0
        End While       '重覆執行

        If ErrCode = 0 Then
            '--郵件傳送---------
            'oCommon.SendMail()
            '
            Dim URL As String = uCommon.GetAppSetting("MessageUrl") & "?pMSGID=C&pFormNo=" & wFormNo & "&pStep=" & wStep & "&pNextGate=" & wNextGate & _
                                                "&pUserID=" & Request.QueryString("pUserID") & "&pApplyID=" & wApplyID
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
        Dim FileName As String



        Dim sql As String = ""
        sql = " Insert into F_RPAStockModifySheet (Sts, CompletedTime, FormNo, FormSno,"
        sql &= " NO,Date,AppName,DepName,DepID,Remark,FILEUPLOAD1,RPAREASON,AMOUNT,"
        sql = sql + "CreateUser, CreateTime, ModifyUser, ModifyTime) "
        sql = sql + "VALUES( "
        sql = sql + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
        sql = sql + " '" + NowDateTime + "', "        '結案日
        sql = sql + " '006003', "                     '表單代號
        sql = sql + " '" + CStr(NewFormSno) + "', "   '表單流水號
        sql = sql + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDate.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DAppName.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDepName.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DDepID.Text) + "', "   'NO  1
        sql = sql + " N'" + YKK.ReplaceString(DRemark.Text) + "', "   'NO  1

        FileName = ""
        If DFileUpload1.PostedFile.FileName <> "" Then
            Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                          CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
            Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                        System.Configuration.ConfigurationManager.AppSettings("RPASTOCKPath")

            '20170912 將檔名修改成不含原始檔名
            Dim fileExtension As String  '副檔名
            fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '取得檔案格式

            FileName = CStr(NewFormSno) + "-RPAFILE-" + UploadDateTime + fileExtension

            DFileUpload1.PostedFile.SaveAs(Path + FileName)

        Else
            FileName = ""
        End If
        sql = sql + "'" + FileName + "'," ' RPA在庫修正表
        sql = sql + "'" + DRPAREASON.SelectedValue + "'," ' RPA在庫修正理由
        If DAMOUNT.Text <> "" Then
            sql = sql + "" + DAMOUNT.Text + "," ' RPA在庫修正表加總金額
        Else
            sql = sql + "0," ' RPA在庫修正表加總金額
        End If


        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' ,"
        sql = sql + "'" & Request.QueryString("pUserID") & "' ,"
        sql = sql + "'" & NowDateTime & "' )"
        uDataBase.ExecuteNonQuery(sql)


        upload()

        '上傳在庫修正表 
        Insert()

        '更新AMOUNT 
        sql = " UPDATE A"
        sql = sql + " SET A.AMOUNT = B.AMOUNT"
        sql = sql + " FROM  F_RPAStockModifySheet A,("
        sql = sql + " select  NO,SUM(AMOUNT)AMOUNT  from F_RPAStockModifySheetdt"
        sql = sql + " WHERE NO ='" + YKK.ReplaceString(DNo.Text) + "'"
        sql = sql + " GROUP BY NO"
        sql = sql + " )B WHERE A.NO =B.NO"
        uDataBase.ExecuteNonQuery(sql)




    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Sub ModifyData(ByVal pFun As String, ByVal pSts As String)



        Dim sql As String

        sql = " Update F_RPAStockModifySheet"
        sql = sql + " Set "
        '##-2
        '##AgentApprov-Start
        'If pFun <> "SAVE" Then
        '    sql = sql + " Sts = '" & pSts & "',"
        '    sql = sql + " CompletedTime = '" & NowDateTime & "',"
        'End If
        '
        If pFun <> "SAVE" And pFun <> "AGENTAPPROVE" Then
            sql = sql + " Sts = '" & pSts & "',"
            sql = sql + " CompletedTime = '" & NowDateTime & "',"
        End If
        '##AgentApprov-End
        '##
        sql = sql + " No = N'" & YKK.ReplaceString(DNo.Text) & "',"
        sql = sql + " Date = N'" & YKK.ReplaceString(DDate.Text) & "',"
        sql = sql + " DepID = N'" & YKK.ReplaceString(DDepID.Text) & "',"
        sql = sql + " DepName = N'" & YKK.ReplaceString(DDepName.Text) & "',"
        sql = sql + " AppName = N'" & YKK.ReplaceString(DAppName.Text) & "',"

        sql = sql + " Remark = N'" & YKK.ReplaceString(DRemark.Text) & "',"
        Dim FileName As String

        FileName = ""
        If DFileUpload1.Visible = True Then
            '   And LCertifcateFile.NavigateUrl = "" Then
            If DFileUpload1.PostedFile.FileName <> "" Or LAttachfile.NavigateUrl <> "" Then
                Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                Dim Path As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                            System.Configuration.ConfigurationManager.AppSettings("RPASTOCKPath")

                '20170912 將檔名修改成不含原始檔名
                Dim fileExtension As String  '副檔名
                fileExtension = IO.Path.GetExtension(DFileUpload1.PostedFile.FileName).ToLower   '取得檔案格式

                FileName = CStr(wFormSno) + "-RPAFILE-" + UploadDateTime + fileExtension

                If DFileUpload1.PostedFile.FileName = "" Then
                    '  FileName = Right(LAttachfile.NavigateUrl, InStr(StrReverse(LAttachfile.NavigateUrl), "\") - 1)
                    DFileUpload1.PostedFile.SaveAs(Path + FileName)
                Else
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime & "-" & Right(DFileUpload1.PostedFile.FileName, InStr(StrReverse(DFileUpload1.PostedFile.FileName), "\") - 1)
                    ' FileName = CStr(wFormSno) + "-DPFILE-" + UploadDateTime + fileExtension
                    DFileUpload1.PostedFile.SaveAs(Path + FileName)
                End If

            Else
                FileName = ""
            End If
            sql = sql + " FILEOUPLOAD1='" + FileName + "'," '報廢證明
        End If
        If DAMOUNT.Text = "" Then
            sql &= " Amount = 0,"
        Else
            sql &= " Amount = '" & DAMOUNT.Text & "',"
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
            '##-3
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "NG2") = 0 Then
                AgentApprovProc("NG2", 2, "3")
            Else
                FlowControl("NG2", 2, "3")
            End If
            '##AgentApprov-End
            '##
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

        DisabledButton()   '停止Button運作
        '##-4
        '##AgentApprov-Start
        If oCommon.UseAgentApprov(wFormNo, wStep, "NG1") = 0 Then
            AgentApprovProc("NG1", 1, "2")
        Else
            FlowControl("NG1", 1, "2")
        End If
        '##AgentApprov-End
        '##
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
            '##-5
            '##AgentApprov-Start
            If oCommon.UseAgentApprov(wFormNo, wStep, "OK") = 0 Then
                AgentApprovProc("OK", 0, "1")
            Else
                FlowControl("OK", 0, "1")
            End If
            '##AgentApprov-End
            '##
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
        Dim Str3 As String = ""
        Dim i As Integer

        'Set當日日期
        Str3 = Mid(CStr(DateTime.Now.Year), 3, 2)  '年

        Str2 = CStr(DateTime.Now.Month)  '月
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str2
        Str2 = CStr(DateTime.Now.Day)    '日
        For i = Str2.Length To 1
            Str2 = "0" + Str2
        Next
        Str = Str3 + Str + Str2
        'Set單號
        Str1 = CStr(Seq)
        For i = Str1.Length To 4 - 1
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
        Dim allowedExtensions As String() = {".xls", ".xlsx", ".xlsm"}   '定義允許的檔案格式
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
            If UPFile.PostedFile.ContentLength <= 3500 * 1024 Then
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

        If ErrCode = 0 Then  ' 判斷檔案上傳
            If DFileUpload1.Visible Then
                If DFileUpload1.PostedFile.FileName <> "" Then
                    ErrCode = UPFileIsNormal(DFileUpload1)

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
                ErrCode = oCommon.CommissionNo("006003", wFormSno, wStep, DNo.Text) '表單號碼, 表單流水號, 工程, 委託書No
                If ErrCode <> 0 Then
                    ErrCode = 9050
                End If
            End If
        End If


        griderror = True
        inserterror = True
        fielderror = True
        qtyerror = True

        If ErrCode = 0 Then
            If wStep = 1 Or wStep = 500 Then
                upload()
                If griderror = False Then

                    ErrCode = 9076

                End If
                If fielderror = False Then
                    ErrCode = 9077
                End If
                If ErrCode = 0 Then

                    If inserterror = False Then
                        ErrCode = 9078
                    End If

                    If qtyerror = False Then
                        ErrCode = 9079
                    End If
                End If

            End If
        End If





        Try

        Catch ex As Exception
            ErrCode = 9075
        End Try

        If ErrCode <> 0 Then
            If ErrCode = 9010 Then Message = "Item資料有誤,請確認!"
            If ErrCode = 9020 Then Message = "上傳檔案格式有誤,請確認!(限EXCEL)"
            If ErrCode = 9030 Then Message = "上傳檔案Size有誤,請確認!(限5MB)"
            If ErrCode = 9040 Then Message = "延遲理由為其他時需填寫說明,請確認!"
            If ErrCode = 9050 Then Message = "委託書No.重覆,請確認委託書No.!"
            If ErrCode = 9060 Then Message = "修改後未重新檢測資料,請確認!"
            If ErrCode = 9070 Then Message = "發現有空白重覆,請重新檢測資料,請確認!"
            If ErrCode = 9071 Then Message = "Item Name(2)字串過長(>34),請重新檢測資料,請確認!"
            If ErrCode = 9072 Then Message = "發現特殊要求中有未排序資料,請重新檢測資料,請確認!"
            If ErrCode = 9073 Then Message = "溫度不正常，請確認!"
            If ErrCode = 9074 Then Message = "溼度不正常，請確認!"
            If ErrCode = 9075 Then Message = "非數字格式,請確認!"
            If ErrCode = 9076 Then Message = "上傳格式有誤gridview,請確認!"
            If ErrCode = 9077 Then Message = "CODE 或 報廢準則 或 部門 不允許空白!"
            If ErrCode = 9078 Then Message = "上傳格式有誤Insert,請確認!"
            If ErrCode = 9079 Then Message = "修正數量大於FREE 數量,請確認!"


            uJavaScript.PopMsg(Me, Message)
        Else
            isOK = True
        End If
        '

        Return isOK

    End Function

    Function GetRelated10(ByVal UserName As String) As String

        Dim sql As String = "select  Userid  from M_Users where username=(SELECT SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data)) FROM M_referp where cat ='6003' and Dkey ='ADMIN1' and Data like N'" & UserName & "%') "
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function

    Function GetRelated20(ByVal UserName As String) As String

        Dim sql As String = "select  Userid  from M_Users where username=(SELECT SUBSTRING(Data,CHARINDEX('-',Data)+1,LEN(Data)-CHARINDEX('-',Data)) FROM M_referp where cat ='6003' and Dkey ='ADMIN2' and Data like N'" & UserName & "%') "
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim NextGate As String = ""
        If dt.Rows.Count > 0 Then
            NextGate = dt.Rows(0)("UserID")
        End If
        Return NextGate
    End Function


    Function upload() As Boolean
        Try
            If DFileUpload1.HasFile Then

                '上傳附檔
                Dim FileName1 As String
                UploadName = DFileUpload1.PostedFile.FileName

                Dim UploadDateTime = CStr(Now.Year) + CStr(Now.Month) + CStr(Now.Day) + _
                                              CStr(Now.Hour) + CStr(Now.Minute) + CStr(Now.Second)     '上傳日期
                ' Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("DISPOSALData1")
                Dim Path1 As String = System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                                System.Configuration.ConfigurationManager.AppSettings("RPAStockData")
                'System.Configuration.ConfigurationManager.AppSettings("SystemPath") & _
                '               System.Configuration.ConfigurationManager.AppSettings("DISPOSALData")

                '20170912 將檔名修改成不含原始檔名
                Dim fileExtension As String  '副檔名
                fileExtension = IO.Path.GetExtension(UploadName).ToLower   '取得檔案格式
                FileName1 = Path1 + CStr(DNo.Text) + fileExtension
                DFileUpload1.PostedFile.SaveAs(FileName1)

                '展開
                Dim FileName As String = Path.GetFileName(DFileUpload1.PostedFile.FileName)
                Dim Extension As String = Path.GetExtension(DFileUpload1.PostedFile.FileName)
                '  Dim FolderPath As String = ConfigurationManager.AppSettings("FolderPath")

                ' FileName1 = "http://10.245.1.6/DASW/Document/006002/" + CStr(DNo.Text) + UploadDateTime + fileExtension
                'Dim FilePath As String = Server.MapPath(FolderPath + FileName)
                ' DFileUpload1.SaveAs(FilePath)
                'Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Text)
                Import_To_Grid(FileName1, Extension, rbHDR.SelectedItem.Text)


            End If
        Catch ex As Exception
            uJavaScript.PopMsg(Me, "上傳檔案格式有誤upload,請確認!")
            griderror = False
        End Try


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
            Case ".xlsm"
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
                If SheetName = "RPA$" Then
                Else
                    SheetName = "RPA$"
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
        GridView1.Caption = Path.GetFileName(FilePath)
        GridView1.DataSource = dt
        GridView1.DataBind()

    End Sub

    Sub Insert()
        '檢查是否有同月份同申請人的檔案，若有則將之前的全刪除
        Dim a As String = ""
        Dim uploadflag As Integer = 1
        Dim nullflag As Integer = 1
        Dim sno As String
        Dim NewFormSno As Integer = wFormSno    '表單流水號
        Dim SQL As String
  
        If wStep = 1 Then
            sno = SetNo(NewFormSno)
        Else
            sno = DNo.Text
        End If

    
        ' If uploadflag = 1 Then
        Try

            '上傳到資料庫
            Dim j As Integer
            Dim jSQL As String
            jSQL = ""

            Dim i As Integer

            SQL = " delete from F_RPAStockModifySheetDT"
            SQL = SQL + " WHERE NO='" + YKK.ReplaceString(DNo.Text) + "'"
            uDataBase.ExecuteNonQuery(SQL)



            For i = 0 To Me.GridView1.Rows.Count - 1 Step i + 1



                SQL = "Insert into F_RPAStockModifySheetDT (Sts,CompletedTime,FormNo,FormSno,Date,No,"
                SQL = SQL + " CODE,LENGTHU,UNIT,COLOR,HISTORYCD,LOCATION,QTY,COMMENT,DEPOT,ACTUAL,FREE,INUSE,USESCHEDULE,ONKEEEP,COSTA,COSTB,AMOUNT,ITEMNAME1,ITEMNAME2,CreateUser,CreateTime,ModifyUser,ModifyTime)"

                SQL = SQL + "VALUES( "
                SQL = SQL + " '0', "                          '狀態(0:未結,1:已結NG,2:已結OK)
                SQL = SQL + " '" + NowDateTime + "', "        '結案日
                SQL = SQL + " '006003', "                     '表單代號
                SQL = SQL + " '" + CStr(NewFormSno) + "', "   '表單流水號
                SQL = SQL + " N'" + YKK.ReplaceString(DDate.Text) + "', "   'NO  1
                SQL = SQL + " N'" + YKK.ReplaceString(DNo.Text) + "', "   'NO  1

                For j = 0 To 18


                    If j = 0 Then
                        If GridView1.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = "''"
                        Else
                            jSQL = "N'" + YKK.ReplaceString(GridView1.Rows(i).Cells(j).Text) + "'"

                        End If

                    Else

                        If GridView1.Rows(i).Cells(j).Text = "&nbsp;" Then
                            jSQL = jSQL + ",N" + "''"
                        Else
                            jSQL = jSQL + ",N'" + YKK.ReplaceString(GridView1.Rows(i).Cells(j).Text) + "'"

                        End If


                    End If

                    'a = GridView1.Rows(i).Cells(0).Text
                    'If a = "&nbsp;" Then  '檢查第一欄是不是NULL
                    '    nullflag = 0
                    'End If

                Next
                SQL = SQL + jSQL + ","
                SQL = SQL + "'" & Request.QueryString("pUserID") & "' ,"
                SQL = SQL + "'" & NowDateTime & "' ,"
                SQL = SQL + "'" & Request.QueryString("pUserID") & "' ,"
                SQL = SQL + "'" & NowDateTime & "' )"
           
                'If CStr(i + 1) = "51" Then
                'uJavaScript.PopMsg(Me, a)
                'End If


                If nullflag = 1 Then '第一欄是空白就不要匯入
                    uDataBase.ExecuteNonQuery(SQL)

                End If


            Next


            uJavaScript.PopMsg(Me, "上傳成功")

         
            '   Send1(wApplyID, Adminmail, wApplyID, "006001", "1", "1", "END")
            '送件者, 收件者, 申請者, 表單號碼, 表單流水號, 工程關卡號碼, 訊息類別

            DInsert.Enabled = False

        Catch ex As Exception
            uJavaScript.PopMsg(Me, "上傳檔案格式有誤Insert,請確認!")
            inserterror = False
        End Try

        ' End If


        GridView1.DataSource = Nothing
        GridView1.DataBind()




    End Sub
    
    Protected Sub DUpload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DUpload.Click
        upload()
    End Sub

    Protected Sub DInsert_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DInsert.Click
        Insert()
    End Sub

     
    Protected Sub DFileUpload1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles DFileUpload1.DataBinding
        'upload()
    End Sub


    '複製資料夾
    Public Shared Sub CopyDir(ByVal srcPath As String, ByVal aimPath As String)
        Try


            ' 檢查目標目錄是否以目錄分割字元結束如果不是則新增之
            If aimPath(aimPath.Length - 1) <> Path.DirectorySeparatorChar Then
                aimPath += Path.DirectorySeparatorChar
            End If
            ' 判斷目標目錄是否存在如果不存在則新建之
            If Not Directory.Exists(aimPath) Then
                Directory.CreateDirectory(aimPath)
            End If
            ' 得到源目錄的檔案列表，該裡面是包含檔案以及目錄路徑的一個數組
            ' 如果你指向copy目標檔案下面的檔案而不包含目錄請使用下面的方法
            ' string[] fileList = Directory.GetFiles(srcPath);
            Dim fileList() As String = Directory.GetFileSystemEntries(srcPath)
            ' 遍歷所有的檔案和目錄
            Dim file As String
            For Each file In fileList
                ' 先當作目錄處理如果存在這個目錄就遞迴Copy該目錄下面的檔案
                If Directory.Exists(file) Then
                    CopyDir(file, aimPath + Path.GetFileName(file))
                Else
                    System.IO.File.Copy(file, aimPath + Path.GetFileName(file), True)
                End If

                ' 否則直接Copy檔案

            Next
        Catch
            Console.WriteLine("無法複製!")
        End Try
    End Sub


    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then

            '檢查欄位名稱

            ' s1 = Trim(e.Row.Cells(26).Text.ToUpper)
            DInsert.Enabled = True

   

            If Trim(e.Row.Cells(0).Text.ToUpper) <> "CODE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(1).Text.ToUpper) <> "LENGTH" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(2).Text.ToUpper) <> "U" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(3).Text.ToUpper) <> "COLOR" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(4).Text.ToUpper) <> "HISTORY CD" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(5).Text.ToUpper) <> "LOCATION" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(6).Text.ToUpper) <> "修正QTY" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(7).Text.ToUpper) <> "COMMENT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(8).Text.ToUpper) <> "DEPOT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(9).Text.ToUpper) <> "ACTUAL" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(10).Text.ToUpper) <> "FREE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(11).Text.ToUpper) <> "IN USE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(12).Text.ToUpper) <> "USE_SCHEDULE" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(13).Text.ToUpper) <> "ON KEEP" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(14).Text.ToUpper) <> "COST A" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(15).Text.ToUpper) <> "COST B" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(16).Text.ToUpper) <> "修正數量_AMOUNT" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(17).Text.ToUpper) <> "ITEM NAME 1" Then
                DInsert.Enabled = False
            ElseIf Trim(e.Row.Cells(18).Text.ToUpper) <> "ITEM NAME 2" Then
                DInsert.Enabled = False
            End If


            'End If

            If DInsert.Enabled = False Then
                Message = "上傳格式有誤gridview,請確認!"
                uJavaScript.PopMsg(Me, Message)
                griderror = False

                GridView1.DataSource = Nothing
                GridView1.DataBind()

            End If
        Else

       
            Dim Qty, Free As Integer
            If Trim(e.Row.Cells(6).Text) = "&nbsp;" Then
                Qty = 0
            Else
                Qty = CInt(Trim(e.Row.Cells(6).Text))
            End If

            If Trim(e.Row.Cells(6).Text) = "&nbsp;" Then
                Free = 0
            Else
                Free = CInt(Trim(e.Row.Cells(10).Text))
            End If



            '2025/5/28 增加修正數量大於FREE 數量
            If Trim(e.Row.Cells(4).Text.ToUpper) = "57" Then

                If Qty > Free Then
                    qtyerror = False

                End If
            End If



        End If


    End Sub
End Class
