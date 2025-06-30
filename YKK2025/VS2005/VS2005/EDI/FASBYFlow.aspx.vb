Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading

Partial Class FASBYFlow
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject             ' 操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uFASCommon As New EDI2011.FCommonService
    Dim NowDateTime As String               ' 現在日時
    Dim xUserID As String                   ' 使用者ID
    Dim xCustomer, xCustomerName As String  ' 客戶名稱
    Dim xOffice2010 As Integer
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        '
        If Not IsPostBack Then
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomer()                               ' 客戶設定
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uFASCommon.Timeout = Timeout.Infinite
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        xCustomer = UCase(Request.QueryString("pCustomer"))
        xOffice2010 = CInt(Request.QueryString("pOffice2010"))
        '
        DLogID.Style("left") = -500 & "px"
        DBuyer.Style("left") = -500 & "px"
        DGRBuyer.Style("left") = -500 & "px"
        DFunList.Style("left") = -500 & "px"
        DUserID.Style("left") = -500 & "px"
        DLastUniqueID.Style("left") = -500 & "px"
        DFileName.Style("left") = -500 & "px"
        DFilePath.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        DFilePath2.Style("left") = -500 & "px"
        DReportFileName.Style("left") = -500 & "px"
        DReportFilePath.Style("left") = -500 & "px"
        DReportFilePath1.Style("left") = -500 & "px"
        '動作按鈕設定
        BBYFMTChange.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟Excel格式轉換或準備BYFCT資料？" + "');if(!ok){return false;} else {CheckAttribute('BYFMTChange')}"
        BBYImport.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟Excel準備資料？" + "');if(!ok){return false;} else {CheckAttribute('BYImport')}"
        BBYDataCheck.Attributes("onclick") = "var ok=window.confirm('" + "是否執行資料檢測作業？" + "');if(!ok){return false;} else {CheckAttribute('BYDataCheck')}"
        BBYReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 報表匯出 作業？" + "');if(!ok){return false;} else {CheckAttribute('BYReport')}"
        BBYConvert.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 BUYER FCT CONVERT 作業？" + "');if(!ok){return false;} else {CheckAttribute('BYConvert')}"
        BBYFASReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 BUYER FCT匯出 作業？" + "');if(!ok){return false;} else {CheckAttribute('BYFASReport')}"
        'Action CheckBox設定
        AtGoImport.Attributes.Add("onclick", "ActionCheckBY('AtGoImport')")
        AtGoReport.Attributes.Add("onclick", "ActionCheckBY('AtGoReport')")
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskStatus)
    '**     設定各動作執行結果狀態
    '**     pAction=動作名稱    pLeft=顯示(>0)或不顯示(<0)  pStatus=顯示圖像
    '**
    '*****************************************************************
    Sub SetTaskStatus(ByVal pAction As String, ByVal pLeft As Integer, ByVal pStatus As Integer)
        Select Case pAction
            Case "StsBYFMTChange"
                StsBYFMTChange.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYFMTChange.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYFMTChange.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBYImport"
                StsBYImport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYImport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYImport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBYDataCheck"
                StsBYDataCheck.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYDataCheck.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYDataCheck.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBYReport"
                StsBYReport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYReport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYReport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBYConvert"
                StsBYConvert.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYConvert.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYConvert.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBYFASReport"
                StsBYFASReport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBYFASReport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBYFASReport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsBYFMTChange.Style("left") = pLeft & "px"
                StsBYImport.Style("left") = pLeft & "px"
                StsBYDataCheck.Style("left") = pLeft & "px"
                StsBYReport.Style("left") = pLeft & "px"
                StsBYConvert.Style("left") = pLeft & "px"
                StsBYFASReport.Style("left") = pLeft & "px"
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskProgress)
    '**     設定各動作執行中Task Bar
    '**     pAction=動作名稱    pShow=顯示(>0)或不顯示(<0)
    '**
    '*****************************************************************
    Sub SetTaskProgress(ByVal pAction As String, ByVal pLeft As Integer)
        Select Case pAction
            Case "ProBYFMTChange"
                ProBYFMTChange.Style("left") = pLeft & "px"
            Case "ProBYImport"
                ProBYImport.Style("left") = pLeft & "px"
            Case "ProBYDataCheck"
                ProBYDataCheck.Style("left") = pLeft & "px"
            Case "ProBYReport"
                ProBYReport.Style("left") = pLeft & "px"
            Case "ProBYConvert"
                ProBYConvert.Style("left") = pLeft & "px"
            Case "ProBYFASReport"
                ProBYFASReport.Style("left") = pLeft & "px"
            Case Else
                ProBYFMTChange.Style("left") = pLeft & "px"
                ProBYImport.Style("left") = pLeft & "px"
                ProBYDataCheck.Style("left") = pLeft & "px"
                ProBYReport.Style("left") = pLeft & "px"
                ProBYConvert.Style("left") = pLeft & "px"
                ProBYFASReport.Style("left") = pLeft & "px"
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetTaskButton)
    '**     設定各動作按鈕
    '**     pAction=動作名稱    pShow=顯示或不顯示
    '**
    '*****************************************************************
    Sub SetTaskButton(ByVal pAction As String, ByVal pShow As Boolean)
        Select Case pAction
            Case "BReset"
                BReset.Visible = pShow
            Case "BBYFMTChange"
                BBYFMTChange.Visible = pShow
            Case "BBYImport"
                BBYImport.Visible = pShow
            Case "BBYDataCheck"
                BBYDataCheck.Visible = pShow
            Case "BBYReport"
                BBYReport.Visible = pShow
            Case "BBYConvert"
                BBYConvert.Visible = pShow
            Case "BBYFASReport"
                BBYFASReport.Visible = pShow
            Case Else
                BReset.Visible = False
                BBYFMTChange.Visible = False
                BBYImport.Visible = False
                BBYDataCheck.Visible = False
                BBYReport.Visible = False
                BBYConvert.Visible = False
                BBYFASReport.Visible = False
                '
                AtGoImport.Visible = False
                AtGoImport.Checked = False
                AtGoReport.Visible = False
                AtGoReport.Checked = False
                AtNotFirst.Visible = False
                AtNotFirst.Checked = False
                '
                'Test-Start
                'BReset.Visible = True
                'BBYFMTChange.Visible = True
                'BBYImport.Visible = True
                'BBYDataCheck.Visible = True
                'BBYReport.Visible = True
                'BBYConvert.Visible = True
                'BBYFASReport.Visible = True

                'AtGoImport.Visible = True
                'AtGoReport.Visible = True
                'AtNotFirst.Visible = True
                'Test-End
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetCustomer)
    '**     客戶設定
    '**
    '*****************************************************************
    Sub SetCustomer()
        '
        Dim sql As String
        Dim errcode As Integer = 0
        '
        sql = "Select * From M_FControlRecord "
        sql = sql & " Where Buyer = '" & xCustomer & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            DLogID.Text = Now.ToString("yyyyMMddHHmmss")
            '
            DBuyer.Text = dt_ControlRecord.Rows(0).Item("Buyer")
            DGRBuyer.Text = dt_ControlRecord.Rows(0).Item("BuyerGroup")
            DFunList.Text = dt_ControlRecord.Rows(0).Item("FunList")
            xCustomerName = dt_ControlRecord.Rows(0).Item("Name")
            DUserID.Text = xUserID
            '
            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "StartBY", 0, "StartBY", 0) = 0 Then
                SetTaskButton("BBYFMTChange", True)
                SetTaskButton("BReset", True)
                '
                AtGoImport.Visible = True
                AtGoReport.Visible = True
                '
                ' 設定 Data or 報表程式
                ' BY FCT 格式轉換
                If xOffice2010 = 1 Then
                    DFileName.Text = xCustomerName + "_BYFormatChange.xlsm"
                Else
                    DFileName.Text = xCustomerName + "_BYFormatChange.xls"
                End If
                DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                ' 準備 BY FCT 資料
                If xOffice2010 = 1 Then
                    DFileName.Text = xCustomerName + "_BYDataPrepare.xlsm"
                Else
                    DFileName.Text = xCustomerName + "_BYDataPrepare.xls"
                End If
                DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text

                If xOffice2010 = 1 Then
                    DFileName.Text = xCustomerName + "_BYDataPrepareUpdate.xlsm"
                Else
                    DFileName.Text = xCustomerName + "_BYDataPrepareUpdate.xls"
                End If
                DFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                ' REPORT
                If xOffice2010 = 1 Then
                    DReportFileName.Text = xCustomerName + "_BYReport.xlsm"
                Else
                    DReportFileName.Text = xCustomerName + "_BYReport.xls"
                End If
                DReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                ' FAS BY FCT
                If xOffice2010 = 1 Then
                    DReportFileName.Text = xCustomerName + "_BYFASReport.xlsm"
                Else
                    DReportFileName.Text = xCustomerName + "_BYFASReport.xls"
                End If
                DReportFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                '
                BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消所有作業重新執行？" + "');if(!ok){return false;}"
            Else
                uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
            End If
        Else
            uJavaScript.PopMsg(Me, "此客戶不存在，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BReset_Click)
    '**     客戶重新選擇
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReset.Click
        '
        If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "ResetBY", 0, "ResetBY", 0) = 0 Then
            ' Clear 各動作 按鈕/結果/跑馬燈等圖像
            ClearKeyData()
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomer()                               ' 客戶設定
        Else
            uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ClearKeyData)
    '**     清除Key欄位資料
    '**
    '*****************************************************************
    Protected Sub ClearKeyData()
        DLogID.Text = ""
        DBuyer.Text = ""
        DGRBuyer.Text = ""
        DFunList.Text = ""
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBYFMTChange_Click)
    '**     BY FCT 格式轉換
    '**
    '*****************************************************************
    Protected Sub BBYFMTChange_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYFMTChange.Click
        '
        '*****************************************************************
        '**  BY FCT 格式轉換
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYFMTChange", 1)

        If AtGoImport.Checked = False And AtGoReport.Checked = False Then
            '
            '判斷是否完成作業
            Dim Code As Integer
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFMTChange", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFMTChange", 3)
            End If
            Do Until Code = 0
                System.Threading.Thread.Sleep(3 * 1000)
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFMTChange", 2)
                If Code <> 0 Then
                    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFMTChange", 3)
                End If
            Loop
        Else
            uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYFMTChange", 2)
        End If
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "BYFMTChange", 2) = 0 Then
            SetTaskProgress("ProBYFMTChange", -500)
            SetTaskButton("BBYFMTChange", False)
            SetTaskButton("BReset", True)
            If AtGoReport.Checked = True Then
                uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYReport", 1)
                SetTaskButton("BBYReport", True)
            Else
                uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYImport", 1)
                SetTaskButton("BBYImport", True)
                AtNotFirst.Visible = True
            End If
            SetTaskStatus("StsBYFMTChange", 212, 0)
        Else
            SetTaskStatus("StsBYFMTChange", 212, 1)
            uJavaScript.PopMsg(Me, "BUYER FCT格式轉換未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BYImport_Click)
    '**     準備 BY FCT 資料
    '**
    '*****************************************************************
    Protected Sub BBYImport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYImport.Click
        '
        '*****************************************************************
        '**  準備 BY FCT 資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYImport", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYImport", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYImport", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYImport", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYImport", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "BYImport", 2) = 0 Then
            SetTaskProgress("ProBYImport", -500)
            SetTaskButton("BBYImport", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BBYDataCheck", True)
            SetTaskStatus("StsBYImport", 165, 0)
        Else
            SetTaskStatus("StsBYImport", 165, 1)
            uJavaScript.PopMsg(Me, "BUYER FCT資料準備未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBYDataCheck_Click)
    '**     資料檢查
    '**
    '*****************************************************************
    Protected Sub BBYDataCheck_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYDataCheck.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  資料檢查
        '*****************************************************************
        ' 刪除執行履歷
        If errcode = 0 Then
            If uFASCommon.DeleteHistory(DBuyer.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        ' 新BY FCT投入
        If errcode = 0 Then
            If AtNotFirst.Checked = False Then
                If uFASCommon.NewBYFCTImport(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
        End If
        ' 資料檢查
        If errcode = 0 Then
            If uFASCommon.BYFCTDataCheck(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                errcode = 9003
            End If
        End If
        ' 更新客戶控制檔(BYDataCheck=2, BYReport=1)
        If errcode = 0 Then
            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "BYDataCheck", 2, "BYReport", 1) <> 0 Then
                errcode = 9011
            End If
        End If
        ' 檢查資料作業是否完成
        If errcode = 0 Then
            If uFASCommon.CheckControlRecord(DBuyer.Text, "BYDataCheck", 2) <> 0 Then
                errcode = 9012
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsBYDataCheck", 552, 1)
            StsBYDataCheck.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
            StsBYDataCheck.Target = "_blank"
            '
            If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
            If errcode = 9002 Then msg = "新BY FCT資料投入異常，請連絡系統人員!"
            If errcode = 9003 Then msg = "資料檢查作業異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9011 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9012 Then msg = "資料檢查作業未完成，請連絡系統人員!"
            uJavaScript.PopMsg(Me, msg)
        Else
            StsBYDataCheck.NavigateUrl = ""
            StsBYDataCheck.Target = ""
            SetTaskProgress("ProBYDataCheck", -500)
            SetTaskButton("BBYDataCheck", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BBYReport", True)
            SetTaskStatus("StsBYDataCheck", 552, 0)
            'Dim Cmd As String
            ''
            'Cmd = "<script>" + _
            '            "window.open('FASBYStatusList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" & xUserID & "','StatusList','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
            '      "</script>"
            ''
            'Response.Write(Cmd)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBYReport_Click)
    '**     報表處理
    '**
    '*****************************************************************
    Protected Sub BBYReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYReport.Click
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYReport", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYReport", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYReport", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYReport", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYReport", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "BYReport", 2) = 0 Then
            SetTaskProgress("ProBYReport", -500)
            SetTaskButton("BBYReport", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BBYConvert", True)
            SetTaskStatus("StsBYReport", 552, 0)
        Else
            SetTaskStatus("StsBYReport", 552, 1)
            uJavaScript.PopMsg(Me, "報表列出未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBYConvert_Click)
    '**     轉FAS使用 BYFCT
    '**
    '*****************************************************************
    Protected Sub BBYConvert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYConvert.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  資料檢查
        '*****************************************************************
        ' 刪除執行履歷
        If errcode = 0 Then
            If uFASCommon.DeleteHistory(DBuyer.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        ' 轉FAS BYFCT資料
        If errcode = 0 Then
            If uFASCommon.ConvertToBYFCT(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                errcode = 9003
            End If
        End If
        ' 更新客戶控制檔(BYConvert=2, BYFASReport=1)
        If errcode = 0 Then
            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "BYConvert", 2, "BYFASReport", 1) <> 0 Then
                errcode = 9011
            End If
        End If
        ' 檢查資料作業是否完成
        If errcode = 0 Then
            If uFASCommon.CheckControlRecord(DBuyer.Text, "BYConvert", 2) <> 0 Then
                errcode = 9012
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsBYConvert", 912, 1)
            StsBYConvert.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
            StsBYConvert.Target = "_blank"
            '
            If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
            If errcode = 9003 Then msg = "資料轉換作業異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9011 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9012 Then msg = "資料檢查作業未完成，請連絡系統人員!"
            uJavaScript.PopMsg(Me, msg)
        Else
            StsBYConvert.NavigateUrl = ""
            StsBYConvert.Target = ""
            SetTaskProgress("ProBYConvert", -500)
            SetTaskButton("BBYConvert", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BBYFASReport", True)
            SetTaskStatus("StsBYConvert", 912, 0)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBYFASReport_Click)
    '**     BY FCT FAS 報表處理
    '**
    '*****************************************************************
    Protected Sub BBYFASReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYFASReport.Click
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BYFASReport", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFASReport", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFASReport", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFASReport", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BYFASReport", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "BYFASReport", 2) = 0 Then
            SetTaskProgress("ProBYFASReport", -500)
            SetTaskButton("BBYFASReport", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsBYFASReport", 912, 0)
        Else
            SetTaskStatus("StsBYFASReport", 912, 1)
            uJavaScript.PopMsg(Me, "報表列出未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BStatus_Click)
    '**     Data Status List 功能
    '**
    '*****************************************************************
    Protected Sub BStatusList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BStatusList.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('FASBYStatusList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" & xUserID & "','StatusList','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BCheckSum_Click)
    '**     Check FCT QTY 功能
    '**
    '*****************************************************************
    Protected Sub BCheckSum_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BCheckSum.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('FASBYCheckSumQty.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" & xUserID & "','CheckSumQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
End Class
