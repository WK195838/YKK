
Partial Class Default2
    Inherits System.Web.UI.Page

    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BBYFMTChange_Click)
    ''**     BY FCT 格式轉換
    ''**
    ''*****************************************************************
    'Protected Sub BBYFMTChange_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBYFMTChange.Click
    '    Dim sql As String
    '    Dim errcode As Integer = 0
    '    Dim msg, mUserID As String
    '    '
    '    sql = "Select * From M_FControlRecord "
    '    sql = sql & " Where Name = '" & xCustomer & "'"
    '    Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
    '    If dt_ControlRecord.Rows.Count > 0 Then
    '        '
    '        DLogID.Text = Now.ToString("yyyyMMddHHmmss")
    '        '
    '        DBuyer.Text = dt_ControlRecord.Rows(0).Item("Buyer")
    '        DGRBuyer.Text = dt_ControlRecord.Rows(0).Item("BuyerGroup")
    '        DFunList.Text = dt_ControlRecord.Rows(0).Item("FunList")
    '        DUserID.Text = xUserID
    '        mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
    '        '
    '        If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
    '            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Start", 0, "Start", 0) = 0 Then
    '                SetTaskButton("BBYFMTChange", False)
    '                SetTaskButton("BReset", True)
    '                SetTaskStatus("StsBYFMTChange", 192, 0)
    '                '
    '                ' 準備FCT資料
    '                SetTaskButton("BBYImport", True)
    '                '
    '                ' 設定 Data or 報表程式
    '                DFileName.Text = xCustomer + "_DataPrepare.xls"              ' 準備FST
    '                DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
    '                '
    '                DFileName.Text = DCustomerBuyer.SelectedValue + "_ImportLSPlan.xls"             ' 匯入LS Plan
    '                DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
    '                '
    '                DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ReportBuilder.xls"      ' 客戶報表資料
    '                DReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
    '                '
    '                DReportFileName.Text = DCustomerBuyer.SelectedValue + "_BuyerReportBuilder.xls"      ' Buyer報表資料
    '                DReportFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
    '                '
    '                BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
    '            Else
    '                errcode = 9003
    '            End If
    '        Else
    '            If mUserID = xUserID Then
    '                errcode = 9004      '同上次使用者(上次不小心)
    '                SetTaskButton("BReset", True)
    '                SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
    '                BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
    '            Else
    '                errcode = 9002      '不同上次使用者
    '            End If
    '        End If
    '    Else
    '        errcode = 9001
    '    End If
    '    '
    '    If errcode <> 0 Then
    '        If errcode = 9001 Then msg = "此客戶不存在，請確認!"
    '        If errcode = 9002 Then msg = "此客戶正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
    '        If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '        If errcode = 9004 Then msg = "懷疑您上次使用有不正常關閉系統，造成此客戶正在使用中! 請[Reset]一次再使用!"
    '        If errcode = 9005 Then msg = "您目前無使用此功能權限，如有疑問請連絡系統人員!"
    '        uJavaScript.PopMsg(Me, msg)
    '    End If
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BExcel_Click)
    ''**     準備客戶資料
    ''**
    ''*****************************************************************
    'Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
    '    '
    '    '*****************************************************************
    '    '**  準備客戶資料
    '    '*****************************************************************
    '    '
    '    uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "DataPrepare", 1)
    '    '判斷是否完成作業
    '    Dim Code As Integer
    '    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
    '    If Code <> 0 Then
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
    '    End If
    '    Do Until Code = 0
    '        System.Threading.Thread.Sleep(3 * 1000)
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
    '        If Code <> 0 Then
    '            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
    '        End If
    '    Loop
    '    '
    '    '檢查是否已將資料上傳完成
    '    If uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2) = 0 Then
    '        SetTaskProgress("ProExcel", -500)
    '        SetTaskButton("BExcel", False)
    '        SetTaskButton("BReset", True)
    '        SetTaskButton("BConvert", True)
    '        AtAddition.Visible = True
    '        SetTaskStatus("StsExcel", 165, 0)
    '    Else
    '        SetTaskStatus("StsExcel", 165, 1)
    '        uJavaScript.PopMsg(Me, "客戶資料準備未完成，請確認!")
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BConvert_Click)
    ''**     ITEM / COLOR 轉換作業
    ''**
    ''*****************************************************************
    'Protected Sub BConvert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BConvert.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  ITEM / COLOR 轉換
    '    '*****************************************************************
    '    ' 刪除執行履歷/FCT Plan/LS Plan
    '    If errcode = 0 Then
    '        If uFASCommon.DeleteHistory(DBuyer.Text) <> 0 Then
    '            errcode = 9001
    '        End If
    '    End If
    '    If errcode = 0 Then
    '        If AtAddition.Checked = False Then
    '            DLastUniqueID.Text = "0"
    '            If uFASCommon.DeleteFCTData(DBuyer.Text) <> 0 Then
    '                errcode = 9002
    '            Else
    '                If uFASCommon.ResetFCTNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
    '                    errcode = 9004
    '                End If
    '            End If
    '        Else
    '            DLastUniqueID.Text = CStr(uFASCommon.GetLastUniqueID(DBuyer.Text, DGRBuyer.Text))
    '        End If
    '    End If
    '    If errcode = 0 Then
    '        If uFASCommon.DeleteLSData(DBuyer.Text) <> 0 Then
    '            errcode = 9003
    '        Else
    '            If uFASCommon.ResetLSNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
    '                errcode = 9005
    '            End If
    '        End If
    '    End If
    '    ' 匯入作業
    '    If errcode = 0 Then
    '        If uFASMapping.Rule2Data(DLogID.Text, DBuyer.Text, xUserID, DFunList.Text) <> 0 Then
    '            errcode = 9011
    '        End If
    '    End If
    '    ' 更新客戶控制檔(DataConvert=2, FCTPlan=1)
    '    If errcode = 0 Then
    '        If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataConvert", 2, "FCTPlan", 1) <> 0 Then
    '            errcode = 9012
    '        End If
    '    End If
    '    ' 檢查匯入作業是否完成
    '    If errcode = 0 Then
    '        If uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 2) <> 0 Then
    '            errcode = 9013
    '        End If
    '    End If
    '    ' 執行結果處理
    '    If errcode <> 0 Then
    '        SetTaskStatus("StsConvert", 192, 1)
    '        StsConvert.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
    '        StsConvert.Target = "_blank"
    '        '
    '        If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
    '        If errcode = 9002 Then msg = "刪除 FCT PLAN 資料異常，請連絡系統人員!"
    '        If errcode = 9003 Then msg = "刪除 LS PLAN 資料異常，請連絡系統人員!"
    '        If errcode = 9004 Then msg = "無法重新設定可使用FCT No.，請連絡系統人員!"
    '        If errcode = 9005 Then msg = "無法重新設定可使用LS No.，請連絡系統人員!"
    '        If errcode = 9011 Then msg = "轉換資料異常，請確認(可點選[ｘ]查詢)!"
    '        If errcode = 9012 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '        If errcode = 9013 Then msg = "轉換資料異常，請連絡系統人員!"
    '        uJavaScript.PopMsg(Me, msg)
    '    Else
    '        StsConvert.NavigateUrl = ""
    '        StsConvert.Target = ""
    '        SetTaskProgress("ProConvert", -500)
    '        SetTaskButton("BConvert", False)
    '        SetTaskButton("BReset", True)
    '        SetTaskButton("BFCTPlan", True)
    '        SetTaskStatus("StsConvert", 192, 0)
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BFCTPlan_Click)
    ''**     Forcast Plan展開
    ''**
    ''*****************************************************************
    'Protected Sub BFCTPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCTPlan.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  FCT Plan展開
    '    '*****************************************************************
    '    If errcode = 0 Then
    '        '
    '        ' 刪除執行履歷(Action=FCTPLAN)
    '        If errcode = 0 Then
    '            If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "FCTPLAN") <> 0 Then
    '                errcode = 9001
    '            End If
    '        End If
    '        '
    '        ' FCTNo展開
    '        If errcode = 0 Then
    '            If uFASCommon.MakeForcastNo(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
    '                errcode = 9002
    '            End If
    '        End If
    '        '
    '        ' Forcast Plan展開
    '        If errcode = 0 Then
    '            If uFASCommon.ForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
    '                errcode = 9003
    '            End If
    '        End If
    '        '
    '        ' 更新客戶控制檔(FCTPlan=2, LSPlan=1)
    '        If errcode = 0 Then
    '            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "FCTPlan", 2, "LSPlan", 1) <> 0 Then
    '                errcode = 9004
    '            End If
    '        End If
    '        '
    '        ' 檢查Forcast Plan展開是否完成
    '        If errcode = 0 Then
    '            If uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 2) <> 0 Then
    '                errcode = 9005
    '            End If
    '        End If
    '        '
    '        ' 執行結果處理
    '        If errcode <> 0 Then
    '            SetTaskStatus("StsFCTPlan", 552, 1)
    '            StsFCTPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
    '            StsFCTPlan.Target = "_blank"
    '            '
    '            If errcode = 9001 Then msg = "刪除執行履歷異常(FCTPLAN)，請連絡系統人員!"
    '            If errcode = 9002 Then msg = "FCT-No.異常，請確認(可點選[ｘ]查詢)!"
    '            If errcode = 9003 Then msg = "FCT Plan展開異常，請確認(可點選[ｘ]查詢)!"
    '            If errcode = 9004 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '            If errcode = 9005 Then msg = "Forcast Plan展開異常，請連絡系統人員!"
    '            uJavaScript.PopMsg(Me, msg)
    '        Else
    '            SetTaskStatus("StsFCTPlan", 552, 0)
    '            SetTaskProgress("ProFCTPlan", -500)
    '            SetTaskButton("BFCTPlan", False)
    '            SetTaskButton("BReset", True)
    '            SetTaskButton("BLSPlan", True)
    '        End If
    '        '
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BLSPlan_Click)
    ''**     Local Stock Plan展開
    ''**
    ''*****************************************************************
    'Protected Sub BLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLSPlan.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  Local Stock Plan展開
    '    '*****************************************************************
    '    If errcode = 0 Then
    '        ' 刪除執行履歷(Action=LSPLAN)
    '        If errcode = 0 Then
    '            If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "LSPLAN") <> 0 Then
    '                errcode = 9001
    '            End If
    '        End If
    '        ' LS Plan展開
    '        If errcode = 0 Then
    '            If uFASCommon.LocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
    '                errcode = 9002
    '            End If
    '        End If
    '        ' 更新客戶控制檔(LSPlan=2, Report=1)
    '        If errcode = 0 Then
    '            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "LSPlan", 2, "Report", 1) <> 0 Then
    '                errcode = 9003
    '            End If
    '        End If
    '        ' LS Plan展開是否完成
    '        If errcode = 0 Then
    '            If uFASCommon.CheckControlRecord(DBuyer.Text, "LSPlan", 2) <> 0 Then
    '                errcode = 9004
    '            End If
    '        End If
    '        ' 執行結果處理
    '        If errcode <> 0 Then
    '            SetTaskStatus("StsLSPlan", 525, 1)
    '            StsLSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
    '            StsLSPlan.Target = "_blank"
    '            '
    '            If errcode = 9001 Then msg = "刪除執行履歷異常(LSPlan)，請連絡系統人員!"
    '            If errcode = 9002 Then msg = "LS Plan展開異常，請確認(可點選[ｘ]查詢)!"
    '            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '            If errcode = 9004 Then msg = "LS Plan展開異常，請連絡系統人員!"
    '            uJavaScript.PopMsg(Me, msg)
    '        Else
    '            SetTaskStatus("StsLSPlan", 525, 0)
    '            SetTaskProgress("ProLSPlan", -500)
    '            SetTaskButton("BLSPlan", False)
    '            SetTaskButton("BReset", True)
    '            SetTaskButton("BReport", True)
    '        End If
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BReport_Click)
    ''**     FCT Plan and LS Plan Report
    ''**
    ''*****************************************************************
    'Protected Sub BReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReport.Click
    '    '
    '    '*****************************************************************
    '    '**  準備報表資料
    '    '*****************************************************************
    '    '
    '    uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "Report", 1)
    '    '判斷是否完成作業
    '    Dim Code As Integer
    '    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2)
    '    If Code <> 0 Then
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 3)
    '    End If
    '    Do Until Code = 0
    '        System.Threading.Thread.Sleep(3 * 1000)
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2)
    '        If Code <> 0 Then
    '            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 3)
    '        End If
    '    Loop
    '    '
    '    '檢查是否已產出報表
    '    If uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2) = 0 Then
    '        SetTaskProgress("ProReport", -500)
    '        SetTaskButton("BReport", False)
    '        SetTaskButton("BReset", True)
    '        SetTaskButton("BIPLSPlan", True)
    '        SetTaskStatus("StsReport", 552, 0)
    '        uJavaScript.PopMsg(Me, "報表製作完成，如果已完成所有作業請按一下[Reset]退出FAS System．")
    '    Else
    '        SetTaskStatus("StsReport", 552, 1)
    '        uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
    '    End If
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BIPLSPlan_Click)
    ''**     Import LS Plan Data
    ''**
    ''*****************************************************************
    'Protected Sub BIPLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BIPLSPlan.Click
    '    '
    '    '*****************************************************************
    '    '**  匯入修改後LS Plan資料
    '    '*****************************************************************
    '    '
    '    uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "IPLSPlan", 1)
    '    '判斷是否完成作業
    '    Dim Code As Integer
    '    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2)
    '    If Code <> 0 Then
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 3)
    '    End If
    '    Do Until Code = 0
    '        System.Threading.Thread.Sleep(3 * 1000)
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2)
    '        If Code <> 0 Then
    '            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 3)
    '        End If
    '    Loop
    '    '
    '    '檢查是否已完成匯入作業
    '    If uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2) = 0 Then
    '        SetTaskProgress("ProIPLSPlan", -500)
    '        SetTaskButton("BIPLSPlan", False)
    '        SetTaskButton("BReset", True)
    '        SetTaskButton("BReport", True)
    '        SetTaskStatus("StsIPLSPlan", 912, 0)
    '        uJavaScript.PopMsg(Me, "匯入作業完成，如果已完成所有作業請按一下[Reset]退出FAS System．")
    '    Else
    '        SetTaskStatus("StsIPLSPlan", 912, 1)
    '        uJavaScript.PopMsg(Me, "匯入作業未完成，請確認!")
    '    End If
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BBULSPlan_Click)
    ''**     Buyer Local Stock Plan展開
    ''**
    ''*****************************************************************
    'Protected Sub BBULSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBULSPlan.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  Buyer Local Stock Plan展開
    '    '*****************************************************************
    '    If errcode = 0 Then
    '        ' 刪除執行履歷(Action=BULSPLAN)
    '        If errcode = 0 Then
    '            If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "BULSPLAN") <> 0 Then
    '                errcode = 9001
    '            End If
    '        End If
    '        ' 刪除BuyerLSPLAN
    '        If errcode = 0 Then
    '            If uFASCommon.DeleteBuyerLSData("BULS", DGRBuyer.Text) <> 0 Then
    '                errcode = 9005
    '            Else
    '                If uFASCommon.ResetBuyerLSNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
    '                    errcode = 9006
    '                End If
    '            End If
    '        End If
    '        ' Buyer LS Plan展開
    '        If errcode = 0 Then
    '            If uFASCommon.BuyerLocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
    '                errcode = 9002
    '            End If
    '        End If
    '        ' 更新客戶控制檔(BULSPlan=2, BUReport=1)
    '        If errcode = 0 Then
    '            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "BULSPlan", 2, "BUReport", 1) <> 0 Then
    '                errcode = 9003
    '            End If
    '        End If
    '        ' Buyer LS Plan展開是否完成
    '        If errcode = 0 Then
    '            If uFASCommon.CheckControlRecord(DBuyer.Text, "BULSPlan", 2) <> 0 Then
    '                errcode = 9004
    '            End If
    '        End If
    '        ' 執行結果處理
    '        If errcode <> 0 Then
    '            SetTaskStatus("StsBULSPlan", 885, 1)
    '            StsBULSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
    '            StsBULSPlan.Target = "_blank"
    '            '
    '            If errcode = 9001 Then msg = "刪除執行履歷異常(BULSPlan)，請連絡系統人員!"
    '            If errcode = 9002 Then msg = "Buyer LS Plan展開異常，請確認(可點選[ｘ]查詢)!"
    '            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '            If errcode = 9004 Then msg = "Buyer LS Plan展開異常，請連絡系統人員!"
    '            If errcode = 9005 Then msg = "刪除 Buyer LS PLAN 資料異常，請連絡系統人員!"
    '            If errcode = 9006 Then msg = "無法重新設定可使用Buyer LS No.，請連絡系統人員!"
    '            uJavaScript.PopMsg(Me, msg)
    '        Else
    '            SetTaskStatus("StsBULSPlan", 885, 0)
    '            SetTaskProgress("ProBULSPlan", -500)
    '            SetTaskButton("BBULSPlan", False)
    '            SetTaskButton("BReset", True)
    '            SetTaskButton("BBUReport", True)
    '        End If
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BBUReport_Click)
    ''**     Buyer LS Plan Report
    ''**
    ''*****************************************************************
    'Protected Sub BBUReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBUReport.Click
    '    '
    '    '*****************************************************************
    '    '**  準備報表資料
    '    '*****************************************************************
    '    '
    '    uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BUReport", 1)
    '    '判斷是否完成作業
    '    Dim Code As Integer
    '    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2)
    '    If Code <> 0 Then
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 3)
    '    End If
    '    Do Until Code = 0
    '        System.Threading.Thread.Sleep(3 * 1000)
    '        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2)
    '        If Code <> 0 Then
    '            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 3)
    '        End If
    '    Loop
    '    '
    '    '檢查是否已產出報表
    '    If uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2) = 0 Then
    '        SetTaskProgress("ProBUReport", -500)
    '        SetTaskButton("BBUReport", False)
    '        SetTaskButton("BReset", True)
    '        SetTaskButton("BLFLSPlan", True)
    '        SetTaskStatus("StsBUReport", 912, 0)
    '    Else
    '        SetTaskStatus("StsBUReport", 912, 1)
    '        uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
    '    End If
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BLFLSPlan_Click)
    ''**     Buyer LS Plan 最終確定
    ''**
    ''*****************************************************************
    'Protected Sub BLFLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLFLSPlan.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  Buyer LS Plan 最終確定
    '    '*****************************************************************
    '    If errcode = 0 Then
    '        ' 刪除執行履歷(Action=LFLSPLAN)
    '        If errcode = 0 Then
    '            If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "LFLSPLAN") <> 0 Then
    '                errcode = 9001
    '            End If
    '        End If
    '        ' 判斷是否可執行
    '        If errcode = 0 Then
    '            If uFASCommon.CanRunLFLocalStockPlan(DBuyer.Text, DGRBuyer.Text) <> 0 Then
    '                errcode = 9005
    '            End If
    '        End If
    '        ' Buyer LS Plan最終確定
    '        If errcode = 0 Then
    '            If uFASCommon.LastFinalLocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
    '                errcode = 9002
    '            End If
    '        End If
    '        ' 更新客戶控制檔(BULSPlan=2, BUReport=1)
    '        If errcode = 0 Then
    '            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "LFLSPlan", 2, "EPEDI", 1) <> 0 Then
    '                errcode = 9003
    '            End If
    '        End If
    '        ' Buyer LS Plan最終確定是否完成
    '        If errcode = 0 Then
    '            If uFASCommon.CheckControlRecord(DBuyer.Text, "LFLSPlan", 2) <> 0 Then
    '                errcode = 9004
    '            End If
    '        End If
    '        ' 執行結果處理
    '        If errcode <> 0 Then
    '            SetTaskStatus("StsLFLSPlan", 1245, 1)
    '            StsLFLSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
    '            StsLFLSPlan.Target = "_blank"
    '            '
    '            If errcode = 9001 Then msg = "刪除執行履歷異常(LFLSPlan)，請連絡系統人員!"
    '            If errcode = 9002 Then msg = "Buyer LS Plan最終確定異常，請確認(可點選[ｘ]查詢)!"
    '            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
    '            If errcode = 9004 Then msg = "Buyer LS Plan最終確定異常，請連絡系統人員!"
    '            If errcode = 9005 Then msg = "異常:懷疑最終確定作業重覆執行，請連絡系統人員!"
    '            uJavaScript.PopMsg(Me, msg)
    '        Else
    '            SetTaskStatus("StsLFLSPlan", 1245, 0)
    '            SetTaskProgress("ProLFLSPlan", -500)
    '            SetTaskButton("BLFLSPlan", False)
    '            SetTaskButton("BReset", True)
    '            SetTaskButton("BEPEDI", True)
    '            uJavaScript.PopMsg(Me, "[FAS]第1Step的線上功能到此為止，請按一下[Reset]退出FAS System．")
    '        End If
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BEPEDI_Click)
    ''**     Export EDI Data
    ''**
    ''*****************************************************************
    'Protected Sub BEPEDI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BEPEDI.Click
    '    Dim msg As String = ""
    '    Dim errcode As Integer = 0
    '    '
    '    '*****************************************************************
    '    '**  Export EDI Data
    '    '*****************************************************************
    '    If errcode = 0 Then
    '        uJavaScript.PopMsg(Me, "此功能目前未上線，如果已完成所有作業請按一下[Reset]退出FAS System．")
    '    End If
    '    '
    'End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BActionLog_Click)
    ''**     查詢ActionLog
    ''**
    ''*****************************************************************
    'Protected Sub BActionLog_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BActionLog.Click
    '    Dim Cmd As String
    '    '
    '    Cmd = "<script>" + _
    '                "window.open('ActionHistory.aspx?pLogID=" & DLogID.Text & "&pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','ActionLog','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
    '          "</script>"
    '    '
    '    Response.Write(Cmd)
    'End Sub


End Class
