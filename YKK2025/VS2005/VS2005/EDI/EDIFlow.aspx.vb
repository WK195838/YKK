Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Partial Class EDIFlow
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
    Dim uEDIMapping As New EDI2011.MappingService
    Dim uEDICommon As New EDI2011.CommonService
    Dim uWFSCommon As New WFS.CommonService
    Dim NowDateTime As String               ' 現在日時
    Dim xUserID As String                   ' 使用者ID
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        If Not IsPostBack Then
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uEDIMapping.Timeout = 900 * 1000
        uEDICommon.Timeout = 900 * 1000
        uWFSCommon.Timeout = 900 * 1000
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        DLogID.Style("left") = -500 & "px"
        DBuyer.Style("left") = -500 & "px"
        DGRBuyer.Style("left") = -500 & "px"
        DUserID.Style("left") = -500 & "px"
        DFileName.Style("left") = -500 & "px"
        DFilePath.Style("left") = -500 & "px"
        Inf_PriceList.Style("left") = -500 & "px"
        '動作按鈕設定
        BCustomer.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此客戶？" + "');if(!ok){return false;}"
        BExcel.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟Excel準備客戶資料？" + "');if(!ok){return false;} else {CheckAttribute('Excel')}"
        BImport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行匯入作業？" + "');if(!ok){return false;} else {CheckAttribute('Import')}"
        BDataCheck.Attributes("onclick") = "var ok=window.confirm('" + "是否執行檢查資料？" + "');if(!ok){return false;} else {CheckAttribute('DataCheck')}"
        BWaves.Attributes("onclick") = "var ok=window.confirm('" + "是否執行匯出至轉Waves？" + "');if(!ok){return false;} else {CheckAttribute('Waves')}"
        BSalesPrice.Attributes("onclick") = "var ok=window.confirm('" + "是否已執行Waves 10-03-07. Receive EDI Order？" + "');if(!ok){return false;} else {CheckAttribute('Price')}"
        BLULUWIP.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟LULU-WIP資料？" + "');if(!ok){return false;} else {CheckAttribute('Excel')}"
        'Action CheckBox設定
        AtOffice2010.Enabled = True
        '公告
        LBoard.Enabled = False

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetCustomerBuyer)
    '**     設定可使用之客戶
    '**
    '*****************************************************************
    Sub SetCustomerBuyer(ByVal pAction As String, ByVal pValue As String)
        Dim sql As String
        Select Case pAction
            Case "Default"
                Dim i As Integer
                '
                DCustomerBuyer.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "001" & "'"
                sql = sql & "   And DKey = '" & xUserID & "'"
                sql = sql & "   And Data Not Like 'F-%' "
                sql = sql & "Order by Data "
                Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp.Rows.Count - 1
                    DCustomerBuyer.Items.Add(dt_Referp.Rows(i).Item("Data"))
                Next
            Case Else
                DCustomerBuyer.Items.Clear()
                DCustomerBuyer.Items.Add(pValue)
        End Select
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
            Case "StsCustomer"
                StsCustomer.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCustomer.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCustomer.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsExcel"
                StsExcel.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsExcel.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsExcel.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsImport"
                StsImport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsImport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsImport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsMakePONO"
                StsMakePONO.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsMakePONO.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsMakePONO.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckPONO"
                StsCheckPONO.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckPONO.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckPONO.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckGRPC"
                StsCheckGRPC.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckGRPC.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckGRPC.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckCompanyCode"
                StsCheckCompanyCode.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckCompanyCode.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckCompanyCode.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckKeepCode"
                StsCheckKeepCode.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckKeepCode.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckKeepCode.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckColorCode"
                StsCheckColorCode.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckColorCode.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckColorCode.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckItemCode"
                StsCheckItemCode.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckItemCode.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckItemCode.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckDuplicateData"
                StsCheckDuplicateData.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckDuplicateData.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckDuplicateData.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsCheckPOStructure"
                StsCheckPOStructure.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsCheckPOStructure.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsCheckPOStructure.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsWaves"
                StsWaves.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsWaves.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsWaves.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsSalesPrice"
                StsSalesPrice.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsSalesPrice.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsSalesPrice.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsCustomer.Style("left") = pLeft & "px"
                StsExcel.Style("left") = pLeft & "px"
                StsImport.Style("left") = pLeft & "px"
                StsMakePONO.Style("left") = pLeft & "px"
                StsCheckPONO.Style("left") = pLeft & "px"
                StsCheckGRPC.Style("left") = pLeft & "px"
                StsCheckCompanyCode.Style("left") = pLeft & "px"
                StsCheckKeepCode.Style("left") = pLeft & "px"
                StsCheckColorCode.Style("left") = pLeft & "px"
                StsCheckItemCode.Style("left") = pLeft & "px"
                StsCheckDuplicateData.Style("left") = pLeft & "px"
                StsCheckPOStructure.Style("left") = pLeft & "px"
                StsWaves.Style("left") = pLeft & "px"
                StsSalesPrice.Style("left") = pLeft & "px"
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
            Case "ProExcel"
                ProExcel.Style("left") = pLeft & "px"
            Case "ProImport"
                ProImport.Style("left") = pLeft & "px"
            Case "ProDataCheck"
                ProDataCheck.Style("left") = pLeft & "px"
            Case "ProWaves"
                ProWaves.Style("left") = pLeft & "px"
            Case "ProSalesPrice"
                ProSalesPrice.Style("left") = pLeft & "px"
            Case Else
                ProExcel.Style("left") = pLeft & "px"
                ProImport.Style("left") = pLeft & "px"
                ProDataCheck.Style("left") = pLeft & "px"
                ProWaves.Style("left") = pLeft & "px"
                ProSalesPrice.Style("left") = pLeft & "px"
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
            Case "BCustomer"
                BCustomer.Visible = pShow
            Case "BExcel"
                BExcel.Visible = pShow
            Case "BImport"
                BImport.Visible = pShow
            Case "BDataCheck"
                BDataCheck.Visible = pShow
            Case "BWaves"
                BWaves.Visible = pShow
            Case "BSalesPrice"
                BSalesPrice.Visible = pShow
            Case "BProgress"
                BProgress.Visible = pShow
            Case "BPackList"
                BPackList.Visible = pShow
            Case "BCheckBuyerPrice"
                BCheckBuyerPrice.Visible = pShow
            Case "BActionLog"
                BActionLog.Visible = pShow
            Case "BSCIJobList"
                BSCIJobList.Visible = pShow
            Case "BDSJobList"
                BDSJobList.Visible = pShow
            Case "BQMIJobList"
                BQMIJobList.Visible = pShow
            Case "BLULUWIP"
                BLULUWIP.Visible = pShow
            Case "BEOES"
                BEOES.Visible = pShow
            Case Else
                BCustomer.Visible = True
                BReset.Visible = False
                BExcel.Visible = False

                BSalesPrice.Visible = False
                BProgress.Visible = False
                BPackList.Visible = False
                BCheckBuyerPrice.Visible = True
                If UCase(xUserID) = "IT003@@" Then
                    BActionLog.Visible = True
                    BImport.Visible = True
                    BDataCheck.Visible = True
                    BWaves.Visible = True
                    BSalesPrice.Visible = True
                Else
                    BActionLog.Visible = False
                    BImport.Visible = False
                    BDataCheck.Visible = False
                    BWaves.Visible = False
                End If
                BSCIJobList.Visible = False
                BDSJobList.Visible = False
                BQMIJobList.Visible = False
                BLULUWIP.Visible = False
                BEOES.Visible = False
                '
                'Test-Start
                'BReset.Visible = True
                'BExcel.Visible = True
                'BImport.Visible = True
                '@@
                'BDataCheck.Visible = True
                'BWaves.Visible = True
                'BSalesPrice.Visible = True
                'BProgress.Visible = True
                'BPackList.Visible = True
                'BCheckBuyerPrice.Visible = True
                'BActionLog.Visible = True
                'BSCIJobList.Visible = True
                'BDSJobList.Visible = True
                'BQMIJobList.Visible = True
                'BLULUWIP.Visible = True
                'BEOES.Visible = True
                'Test-End
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BCustomer_Click)
    '**     客戶選擇
    '**
    '*****************************************************************
    Protected Sub BCustomer_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BCustomer.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String
        '
        sql = "Select * From M_ControlRecord "
        sql = sql & " Where Name = '" & DCustomerBuyer.SelectedValue & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            DLogID.Text = Now.ToString("yyyyMMddHHmmss")
            'Test-Start
            'DLogID.Text = "20130104114533"
            'Test-End

            DBuyer.Text = dt_ControlRecord.Rows(0).Item("Buyer")
            DGRBuyer.Text = dt_ControlRecord.Rows(0).Item("BuyerGroup")
            DUserID.Text = xUserID
            mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
            '
            If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Start", 0, "Start", 0) = 0 Then
                    SetTaskButton("BCustomer", False)
                    SetTaskButton("BReset", True)
                    SetTaskButton("BExcel", True)
                    SetTaskStatus("StsCustomer", 192, 0)
                    SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
                    '
                    If AtOffice2010.Checked = True Then
                        DFileName.Text = Trim(DCustomerBuyer.SelectedValue) + "_DataPrepare.xlsm"
                    Else
                        DFileName.Text = Trim(DCustomerBuyer.SelectedValue) + "_DataPrepare.xls"
                    End If
                    DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                    '
                    'EOES 複製 EOES_DataPrepare --> EOES-[USERID]_DataPrepare
                    If InStr(DBuyer.Text, "EOES-") > 0 Then
                        If Not File.Exists(DFilePath.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "EOES_DataPrepare.xlsm", DFilePath.Text)
                        End If
                    End If
                    '
                    'LULUWIP 複製 LULUWIP_DataPrepare --> LULUWIP-[USERID]_DataPrepare
                    If InStr(DBuyer.Text, "LULUWIP-") > 0 Then
                        If Not File.Exists(DFilePath.Text) Then
                            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "LULUWIP_DataPrepare.xlsm", DFilePath.Text)
                        End If
                        BExcel.Style("left") = -500 & "px"
                    End If
                    '
                    '使用訂單進度--> GRBuyer(8)=O
                    If fpObj.GetFunctionCode(DGRBuyer.Text, 8) = "O" Then
                        SetTaskButton("BProgress", True)
                    End If
                    '
                    '使用PackList--> GRBuyer(9)=O
                    If fpObj.GetFunctionCode(DGRBuyer.Text, 9) = "O" Then
                        SetTaskButton("BPackList", True)
                    End If
                    '
                    ' 客戶別專屬功能
                    Select Case dt_ControlRecord.Rows(0).Item("ExclusiveFunction")
                        Case 1
                            SetTaskButton("BSCIJobList", True)
                        Case 2
                            SetTaskButton("BDSJobList", True)
                        Case 3
                            SetTaskButton("BQMIJobList", True)
                        Case 4
                            SetTaskButton("BLULUWIP", True)
                        Case 9
                            SetTaskButton("BEOES", True)
                        Case Else
                    End Select
                    '
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
                Else
                    errcode = 9003
                End If
            Else
                If mUserID = xUserID Then
                    errcode = 9004      '同上次使用者(上次不小心)
                    SetTaskButton("BReset", True)
                    SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
                Else
                    errcode = 9002      '不同上次使用者
                End If
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此客戶不存在，請確認!"
            If errcode = 9002 Then msg = "此客戶正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9004 Then msg = "懷疑您上次使用有不正常關閉系統，造成此客戶正在使用中! 請[Reset]一次再使用!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BReset_Click)
    '**     客戶重新選擇
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReset.Click
        '
        If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            ' Clear 各動作 按鈕/結果/跑馬燈等圖像
            ClearKeyData()
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
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
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BExcel_Click)
    '**     準備客戶資料
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        If File.Exists(DFilePath.Text) Then
            '判斷是否完成作業
            uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "DataPrepare", 1)
            '
            Dim Code As Integer
            Code = uEDICommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
            If Code <> 0 Then
                Code = uEDICommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
            End If
            Do Until Code = 0
                System.Threading.Thread.Sleep(3 * 1000)
                Code = uEDICommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
                If Code <> 0 Then
                    Code = uEDICommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
                End If
            Loop
            '
            '檢查是否已將資料上傳完成
            If uEDICommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2) = 0 Then
                SetTaskProgress("ProExcel", -500)
                SetTaskButton("BExcel", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BImport", True)
                SetTaskStatus("StsExcel", 192, 0)
            Else
                SetTaskStatus("StsExcel", 192, 1)
                uJavaScript.PopMsg(Me, "客戶資料準備未完成，請確認!")
            End If
        Else
            SetTaskStatus("StsExcel", 192, 1)
            uJavaScript.PopMsg(Me, "客戶EXCEL SHEET不存在，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BImport_Click)
    '**     匯入作業
    '**
    '*****************************************************************
    Protected Sub BImport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BImport.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        ' 刪除執行履歷/S5K00/S5L00/SC760W1(NIKE-VDP)
        If errcode = 0 Then
            If uEDICommon.DeleteHistory(DBuyer.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        If errcode = 0 Then
            If uEDICommon.DeleteS5K00Data(DBuyer.Text) <> 0 Then
                errcode = 9002
            End If
        End If
        If errcode = 0 Then
            If uEDICommon.DeleteS5L00Data(DBuyer.Text) <> 0 Then
                errcode = 9003
            End If
        End If
        If errcode = 0 Then
            If uEDICommon.DeleteSC760W1Data(DBuyer.Text) <> 0 Then
                errcode = 9004
            End If
        End If
        ' 匯入作業
        If errcode = 0 Then
            If InStr(DBuyer.Text, "EOES-") > 0 Then
                ' **EOES-START
                If uEDIMapping.Rule2DataEOES(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9011
                End If
                ' **EOES-END
            Else
                ' **原程式-START
                If uEDIMapping.Rule2Data(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9011
                End If
                ' **原程式-END
            End If
        End If
        ' 更新客戶控制檔(DataImport=2, CreatePONO=1)
        If errcode = 0 Then
            If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataImport", 2, "CreatePONO", 1) <> 0 Then
                errcode = 9012
            End If
        End If
        ' 檢查匯入作業是否完成
        If errcode = 0 Then
            If uEDICommon.CheckControlRecord(DBuyer.Text, "DataImport", 2) <> 0 Then
                errcode = 9013
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsImport", 192, 1)
            StsImport.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
            StsImport.Target = "_blank"
            '
            If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
            If errcode = 9002 Then msg = "刪除客戶資料異常(S5K00)，請連絡系統人員!"
            If errcode = 9003 Then msg = "刪除客戶資料異常(S5L00)，請連絡系統人員!"
            If errcode = 9004 Then msg = "刪除客戶資料異常(SC760W1)，請連絡系統人員!"
            If errcode = 9011 Then msg = "匯入資料異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9012 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9013 Then msg = "匯入資料異常，請連絡系統人員!"
            uJavaScript.PopMsg(Me, msg)
        Else
            StsImport.NavigateUrl = ""
            StsImport.Target = ""
            SetTaskProgress("ProImport", -500)
            SetTaskButton("BImport", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BDataCheck", True)
            SetTaskStatus("StsImport", 192, 0)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BDataCheck_Click)
    '**     檢查資料
    '**
    '*****************************************************************
    Protected Sub BDataCheck_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDataCheck.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  製作採購號碼(MakePONO)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=MakePONO)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "MakePONO") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 製作採購號碼
            If errcode = 0 Then
                If fpObj.GetFunctionCode(DGRBuyer.Text, 2) = "P" Then        ' 製作採購號碼 --> GRBuyer(2)=P 
                    If uEDICommon.MakePONO(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                        errcode = 9002
                    End If
                End If
            End If
            ' 更新客戶控制檔(MakePONO=2, OrderNo=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "CreatePONO", 2, "OrderNo", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢查製作採購號碼是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "CreatePONO", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsMakePONO", 520, 1)
                StsMakePONO.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsMakePONO.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(MakePONO)，請連絡系統人員!"
                If errcode = 9002 Then msg = "訂單號碼異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "產生訂單號碼異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsMakePONO", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測採購號碼(CheckPONO)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckPONO)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckPONO") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 檢測採購號碼
            If errcode = 0 Then
                If uEDICommon.CheckPONO(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(CheckPONO=2, GRPC=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "OrderNo", 2, "GRPC", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測採購號碼是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "OrderNo", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckPONO", 520, 1)
                StsCheckPONO.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckPONO.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckPONO)，請連絡系統人員!"
                If errcode = 9002 Then msg = "訂單號碼異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測訂單號碼異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckPONO", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  製作Group Code(MakeGRPC)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=MakeGRPC)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "MakeGRPC") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 製作Group Code
            If errcode = 0 Then
                If uEDICommon.MakeGRPC(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(GRPC=2, Company=1)
            'If errcode = 0 Then
            '    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "GRPC", 2, "Company", 1) <> 0 Then
            '        errcode = 9003
            '    End If
            'End If
            ' 製作Group Code是否完成
            'If errcode = 0 Then
            '    If uEDICommon.CheckControlRecord(DBuyer.Text, "GRPC", 2) <> 0 Then
            '        errcode = 9004
            '    End If
            'End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckGRPC", 520, 1)
                StsCheckGRPC.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckGRPC.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(MakeGRPC)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Group Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "製作Group Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckGRPC", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測Group Code(CheckGRPC)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckGRPC)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckGRPC") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 檢測Group Code
            If errcode = 0 Then
                If uEDICommon.CheckGRPC(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 檢測SampleQty
            If errcode = 0 Then
                If uEDICommon.CheckSampleQty(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9005
                End If
            End If
            ' 更新客戶控制檔(GRPC=2, Company=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "GRPC", 2, "Company", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測Group Code是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "GRPC", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckGRPC", 520, 1)
                StsCheckGRPC.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckGRPC.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckGRPC)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Group Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測Group Code異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "樣品數量異常，請確認(可點選[ｘ]查詢)!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckGRPC", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測Company Code(CheckCompanyCode)
        '**     對象:非LS Order
        '*****************************************************************
        If errcode = 0 Then
            '
            ' 刪除執行履歷(Action=CheckCompanyCode)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckCompanyCode") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 檢測Company Code
            If errcode = 0 Then
                If uEDICommon.CheckCompanyCode(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Company=2, Color=1)
            'If errcode = 0 Then
            '    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Company", 2, "Color", 1) <> 0 Then
            '        errcode = 9003
            '    End If
            'End If
            ' 檢測Company Code是否完成
            'If errcode = 0 Then
            '    If uEDICommon.CheckControlRecord(DBuyer.Text, "Company", 2) <> 0 Then
            '        errcode = 9004
            '    End If
            'End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckCompanyCode", 520, 1)
                StsCheckCompanyCode.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckCompanyCode.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckCompanyCode)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Company Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測Company Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckCompanyCode", 520, 0)
            End If
        End If

        '
        '*****************************************************************
        '**  檢測Keep Code(CheckKeepCode)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckKeepCode)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckKeepCode") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' @@檢測Keep Code
            If errcode = 0 Then
                If uEDICommon.CheckKeepCode(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Company=2, Color=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Company", 2, "Color", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測Keep Code是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "Company", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckKeepCode", 520, 1)
                StsCheckKeepCode.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckKeepCode.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckKeepCode)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Keep Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測Keep Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckKeepCode", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  製作Color Code(MakeColorCode)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=MakeColorCode)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "MakeColorCode") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 製作Color Code
            If errcode = 0 Then
                If uEDICommon.MakeColorCode(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Color=2, Item=1)
            'If errcode = 0 Then
            '    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Color", 2, "Item", 1) <> 0 Then
            '        errcode = 9003
            '    End If
            'End If
            ' 製作Color Code是否完成
            'If errcode = 0 Then
            '    If uEDICommon.CheckControlRecord(DBuyer.Text, "Color", 2) <> 0 Then
            '        errcode = 9004
            '    End If
            'End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckColorCode", 520, 1)
                StsCheckColorCode.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckColorCode.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(MakeColorCode)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Color Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "製作Color Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckColorCode", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測Color Code(CheckColorCode)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckColorCode)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckColorCode") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' @@檢測Color Code
            If errcode = 0 Then
                If uEDICommon.CheckColorCode(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Color=2, Item=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Color", 2, "Item", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測Color Code是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "Color", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckColorCode", 520, 1)
                StsCheckColorCode.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckColorCode.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckColorCode)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Color Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測Color Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckColorCode", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測Item Code(CheckItemCode)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckItemCode)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckItemCode") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' @@檢測Item Code
            If errcode = 0 Then
                If uEDICommon.CheckItemCode(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Item=2, Duplication=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Item", 2, "Duplication", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測Item Code是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "Item", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckItemCode", 520, 1)
                StsCheckItemCode.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckItemCode.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckItemCode)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Item Code異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測Item Code異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckItemCode", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測Duplicate(CheckDuplicateData)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckDuplicateData)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckDuplicateData") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' @@檢測Duplicate
            If errcode = 0 Then
                If uEDICommon.CheckDuplicateData(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(Duplication=2, DataBalance=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Duplication", 2, "DataBalance", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測Duplicate是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "Duplication", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckDuplicateData", 520, 1)
                StsCheckDuplicateData.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckDuplicateData.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckDuplicateData)，請連絡系統人員!"
                If errcode = 9002 Then msg = "訂單有重覆異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測訂單重覆異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckDuplicateData", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測NIKE VDP(CheckNikeVDP)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckNikeVDP)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckNikeVDP") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 檢測NIKE VDP---2012/1/19 支援NIKE以外VDP時廢除
            'If errcode = 0 Then
            '    If uEDICommon.CheckNikeVDP(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
            '        errcode = 9002
            '    End If
            'End If
            ' 更新客戶控制檔(DataBalance=2, ToWaves=1)
            'If errcode = 0 Then
            '    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataBalance", 2, "ToWaves", 1) <> 0 Then
            '        errcode = 9003
            '    End If
            'End If
            ' 檢測NIKE VDP是否完成
            'If errcode = 0 Then
            '    If uEDICommon.CheckControlRecord(DBuyer.Text, "DataBalance", 2) <> 0 Then
            '        errcode = 9004
            '    End If
            'End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckPOStructure", 520, 1)
                StsCheckPOStructure.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckPOStructure.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckNikeVDP)，請連絡系統人員!"
                If errcode = 9002 Then msg = "訂單資料不齊異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測NIKE VDP異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckPOStructure", 520, 0)
            End If
        End If
        '
        '*****************************************************************
        '**  檢測客戶PO結構資料(CheckPOStructure)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=CheckPOStructure)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "CheckPOStructure") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 檢測客戶PO結構資料
            If errcode = 0 Then
                If uEDICommon.CheckPOStructure(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(DataBalance=2, ToWaves=1)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataBalance", 2, "ToWaves", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 檢測客戶PO結構資料是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "DataBalance", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsCheckPOStructure", 520, 1)
                StsCheckPOStructure.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsCheckPOStructure.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(CheckPOStructure)，請連絡系統人員!"
                If errcode = 9002 Then msg = "訂單資料不齊異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "檢測客戶PO結構資料異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsCheckPOStructure", 520, 0)
            End If
        End If
        '
        ' 執行結果處理
        If errcode = 0 Then
            SetTaskProgress("ProDataCheck", -500)
            SetTaskButton("BDataCheck", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BWaves", True)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BWaves_Click)
    '**     EDI Data To Waves System
    '**
    '*****************************************************************
    Protected Sub BWaves_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BWaves.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  匯出至轉Waves(EDI2Waves)
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=EDI2Waves)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "EDI2Waves") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 匯出至轉Waves
            If errcode = 0 Then
                If uEDICommon.EDI2Waves(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' Keep進Wave's前資料(B2B對應)
            'GRBuyer(6)=    [T] --> PRICE & SHIP / [X] --> 不使用
            If errcode = 0 Then
                If fpObj.GetFunctionCode(DGRBuyer.Text, 6) <> "X" Then
                    If uEDICommon.EDI2B2B(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                        errcode = 9005
                    End If
                End If
            End If
            ' 更新客戶控制檔
            If errcode = 0 Then
                If fpObj.GetFunctionCode(DGRBuyer.Text, 7) = "P" Then        ' GRBuyer(7)=P 
                    ' (ToWaves=2, GetSalesPrice=1)
                    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "ToWaves", 2, "GetSalesPrice", 1) <> 0 Then
                        errcode = 9003
                    End If
                Else
                    '(指定只更新Action2, ToWaves=2)
                    If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "ToWaves", 2) <> 0 Then
                        errcode = 9003
                    End If
                End If
            End If
            ' 匯出至轉Waves是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "ToWaves", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskButton("BWaves", False)      ' 2014/5/2 追加 
                SetTaskStatus("StsWaves", 917, 1)
                StsWaves.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsWaves.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(EDI2Waves)，請連絡系統人員!"
                If errcode = 9002 Then msg = "轉Waves異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "匯出至轉Waves異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "匯出至B2B連接檔異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskProgress("ProWaves", -500)
                SetTaskButton("BWaves", False)
                SetTaskButton("BReset", True)
                SetTaskStatus("StsWaves", 917, 0)
                '
                If fpObj.GetFunctionCode(DGRBuyer.Text, 7) = "P" Then        ' GRBuyer(7)=P 
                    SetTaskButton("BSalesPrice", True)
                End If
                '
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "End", 0, "End", 0) = 0 Then
                Else
                    uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
                End If
                '
            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSalesPrice_Click)
    '**     Get Sales Price Inf
    '**
    '*****************************************************************
    Protected Sub BSalesPrice_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSalesPrice.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  取得Waves Sales Price Inf
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=GetSalesPrice)
            If errcode = 0 Then
                If uEDICommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "GetSalesPrice") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 展開PriceList資料a(Wave's Price對應)
            If errcode = 0 Then
                If uEDICommon.PriceList(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(指定只更新Action2, GetSalesPrice=2)
            If errcode = 0 Then
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "GetSalesPrice", 2) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 匯出至轉Waves是否完成
            If errcode = 0 Then
                If uEDICommon.CheckControlRecord(DBuyer.Text, "GetSalesPrice", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsSalesPrice", 917, 1)
                StsSalesPrice.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsSalesPrice.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(GetSalesPrice)，請連絡系統人員!"
                If errcode = 9002 Then msg = "取得 Price Inf.異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "取得 Price Inf.異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskProgress("ProSalesPrice", -500)
                SetTaskButton("BSalesPrice", False)
                SetTaskButton("BReset", True)
                SetTaskStatus("StsSalesPrice", 917, 0)
                '
                Inf_PriceList.Style("left") = 732 & "px"
                Inf_PriceList.NavigateUrl = "InfPriceList.aspx?pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                '
                If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "End", 0, "End", 0) = 0 Then
                Else
                    uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
                End If
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BProgressr_Click)
    '**     查詢進度
    '**
    '*****************************************************************
    Protected Sub BProgress_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BProgress.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfOrderProgress.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','OrderProgress','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BCheckBuyerPrice_Click)
    '**     查詢BuyerPrice
    '**
    '*****************************************************************
    Protected Sub BCheckBuyerPrice_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BCheckBuyerPrice.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/IRW/OrderBuyerPriceCheckList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','OrderProgress','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BActionLog_Click)
    '**     查詢ActionLog
    '**
    '*****************************************************************
    Protected Sub BActionLog_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BActionLog.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('ActionHistory.aspx?pLogID=" & DLogID.Text & "&pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','ActionLog','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BPackList_Click)
    '**     查詢PackList
    '**
    '*****************************************************************
    Protected Sub BPackList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPackList.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfPackingList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','PackList','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSCIJobList_Click)
    '**     SCI客戶-專用功能 
    '**
    '*****************************************************************
    Protected Sub BSCIJobList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSCIJobList.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfSCIJobList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','SCIJobList','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BDSJobList_Click)
    '**     DS客戶-專用功能 
    '**
    '*****************************************************************
    Protected Sub BDSJobList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDSJobList.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfDSJobList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "','DSJobList','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BQMIJobList_Click)
    '**     QMI客戶-專用功能 
    '**
    '*****************************************************************
    Protected Sub BQMIJobList_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BQMIJobList.Click
        Dim Cmd As String
        '
        If AtOffice2010.Checked = True Then
            Cmd = "<script>" + _
                        "window.open('InfQMIJobList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "&p2010=" + "1" & "','QMIJobList','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        Else
            Cmd = "<script>" + _
                        "window.open('InfQMIJobList.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "&p2010=" + "0" & "','QMIJobList','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If

        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BLULUWIP_Click)
    '**     LULU客戶-專用功能 
    '**
    '*****************************************************************
    Protected Sub BLULUWIP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLULUWIP.Click
        '
        If uEDICommon.UpdateControlRecord(DBuyer.Text, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            ' Clear 各動作 按鈕/結果/跑馬燈等圖像
            ClearKeyData()
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
        Else
            uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BEOES)
    '**     EOES客戶-專用功能 
    '**
    '*****************************************************************
    Protected Sub BEOES_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BEOES.Click
        Dim Cmd As String
        '
        If AtOffice2010.Checked = True Then
            Cmd = "<script>" + _
                        "window.open('EOESMenu.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "&p2010=" + "1" & "','EOESMenu','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        Else
            Cmd = "<script>" + _
                        "window.open('EOESMenu.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" + xUserID & "&p2010=" + "0" & "','EOESMenu','status=1,toolbar=0,width=1000,height=600,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If

        '
        Response.Write(Cmd)
    End Sub

End Class
