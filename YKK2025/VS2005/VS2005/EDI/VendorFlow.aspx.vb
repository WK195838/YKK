Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading

Partial Class VendorFlow
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
    Dim uFASMapping As New EDI2011.FMappingService
    Dim uFASCommon As New EDI2011.FCommonService
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
            SetSalesMan("Default", "")                  ' 設定可使用業務員
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        DLogID.Style("left") = -500 & "px"
        DBuyer.Style("left") = -500 & "px"
        DGRBuyer.Style("left") = -500 & "px"
        DFunList.Style("left") = -500 & "px"
        DUserID.Style("left") = -500 & "px"
        DFileName.Style("left") = -500 & "px"
        DFilePath.Style("left") = -500 & "px"
        DFileName1.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        '動作按鈕設定
        BSalesMan.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此業務員？" + "');if(!ok){return false;}"
        BUpdateFC.Attributes("onclick") = "var ok=window.confirm('" + "是否取得[FC]資料？" + "');if(!ok){return false;} else {CheckAttribute('UpdateFC')}"
        BAppendFC.Attributes("onclick") = "var ok=window.confirm('" + "是否準備[FC]資料？" + "');if(!ok){return false;} else {CheckAttribute('AppendFC')}"
        BConfirm.Attributes("onclick") = "var ok=window.confirm('" + "是否確認[FC]資料？" + "');if(!ok){return false;} else {CheckAttribute('Confirm')}"
        BFinal.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[移送管理人]作業？" + "');if(!ok){return false;} else {CheckAttribute('Final')}"
        'Action CheckBox設定
        AtAppendFC.Attributes.Add("onclick", "ActionCheck('AtAppendFC')")
        AtConfirm.Attributes.Add("onclick", "ActionCheck('AtConfirm')")
        AtFinal.Attributes.Add("onclick", "ActionCheck('AtFinal')")
        AtOffice2010.Enabled = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSalesMan)
    '**     設定可使用之業務員
    '**
    '*****************************************************************
    Sub SetSalesMan(ByVal pAction As String, ByVal pValue As String)
        Dim sql As String
        Select Case pAction
            Case "Default"
                Dim i As Integer
                '
                DSalesMan.Items.Clear()
                sql = "Select * From M_Referp "
                sql = sql & " Where Cat  = '" & "001" & "' "
                sql = sql & "   And DKey = '" & xUserID & "' "
                sql = sql & "   And Data Like 'VF-%' "
                sql = sql & "Order by Data "
                Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
                For i = 0 To dt_Referp.Rows.Count - 1
                    DSalesMan.Items.Add(dt_Referp.Rows(i).Item("Data"))
                Next
            Case Else
                DSalesMan.Items.Clear()
                DSalesMan.Items.Add(pValue)
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
            Case "StsSalesMan"
                StsSalesMan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsSalesMan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsSalesMan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsUpdateFC"
                StsUpdateFC.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsUpdateFC.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsUpdateFC.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsAppendFC"
                StsAppendFC.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsAppendFC.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsAppendFC.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsConfirm"
                StsConfirm.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsConfirm.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsConfirm.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsFinal"
                StsFinal.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsFinal.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsFinal.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsSalesMan.Style("left") = pLeft & "px"
                StsUpdateFC.Style("left") = pLeft & "px"
                StsAppendFC.Style("left") = pLeft & "px"
                StsConfirm.Style("left") = pLeft & "px"
                StsFinal.Style("left") = pLeft & "px"
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
            Case "ProUpdateFC"
                ProUpdateFC.Style("left") = pLeft & "px"
            Case "ProAppendFC"
                ProAppendFC.Style("left") = pLeft & "px"
            Case "ProConfirm"
                ProConfirm.Style("left") = pLeft & "px"
            Case "ProFinal"
                ProFinal.Style("left") = pLeft & "px"
            Case Else
                ProUpdateFC.Style("left") = pLeft & "px"
                ProAppendFC.Style("left") = pLeft & "px"
                ProConfirm.Style("left") = pLeft & "px"
                ProFinal.Style("left") = pLeft & "px"
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
            Case "BSalesMan"
                BSalesMan.Visible = pShow
            Case "BUpdateFC"
                BUpdateFC.Visible = pShow
            Case "BAppendFC"
                BAppendFC.Visible = pShow
            Case "BConfirm"
                BConfirm.Visible = pShow
            Case "BFinal"
                BFinal.Visible = pShow
            Case Else
                BReset.Visible = False
                BSalesMan.Visible = True
                BUpdateFC.Visible = False
                BAppendFC.Visible = False
                BConfirm.Visible = False
                BFinal.Visible = False
                '
                AtAppendFC.Checked = False
                AtConfirm.Checked = False
                AtFinal.Checked = False
                '
                'Test-Start
                'BReset.Visible = True
                'BSalesMan.Visible = True
                'BUpdateFC.Visible = True
                'BAppendFC.Visible = True
                'BConfirm.Visible = True
                'BFinal.Visible = True
                'AtAppendFC.Checked = True
                'AtConfirm.Checked = True
                'AtFinal.Checked = True
                'Test-End
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSalesMan_Click)
    '**     業務員選擇
    '**
    '*****************************************************************
    Protected Sub BSalesMan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSalesMan.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg, mUserID As String
        '
        sql = "Select * From M_FControlRecord "
        sql = sql & " Where Name = '" & DSalesMan.SelectedValue & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            DLogID.Text = Now.ToString("yyyyMMddHHmmss")
            '
            DBuyer.Text = dt_ControlRecord.Rows(0).Item("Buyer")
            DGRBuyer.Text = dt_ControlRecord.Rows(0).Item("BuyerGroup")
            DFunList.Text = dt_ControlRecord.Rows(0).Item("FunList")
            DUserID.Text = xUserID
            mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
            '
            If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
                ' FCT資料作業者 
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Start", 0, "Start", 0) = 0 Then
                    SetTaskButton("BSalesMan", False)
                    SetTaskButton("BReset", True)
                    SetTaskStatus("StsSalesMan", 127, 0)
                    SetSalesMan("SalesMan", DSalesMan.SelectedValue)
                    ' --------------------------------------------------------------------------
                    ' FunctionList(1)   判斷是否跳躍功能 
                    ' --------------------------------------------------------------------------
                    ' 準備FCT資料
                    If AtAppendFC.Checked = False And AtConfirm.Checked = False And AtFinal.Checked = False Then
                        SetTaskButton("BUpdateFC", True)
                    End If
                    ' 
                    ' 新FC INF.
                    If AtAppendFC.Checked = True Then
                        SetTaskButton("BAppendFC", True)
                    End If
                    '
                    ' 確認FC INF.
                    If AtConfirm.Checked = True Then
                        SetTaskButton("BConfirm", True)
                    End If
                    '
                    ' 移送管理人
                    If AtFinal.Checked = True Then
                        SetTaskButton("BFinal", True)
                    End If
                    ' --------------------------------------------------------------------------
                    ' 設定 Data or 報表程式
                    ' --------------------------------------------------------------------------
                    '
                    ' 更新FC INF.
                    If AtOffice2010.Checked = True Then
                        DFileName.Text = DSalesMan.SelectedValue + "_UpdateFC.xlsm"
                    Else
                        DFileName.Text = DSalesMan.SelectedValue + "_UpdateFC.xls"
                    End If
                    DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                    '
                    ' 新FC INF.
                    If AtOffice2010.Checked = True Then
                        DFileName1.Text = DSalesMan.SelectedValue + "_AppendFC.xlsm"
                    Else
                        DFileName1.Text = DSalesMan.SelectedValue + "_AppendFC.xls"
                    End If
                    DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName1.Text
                    '
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消[" + DSalesMan.SelectedValue + "]，" + "並且重新選擇業務員？" + "');if(!ok){return false;}"
                Else
                    errcode = 9003
                End If
            Else
                If mUserID = xUserID Then
                    errcode = 9004      '同上次使用者(上次不小心)
                    SetTaskButton("BReset", True)
                    SetSalesMan("SalesMan", DSalesMan.SelectedValue)
                    BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消[" + DSalesMan.SelectedValue + "]，" + "並且重新選擇業務員？" + "');if(!ok){return false;}"
                Else
                    errcode = 9002      '不同上次使用者
                End If
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此業務員不存在，請確認!"
            If errcode = 9002 Then msg = "此業務員正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9004 Then msg = "懷疑您上次使用有不正常關閉系統，造成此業務員正在使用中! 請[Reset]一次再使用!"
            If errcode = 9005 Then msg = "您目前無使用此功能權限，如有疑問請連絡系統人員!"
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
        If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            ' Clear 各動作 按鈕/結果/跑馬燈等圖像
            ClearKeyData()
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetSalesMan("Default", "")             ' 設定可使用客戶
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
    '**(BUpdateFC_Click)
    '**     更新客戶資料
    '**
    '*****************************************************************
    Protected Sub BUpdateFC_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BUpdateFC.Click
        '
        '*****************************************************************
        '**  更新客戶資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "DataPrepare", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "DataPrepare", 2) = 0 Then
            SetTaskProgress("ProUpdateFC", -500)
            SetTaskButton("BUpdateFC", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BAppendFC", True)
            SetTaskStatus("StsUpdateFC", 165, 0)
        Else
            SetTaskStatus("StsUpdateFC", 165, 1)
            uJavaScript.PopMsg(Me, "客戶資料更新未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BAppendFC_Click)
    '**     準備客戶資料
    '**
    '*****************************************************************
    Protected Sub BAppendFC_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BAppendFC.Click
        '
        '*****************************************************************
        '**  準備客戶資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataPrepare", 2, "DataConvert", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 2) = 0 Then
            SetTaskProgress("ProAppendFC", -500)
            SetTaskButton("BAppendFC", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BConfirm", True)
            SetTaskStatus("StsAppendFC", 165, 0)
        Else
            SetTaskStatus("StsAppendFC", 165, 1)
            uJavaScript.PopMsg(Me, "客戶資料準備未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BConfirm_Click)
    '**     確認FC INF.
    '**
    '*****************************************************************
    Protected Sub BConfirm_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BConfirm.Click
        Dim Cmd As String
        '
        '*****************************************************************
        '**  確認FC INF.
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataConvert", 2, "FCTPlan", 1)
        '
        Cmd = "<script>" + _
                    "window.open('VendorFCConfirm.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','VendorFCConfirm','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 2) = 0 Then
            SetTaskProgress("ProConfirm", -500)
            SetTaskButton("BConfirm", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BFinal", True)
            SetTaskStatus("StsConfirm", 165, 0)
        Else
            SetTaskStatus("StsConfirm", 165, 1)
            uJavaScript.PopMsg(Me, "客戶資料確認未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFinal_Click)
    '**     移送管理人
    '**
    '*****************************************************************
    Protected Sub BFinal_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFinal.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  移送管理人
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=VF-FINAL)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "VF-FINAL") <> 0 Then
                    errcode = 9001
                End If
            End If
            '' 判斷是否可執行
            'If errcode = 0 Then
            '    If uFASCommon.CanRunVendorFCFinal(DBuyer.Text, DGRBuyer.Text) <> 0 Then
            '        errcode = 9005
            '    End If
            'End If
            '' 移送管理人
            'If errcode = 0 Then
            '    If uFASCommon.VendorFCFinal(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
            '        errcode = 9002
            '    End If
            'End If
            ' 更新客戶控制檔(LSPlan=1)
            If errcode = 0 Then
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "LSPlan", 2) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' 移送管理人是否完成
            If errcode = 0 Then
                If uFASCommon.CheckControlRecord(DBuyer.Text, "LSPlan", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If uFASCommon.CheckControlRecord(DBuyer.Text, "LSPlan", 2) = 0 Then
                SetTaskProgress("ProFinal", -500)
                SetTaskButton("BFinal", False)
                SetTaskButton("BReset", True)
                SetTaskStatus("StsFinal", 1245, 0)

                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "End", 0, "End", 0) = 0 Then
                Else
                    uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
                End If
            Else
                SetTaskStatus("StsFinal", 1245, 1)
                uJavaScript.PopMsg(Me, "移送管理人未完成，請確認!")
            End If
        End If
        '
    End Sub

End Class
