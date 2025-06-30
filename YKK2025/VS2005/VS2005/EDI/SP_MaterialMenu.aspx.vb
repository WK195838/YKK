Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO
Imports System.Drawing

Partial Class SP_MaterialMenu
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
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
            SetAtFunction("Default", True)              ' 設定跳躍機能
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
        '隱藏
        DLogID.Style("left") = -500 & "px"
        DCode.Style("left") = -500 & "px"
        DCustomerGr.Style("left") = -500 & "px"
        DFunList.Style("left") = -500 & "px"
        DUserID.Style("left") = -500 & "px"

        DFileName.Style("left") = -500 & "px"
        DPathImport.Style("left") = -500 & "px"
        DPathActionPlan.Style("left") = -500 & "px"
        DPathActionPlanImport.Style("left") = -500 & "px"
        DPathActionPlanYOC.Style("left") = -500 & "px"
        DPathKPIFSheet.Style("left") = -500 & "px"
        DPathWINGS.Style("left") = -500 & "px"
        DPathPIL.Style("left") = -500 & "px"
        DPathChangeFinal.Style("left") = -500 & "px"
        DPathPILFinal.Style("left") = -500 & "px"
        LUpdateWINGS.Style("left") = -500 & "px"
        LPendingFinal.Style("left") = -500 & "px"
        GridView2.Style("left") = 210 & "px"
        GridView3.Style("left") = 210 & "px"


        '動作按鈕設定
        BCustomer.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此客戶？" + "');if(!ok){return false;}"
        BImport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行匯入客戶資料？" + "');if(!ok){return false;} else {CheckAttribute('Import')}"
        BDemand.Attributes("onclick") = "var ok=window.confirm('" + "是否執行需求量計算？　注意:前回SHOPPING DATA會刪除！" + "');if(!ok){return false;} else {CheckAttribute('Demand')}"
        BActionPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否取得需求量報表？" + "');if(!ok){return false;} else {CheckAttribute('ActionPlan')}"
        BActionPlanImport.Attributes("onclick") = "var ok=window.confirm('" + "是否上傳Action Plan？" + "');if(!ok){return false;} else {CheckAttribute('ActionPlanImport')}"
        BActionPlanYOC.Attributes("onclick") = "var ok=window.confirm('" + "是否取得[YOC]回覆報報表？" + "');if(!ok){return false;} else {CheckAttribute('ActionPlanYOC')}"
        BKPIFSheet.Attributes("onclick") = "var ok=window.confirm('" + "是否取得[SSO-KP]上傳資料？" + "');if(!ok){return false;} else {CheckAttribute('KPIFSheet')}"
        BWINGS.Attributes("onclick") = "var ok=window.confirm('" + "是否回報[WINGS]資料？" + "');if(!ok){return false;} else {CheckAttribute('WINGS')}"
        BPIL.Attributes("onclick") = "var ok=window.confirm('" + "是否取得 Pachase Information List？" + "');if(!ok){return false;} else {CheckAttribute('PIL')}"
        BFinal.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 Final Proc.？" + "');if(!ok){return false;} else {CheckAttribute('Final')}"
        BChangeFinal.Attributes("onclick") = "var ok=window.confirm('" + "是否更新 Final Data？" + "');if(!ok){return false;} else {CheckAttribute('ChangeFinal')}"
        '
        BProgress.Attributes("onclick") = "var ok=window.confirm('" + "是否結束？" + "');if(!ok){return false;} else {window.close();}"
        BToolBox.Attributes.Add("onclick", "window.open('" + "file://10.245.0.186/Share/SystemSupport/[ISOS]ToolBox/Shopping" + "','_blank');return false;")

        GridView1.Visible = False
        GridView2.Visible = False
        GridView3.Visible = False

        'Action CheckBox設定
        'AtDemand.Attributes.Add("onclick", "ActionCheck('AtDemand')")
        'AtActionPlan.Attributes.Add("onclick", "ActionCheck('AtActionPlan')")
        'AtActionPlanImport.Attributes.Add("onclick", "ActionCheck('AtActionPlanImport')")
        'AtActionPlanYOC.Attributes.Add("onclick", "ActionCheck('AtActionPlanYOC')")
        'AtKPIFSheet.Attributes.Add("onclick", "ActionCheck('AtKPIFSheet')")
        'AtWINGS.Attributes.Add("onclick", "ActionCheck('AtWINGS')")
        'AtPIL.Attributes.Add("onclick", "ActionCheck('AtPIL')")
        'AtFinal.Attributes.Add("onclick", "ActionCheck('AtFinal')")
        'AtChangeFinal.Attributes.Add("onclick", "ActionCheck('AtChangeFinal')")
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
                sql = sql & " Where Cat  = '" & "001" & "' "
                sql = sql & "   And DKey = '" & xUserID & "-SP" & "' "
                sql = sql & "   And RIGHT(DKEY,2) = 'SP' "
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
            Case "StsImport"
                StsImport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsImport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsImport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsDemand"
                StsDemand.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsDemand.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsDemand.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsActionPlan"
                StsActionPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsActionPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsActionPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsActionPlanImport"
                StsActionPlanImport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsActionPlanImport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsActionPlanImport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsActionPlanYOC"
                StsActionPlanYOC.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsActionPlanYOC.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsActionPlanYOC.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsKPIFSheet"
                StsKPIFSheet.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsKPIFSheet.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsKPIFSheet.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsWINGS"
                StsWINGS.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsWINGS.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsWINGS.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsPIL"
                StsPIL.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsPIL.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsPIL.ImageUrl = "iMages/NG.jpg"
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
            Case "StsChangeFinal"
                StsChangeFinal.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsChangeFinal.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsChangeFinal.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsCustomer.Style("left") = pLeft & "px"
                StsImport.Style("left") = pLeft & "px"
                StsDemand.Style("left") = pLeft & "px"
                StsActionPlan.Style("left") = pLeft & "px"
                StsActionPlanImport.Style("left") = pLeft & "px"
                StsActionPlanYOC.Style("left") = pLeft & "px"
                StsKPIFSheet.Style("left") = pLeft & "px"
                StsWINGS.Style("left") = pLeft & "px"
                StsPIL.Style("left") = pLeft & "px"
                StsFinal.Style("left") = pLeft & "px"
                StsChangeFinal.Style("left") = pLeft & "px"
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
            Case "ProImport"
                ProImport.Style("left") = pLeft & "px"
            Case "ProDemand"
                ProDemand.Style("left") = pLeft & "px"
            Case "ProActionPlan"
                ProActionPlan.Style("left") = pLeft & "px"
            Case "ProActionPlanImport"
                ProActionPlanImport.Style("left") = pLeft & "px"
            Case "ProActionPlanYOC"
                ProActionPlanYOC.Style("left") = pLeft & "px"
            Case "ProKPIFSheet"
                ProKPIFSheet.Style("left") = pLeft & "px"
            Case "ProWINGS"
                ProWINGS.Style("left") = pLeft & "px"
            Case "ProPIL"
                ProPIL.Style("left") = pLeft & "px"
            Case "ProFinal"
                ProFinal.Style("left") = pLeft & "px"
            Case "ProChangeFinal"
                ProChangeFinal.Style("left") = pLeft & "px"
            Case Else
                ProImport.Style("left") = pLeft & "px"
                ProDemand.Style("left") = pLeft & "px"
                ProActionPlan.Style("left") = pLeft & "px"
                ProActionPlanImport.Style("left") = pLeft & "px"
                ProActionPlanYOC.Style("left") = pLeft & "px"
                ProKPIFSheet.Style("left") = pLeft & "px"
                ProWINGS.Style("left") = pLeft & "px"
                ProPIL.Style("left") = pLeft & "px"
                ProFinal.Style("left") = pLeft & "px"
                ProChangeFinal.Style("left") = pLeft & "px"
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
            Case "BImport"
                BImport.Visible = pShow
            Case "BDemand"
                BDemand.Visible = pShow
            Case "BActionPlan"
                BActionPlan.Visible = pShow
            Case "BActionPlanImport"
                BActionPlanImport.Visible = pShow
            Case "BActionPlanYOC"
                BActionPlanYOC.Visible = pShow
            Case "BKPIFSheet"
                BKPIFSheet.Visible = pShow
            Case "BWINGS"
                BWINGS.Visible = pShow
            Case "BPIL"
                BPIL.Visible = pShow
            Case "BFinal"
                BFinal.Visible = pShow
            Case "BChangeFinal"
                BChangeFinal.Visible = pShow
            Case "BProgress"
                BProgress.Visible = pShow
            Case Else
                BReset.Visible = False
                BCustomer.Visible = True
                BImport.Visible = False
                BDemand.Visible = False
                BActionPlan.Visible = False
                BActionPlanImport.Visible = False
                BActionPlanYOC.Visible = False
                BKPIFSheet.Visible = False
                BWINGS.Visible = False
                BPIL.Visible = False
                BFinal.Visible = False
                BChangeFinal.Visible = False
                BProgress.Visible = False
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetAtFunction)
    '**     設定跳躍機能
    '**
    '*****************************************************************
    Sub SetAtFunction(ByVal pAction As String, ByVal pShow As Boolean)
        '
        AtDemand.Checked = False
        AtActionPlan.Checked = False
        AtActionPlanImport.Checked = False
        AtActionPlanYOC.Checked = False
        AtKPIFSheet.Checked = False
        AtWINGS.Checked = False
        AtPIL.Checked = False
        AtFinal.Checked = False
        AtChangeFinal.Checked = False
        AtPending.Checked = False
        AtPending.Enabled = False
        '
        Select Case pAction
            Case "AtDemand"
                AtDemand.Checked = pShow
                AtDemand.Enabled = pShow
            Case "AtActionPlan"
                AtActionPlan.Checked = pShow
                AtActionPlan.Enabled = pShow
            Case "AtActionPlanImport"
                AtActionPlanImport.Checked = pShow
                AtActionPlanImport.Enabled = pShow
            Case "AtActionPlanYOC"
                AtActionPlanYOC.Checked = pShow
                AtActionPlanYOC.Enabled = pShow
            Case "AtKPIFSheet"
                AtKPIFSheet.Checked = pShow
                AtKPIFSheet.Enabled = pShow
            Case "AtWINGS"
                AtWINGS.Checked = pShow
                AtWINGS.Enabled = pShow
            Case "AtPIL"
                AtPIL.Checked = pShow
                AtPIL.Enabled = pShow
            Case "AtFinal"
                AtFinal.Checked = pShow
                AtFinal.Enabled = pShow
            Case "AtChangeFinal"
                AtChangeFinal.Checked = pShow
                AtChangeFinal.Enabled = pShow
            Case "AtPending"
                'AtPending.Checked = pShow
                AtPending.Enabled = pShow
            Case "Customer"
                AtDemand.Enabled = False
                AtActionPlan.Enabled = False
                AtActionPlanImport.Enabled = False
                AtActionPlanYOC.Enabled = False
                AtKPIFSheet.Enabled = False
                AtWINGS.Enabled = False
                AtPIL.Enabled = False
                AtFinal.Enabled = False
                AtChangeFinal.Enabled = False
            Case Else
                AtDemand.Enabled = pShow
                AtActionPlan.Enabled = pShow
                'left: 232px
                AtActionPlan.Style("left") = -500 & "px"
                AtActionPlanImport.Enabled = pShow
                AtActionPlanYOC.Enabled = pShow
                AtActionPlanYOC.Style("left") = -500 & "px"
                AtKPIFSheet.Enabled = pShow
                AtWINGS.Enabled = pShow
                AtPIL.Enabled = pShow
                AtPIL.Style("left") = -500 & "px"
                AtFinal.Enabled = pShow
                AtChangeFinal.Enabled = pShow
        End Select
        '
    End Sub
    '
    Protected Sub AtDemand_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtDemand.CheckedChanged
        SetAtFunction("AtDemand", AtDemand.Checked)
        AtDemand.Enabled = True
    End Sub
    Protected Sub AtActionPlan_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtActionPlan.CheckedChanged
        SetAtFunction("AtActionPlan", AtActionPlan.Checked)
        AtActionPlan.Enabled = True
    End Sub
    Protected Sub AtActionPlanImport_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtActionPlanImport.CheckedChanged
        SetAtFunction("AtActionPlanImport", AtActionPlanImport.Checked)
        AtActionPlanImport.Enabled = True
    End Sub
    Protected Sub AtActionPlanYOC_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtActionPlanYOC.CheckedChanged
        SetAtFunction("AtActionPlanYOC", AtActionPlanYOC.Checked)
        AtActionPlanYOC.Enabled = True
    End Sub
    Protected Sub AtKPIFSheet_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtKPIFSheet.CheckedChanged
        SetAtFunction("AtKPIFSheet", AtKPIFSheet.Checked)
        AtKPIFSheet.Enabled = True
    End Sub
    Protected Sub AtWINGS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtWINGS.CheckedChanged
        SetAtFunction("AtWINGS", AtWINGS.Checked)
        AtWINGS.Enabled = True
    End Sub
    Protected Sub AtPIL_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtPIL.CheckedChanged
        SetAtFunction("AtPIL", AtPIL.Checked)
        AtPIL.Enabled = True
    End Sub
    Protected Sub AtFinal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtFinal.CheckedChanged
        SetAtFunction("AtFinal", AtFinal.Checked)
        AtFinal.Enabled = True
    End Sub
    Protected Sub AtChangeFinal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtChangeFinal.CheckedChanged
        SetAtFunction("AtChangeFinal", AtChangeFinal.Checked)
        AtChangeFinal.Enabled = True
    End Sub
    '
    '=========================================================================================================================
    '==PROC 
    '==
    '*****************************************************************
    '**(BReset_Click)
    '**     客戶重新選擇
    '**
    '*****************************************************************
    Protected Sub BReset_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReset.Click
        '
        If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            ' Clear 各動作 按鈕/結果/跑馬燈等圖像
            DLogID.Text = ""
            DCode.Text = ""
            DCustomerGr.Text = ""
            DFunList.Text = ""
            GridView1.Visible = False
            GridView2.Visible = False
            GridView3.Visible = False

            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
            SetAtFunction("Default", True)              ' 設定跳躍機能
        Else
            uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
        End If
        '
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
        sql = "Select * From M_SPControlRecord "
        sql = sql & " Where Name = '" & DCustomerBuyer.SelectedValue & "'"
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            '
            DLogID.Text = Now.ToString("yyyyMMddHHmmss")
            '
            DCode.Text = dt_ControlRecord.Rows(0).Item("Code")
            DCustomerGr.Text = dt_ControlRecord.Rows(0).Item("CustomerGr")
            DFunList.Text = dt_ControlRecord.Rows(0).Item("FunList")
            DUserID.Text = xUserID
            mUserID = UCase(dt_ControlRecord.Rows(0).Item("ModifyUser"))
            '
            If dt_ControlRecord.Rows(0).Item("Yobi1") = 1 Then
                '
                '跳躍機能=沒選擇
                If AtDemand.Checked = False And AtActionPlan.Checked = False And _
                   AtActionPlanImport.Checked = False And AtActionPlanYOC.Checked = False And _
                   AtKPIFSheet.Checked = False And AtWINGS.Checked = False And _
                   AtPIL.Checked = False And AtFinal.Checked = False And AtChangeFinal.Checked = False Then
                    SetAtFunction("Default", True)
                    errcode = 9005
                End If
                '
                '跳躍機能=選擇KP前
                If errcode = 0 Then
                    If AtDemand.Checked = True Or AtActionPlan.Checked = True Or _
                       AtActionPlanImport.Checked = True Or AtActionPlanYOC.Checked = True Or AtKPIFSheet.Checked = True Then
                        SetAtFunction("Default", True)
                        errcode = 9005
                    End If
                End If
            Else
                '
                '跳躍機能=選擇KP後
                If errcode = 0 Then
                    If AtWINGS.Checked = True Or AtPIL.Checked = True Or AtFinal.Checked = True Then
                        SetAtFunction("Default", True)
                        errcode = 9006
                    End If
                End If
            End If
            '
            If errcode = 0 Then
                '
                If dt_ControlRecord.Rows(0).Item("Active") = 0 Then
                    ' FCT資料作業者 
                    If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Start", 0, "Start", 0) = 0 Then

                        SetTaskButton("BCustomer", False)
                        SetTaskButton("BReset", True)
                        SetTaskStatus("StsCustomer", 176, 0)
                        SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
                        ' --------------------------------------------------------------------------
                        ' 跳躍功能 
                        If AtDemand.Checked = False And AtActionPlan.Checked = False And _
                           AtActionPlanImport.Checked = False And AtActionPlanYOC.Checked = False And _
                           AtKPIFSheet.Checked = False And AtWINGS.Checked = False And _
                           AtPIL.Checked = False And AtFinal.Checked = False And AtChangeFinal.Checked = False Then
                            SetTaskButton("BImport", True)
                        End If

                        If AtDemand.Checked = True Then
                            SetTaskButton("BDemand", True)
                        End If

                        If AtActionPlan.Checked = True Then
                            SetTaskButton("BActionPlan", True)
                        End If

                        If AtActionPlanImport.Checked = True Then
                            SetTaskButton("BActionPlanImport", True)
                        End If

                        If AtActionPlanYOC.Checked = True Then
                            SetTaskButton("BActionPlanYOC", True)
                        End If

                        If AtKPIFSheet.Checked = True Then
                            SetTaskButton("BKPIFSheet", True)
                        End If

                        If AtWINGS.Checked = True Then
                            SetTaskButton("BWINGS", True)
                        End If

                        If AtPIL.Checked = True Then
                            SetTaskButton("BPIL", True)
                        End If

                        If AtFinal.Checked = True Then
                            SetTaskButton("BFinal", True)
                            SetAtFunction("AtPending", True)
                        End If

                        If AtChangeFinal.Checked = True Then
                            SetTaskButton("BChangeFinal", True)
                        End If
                        ' --------------------------------------------------------------------------
                        ' FunctionList(1)   判斷是否跳躍功能 
                        ' FunctionList(2)   判斷是否有BUYER FCT功能 
                        ' FunctionList(3)   判斷是否有ITEM CHECK功能 
                        ' FunctionList(4)   判斷是否有管理分析功能 / 限定FC-VENDOR
                        ' --------------------------------------------------------------------------
                        ' 設定 Data or 報表程式
                        ' --------------------------------------------------------------------------
                        '
                        'SP_YOC-BUYER-SPTYPE_Import.xlsm --> SP_SHANGHAI-ADIDAS-BY_Import.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_Import.xlsm"
                        DPathImport.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathImport.Text) Then
                            If InStr(DCustomerBuyer.SelectedValue, "YOC-Y") > 0 Then
                                File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_ImportTrade.xlsm", DPathImport.Text)
                            Else
                                File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_Import.xlsm", DPathImport.Text)
                            End If
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_ActionPlan.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_ActionPlan.xlsm"
                        DPathActionPlan.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathActionPlan.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_ActionPlan.xlsm", DPathActionPlan.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_ActionPlanImport.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_ActionPlanImport.xlsm"
                        DPathActionPlanImport.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathActionPlanImport.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_ActionPlanImport.xlsm", DPathActionPlanImport.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_ActionPlanYOC.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_ActionPlanYOC.xlsm"
                        DPathActionPlanYOC.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathActionPlanYOC.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_ActionPlanYOC.xlsm", DPathActionPlanYOC.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_KPIFSheet.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_KPIFSheet.xlsm"
                        DPathKPIFSheet.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathKPIFSheet.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_KPIFSheet.xlsm", DPathKPIFSheet.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_WINGS.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_WINGS.xlsm"
                        DPathWINGS.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathWINGS.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_WINGS.xlsm", DPathWINGS.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_PIL.xlsm
                        DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_PIL.xlsm"
                        DPathPIL.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        If Not File.Exists(DPathPIL.Text) Then
                            File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_PIL.xlsm", DPathPIL.Text)
                        End If
                        '
                        'SP_YOC-BUYER-SPTYPE_ChangeFinal.xlsm
                        'DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_ChangeFinal.xlsm"
                        'DPathChangeFinal.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        'If Not File.Exists(DPathChangeFinal.Text) Then
                        '    File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_ChangeFinal.xlsm", DPathChangeFinal.Text)
                        'End If
                        '
                        'SP_YOC-BUYER-SPTYPE_PILFinal.xlsm
                        'DFileName.Text = "SP_" + DCustomerBuyer.SelectedValue + "_PILFinal.xlsm"
                        'DPathPILFinal.Text = uCommon.GetAppSetting("SPDataPrepareFile") + DFileName.Text
                        'If Not File.Exists(DPathPILFinal.Text) Then
                        '    File.Copy(uCommon.GetAppSetting("SPDataPrepareFile") + "SP_YOC-BUYER-SPTYPE_PILFinal.xlsm", DPathPILFinal.Text)
                        'End If
                        '
                        BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
                        '
                        'Pending Final Error > 31
                        sql = "SELECT "
                        sql = sql + "CUSTOMER as SP_Code"
                        sql = sql + ",(SELECT TOP 1 [Name] FROM M_SPControlRecord WHERE CODE=PD_SPActionPlan.CUSTOMER) as SP_Name "
                        sql = sql + ",SPNo as SP_No "
                        sql = sql + ",convert(varchar, convert(datetime, SUBSTRING(SPNo, 3,4) + '/' + SUBSTRING(SPNo, 7,2) +'/'+ SUBSTRING(SPNo, 9,2), 111), 111) as SP_Date "
                        sql = sql + ",convert(varchar, GETDATE(), 111) as Now_Date "
                        '
                        sql = sql + ",ltrim(rtrim( "
                        sql = sql + "case when datediff(day, getdate(), convert(datetime, SUBSTRING(SPNo, 3,4) + '/' + SUBSTRING(SPNo, 7,2) +'/'+ SUBSTRING(SPNo, 9,2), 111))>=0 then 0 "
                        sql = sql + "      else datediff(day, getdate(), convert(datetime, SUBSTRING(SPNo, 3,4) + '/' + SUBSTRING(SPNo, 7,2) +'/'+ SUBSTRING(SPNo, 9,2), 111)) * -1 "
                        sql = sql + "end "
                        sql = sql + ")) + ' Days.' as Diff_Date "
                        '
                        sql = sql + ",'.....' as PDL "
                        sql = sql + ",'http://10.245.0.205/EDI/SP_ShoppingListInf.aspx?pUserID=" & xUserID & "&pSPName=" & DCustomerBuyer.SelectedValue & "&pSPNo='+SPNo as PDLUrl "
                        '
                        sql = sql + "FROM PD_SPActionPlan "
                        sql = sql + "Where CUSTOMER = '" & DCode.Text & "' "
                        sql = sql + "and [Version]=99 "
                        sql = sql + "and datediff(day, getdate(), convert(datetime, SUBSTRING(SPNo, 3,4) + '/' + SUBSTRING(SPNo, 7,2) +'/'+ SUBSTRING(SPNo, 9,2), 111)) < -31 "
                        sql = sql + "group by CUSTOMER, SPNo "
                        sql = sql + "order by SPNo "
                        Dim dt As DataTable = uDataBase.GetDataTable(sql)
                        If dt.Rows.Count > 0 Then
                            GridView2.Visible = True
                            GridView2.DataSource = dt
                            GridView2.DataBind()
                        End If
                        '
                        'Wings Order Error > 7
                        sql = "SELECT "
                        sql = sql + "SPCode, SPName, SPTime, SPNo, F_ORNo, F_COrderNo, F_Time, F_User, F_ItemInf, DiffDaysDesc "
                        '
                        sql = sql + ",'.....' as ChgFinal "
                        sql = sql + ",'http://10.245.0.205/EDI/SP_ShoppingListInf.aspx?pUserID=" & xUserID & "&pSPName=" & DCustomerBuyer.SelectedValue & "&pSPNo='+SPNo as ChgFinalUrl "
                        '
                        sql = sql + "FROM V_SPWINGSOrderFail "
                        sql = sql + "Where SPCode = '" & DCode.Text & "' "
                        sql = sql + "and DiffDays > 7 "
                        sql = sql + "order by SPNo "
                        '
                        Dim dtWings As DataTable = uDataBase.GetDataTable(sql)
                        If dtWings.Rows.Count > 0 Then
                            If GridView2.Visible = True Then
                                GridView3.Style("left") = 750 & "px"
                            End If
                            '
                            GridView3.Visible = True
                            GridView3.DataSource = dtWings
                            GridView3.DataBind()
                        End If
                    Else
                        errcode = 9003
                    End If
                Else
                    If mUserID = xUserID Then

                        errcode = 9004      '同上次使用者(上次不小心)
                        SetTaskButton("BReset", True)
                        SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
                        BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
                        '
                    Else
                        errcode = 9002      '不同上次使用者
                    End If
                    '
                    'access log
                    sql = "SELECT Top 100 "
                    sql = sql + "[User],convert(varchar, AccessTime, 120) as AccessTime,Cat,Active, Code, Name, "
                    sql = sql + "Customer,Import,Demand,ActPlan,ImpActPlan,RspActPlan,KPInterface,RspWINGS,PILSheet,Final,ChgFinal,Progress, Status "
                    sql = sql + "FROM V_SPControlRecordLog "
                    sql = sql + "Where Code = '" & DCode.Text & "' "
                    sql = sql + "Order by User, AccessTime desc, Cat "
                    Dim dt As DataTable = uDataBase.GetDataTable(sql)
                    If dt.Rows.Count > 0 Then
                        GridView1.Visible = True
                        GridView1.DataSource = dt
                        GridView1.DataBind()
                    Else
                        sql = "SELECT "
                        sql = sql + "null as [user],null as AccessTime,null as [Cat],null as [Active], null as [Code], null as [Name], "
                        sql = sql + "null as [Customer],null as [Import],null as [Demand],null as [ActPlan],null as [ImpActPlan],null as [RspActPlan],null as [KPInterface], "
                        sql = sql + "null as [RspWINGS],null as [PILSheet],null as [Final],null as [ChgFinal],null as [Progress], 'No Data' as [Status] "
                        Dim dt1 As DataTable = uDataBase.GetDataTable(sql)
                        If dt1.Rows.Count > 0 Then
                            GridView1.Visible = True
                            GridView1.DataSource = dt1
                            GridView1.DataBind()
                        End If
                    End If
                    '
                End If
                '
                'Pending Final Proc
                LPendingFinal.Style("left") = 344 & "px"
                sql = "SELECT Customer, SPNo "
                sql = sql + "FROM PD_SPActionPlan "
                sql = sql + "Where Customer = '" & DCode.Text & "' "
                sql = sql + "  And Version = 99 "
                sql = sql + "Group by Customer, SPNo "
                sql = sql + "Order by Customer, SPNo "
                Dim dt_PDSPActionPlan As DataTable = uDataBase.GetDataTable(sql)
                If dt_PDSPActionPlan.Rows.Count > 0 Then
                    LPendingFinal.Text = "Pending Final Count=(" & dt_PDSPActionPlan.Rows.Count.ToString & ")"
                Else
                    LPendingFinal.Text = "Pending Final Count=(0)"
                End If
                LPendingFinal.NavigateUrl = "http://10.245.0.205/EDI/SP_ShoppingListInf.aspx?pUserID=" & xUserID & "&pSPName=" & DCustomerBuyer.SelectedValue & ""
                '
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "此客戶不存在，請確認!"
            If errcode = 9002 Then msg = "此客戶正由 [" + uWFSCommon.GetUserName(mUserID) + "] 使用中，請確認!"
            If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9004 Then msg = "此客戶正在使用中, 您可能上次使用至途中，請[Reset]一次再使用!"
            If errcode = 9005 Then msg = "懷疑您上次發行的[Shopping]未完成[FINNAL]，請繼續執行至[FINNAL] !"
            If errcode = 9006 Then msg = "懷疑您上次未完成[KP I/F]作業，請繼續執行至[KP I/F] !"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BImport_Click)
    ''**     Import客戶資料
    ''**
    ''*****************************************************************
    Protected Sub BImport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BImport.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "Import", 1)
        '
        '分離處理
        If InStr(DCode.Text, "YOC-Y") <= 0 Then
            '
            'BUYER
            '判斷是否完成作業
            Dim Code As Integer
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 3)
            End If
            Do Until Code = 0
                System.Threading.Thread.Sleep(3 * 1000)
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2)
                If Code <> 0 Then
                    Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 3)
                End If
            Loop
            '
            '檢查是否已將資料上傳完成
            If uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2) = 0 Then
                SetTaskProgress("ProImport", -500)
                SetTaskButton("BImport", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BDemand", True)
                '
                SetTaskStatus("StsImport", 176, 0)
            Else
                SetTaskStatus("StsImport", 176, 1)
                uJavaScript.PopMsg(Me, "未完成，請確認!")
            End If
            '
        Else
            '
            'YOC
            '判斷是否完成作業
            Dim Code As Integer
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 3)
            End If
            Do Until Code = 0
                System.Threading.Thread.Sleep(3 * 1000)
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2)
                If Code <> 0 Then
                    Code = uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 3)
                End If
            Loop
            '
            '檢查是否已將資料上傳完成 / JOYJOY
            If uFASCommon.SPCheckControlRecord(DCode.Text, "Import", 2) = 0 Then
                SetTaskProgress("ProImport", -500)
                '
                SetTaskButton("BImport", False)
                '
                SetTaskStatus("StsImport", 176, 0)
                SetTaskStatus("StsDemand", 176, 0)
                SetTaskStatus("StsActionPlan", 176, 0)
                SetTaskStatus("StsActionPlanImport", 632, 0)
                SetTaskStatus("StsActionPlanYOC", 632, 0)
                SetTaskStatus("StsKPIFSheet", 632, 0)
                SetTaskStatus("StsWINGS", 1136, 0)
                SetTaskStatus("StsPIL", 1248, 0)
                SetTaskStatus("StsFinal", 1064, 0)
                SetTaskButton("BProgress", True)
            Else
                SetTaskStatus("StsImport", 176, 1)
                '
                'CHECK DEMAND
                If uFASCommon.SPCheckControlRecord(DCode.Text, "Demand", 3) = 0 Then
                    SetTaskStatus("StsDemand", 176, 1)
                    SetTaskStatus("StsActionPlan", 176, 1)
                    SetTaskStatus("StsActionPlanImport", 632, 1)
                    SetTaskStatus("StsActionPlanYOC", 632, 1)
                End If
                '
                'CHECK KPInterface
                If uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 3) = 0 Then
                    SetTaskStatus("StsKPIFSheet", 632, 1)
                    SetTaskStatus("StsWINGS", 1136, 1)
                    SetTaskStatus("StsPIL", 1248, 1)
                    SetTaskStatus("StsFinal", 1064, 1)
                End If
                '
                uJavaScript.PopMsg(Me, "未完成，請確認!")
            End If
            '
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BDemand_Click)
    ''**     需求量計算
    ''**
    ''*****************************************************************
    Protected Sub BDemand_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDemand.Click
        Dim msg As String = ""
        Dim msgcontent As String = ""
        Dim sql As String = ""
        Dim xBuyerConvert As String = ""
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        '*****************************************************************
        '**  需求量計算
        '*****************************************************************
        '
        ' SP-NO置入IMPORT FILE(E_InputSheet) SPNO-->BK1
        ' BACKUP前次 PLAN DATA
        If errcode = 0 Then
            If uFASCommon.SPNo2Import(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text, 0) <> 0 Then
                errcode = 9011
            End If
        End If
        ' ---------------
        ' 刪除執行履歷/FCT Plan/LS Plan
        If errcode = 0 Then
            If uFASCommon.DeleteHistory(DCode.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        ' 刪除FCT Plan上回資料
        If errcode = 0 Then
            If uFASCommon.DeleteFCTData(DCode.Text) <> 0 Then
                errcode = 9002
            End If
        End If
        '
        ' 刪除LS Plan上回資料
        If errcode = 0 Then
            If uFASCommon.DeleteLSData(DCode.Text) <> 0 Then
                errcode = 9003
            End If
        End If
        '
        ' 刪除Action Plan上回資料
        If errcode = 0 Then
            If uFASCommon.DeleteActionData(DCode.Text) <> 0 Then
                errcode = 9004
            End If
        End If
        ' ---------------
        ' FC匯入作業
        If errcode = 0 Then
            If uFASMapping.Rule2Data(DLogID.Text, DCode.Text, xUserID, DFunList.Text) <> 0 Then
                errcode = 9012
            End If
        End If
        ' FCNo
        If errcode = 0 Then
            If uFASCommon.SPMakeForcastNo(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text, 0) <> 0 Then
                errcode = 9013
            End If
        End If
        '
        ' FC Plan展開
        If errcode = 0 Then
            If uFASCommon.SPForcastPlan(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text, 0) <> 0 Then
                errcode = 9021
            End If
        End If
        '
        ' LS Plan展開
        If errcode = 0 Then
            If uFASCommon.SPLocalStockPlan(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text) <> 0 Then
                errcode = 9031
            End If
        End If
        '
        ' Action Plan展開
        If errcode = 0 Then
            If uFASCommon.SPActionkPlan(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text) <> 0 Then
                errcode = 9032
            End If
        End If
        '
        ' 更新客戶控制檔
        If errcode = 0 Then
            If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Demand", 2, "ActPlan", 1) <> 0 Then
                errcode = 9041
            End If
        End If
        ' 檢查匯入作業是否完成
        If errcode = 0 Then
            If uFASCommon.SPCheckControlRecord(DCode.Text, "Demand", 2) <> 0 Then
                errcode = 9042
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsDemand", 176, 1)
            StsDemand.NavigateUrl = "SPSActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DCode.Text + "&pUserID=" + xUserID
            StsDemand.Target = "_blank"
            '
            If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
            If errcode = 9002 Then msg = "刪除 FCT Plan 資料異常，請連絡系統人員!"
            If errcode = 9003 Then msg = "刪除 LS Plan 資料異常，請連絡系統人員!"
            If errcode = 9004 Then msg = "刪除 Action Plan 資料異常，請連絡系統人員!"
            If errcode = 9011 Then msg = "轉換SP-NO轉入IMPORT異常，請連絡系統人員!"
            If errcode = 9012 Then msg = "轉換資料異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9013 Then msg = "FCT-No.異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9021 Then msg = "FCT Plan展開異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9031 Then msg = "LS Plan展開異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9032 Then msg = "Action Plan展開異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9041 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9042 Then msg = "轉換資料異常，請連絡系統人員!"
            '
            uJavaScript.PopMsg(Me, msg)
        Else
            StsDemand.NavigateUrl = ""
            StsDemand.Target = ""
            '
            SetTaskProgress("ProDemand", -500)
            SetTaskButton("BDemand", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BActionPlan", True)
            SetTaskStatus("StsDemand", 176, 0)
            '
            StsDemand.NavigateUrl = "SPSActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DCode.Text + "&pUserID=" + xUserID
            StsDemand.Target = "_blank"
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BActionPlan_Click)
    ''**     Action Report
    ''**
    ''*****************************************************************
    Protected Sub BActionPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BActionPlan.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "ActPlan", 1)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ActPlan", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ActPlan", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ActPlan", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ActPlan", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "ActPlan", 2) = 0 Then

            SetTaskProgress("ProActPlan", -500)
            SetTaskButton("BActPlan", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BActionPlanImport", True)
            SetTaskStatus("StsActionPlan", 176, 0)
        Else
            SetTaskStatus("StsActionPlan", 176, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BActionPlanImport_Click)
    ''**     Action Report Import
    ''**
    ''*****************************************************************
    Protected Sub BActionPlanImport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BActionPlanImport.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "ImpActPlan", 1)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ImpActPlan", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ImpActPlan", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ImpActPlan", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ImpActPlan", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "ImpActPlan", 2) = 0 Then
            SetTaskProgress("ProActionPlanImport", -500)
            SetTaskButton("BActionPlanImport", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BActionPlanYOC", True)
            SetTaskStatus("StsActionPlanImport", 632, 0)
        Else
            SetTaskStatus("StsActionPlanImport", 632, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BActionPlanYOC_Click)
    ''**     YOC Confirm Action Report
    ''**
    ''*****************************************************************
    Protected Sub BActionPlanYOC_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BActionPlanYOC.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "RspActPlan", 1)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspActPlan", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspActPlan", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspActPlan", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspActPlan", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "RspActPlan", 2) = 0 Then
            SetTaskProgress("ProActionPlanYOC", -500)
            SetTaskButton("BActionPlanYOC", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BKPIFSheet", True)
            SetTaskStatus("StsActionPlanYOC", 632, 0)
        Else
            SetTaskStatus("StsActionPlanYOC", 632, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BKPIFSheet_Click)
    ''**     KP I/F SHEET
    ''**
    ''*****************************************************************
    Protected Sub BKPIFSheet_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BKPIFSheet.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '*****************************************************************
        '**  KP I/F Data
        '*****************************************************************
        If errcode = 0 Then
            '
            ' 刪除執行履歷(Action=SPEDITRANSFER)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DCode.Text, "SPEDITRANSFER") <> 0 Then
                    errcode = 9001
                End If
            End If
            '
            ' 刪除EDI Data
            If errcode = 0 Then
                If uFASCommon.DeleteLS2EDIInterface(DCode.Text, DCode.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            '
            ' KP I/F Data
            If errcode = 0 Then
                If uFASCommon.SPKPInterface(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text) <> 0 Then
                    errcode = 0
                End If
            End If
            '
            ' KP I/F REPORT
            If errcode = 0 Then
                '
                uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "KPInterface", 1)
                '
                '判斷是否完成作業
                Dim Code As Integer
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 2)
                If Code <> 0 Then
                    Code = uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 3)
                End If
                Do Until Code = 0
                    System.Threading.Thread.Sleep(3 * 1000)
                    Code = uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 2)
                    If Code <> 0 Then
                        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 3)
                    End If
                Loop
                '
                '檢查是否已將資料上傳完成
                If uFASCommon.SPCheckControlRecord(DCode.Text, "KPInterface", 2) = 0 Then
                    '
                    uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "YOBI1", 1)
                    '
                    SetTaskProgress("ProKPIFSheet", -500)
                    SetTaskButton("BKPIFSheet", False)
                    SetTaskButton("BReset", True)
                    SetTaskButton("BWINGS", True)
                    SetTaskStatus("StsKPIFSheet", 632, 0)
                Else
                    SetTaskStatus("StsKPIFSheet", 632, 1)
                    uJavaScript.PopMsg(Me, "未完成，請確認!")
                End If
            Else
                SetTaskStatus("StsKPIFSheet", 632, 1)
                StsKPIFSheet.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DCode.Text + "&pUserID=" + xUserID
                StsKPIFSheet.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(KPI/F)，請連絡系統人員!"
                If errcode = 9002 Then msg = "刪除 KP I/F 資料異常，請連絡系統人員!"
                If errcode = 9003 Then msg = "KP I/F資料轉換異常，請確認(可點選[ｘ]查詢)!"
                uJavaScript.PopMsg(Me, msg)
            End If
            '
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BWINGS_Click)
    ''**     REPONSE WINGS INF.
    ''**
    ''*****************************************************************
    Protected Sub BWINGS_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BWINGS.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "RspWINGS", 1)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspWINGS", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspWINGS", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspWINGS", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "RspWINGS", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "RspWINGS", 2) = 0 Then
            SetTaskProgress("ProWINGS", -500)
            SetTaskButton("BWINGS", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BPIL", True)
            SetTaskStatus("StsWINGS", 1136, 0)
        Else
            SetTaskStatus("StsWINGS", 1136, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BPIL_Click)
    ''**     PURCHASE INFORMATION LIST
    ''**
    ''*****************************************************************
    Protected Sub BPIL_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPIL.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "PILSheet", 1)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "PILSheet", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "PILSheet", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "PILSheet", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "PILSheet", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "PILSheet", 2) = 0 Then
            SetTaskProgress("ProPIL", -500)
            SetTaskButton("BPIL", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BFinal", True)
            SetTaskStatus("StsPIL", 1248, 0)
            SetAtFunction("AtPending", True)
        Else
            SetTaskStatus("StsPIL", 1248, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BFinal_Click)
    ''**     SHOPPING PLAN FINAL PROCESS
    ''**
    ''*****************************************************************
    Protected Sub BFinal_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFinal.Click
        Dim msg As String = ""
        Dim msgcontent As String = ""
        Dim sql As String = ""
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        ' FINAL PROCESS
        If errcode = 0 Then
            If AtPending.Checked = True Then
                ' Pending Process
                If uFASCommon.SPPendingPlan(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text) <> 0 Then
                    errcode = 9031
                End If
            Else
                ' Final Process
                If uFASCommon.SPFinalPlan(DLogID.Text, DCode.Text, xUserID, DCustomerGr.Text, DFunList.Text) <> 0 Then
                    errcode = 9032
                End If
            End If
        End If
        '
        ' 更新客戶控制檔
        If errcode = 0 Then
            If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Final", 2, "Progress", 1) <> 0 Then
                errcode = 9041
            End If
        End If
        ' 檢查匯入作業是否完成
        If errcode = 0 Then
            If uFASCommon.SPCheckControlRecord(DCode.Text, "Final", 2) <> 0 Then
                errcode = 9042
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsFinal", 1064, 1)
            StsFinal.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DCode.Text + "&pUserID=" + xUserID
            StsFinal.Target = "_blank"
            '
            If errcode = 9031 Then msg = "WINGS回報(KP/OR?)不正確，Pending無法處理，請確認(可點選[ｘ]查詢)!"
            If errcode = 9032 Then msg = "WINGS回報(TBA有?)不正確，Final無法處理，請確認(可點選[ｘ]查詢)!"
            If errcode = 9041 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9042 Then msg = "轉換資料異常，請連絡系統人員!"
            '
            uJavaScript.PopMsg(Me, msg)
        Else
            StsFinal.NavigateUrl = ""
            StsFinal.Target = ""
            SetTaskProgress("ProFinal", -500)
            SetTaskButton("BFinal", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsFinal", 1064, 0)
            '
            uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "YOBI1", 0)
            '
            If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "End", 0, "End", 0) = 0 Then
                If AtPending.Checked = False Then
                    LUpdateWINGS.Style("left") = 1100 & "px"
                    LUpdateWINGS.NavigateUrl = "http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=WINGS&pSPCode=" & DCode.Text & ""
                End If
                '
                SetTaskButton("BProgress", True)
            Else
                uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
            End If
            '
        End If
        '
    End Sub
    ''---------------------------------------------------------------------------------------------------
    ''*****************************************************************
    ''**(BChangeFinal_Click)
    ''**     CHANGE FINAL DATA
    ''**
    ''*****************************************************************
    Protected Sub BChangeFinal_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BChangeFinal.Click
        '
        uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "Action2", 2, "ChgFinal", 2)
        '
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ChgFinal", 2)
        If Code <> 0 Then
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ChgFinal", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ChgFinal", 2)
            If Code <> 0 Then
                Code = uFASCommon.SPCheckControlRecord(DCode.Text, "ChgFinal", 3)
            End If
        Loop
        '
        '檢查是否已將資料上傳完成
        If uFASCommon.SPCheckControlRecord(DCode.Text, "ChgFinal", 2) = 0 Then
            SetTaskProgress("ProChangeFinal", -500)
            SetTaskButton("BChangeFinal", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsChangeFinal", 1192, 0)
            '
            If uFASCommon.SPUpdateControlRecord(DCode.Text, xUserID, "End", 0, "End", 0) = 0 Then
                SetTaskButton("BProgress", True)
                '
                Dim Cmd As String
                Cmd = "<script>" + _
                            "window.open('http://10.245.0.205/EDI/SP_ShoppingListInf.aspx?pUserID=" & xUserID & "&pSPName=" & DCustomerBuyer.SelectedValue & "','',''); " + _
                      "</script>"
                Response.Write(Cmd)
            Else
                uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
            End If
            '
        Else
            SetTaskStatus("StsChangeFinal", 1192, 1)
            uJavaScript.PopMsg(Me, "未完成，請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BUTTON-GO New Customer Automate)
    '**     
    '**
    '*****************************************************************
    Protected Sub BO365_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BO365.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('https://forms.office.com/r/jhFXy5y3sz');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        ''
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()

            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "User"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "AccessTime"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Cat"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Active"
            tcl(i).BackColor = Color.Blue
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Name"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            ' -----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Status"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Customer"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Import"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Plan" & "<br>" & "Proc."
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Plan" & "<br>" & "Report"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "YOC" & "<br>" & "Confirm"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "KP"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "WINGS"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "PIL"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Final"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Change" & "<br>" & "Final"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Progress"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '
            '** line *****************
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = ""
            H3tc1.ColumnSpan = 7
            H3tc1.BackColor = Color.LightGray
            H3row.Cells.Add(H3tc1)
            '** 
            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "  * : Processing completed."
            'H3tc2.HorizontalAlign = HorizontalAlign.Left
            H3tc2.ColumnSpan = 12
            H3tc2.Font.Bold = False
            H3tc2.ForeColor = Color.Red
            H3tc2.BackColor = Color.White
            H3row.Cells.Add(H3tc2)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            If e.Row.Cells(0).Text = "00000" Then
                e.Row.Cells(i).ForeColor = Color.Red
                For i = 0 To 18
                    Select Case i
                        Case 5
                            e.Row.Cells(0).Text = e.Row.Cells(i).Text
                            e.Row.Cells(i).Text = ""
                        Case Else
                            e.Row.Cells(i).Text = ""
                    End Select
                Next
                e.Row.Cells(0).ColumnSpan = 10
            End If

            For i = 0 To 18
                Select Case i
                    Case 5
                    Case Else
                End Select
            Next
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView2 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        Dim i As Integer
        ''
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()

            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Code"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Name"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "No"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Import"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Now"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Now-Import"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Link"
            tcl(i).BackColor = Color.Red
            i = i + 1
            '
            '** line *****************
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Unable to close then case. (>31 Calenday days) "
            H3tc1.ColumnSpan = 7
            H3tc1.ForeColor = Color.White
            H3tc1.BackColor = Color.Black
            H3row.Cells.Add(H3tc1)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView3 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        Dim i As Integer
        ''
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()

            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Code"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Name"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Time"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "No"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "OR"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "C.Order No"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Time"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "User"
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Item Inf."
            tcl(i).BackColor = Color.Purple
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Days"
            tcl(i).BackColor = Color.Red
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Link"
            tcl(i).BackColor = Color.Red
            i = i + 1
            '
            '** line *****************
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Shopping"
            H3tc1.ColumnSpan = 4
            H3tc1.ForeColor = Color.White
            H3tc1.BackColor = Color.Black
            H3row.Cells.Add(H3tc1)
            '
            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "WINGS Order fail"
            H3tc2.ColumnSpan = 5
            H3tc2.ForeColor = Color.White
            H3tc2.BackColor = Color.Black
            H3row.Cells.Add(H3tc2)
            '
            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "Pls modify"
            H3tc3.ColumnSpan = 2
            H3tc3.ForeColor = Color.White
            H3tc3.BackColor = Color.Black
            H3row.Cells.Add(H3tc3)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)
            '
            '** line *****************
            Dim H3tc9 As TableCell = New TableCell
            H3tc9.Text = "WINGS Stock Order(LS) is fail. (>7 Calenday days) "
            H3tc9.ColumnSpan = 11
            H3tc9.ForeColor = Color.White
            H3tc9.BackColor = Color.Black
            H4row.Cells.Add(H3tc9)
            '
            gv.Controls(0).Controls.AddAt(0, H4row)

        End If
    End Sub
End Class

