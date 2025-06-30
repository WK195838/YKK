Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO

Partial Class FASFlow
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
        DLastUniqueID.Style("left") = -500 & "px"
        DFileName.Style("left") = -500 & "px"
        DFilePath.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        DReportFileName.Style("left") = -500 & "px"
        DReportFilePath.Style("left") = -500 & "px"
        DReportFilePath1.Style("left") = -500 & "px"
        DReportFilePath2.Style("left") = -500 & "px"
        DPLANReportFilePath.Style("left") = -500 & "px"
        DFCT2ACTReportFilePath.Style("left") = -500 & "px"
        DFCT2ACTMMReportFilePath.Style("left") = -500 & "px"
        DFCT2ACTVDPReportFilePath.Style("left") = -500 & "px"
        DMAT2SEAReportFilePath.Style("left") = -500 & "px"
        DMAT2MONReportFilePath.Style("left") = -500 & "px"
        DLSInputFilePath.Style("left") = -500 & "px"
        DLastVersion.Style("left") = -500 & "px"
        '動作按鈕設定
        BCustomer.Attributes("onclick") = "var ok=window.confirm('" + "是否確定此客戶？" + "');if(!ok){return false;}"
        BExcel.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟Excel準備客戶資料？" + "');if(!ok){return false;} else {CheckAttribute('FASExcel')}"
        BConvert.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 ITEM/COLOR轉換 作業？" + "');if(!ok){return false;} else {CheckAttribute('Convert')}"
        BFCTPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 FCT PLAN 作業？" + "');if(!ok){return false;} else {CheckAttribute('FCTPlan')}"
        BLSPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 LS PLAN 作業？" + "');if(!ok){return false;} else {CheckAttribute('LSPlan')}"
        BReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 報表匯出 作業？" + "');if(!ok){return false;} else {CheckAttribute('Report')}"
        BIPLSPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 匯入 LS PLAN 作業？" + "');if(!ok){return false;} else {CheckAttribute('IPLSPlan')}"
        BBULSPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 LS PLAN (BUYER) 作業？" + "');if(!ok){return false;} else {CheckAttribute('BULSPlan')}"
        BBUReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 報表匯出 (BUYER) 作業？" + "');if(!ok){return false;} else {CheckAttribute('BUReport')}"
        BLFLSPlan.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 LS PLAN最終確定 作業？" + "');if(!ok){return false;} else {CheckAttribute('LFLSPlan')}"
        BEPEDI.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 EDI資料轉換 作業？" + "');if(!ok){return false;} else {CheckAttribute('EPEDI')}"
        BEDIReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 報表匯出 (EDI) 作業？" + "');if(!ok){return false;} else {CheckAttribute('EDIReport')}"

        BPLANR.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 PLAN報表 作業？" + "');if(!ok){return false;} else {CheckAttribute('PLANReport')}"
        BFCT2ACTR.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 比較分析報表 作業？" + "');if(!ok){return false;} else {CheckAttribute('FCT2ACTReport')}"
        BFCT2ACTMMR.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 比較分析報表(MONTH) 作業？" + "');if(!ok){return false;} else {CheckAttribute('FCT2ACTMMReport')}"
        BFCT2ACTVDPR.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 比較分析報表(VDP) 作業？" + "');if(!ok){return false;} else {CheckAttribute('FCT2ACTVDPReport')}"
        BMAT2SEA.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 材料分析報表(季) 作業？" + "');if(!ok){return false;} else {CheckAttribute('MAT2SEAReport')}"
        BMAT2MON.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 材料分析報表(月) 作業？" + "');if(!ok){return false;} else {CheckAttribute('MAT2MONReport')}"

        BLSInput.Attributes("onclick") = "var ok=window.confirm('" + "是否執行 備料輸入作業？" + "');if(!ok){return false;} else {CheckAttribute('LSInput')}"
        'Action CheckBox設定
        AtReport.Attributes.Add("onclick", "ActionCheck('AtReport')")
        AtIPLSPlan.Attributes.Add("onclick", "ActionCheck('AtIPLSPlan')")
        AtBULSPlan.Attributes.Add("onclick", "ActionCheck('AtBULSPlan')")
        AtConvert.Attributes.Add("onclick", "ActionCheck('AtConvert')")
        AtFCTPlan.Attributes.Add("onclick", "ActionCheck('AtFCTPlan')")
        AtLSPlan.Attributes.Add("onclick", "ActionCheck('AtLSPlan')")
        AtLFLSPlan.Attributes.Add("onclick", "ActionCheck('AtLFLSPlan')")
        AtEPEDI.Attributes.Add("onclick", "ActionCheck('AtEPEDI')")
        AtOffice2010.Enabled = True
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
                sql = sql & "   And DKey = '" & xUserID & "' "
                sql = sql & "   And Data Like 'F-%' "
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
            Case "StsConvert"
                StsConvert.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsConvert.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsConvert.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsFCTPlan"
                StsFCTPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsFCTPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsFCTPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsLSPlan"
                StsLSPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsLSPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsLSPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsReport"
                StsReport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsReport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsReport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsIPLSPlan"
                StsIPLSPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsIPLSPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsIPLSPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBULSPlan"
                StsBULSPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBULSPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBULSPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsBUReport"
                StsBUReport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsBUReport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsBUReport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsLFLSPlan"
                StsLFLSPlan.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsLFLSPlan.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsLFLSPlan.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsEPEDI"
                StsEPEDI.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsEPEDI.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsEPEDI.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case "StsEDIReport"
                StsEDIReport.Style("left") = pLeft & "px"
                If pLeft > 0 Then
                    If pStatus = 0 Then
                        StsEDIReport.ImageUrl = "iMages/OK.jpg"
                    Else
                        StsEDIReport.ImageUrl = "iMages/NG.jpg"
                    End If
                End If
            Case Else
                StsCustomer.Style("left") = pLeft & "px"
                StsExcel.Style("left") = pLeft & "px"
                StsConvert.Style("left") = pLeft & "px"
                StsFCTPlan.Style("left") = pLeft & "px"
                StsLSPlan.Style("left") = pLeft & "px"
                StsReport.Style("left") = pLeft & "px"
                StsIPLSPlan.Style("left") = pLeft & "px"
                StsBULSPlan.Style("left") = pLeft & "px"
                StsBUReport.Style("left") = pLeft & "px"
                StsLFLSPlan.Style("left") = pLeft & "px"
                StsEPEDI.Style("left") = pLeft & "px"
                StsEDIReport.Style("left") = pLeft & "px"
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
            Case "ProConvert"
                ProConvert.Style("left") = pLeft & "px"
            Case "ProFCTPlan"
                ProFCTPlan.Style("left") = pLeft & "px"
            Case "ProLSPlan"
                ProLSPlan.Style("left") = pLeft & "px"
            Case "ProReport"
                ProReport.Style("left") = pLeft & "px"
            Case "ProIPLSPlan"
                ProIPLSPlan.Style("left") = pLeft & "px"
            Case "ProBULSPlan"
                ProBULSPlan.Style("left") = pLeft & "px"
            Case "ProBUReport"
                ProBUReport.Style("left") = pLeft & "px"
            Case "ProLFLSPlan"
                ProLFLSPlan.Style("left") = pLeft & "px"
            Case "ProEPEDI"
                ProEPEDI.Style("left") = pLeft & "px"
            Case "ProEDIReport"
                ProEDIReport.Style("left") = pLeft & "px"
            Case Else
                ProExcel.Style("left") = pLeft & "px"
                ProConvert.Style("left") = pLeft & "px"
                ProFCTPlan.Style("left") = pLeft & "px"
                ProLSPlan.Style("left") = pLeft & "px"
                ProReport.Style("left") = pLeft & "px"
                ProIPLSPlan.Style("left") = pLeft & "px"
                ProBULSPlan.Style("left") = pLeft & "px"
                ProBUReport.Style("left") = pLeft & "px"
                ProLFLSPlan.Style("left") = pLeft & "px"
                ProEPEDI.Style("left") = pLeft & "px"
                ProEDIReport.Style("left") = pLeft & "px"
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
            Case "BConvert"
                BConvert.Visible = pShow
            Case "BFCTPlan"
                BFCTPlan.Visible = pShow
            Case "BLSPlan"
                BLSPlan.Visible = pShow
            Case "BReport"
                BReport.Visible = pShow
            Case "BIPLSPlan"
                BIPLSPlan.Visible = pShow
            Case "BBULSPlan"
                BBULSPlan.Visible = pShow
            Case "BBUReport"
                BBUReport.Visible = pShow
            Case "BLFLSPlan"
                BLFLSPlan.Visible = pShow
            Case "BEPEDI"
                BEPEDI.Visible = pShow
            Case "BEDIReport"
                BEDIReport.Visible = pShow
            Case "BActionLog"
                BActionLog.Visible = pShow
            Case "BBuyerFCT"
                If DBuyer.Text = "FALL-VENDOR" Or InStr(DBuyer.Text, "F-VENDOR") > 0 Then
                    BLSInput.Visible = pShow
                Else
                    BBuyerFCT.Visible = pShow
                End If
            Case "BItemCheck"
                BItemCheck.Visible = pShow

            Case "BLEFT"
                BLEFT.Visible = pShow
            Case "BRIGHT"
                BRIGHT.Visible = pShow

            Case "BPLANR"
                BPLANR.Visible = pShow
            Case "BFCT2ACTR"
                BFCT2ACTR.Visible = pShow
            Case "BFCT2ACTMMR"
                BFCT2ACTMMR.Visible = pShow
            Case "BFCT2ACTVDPR"
                BFCT2ACTVDPR.Visible = pShow
            Case "BMAT2SEA"
                'BMAT2SEA.Visible = pShow
            Case "BMAT2MON"
                'BMAT2MON.Visible = pShow
            Case "BKPI"
                If DBuyer.Text = "FALL-000001" Or DBuyer.Text = "FALL-000001-V" Then
                    BKPI.Visible = pShow
                End If
            Case "BFCT2ORD"
                BFCT2ORD.Visible = pShow
            Case "BFCT2FCT"
                BFCT2FCT.Visible = pShow
            Case "BFCT2ACT"
                BFCT2ACT.Visible = pShow
            Case "BFCT2ACTVDP"
                BFCT2ACTVDP.Visible = pShow
            Case "BSTOCK"
                BSTOCK.Visible = pShow
            Case "BALERT"
                BALERT.Visible = pShow
            Case "BVDPERR"
                BVDPERR.Visible = pShow
            Case "BVDPEDIT"
                BVDPEDIT.Visible = pShow

            Case "BGotoEDI"
                BGotoEDI.Visible = pShow
            Case Else
                BReset.Visible = False
                BCustomer.Visible = True
                BExcel.Visible = False
                BConvert.Visible = False
                BFCTPlan.Visible = False
                BLSPlan.Visible = False
                BReport.Visible = False
                BIPLSPlan.Visible = False
                BBULSPlan.Visible = False
                BBUReport.Visible = False
                BLFLSPlan.Visible = False
                BEPEDI.Visible = False
                BEDIReport.Visible = False
                '
                If UCase(xUserID) = "IT003" Then
                    BActionLog.Visible = True
                Else
                    BActionLog.Visible = False
                End If

                BLEFT.Visible = False
                BRIGHT.Visible = False

                BBuyerFCT.Visible = False
                BLSInput.Visible = False
                BItemCheck.Visible = False
                BPLANR.Visible = False
                BFCT2ACTR.Visible = False
                BFCT2ACTMMR.Visible = False
                BFCT2ACTVDPR.Visible = False
                BMAT2SEA.Visible = False
                BMAT2MON.Visible = False
                BKPI.Visible = False
                BFCT2ORD.Visible = False
                BFCT2FCT.Visible = False
                BFCT2ACT.Visible = False
                BFCT2ACTVDP.Visible = False
                BSTOCK.Visible = False
                BALERT.Visible = False
                BVDPERR.Visible = False
                BVDPEDIT.Visible = False
                BGotoEDI.Visible = False
                '
                AtReport.Checked = False
                AtIPLSPlan.Checked = False
                AtBULSPlan.Checked = False
                AtConvert.Checked = False
                AtFCTPlan.Checked = False
                AtLSPlan.Checked = False
                AtLFLSPlan.Checked = False
                AtEPEDI.Checked = False
                '
                AtAddition.Visible = False
                AtAddition.Checked = False
                '
                'Test-Start
                'BReset.Visible = True
                'BExcel.Visible = True
                'BConvert.Visible = True
                'AtAddition.Visible = True
                'BFCTPlan.Visible = True
                'BLSPlan.Visible = True
                'BReport.Visible = True
                'BIPLSPlan.Visible = True
                'BBULSPlan.Visible = True
                'BBUReport.Visible = True
                'BLFLSPlan.Visible = True
                'BEPEDI.Visible = True
                'BEDIReport.Visible = True
                'BActionLog.Visible = True
                'BBuyerFCT.Visible = True
                'BLSInput.Visible = True
                'BItemCheck.Visible = True

                'BLEFT.Visible = True
                'BRIGHT.Visible = True
                'BPLANR.Visible = True
                'BFCT2ACTR.Visible = True
                'BFCT2ACTMMR.Visible = True
                'BFCT2ACTVDPR.Visible = True
                'BMAT2SEA.Visible = True
                'BMAT2MON.Visible = True
                'BKPI.Visible = True
                'BFCT2ORD.Visible = True
                'BFCT2FCT.Visible = True
                'BFCT2ACT.Visible = True
                'BFCT2ACTVDP.Visible = True
                'BSTOCK.Visible = True
                'BALERT.Visible = True
                'BVDPERR.Visible = True
                'BVDPEDIT.Visible = True

                'BGotoEDI.visible=true
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
        sql = "Select * From M_FControlRecord "
        sql = sql & " Where Name = '" & DCustomerBuyer.SelectedValue & "'"
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
                ' --------------------------------------------------------------------------
                ' 區分FCT資料作業者 & 統計分析者 
                ' --------------------------------------------------------------------------
                'If InStr(DBuyer.Text, "-V") > 0 Then
                If Right(DBuyer.Text, 2) = "-V" Then
                    ' 統計分析者 
                    ' --------------------------------------------------------------------------
                    ' FunctionList(4)   判斷是否有管理分析功能 
                    ' --------------------------------------------------------------------------
                    If fpObj.GetFunctionCode(DFunList.Text, 4) = "1" Then
                        BPLANR.Style("left") = 254 & "px"
                        BFCT2ACTR.Style("left") = 339 & "px"
                        BFCT2ACTMMR.Style("left") = 424 & "px"
                        BFCT2ACTVDPR.Style("left") = 509 & "px"
                        BMAT2SEA.Style("left") = 594 & "px"
                        BMAT2MON.Style("left") = 679 & "px"
                        BKPI.Style("left") = 594 & "px"
                        BFCT2ORD.Style("left") = 764 & "px"
                        BFCT2FCT.Style("left") = 849 & "px"
                        BFCT2ACT.Style("left") = 934 & "px"
                        BFCT2ACTVDP.Style("left") = 1019 & "px"

                        SetTaskButton("BLEFT", False)
                        SetTaskButton("BPLANR", True)
                        SetTaskButton("BFCT2ACTR", True)
                        SetTaskButton("BFCT2ACTMMR", True)
                        SetTaskButton("BFCT2ACTVDPR", True)
                        SetTaskButton("BMAT2SEA", True)
                        SetTaskButton("BMAT2MON", True)
                        SetTaskButton("BKPI", True)
                        SetTaskButton("BFCT2ORD", True)
                        SetTaskButton("BFCT2FCT", True)
                        SetTaskButton("BFCT2ACT", True)
                        SetTaskButton("BFCT2ACTVDP", True)
                        SetTaskButton("BSTOCK", False)
                        SetTaskButton("BALERT", False)
                        SetTaskButton("BVDPERR", False)
                        SetTaskButton("BVDPEDIT", False)
                        SetTaskButton("BRIGHT", True)
                    End If
                    ' --------------------------------------------------------------------------
                    ' 設定 Data or 報表程式
                    ' --------------------------------------------------------------------------
                    '-- 管理Report
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_PlanReport.xlsm"        ' PLAN 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_PlanReport.xls"         ' PLAN 報表資料
                    End If
                    DPLANReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                    '
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTReport.xlsm"     ' FCT2ACT 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTReport.xls"      ' FCT2ACT 報表資料
                    End If
                    DFCT2ACTReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                    '
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTMMReport.xlsm"    ' FCT2ACTMM 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTMMReport.xls"     ' FCT2ACTMM 報表資料
                    End If
                    DFCT2ACTMMReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                    '
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTVDPReport.xlsm"  ' FCT2ACTVDP 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_FCT2ACTVDPReport.xls"   ' FCT2ACTVDP 報表資料
                    End If
                    DFCT2ACTVDPReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                    '
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_MAT2SEASON.xlsm"        ' 材料分析(季) 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_MAT2SEASON.xls"         ' 材料分析(季) 報表資料
                    End If
                    DMAT2SEAReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                    '
                    If AtOffice2010.Checked = True Then
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_MAT2MONTH.xlsm"        ' 材料分析(月) 報表資料
                    Else
                        DReportFileName.Text = Mid(DCustomerBuyer.SelectedValue, 1, InStr(DCustomerBuyer.SelectedValue, "-V") - 1) + "_ANL_MAT2MONTH.xls"         ' 材料分析(月) 報表資料
                    End If
                    DMAT2MONReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                Else
                    ' FCT資料作業者 
                    If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Start", 0, "Start", 0) = 0 Then
                        SetTaskButton("BCustomer", False)
                        SetTaskButton("BReset", True)
                        SetTaskStatus("StsCustomer", 127, 0)
                        SetCustomerBuyer("Customer", DCustomerBuyer.SelectedValue)
                        ' --------------------------------------------------------------------------
                        ' FunctionList(1)   判斷是否跳躍功能 
                        ' --------------------------------------------------------------------------
                        ' 準備FCT資料
                        If AtReport.Checked = False And AtIPLSPlan.Checked = False And _
                           AtBULSPlan.Checked = False And AtConvert.Checked = False And _
                           AtFCTPlan.Checked = False And AtLSPlan.Checked = False And _
                           AtLFLSPlan.Checked = False And AtEPEDI.Checked = False Then
                            SetTaskButton("BExcel", True)
                        End If
                        ' 報表
                        If AtReport.Checked = True Then
                            SetTaskButton("BReport", True)
                        End If
                        '
                        ' 匯入LS Plan
                        If AtIPLSPlan.Checked = True Then
                            SetTaskButton("BIPLSPlan", True)
                        End If
                        '
                        ' Buyer Proc.
                        If AtBULSPlan.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=1
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "1" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                SetTaskButton("BBULSPlan", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        '
                        ' Convert
                        If AtConvert.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=2
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "2" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                SetTaskButton("BConvert", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        '
                        ' FCT Plan
                        If AtFCTPlan.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=3
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "3" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                DLastUniqueID.Text = "0"
                                SetTaskButton("BFCTPlan", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        '
                        ' LS Plan
                        If AtLSPlan.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=4
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "4" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                SetTaskButton("BLSPlan", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        '
                        ' 最終確定LS Plan
                        If AtLFLSPlan.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=5
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "5" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                SetTaskButton("BLFLSPlan", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        '
                        ' EDI Convert
                        If AtEPEDI.Checked = True Then
                            ' 判斷是否有使用權限  FunList(1)=6
                            If fpObj.GetFunctionCode(DFunList.Text, 1) = "6" Or fpObj.GetFunctionCode(DFunList.Text, 1) = "9" Then
                                DLastVersion.Text = "Y"
                                SetTaskButton("BEPEDI", True)
                            Else
                                errcode = 9005
                            End If
                        End If
                        ' --------------------------------------------------------------------------
                        ' FunctionList(2)   判斷是否有BUYER FCT功能 
                        ' --------------------------------------------------------------------------
                        If fpObj.GetFunctionCode(DFunList.Text, 2) = "1" Then
                            SetTaskButton("BBuyerFCT", True)
                        End If
                        ' --------------------------------------------------------------------------
                        ' FunctionList(3)   判斷是否有ITEM CHECK功能 
                        ' --------------------------------------------------------------------------
                        If fpObj.GetFunctionCode(DFunList.Text, 3) = "1" Then
                            SetTaskButton("BItemCheck", True)
                        End If
                        ' --------------------------------------------------------------------------
                        ' FunctionList(4)   判斷是否有管理分析功能 
                        ' --------------------------------------------------------------------------
                        If fpObj.GetFunctionCode(DFunList.Text, 4) = "1" Then
                            BPLANR.Style("left") = 254 & "px"
                            BFCT2ACTR.Style("left") = 339 & "px"
                            BFCT2ACTMMR.Style("left") = 424 & "px"
                            BFCT2ACTVDPR.Style("left") = 509 & "px"
                            BMAT2SEA.Style("left") = 594 & "px"
                            BMAT2MON.Style("left") = 679 & "px"
                            BKPI.Style("left") = 594 & "px"
                            BFCT2ORD.Style("left") = 764 & "px"
                            BFCT2FCT.Style("left") = 849 & "px"
                            BFCT2ACT.Style("left") = 934 & "px"
                            BFCT2ACTVDP.Style("left") = 1019 & "px"

                            SetTaskButton("BLEFT", False)
                            SetTaskButton("BPLANR", True)
                            SetTaskButton("BFCT2ACTR", True)
                            SetTaskButton("BFCT2ACTMMR", True)
                            SetTaskButton("BFCT2ACTVDPR", True)
                            SetTaskButton("BMAT2SEA", True)
                            SetTaskButton("BMAT2MON", True)
                            SetTaskButton("BKPI", True)
                            SetTaskButton("BFCT2ORD", True)
                            SetTaskButton("BFCT2FCT", True)
                            SetTaskButton("BFCT2ACT", True)
                            SetTaskButton("BFCT2ACTVDP", True)
                            SetTaskButton("BSTOCK", False)
                            SetTaskButton("BALERT", False)
                            SetTaskButton("BVDPERR", False)
                            SetTaskButton("BVDPEDIT", False)
                            SetTaskButton("BRIGHT", True)
                        Else
                            If fpObj.GetFunctionCode(DFunList.Text, 4) = "2" Then       ' 限定FC-VENDOR
                                BPLANR.Style("left") = 254 & "px"
                                BSTOCK.Style("left") = 339 & "px"
                                SetTaskButton("BLEFT", False)
                                SetTaskButton("BPLANR", True)
                                SetTaskButton("BFCT2ACTR", False)
                                SetTaskButton("BFCT2ACTMMR", False)
                                SetTaskButton("BFCT2ACTVDPR", False)
                                SetTaskButton("BMAT2SEA", False)
                                SetTaskButton("BMAT2MON", False)
                                SetTaskButton("BKPI", False)
                                SetTaskButton("BFCT2ORD", False)
                                SetTaskButton("BFCT2FCT", False)
                                SetTaskButton("BFCT2ACT", False)
                                SetTaskButton("BFCT2ACTVDP", False)
                                SetTaskButton("BSTOCK", True)
                                SetTaskButton("BALERT", False)
                                SetTaskButton("BVDPERR", False)
                                SetTaskButton("BVDPEDIT", False)
                                SetTaskButton("BRIGHT", False)
                            End If
                        End If
                        ' --------------------------------------------------------------------------
                        ' 設定 Data or 報表程式
                        ' --------------------------------------------------------------------------
                        If AtOffice2010.Checked = True Then
                            DFileName.Text = DCustomerBuyer.SelectedValue + "_DataPrepare.xlsm"             ' 準備FST
                        Else
                            DFileName.Text = DCustomerBuyer.SelectedValue + "_DataPrepare.xls"              ' 準備FST
                        End If
                        DFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DFileName.Text = DCustomerBuyer.SelectedValue + "_ImportLSPlan.xlsm"            ' 匯入LS Plan
                        Else
                            DFileName.Text = DCustomerBuyer.SelectedValue + "_ImportLSPlan.xls"             ' 匯入LS Plan
                        End If
                        DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ReportBuilder.xlsm"     ' 客戶報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ReportBuilder.xls"      ' 客戶報表資料
                        End If
                        DReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_BuyerReportBuilder.xlsm"    ' Buyer報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_BuyerReportBuilder.xls"     ' Buyer報表資料
                        End If
                        DReportFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_EDIReportBuilder.xlsm"      ' EDI報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_EDIReportBuilder.xls"       ' EDI報表資料
                        End If
                        DReportFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        '-- 管理Report
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_PlanReport.xlsm"        ' PLAN 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_PlanReport.xls"         ' PLAN 報表資料
                        End If
                        DPLANReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTReport.xlsm"     ' FCT2ACT 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTReport.xls"      ' FCT2ACT 報表資料
                        End If
                        DFCT2ACTReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTMMReport.xlsm"    ' FCT2ACTMM 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTMMReport.xls"     ' FCT2ACTMM 報表資料
                        End If
                        DFCT2ACTMMReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTVDPReport.xlsm"  ' FCT2ACTVDP 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_FCT2ACTVDPReport.xls"   ' FCT2ACTVDP 報表資料
                        End If
                        DFCT2ACTVDPReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_MAT2SEASON.xlsm"  ' 材料分析(季) 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_MAT2SEASON.xls"   ' 材料分析(季) 報表資料
                        End If
                        DMAT2SEAReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        If AtOffice2010.Checked = True Then
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_MAT2MONTH.xlsm"  ' 材料分析(月) 報表資料
                        Else
                            DReportFileName.Text = DCustomerBuyer.SelectedValue + "_ANL_MAT2MONTH.xls"   ' 材料分析(月) 報表資料
                        End If
                        DMAT2MONReportFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                        '
                        ' VENDOR-FC
                        If DBuyer.Text = "FALL-VENDOR" Or InStr(DBuyer.Text, "F-VENDOR") > 0 Then
                            If AtOffice2010.Checked = True Then
                                DReportFileName.Text = DCustomerBuyer.SelectedValue + "_InputSheet.xlsm"      ' 準備FST
                            Else
                                DReportFileName.Text = DCustomerBuyer.SelectedValue + "_InputSheet.xls"       ' 準備FST
                            End If
                            DLSInputFilePath.Text = uCommon.GetAppSetting("DataPrepareFile") + DReportFileName.Text
                            'F-VENDOR-INPUT 複製 InputSheet --> F-VENDOR-INPUT-[USERID]_InputSheet
                            If Not File.Exists(DLSInputFilePath.Text) Then
                                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "F-VENDOR-INPUT.xlsm", DLSInputFilePath.Text)
                            End If
                        End If
                        '
                        BReset.Attributes("onclick") = "var ok=window.confirm('" + "是否取消客戶[" + DCustomerBuyer.SelectedValue + "]，" + "並且重新選擇客戶？" + "');if(!ok){return false;}"
                    Else
                        errcode = 9003
                    End If
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
        DFunList.Text = ""
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BExcel_Click)
    '**     準備客戶資料
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        '*****************************************************************
        '**  準備客戶資料
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
            SetTaskProgress("ProExcel", -500)
            SetTaskButton("BExcel", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BConvert", True)
            AtAddition.Visible = True
            SetTaskStatus("StsExcel", 165, 0)
        Else
            SetTaskStatus("StsExcel", 165, 1)
            uJavaScript.PopMsg(Me, "客戶資料準備未完成，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BConvert_Click)
    '**     ITEM / COLOR 轉換作業
    '**
    '*****************************************************************
    Protected Sub BConvert_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BConvert.Click
        Dim msg As String = ""
        Dim msgcontent As String = ""
        Dim sql As String = ""
        Dim i As Integer
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        '*****************************************************************
        '**  ITEM / COLOR 轉換
        '*****************************************************************
        ' 刪除執行履歷/FCT Plan/LS Plan
        If errcode = 0 Then
            If uFASCommon.DeleteHistory(DBuyer.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        ' 刪除FCT Plan上回資料 & Reset FCTNo
        If errcode = 0 Then
            If AtAddition.Checked = False Then
                DLastUniqueID.Text = "0"
                If uFASCommon.DeleteFCTData(DBuyer.Text) <> 0 Then
                    errcode = 9002
                End If
            Else
                DLastUniqueID.Text = CStr(uFASCommon.GetLastUniqueID(DBuyer.Text, DGRBuyer.Text))
            End If
        End If
        ' 匯入作業
        If errcode = 0 Then
            If uFASMapping.Rule2Data(DLogID.Text, DBuyer.Text, xUserID, DFunList.Text) <> 0 Then
                errcode = 9011
            Else
                '* ADD-START 2021/6/25  PBR切替-中止使用
                'sql = "SELECT Y_ItemCode FROM ForcastPlan "
                'sql &= "Where Buyer = '" & DBuyer.Text & "' "
                'sql &= "GROUP BY Y_ItemCode "
                'sql &= "Order by Y_ItemCode "
                'Dim dt_ForcastPlan As DataTable = uDataBase.GetDataTable(sql)
                'For i = 0 To dt_ForcastPlan.Rows.Count - 1
                '    '
                '    sql = "SELECT Unique_ID, YCode, PCode FROM M_PBRItemConvert "
                '    sql &= "Where YCode = '" & dt_ForcastPlan.Rows(i).Item("Y_ItemCode") & "' "
                '    Dim dt_PBRITEM As DataTable = uDataBase.GetDataTable(sql)
                '    If dt_PBRITEM.Rows.Count > 0 Then
                '        msgcontent = "FAS-ITEM/PBR-ITEM/ID" + ":" + _
                '                     dt_PBRITEM.Rows(0).Item("YCode").ToString + "/" + _
                '                     dt_PBRITEM.Rows(0).Item("PCode").ToString + "/" + _
                '                     dt_PBRITEM.Rows(0).Item("Unique_ID").ToString
                '        errcode = 9014
                '        Exit For
                '    End If
                'Next
                '* ADD-END  
            End If
        End If
        ' 更新客戶控制檔(DataConvert=2, FCTPlan=1)
        If errcode = 0 Then
            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "DataConvert", 2, "FCTPlan", 1) <> 0 Then
                errcode = 9012
            End If
        End If
        ' 檢查匯入作業是否完成
        If errcode = 0 Then
            If uFASCommon.CheckControlRecord(DBuyer.Text, "DataConvert", 2) <> 0 Then
                errcode = 9013
            End If
        End If
        ' 執行結果處理
        If errcode <> 0 Then
            SetTaskStatus("StsConvert", 192, 1)
            StsConvert.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
            StsConvert.Target = "_blank"
            '
            If errcode = 9001 Then msg = "刪除執行履歷異常(ActionHistory)，請連絡系統人員!"
            If errcode = 9002 Then msg = "刪除 FCT PLAN 資料異常，請連絡系統人員!"
            If errcode = 9011 Then msg = "轉換資料異常，請確認(可點選[ｘ]查詢)!"
            If errcode = 9012 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
            If errcode = 9013 Then msg = "轉換資料異常，請連絡系統人員!"
            If errcode = 9014 Then msg = "PBR ITEM異常!  " & msgcontent
            '
            uJavaScript.PopMsg(Me, msg)
        Else
            StsConvert.NavigateUrl = ""
            StsConvert.Target = ""
            SetTaskProgress("ProConvert", -500)
            SetTaskButton("BConvert", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BFCTPlan", True)
            SetTaskStatus("StsConvert", 192, 0)

            '2017/8/24 ADIDAS 備料問題發生
            Dim Cmd As String
            '
            Cmd = "<script>" + _
                        "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=CONVERT" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                  "</script>"
            '
            Response.Write(Cmd)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCTPlan_Click)
    '**     Forcast Plan展開
    '**
    '*****************************************************************
    Protected Sub BFCTPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCTPlan.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        '*****************************************************************
        '**  FCT Plan展開
        '*****************************************************************
        If errcode = 0 Then
            '
            ' 刪除執行履歷(Action=FCTPLAN)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "FCTPLAN") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 刪除FCT Plan上回資料(LEVEL<>0) 
            If errcode = 0 Then
                If uFASCommon.DeleteFCTLevelData(DBuyer.Text) <> 0 Then
                    errcode = 9006
                Else
                    ' Reset FCT Plan FCTNo
                    If uFASCommon.ResetPlanFCTNo(DBuyer.Text) <> 0 Then
                        errcode = 9007
                    Else
                        ' Reset Control FCTNo
                        If uFASCommon.ResetFCTNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
                            errcode = 9008
                        End If
                    End If
                End If
            End If
            '
            ' FCTNo展開
            If errcode = 0 Then
                If uFASCommon.MakeForcastNo(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                    errcode = 9002
                End If
            End If
            '
            ' Forcast Plan展開
            If errcode = 0 Then
                ' MODIFY-START BY JOY 2018/1/26
                'If errcode = 0 Then
                '    ' [FAS-PJ] MODIFY-START BY JOY 2017/9/27
                '    'If uFASCommon.NewForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                '    '    errcode = 9003
                '    'End If

                '    If DBuyer.Text = "FALL-VENDOR" Or InStr(DBuyer.Text, "F-VENDOR") > 0 Then
                '        If uFASCommon.VendorForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                '            errcode = 9003
                '        End If
                '    Else
                '        If uFASCommon.NewForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                '            errcode = 9003
                '        End If
                '    End If
                '    ' [FAS-PJ] MODIFY-END
                'End If
                '
                'FUN LIST 5 = [1]→BUYER/ [2]→YKK扣具/ [3]→VENDOR/ [4]→材料
                If fpObj.GetFunctionCode(DFunList.Text, 5) = "1" Then
                    If uFASCommon.NewForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                        errcode = 9003
                    End If
                Else
                    If fpObj.GetFunctionCode(DFunList.Text, 5) = "2" Then
                        If uFASCommon.TPForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                            errcode = 9003
                        End If
                    Else
                        If fpObj.GetFunctionCode(DFunList.Text, 5) = "3" Then
                            If uFASCommon.VendorForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                                errcode = 9003
                            End If
                        Else
                            If fpObj.GetFunctionCode(DFunList.Text, 5) = "4" Then
                                If uFASCommon.MatForcastPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                                    errcode = 9003
                                End If
                            Else
                                If fpObj.GetFunctionCode(DFunList.Text, 5) = "5" Then
                                    If uFASCommon.MatForcastPlanUABAG(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text, CInt(DLastUniqueID.Text)) <> 0 Then
                                        errcode = 9003
                                    End If
                                Else
                                    errcode = 9003
                                End If
                            End If
                        End If
                    End If
                End If
                ' MODIFY-END
            End If
                '
                ' 更新客戶控制檔(FCTPlan=2, LSPlan=1)
                If errcode = 0 Then
                    If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "FCTPlan", 2, "LSPlan", 1) <> 0 Then
                        errcode = 9004
                    End If
                End If
                '
                ' 檢查Forcast Plan展開是否完成
                If errcode = 0 Then
                    If uFASCommon.CheckControlRecord(DBuyer.Text, "FCTPlan", 2) <> 0 Then
                        errcode = 9005
                    End If
                End If
                '
                ' 執行結果處理
                If errcode <> 0 Then
                    SetTaskStatus("StsFCTPlan", 552, 1)
                    StsFCTPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                    StsFCTPlan.Target = "_blank"
                    '
                    If errcode = 9001 Then msg = "刪除執行履歷異常(FCTPLAN)，請連絡系統人員!"
                    If errcode = 9002 Then msg = "FCT-No.異常，請確認(可點選[ｘ]查詢)!"
                    If errcode = 9003 Then msg = "FCT Plan展開異常，請確認(可點選[ｘ]查詢)!"
                    If errcode = 9004 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                    If errcode = 9005 Then msg = "Forcast Plan展開異常，請連絡系統人員!"
                    If errcode = 9006 Then msg = "刪除 FCT PLAN 資料異常(LEVEL<>0)，請連絡系統人員!"
                    If errcode = 9007 Then msg = "無法重新設定FCT-PLAN FCT No.，請連絡系統人員!"
                    If errcode = 9008 Then msg = "無法重新設定可使用FCT No.，請連絡系統人員!"
                    uJavaScript.PopMsg(Me, msg)
                Else
                    SetTaskStatus("StsFCTPlan", 552, 0)
                    SetTaskProgress("ProFCTPlan", -500)
                    SetTaskButton("BFCTPlan", False)
                    SetTaskButton("BReset", True)
                    SetTaskButton("BLSPlan", True)

                    '2017/8/24 ADIDAS 備料問題發生
                    Dim Cmd As String
                    '
                    Cmd = "<script>" + _
                            "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=FCTPLAN" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
                '
            End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BLSPlan_Click)
    '**     Local Stock Plan展開
    '**
    '*****************************************************************
    Protected Sub BLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLSPlan.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        '*****************************************************************
        '**  Local Stock Plan展開
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=LSPLAN)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "LSPLAN") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 刪除LS Plan上回資料 & Reset LSNo
            If errcode = 0 Then
                If uFASCommon.DeleteLSData(DBuyer.Text) <> 0 Then
                    errcode = 9003
                Else
                    If uFASCommon.ResetLSNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
                        errcode = 9005
                    End If
                End If
            End If
            ' LS Plan展開
            If errcode = 0 Then
                If uFASCommon.NewLocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                    errcode = 9002
                Else
                    'If uFASCommon.LocalStockPlanProdInf(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                    '    errcode = 9002
                    'End If
                End If
            End If
            ' 更新客戶控制檔(LSPlan=2, Report=1)
            If errcode = 0 Then
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "LSPlan", 2, "Report", 1) <> 0 Then
                    errcode = 9006
                End If
            End If
            ' LS Plan展開是否完成
            If errcode = 0 Then
                If uFASCommon.CheckControlRecord(DBuyer.Text, "LSPlan", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsLSPlan", 525, 1)
                StsLSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsLSPlan.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(LSPlan)，請連絡系統人員!"
                If errcode = 9002 Then msg = "LS Plan展開異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "刪除 LS PLAN 資料異常，請連絡系統人員!"
                If errcode = 9004 Then msg = "LS Plan展開異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "無法重新設定可使用LS No.，請連絡系統人員!"
                If errcode = 9006 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsLSPlan", 525, 0)
                SetTaskProgress("ProLSPlan", -500)
                SetTaskButton("BLSPlan", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BReport", True)

                '2017/8/24 ADIDAS 備料問題發生
                Dim Cmd As String
                '
                Cmd = "<script>" + _
                            "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=LSPLAN" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                      "</script>"
                '
                Response.Write(Cmd)
            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BReport_Click)
    '**     FCT Plan and LS Plan Report
    '**
    '*****************************************************************
    Protected Sub BReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BReport.Click
        '
        '*****************************************************************
        '**  準備報表資料
        '*****************************************************************
        '
        '17/10/23 JOY  F-VENDOR
        If InStr(DBuyer.Text, "F-VENDOR") <= 0 Then
            uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "Report", 1)
            '判斷是否完成作業
            Dim Code As Integer
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 3)
            End If
            Do Until Code = 0
                System.Threading.Thread.Sleep(3 * 1000)
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2)
                If Code <> 0 Then
                    Code = uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 3)
                End If
            Loop
            '
            '檢查是否已產出報表
            If uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 2) = 0 Then
                SetTaskProgress("ProReport", -500)
                SetTaskButton("BReport", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BIPLSPlan", True)
                SetTaskStatus("StsReport", 552, 0)
            Else
                SetTaskStatus("StsReport", 552, 1)
                uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
            End If
        Else
            '
            '檢查是否已產出報表
            ClearKeyData()
            SetTaskButton("Default", False)             ' 動作按鈕圖像
            SetTaskStatus("Default", -500, 1)           ' 動作結果圖像
            SetTaskProgress("Default", -500)            ' 動作執行中跑馬燈
            SetCustomerBuyer("Default", "")             ' 設定可使用客戶
            If uFASCommon.CheckControlRecord(DBuyer.Text, "Report", 3) = 0 Then
                uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BIPLSPlan_Click)
    '**     Import LS Plan Data
    '**
    '*****************************************************************
    Protected Sub BIPLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BIPLSPlan.Click
        '
        '*****************************************************************
        '**  匯入修改後LS Plan資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "IPLSPlan", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 3)
            End If
        Loop
        '
        '檢查是否已完成匯入作業
        If uFASCommon.CheckControlRecord(DBuyer.Text, "IPLSPlan", 2) = 0 Then
            SetTaskProgress("ProIPLSPlan", -500)
            SetTaskButton("BIPLSPlan", False)
            SetTaskButton("BReset", True)
            ' 判斷是否 BUYER-FCT
            If Mid(DGRBuyer.Text, 7, 1) = "B" Or Mid(DGRBuyer.Text, 7, 1) = "T" Then
                SetTaskButton("BBULSPlan", True)
            End If
            SetTaskStatus("StsIPLSPlan", 912, 0)

            '2017/8/24 ADIDAS 備料問題發生
            Dim Cmd As String
            '
            Cmd = "<script>" + _
                    "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=IPLSPlan" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                  "</script>"
            '
            Response.Write(Cmd)
        Else
            SetTaskStatus("StsIPLSPlan", 912, 1)
            uJavaScript.PopMsg(Me, "匯入作業未完成，請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBULSPlan_Click)
    '**     Buyer Local Stock Plan展開
    '**
    '*****************************************************************
    Protected Sub BBULSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBULSPlan.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  Buyer Local Stock Plan展開
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=BULSPLAN)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "BULSPLAN") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 刪除BuyerLSPLAN
            If errcode = 0 Then
                If uFASCommon.DeleteBuyerLSData("BULS", DGRBuyer.Text) <> 0 Then
                    errcode = 9005
                Else
                    If uFASCommon.ResetBuyerLSNo(DBuyer.Text, DGRBuyer.Text) <> 0 Then
                        errcode = 9006
                    End If
                End If
            End If
            ' Buyer LS Plan展開
            If errcode = 0 Then
                If uFASCommon.BuyerLocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(BULSPlan=2, BUReport=1)
            If errcode = 0 Then
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "BULSPlan", 2, "BUReport", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' Buyer LS Plan展開是否完成
            If errcode = 0 Then
                If uFASCommon.CheckControlRecord(DBuyer.Text, "BULSPlan", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsBULSPlan", 885, 1)
                StsBULSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsBULSPlan.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(BULSPlan)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Buyer LS Plan展開異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "Buyer LS Plan展開異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "刪除 Buyer LS PLAN 資料異常，請連絡系統人員!"
                If errcode = 9006 Then msg = "無法重新設定可使用Buyer LS No.，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsBULSPlan", 885, 0)
                SetTaskProgress("ProBULSPlan", -500)
                SetTaskButton("BBULSPlan", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BBUReport", True)

                '2017/8/24 ADIDAS 備料問題發生
                Dim Cmd As String
                '
                Cmd = "<script>" + _
                        "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=BYLSPLAN" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                      "</script>"
                '
                Response.Write(Cmd)

            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBUReport_Click)
    '**     Buyer LS Plan Report
    '**
    '*****************************************************************
    Protected Sub BBUReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBUReport.Click
        '
        '*****************************************************************
        '**  準備報表資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "BUReport", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 3)
            End If
        Loop
        '
        '檢查是否已產出報表
        If uFASCommon.CheckControlRecord(DBuyer.Text, "BUReport", 2) = 0 Then
            SetTaskProgress("ProBUReport", -500)
            SetTaskButton("BBUReport", False)
            SetTaskButton("BReset", True)
            SetTaskButton("BLFLSPlan", True)
            SetTaskStatus("StsBUReport", 912, 0)
        Else
            SetTaskStatus("StsBUReport", 912, 1)
            uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BLFLSPlan_Click)
    '**     Buyer LS Plan 最終確定
    '**
    '*****************************************************************
    Protected Sub BLFLSPlan_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLFLSPlan.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  Buyer LS Plan 最終確定
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=LFLSPLAN)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "LFLSPLAN") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 判斷是否可執行
            If errcode = 0 Then
                If uFASCommon.CanRunLFLocalStockPlan(DBuyer.Text, DGRBuyer.Text) <> 0 Then
                    errcode = 9005
                End If
            End If
            ' Buyer LS Plan最終確定
            If errcode = 0 Then
                If uFASCommon.LastFinalLocalStockPlan(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                    errcode = 9002
                End If
            End If
            ' 更新客戶控制檔(LFLSPlan=2, EPEDI=1)
            If errcode = 0 Then
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "LFLSPlan", 2, "EPEDI", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' Buyer LS Plan最終確定是否完成
            If errcode = 0 Then
                If uFASCommon.CheckControlRecord(DBuyer.Text, "LFLSPlan", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsLFLSPlan", 1245, 1)
                StsLFLSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsLFLSPlan.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(LFLSPlan)，請連絡系統人員!"
                If errcode = 9002 Then msg = "Buyer LS Plan最終確定異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "Buyer LS Plan最終確定異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "異常:懷疑最終確定作業重覆執行，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsLFLSPlan", 1245, 0)
                SetTaskProgress("ProLFLSPlan", -500)
                SetTaskButton("BLFLSPlan", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BEPEDI", True)
            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BEPEDI_Click)
    '**     EDI轉換
    '**
    '*****************************************************************
    Protected Sub BEPEDI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BEPEDI.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0
        '
        '*****************************************************************
        '**  EDI Data 轉換
        '*****************************************************************
        If errcode = 0 Then
            ' 刪除執行履歷(Action=LFLSPLAN)
            If errcode = 0 Then
                If uFASCommon.DeleteActionHistory(DLogID.Text, DBuyer.Text, "EDITRANSFER") <> 0 Then
                    errcode = 9001
                End If
            End If
            ' 刪除EDI Data
            If errcode = 0 Then
                If uFASCommon.DeleteLS2EDIInterface("BULS", DGRBuyer.Text) <> 0 Then
                    errcode = 9005
                End If
            End If
            ' EDI Data 轉換
            If errcode = 0 Then
                If DLastVersion.Text = "Y" Then
                    ' 跳躍
                    If uFASCommon.EDITransfer(uFASCommon.GetLastPlanVersion(DGRBuyer.Text), DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                        errcode = 9002
                    End If
                Else
                    If uFASCommon.EDITransfer(DLogID.Text, DBuyer.Text, xUserID, DGRBuyer.Text, DFunList.Text) <> 0 Then
                        errcode = 9002
                    End If
                End If
            End If
            ' 更新客戶控制檔(EPEDI=2, EDIReport=1)
            If errcode = 0 Then
                If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "EPEDI", 2, "EDIReport", 1) <> 0 Then
                    errcode = 9003
                End If
            End If
            ' EDI-Data轉換是否完成
            If errcode = 0 Then
                If uFASCommon.CheckControlRecord(DBuyer.Text, "EPEDI", 2) <> 0 Then
                    errcode = 9004
                End If
            End If
            ' 執行結果處理
            If errcode <> 0 Then
                SetTaskStatus("StsEPEDI", 1245, 1)
                StsLFLSPlan.NavigateUrl = "ActionHistory.aspx?pLogID=" + DLogID.Text + "&pBuyer=" + DBuyer.Text + "&pUserID=" + xUserID
                StsLFLSPlan.Target = "_blank"
                '
                If errcode = 9001 Then msg = "刪除執行履歷異常(EPEDI)，請連絡系統人員!"
                If errcode = 9002 Then msg = "EDI資料轉換異常，請確認(可點選[ｘ]查詢)!"
                If errcode = 9003 Then msg = "無法更新客戶控制檔(Flag)，請連絡系統人員!"
                If errcode = 9004 Then msg = "EDI資料轉換異常異常，請連絡系統人員!"
                If errcode = 9005 Then msg = "刪除 EDI Interface 資料異常，請連絡系統人員!"
                uJavaScript.PopMsg(Me, msg)
            Else
                SetTaskStatus("StsEPEDI", 1245, 0)
                SetTaskProgress("ProEPEDI", -500)
                SetTaskButton("BEPEDI", False)
                SetTaskButton("BReset", True)
                SetTaskButton("BEDIReport", True)

                '2017/8/24 ADIDAS 備料問題發生
                '2018/1/30 拿掉
                'Dim Cmd As String
                '
                'Cmd = "<script>" + _
                '        "window.open('FASBYCheckPlanQty.aspx?pBuyer=" & DBuyer.Text & "&pBuyerGroup=" & DGRBuyer.Text & "&pLogID=" & DLogID.Text & "&pUserID=" & xUserID & "&pFun=EDI" & "','CheckPlanQty','status=1,toolbar=0,top=100,left=100,width=900,height=750,resizable=yes,scrollbars=yes');" + _
                '      "</script>"
                '
                'Response.Write(Cmd)

            End If
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BEDIReport_Click)
    '**     EDI報表
    '**
    '*****************************************************************
    Protected Sub BEDIReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BEDIReport.Click
        '
        '*****************************************************************
        '**  準備報表資料
        '*****************************************************************
        '
        uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "Action2", 2, "EDIReport", 1)
        '判斷是否完成作業
        Dim Code As Integer
        Code = uFASCommon.CheckControlRecord(DBuyer.Text, "EDIReport", 2)
        If Code <> 0 Then
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "EDIReport", 3)
        End If
        Do Until Code = 0
            System.Threading.Thread.Sleep(3 * 1000)
            Code = uFASCommon.CheckControlRecord(DBuyer.Text, "EDIReport", 2)
            If Code <> 0 Then
                Code = uFASCommon.CheckControlRecord(DBuyer.Text, "EDIReport", 3)
            End If
        Loop
        '
        '檢查是否已產出報表
        If uFASCommon.CheckControlRecord(DBuyer.Text, "EDIReport", 2) = 0 Then
            SetTaskProgress("ProEDIReport", -500)
            SetTaskButton("BEDIReport", False)
            SetTaskButton("BReset", True)
            SetTaskStatus("StsEDIReport", 1245, 0)

            If uFASCommon.UpdateControlRecord(DBuyer.Text, xUserID, "End", 0, "End", 0) = 0 Then
                SetTaskButton("BGotoEDI", True)
            Else
                uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag-EndProcess)，請連絡系統人員!")
            End If
        Else
            SetTaskStatus("StsEDIReport", 1245, 1)
            uJavaScript.PopMsg(Me, "報表製作未完成，請確認!")
        End If
        '
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
                    "window.open('ActionHistory.aspx?pLogID=" & DLogID.Text & "&pBuyer=" & DBuyer.Text & "&pUserID=" & xUserID & "','ActionLog','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BBuyerFCT_Click)
    '**     BUYER FCT 功能
    '**
    '*****************************************************************
    Protected Sub BBuyerFCT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BBuyerFCT.Click
        Dim Cmd As String
        Dim xOffice2010 As Integer = 0
        If AtOffice2010.Checked = True Then
            xOffice2010 = 1
        End If
        '
        ' ADIDAS & REEBOK-PROC
        If DBuyer.Text = "FALL-000001" Or DBuyer.Text = "FALL-000016" Then
            Cmd = "<script>" + _
                        "window.open('FASBYFlow.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "&pOffice2010=" & CStr(xOffice2010) & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        ' NIKE-PROC
        If DBuyer.Text = "FALL-000013" Then
            Cmd = "<script>" + _
                        "window.open('NIKEBYFlow.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "&pOffice2010=" & CStr(xOffice2010) & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        ' TNF-PROC
        If DBuyer.Text = "FALL-000021" Then
            Cmd = "<script>" + _
                        "window.open('TNFBYFlow.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "&pOffice2010=" & CStr(xOffice2010) & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        ' COLUMBIA-PROC
        If DBuyer.Text = "FALL-000003" Then
            Cmd = "<script>" + _
                        "window.open('COLUMBIABYFlow.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "&pOffice2010=" & CStr(xOffice2010) & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        ' UNDERARMOUR & T&P NIKE PROC
        If DBuyer.Text = "FALL-TW0371" Or DBuyer.Text = "FALL-TW0371T" Or DBuyer.Text = "FALL-TP000013" Or DBuyer.Text = "FALL-TW1741" Or _
           DBuyer.Text = "FALL-TW0655" Or DBuyer.Text = "FALL-000053" Or DBuyer.Text = "FALL-000141" Then
            Cmd = "<script>" + _
                        "window.open('UNDERARMOURBYFlow.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "&pOffice2010=" & CStr(xOffice2010) & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BItemCheck_Click)
    '**     ITEM CHECK 功能
    '**
    '*****************************************************************
    Protected Sub BItemCheck_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BItemCheck.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('ItemCheck.aspx?pBuyer=" & DBuyer.Text & "&pUserID=" & xUserID & "','ItemCheck','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BGotoEDI_Click)
    '**     Go To EDI PROCESS
    '**
    '*****************************************************************
    Protected Sub BGotoEDI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BGotoEDI.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.location.href('EDIFlow.aspx?pUserID=" & xUserID & "');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BPLANR_Click)
    '**     FCT & LS BLS REPOPRT
    '**
    '*****************************************************************
    Protected Sub BPLANR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPLANR.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ACTR_Click)
    '**     FCT & ACT REPORT(ORDER YYYYMM)
    '**
    '*****************************************************************
    Protected Sub BFCT2ACTR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ACTR.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ACTMMR_Click)
    '**     FCT & ACT MONTH REPORT(ORDER YYYYMM)
    '**
    '*****************************************************************
    Protected Sub BFCT2ACTMMR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ACTMMR.Click

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ACTVDPR_Click)
    '**     FCT & ACT REPORT(VDP)
    '**
    '*****************************************************************
    Protected Sub BFCT2ACTVDPR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ACTVDPR.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BMAT2SEA_Click)
    '**     材料分析 REPORT(季)
    '**
    '*****************************************************************
    Protected Sub BMAT2SEA_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMAT2SEA.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BMAT2MON_Click)
    '**     材料分析 REPORT(月)
    '**
    '*****************************************************************
    Protected Sub BMAT2MON_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMAT2MON.Click
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ORD_Click)
    '**     FCT - ORDER 比較分析功能
    '**
    '*****************************************************************
    Protected Sub BFCT2ORD_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ORD.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_OrderProgress.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','FCT2ORD','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2FCT_Click)
    '**     FCT - FCT 比較分析功能
    '**
    '*****************************************************************
    Protected Sub BFCT2FCT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2FCT.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_FCTAnalysis.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','FCT2FCT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ACT_Click)
    '**     FCT - ACT 比較分析功能
    '**
    '*****************************************************************
    Protected Sub BFCT2ACT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ACT.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_FCTACTAnalysis.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','FCT2ACT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BFCT2ACTVDP_Click)
    '**     FCT - ACT 比較分析功能(VDP)
    '**
    '*****************************************************************
    Protected Sub BFCT2ACTVDP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BFCT2ACTVDP.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_FCTACTVDPAnalysis.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','FCT2ACTVDP','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSTOCK_Click)
    '**     STOCK 分析
    '**
    '*****************************************************************
    Protected Sub BSTOCK_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSTOCK.Click
        Dim Cmd As String
        '
        If InStr(DBuyer.Text, "TP000013") > 0 Then
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/EDI/InfF_StockProdFreeList.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','STOCK','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"

        Else
            Cmd = "<script>" + _
                        "window.open('http://10.245.1.6/EDI/InfF_StockProdList.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','STOCK','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
                  "</script>"
        End If
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BALERT_Click)
    '**     ALERT ACT - FCT 分析
    '**
    '*****************************************************************
    Protected Sub BALERT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BALERT.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_FCTACTAlertList.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','ALERT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BVDPERR_Click)
    '**     VDP ERROR INF.
    '**
    '*****************************************************************
    Protected Sub BVDPERR_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BVDPERR.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_VDPErrorList.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','VDPERR','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BVDPEDIT_Click)
    '**     VDP INF. EDIT 
    '**
    '*****************************************************************
    Protected Sub BVDPEDIT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BVDPEDIT.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('InfF_VDPErrorMaint.aspx?pBuyer=" & DGRBuyer.Text & "&pUserID=" & xUserID & "','VDPEDIT','status=1,toolbar=0,top=100,left=100,width=1000,height=500,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BRIGHT_Click)
    '**     分析功能右移 
    '**
    '*****************************************************************
    Protected Sub BRIGHT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRIGHT.Click
        BSTOCK.Style("left") = 254 & "px"
        BALERT.Style("left") = 339 & "px"
        BVDPERR.Style("left") = 424 & "px"
        BVDPEDIT.Style("left") = 509 & "px"

        SetTaskButton("BLEFT", True)
        SetTaskButton("BPLANR", False)
        SetTaskButton("BFCT2ACTR", False)
        SetTaskButton("BFCT2ACTMMR", False)
        SetTaskButton("BFCT2ACTVDPR", False)
        SetTaskButton("BMAT2SEA", False)
        SetTaskButton("BMAT2MON", False)
        SetTaskButton("BKPI", False)
        SetTaskButton("BFCT2ORD", False)
        SetTaskButton("BFCT2FCT", False)
        SetTaskButton("BFCT2ACT", False)
        SetTaskButton("BFCT2ACTVDP", False)
        SetTaskButton("BSTOCK", True)
        SetTaskButton("BALERT", True)
        SetTaskButton("BVDPERR", True)
        SetTaskButton("BVDPEDIT", True)
        SetTaskButton("BRIGHT", False)
    End Sub
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BLEFT_Click)
    '**     分析功能左移 
    '**
    '*****************************************************************
    Protected Sub BLEFT_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BLEFT.Click
        BPLANR.Style("left") = 254 & "px"
        BFCT2ACTR.Style("left") = 339 & "px"
        BFCT2ACTMMR.Style("left") = 424 & "px"
        BFCT2ACTVDPR.Style("left") = 509 & "px"
        BMAT2SEA.Style("left") = 594 & "px"
        BMAT2MON.Style("left") = 679 & "px"
        BKPI.Style("left") = 594 & "px"
        BFCT2ORD.Style("left") = 764 & "px"
        BFCT2FCT.Style("left") = 849 & "px"
        BFCT2ACT.Style("left") = 934 & "px"
        BFCT2ACTVDP.Style("left") = 1019 & "px"

        SetTaskButton("BLEFT", False)
        SetTaskButton("BPLANR", True)
        SetTaskButton("BFCT2ACTR", True)
        SetTaskButton("BFCT2ACTMMR", True)
        SetTaskButton("BFCT2ACTVDPR", True)
        SetTaskButton("BMAT2SEA", True)
        SetTaskButton("BMAT2MON", True)
        SetTaskButton("BKPI", True)
        SetTaskButton("BFCT2ORD", True)
        SetTaskButton("BFCT2FCT", True)
        SetTaskButton("BFCT2ACT", True)
        SetTaskButton("BFCT2ACTVDP", True)
        SetTaskButton("BSTOCK", False)
        SetTaskButton("BALERT", False)
        SetTaskButton("BVDPERR", False)
        SetTaskButton("BVDPEDIT", False)
        SetTaskButton("BRIGHT", True)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BKPI_Click)
    '**     e KPI System
    '**
    '*****************************************************************
    Protected Sub BKPI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BKPI.Click
        Dim Cmd As String
        '
        Cmd = "<script>" + _
                    "window.open('eKPIMenu.aspx?pCustomer=" & DBuyer.Text & "&pUserID=" & xUserID & "','BuyerFCT','status=1,toolbar=0,top=100,left=100,width=1100,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
End Class
