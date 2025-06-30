Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO

Partial Class eKPIMenu
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
    Dim xCustomer As String                 ' 客戶名稱
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        If Not IsPostBack Then
            SetExcelPgm()                               ' 設定各Excel
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uFASMapping.Timeout = 900 * 1000
        uFASCommon.Timeout = 900 * 1000
        uWFSCommon.Timeout = 900 * 1000
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        xCustomer = UCase(Request.QueryString("pCustomer"))
        '
        DLogID.Style("left") = -500 & "px"
        DBuyer.Style("left") = -500 & "px"
        DUserID.Style("left") = -500 & "px"
        DGRBuyer.Style("left") = -500 & "px"
        DFunList.Style("left") = -500 & "px"
        '
        DFileName.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        DFilePath2.Style("left") = -500 & "px"
        DFilePath3.Style("left") = -500 & "px"
        DFilePath4.Style("left") = -500 & "px"
        DFilePath5.Style("left") = -500 & "px"
        DFilePath6.Style("left") = -500 & "px"
        DFilePath7.Style("left") = -500 & "px"
        '
        ProDecision.Style("left") = -500 & "px"
        '動作按鈕設定
        BMSTMaint.Attributes("onclick") = "var ok=window.confirm('" + "是否維護MST作業？" + "');if(!ok){return false;} else {CheckAttribute('KPIMSTM')}"
        BMSTExpand.Attributes("onclick") = "var ok=window.confirm('" + "是否展開MST？" + "');if(!ok){return false;} else {CheckAttribute('KPIMSTE')}"
        BOFCT.Attributes("onclick") = "var ok=window.confirm('" + "是否上傳Overflow FCT？" + "');if(!ok){return false;} else {CheckAttribute('KPIOFCT')}"
        BOFCTCAL.Attributes("onclick") = "var ok=window.confirm('" + "是否計算Overflow FCT？" + "');if(!ok){return false;} else {CheckAttribute('KPIOFCTCAL')}"
        BDecision.Attributes("onclick") = "var ok=window.confirm('" + "是否執行判定[Delay Reason]？" + "');if(!ok){return false;} else {CheckAttribute('KPIDELAYREASON')}"
        BFINAL.Attributes("onclick") = "var ok=window.confirm('" + "是否最終決定？" + "');if(!ok){return false;} else {CheckAttribute('KPIFINAL')}"
        BREPORT.Attributes("onclick") = "var ok=window.confirm('" + "是否列表？" + "');if(!ok){return false;} else {CheckAttribute('KPIREPORT')}"
        BADCHECK.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟[AD CHECK TOOL]？" + "');if(!ok){return false;} else {CheckAttribute('KPIADCHECK')}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetExcelPgm)
    '**     設定各Excel
    '**
    '*****************************************************************
    Sub SetExcelPgm()
        '
        If xUserID <> "" Then
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
                DBuyer.Text = dt_ControlRecord.Rows(0).Item("Buyer")
                DUserID.Text = xUserID
                DGRBuyer.Text = dt_ControlRecord.Rows(0).Item("BuyerGroup")
                DFunList.Text = dt_ControlRecord.Rows(0).Item("FunList")
                '
                'MST
                DFileName.Text = "eKPI_Master.xlsm"
                DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'MST展開 
                DFileName.Text = "eKPI_DataPrepare.xlsm"
                DFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'OVERFLOW FCT 
                DFileName.Text = "eKPI_OFCTDataPrepare.xlsm"
                DFilePath3.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'FINAL
                DFileName.Text = "eKPI_DataLastFinal.xlsm"
                DFilePath4.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'REPORT
                DFileName.Text = "eKPI_ReportBuilder.xlsm"
                DFilePath5.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'AD CHECK TOOL
                DFileName.Text = "eKPI_ADChecking.xlsm"
                DFilePath6.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                '
                'OVERFLOW FCT(CAL) 
                DFileName.Text = "eKPI_OFCTCALDataPrepare.xlsm"
                DFilePath7.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            Else
                uJavaScript.PopMsg(Me, "此客戶不存在，請確認!")
            End If
        Else
            uJavaScript.PopMsg(Me, "無使用權限，請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BDecision_Click)
    '**     判定Delay Reason
    '**
    '*****************************************************************
    Protected Sub BDecision_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BDecision.Click
        Dim msg As String = ""
        Dim errcode As Integer = 0

        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        '*****************************************************************
        '**  判定Delay Reason
        '*****************************************************************
        '
        If errcode = 0 Then
            If uFASCommon.KPIExpansion(DLogID.Text, DBuyer.Text, DUserID.Text, DGRBuyer.Text, DFunList.Text) <> 0 Then
                errcode = 9001
            End If
        End If
        '
        ' 執行結果處理
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "Delay Reason 判定異常，請連絡系統人員!"
        Else
            msg = "Delay Reason 判定處理完成!"
        End If
        ProDecision.Style("left") = -500 & "px"
        uJavaScript.PopMsg(Me, msg)
        '
    End Sub
End Class
