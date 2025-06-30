Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Partial Class OPPIMenu
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
        uEDIMapping.Timeout = 900 * 1000
        uEDICommon.Timeout = 900 * 1000
        uWFSCommon.Timeout = 900 * 1000
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        DUserID.Style("left") = -500 & "px"
        DUserID.Text = UCase(Request.QueryString("pUserID"))
        DFileName.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        DFilePath2.Style("left") = -500 & "px"
        DFilePath3.Style("left") = -500 & "px"
        '動作按鈕設定
        BMCustomer.Attributes("onclick") = "var ok=window.confirm('" + "是否執行維護顧客？" + "');if(!ok){return false;} else {CheckAttribute('CUSTOMER')}"
        BMCustomerHistory.Attributes("onclick") = "var ok=window.confirm('" + "是否執行查詢顧客履歷？" + "');if(!ok){return false;}"
        BMSystem.Attributes("onclick") = "var ok=window.confirm('" + "是否執行查詢系統參數？" + "');if(!ok){return false;}"
        BInqReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行調閱報表？" + "');if(!ok){return false;}"
        BInqSendOffMail.Attributes("onclick") = "var ok=window.confirm('" + "是否執行調閱郵件？" + "');if(!ok){return false;}"
        BPIReport.Attributes("onclick") = "var ok=window.confirm('" + "是否執行PI報表(未PACKCP)？" + "');if(!ok){return false;} else {CheckAttribute('PIREPORT')}"
        BOPReportSys.Attributes("onclick") = "var ok=window.confirm('" + "是否執行OP報表？" + "');if(!ok){return false;} else {CheckAttribute('OPREPORT')}"

        BSpoolStore.Attributes("onclick") = "var ok=window.confirm('" + "是否連結 [YKK Spool Store System]？" + "');if(!ok){return false;}"
        '未使用
        BRigisterSheet.Style("left") = -500 & "px"
        BRigisterSheet.Attributes("onclick") = "var ok=window.confirm('" + "是否下載顧客登錄單？" + "');if(!ok){return false;}"
        '
        BChangeSheet.Style("left") = -500 & "px"
        BChangeSheet.Attributes("onclick") = "var ok=window.confirm('" + "是否下載顧客變更單？" + "');if(!ok){return false;}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetExcelPgm)
    '**     設定各Excel
    '**
    '*****************************************************************
    Sub SetExcelPgm()
        If DUserID.Text <> "" Then
            '
            '設定顧客主檔維護 
            DFileName.Text = DUserID.Text + "_CustomerMasterMaint.xlsm"
            DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            '
            '設定PI報表 
            DFileName.Text = DUserID.Text + "_PackingInformation.xlsm"
            DFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            '
            '設定OP報表 
            DFileName.Text = DUserID.Text + "_OPReportSys.xlsm"
            DFilePath3.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            ''
            ''複製 顧客變更單 
            'DFileName.Text = DUserID.Text + "_CustomerChangeSheet.xlsx"
            'DFilePath3.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            'If Not File.Exists(DFilePath3.Text) Then
            '    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "CustomerChangeSheet.xlsx", DFilePath3.Text)
            'End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BMCustomer_Click)
    '**     維護顧客MST
    '**
    '*****************************************************************
    Protected Sub BMCustomer_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMCustomer.Click
        If DUserID.Text <> "" Then
            '
            '複製 顧客主檔維護 
            If Not File.Exists(DFilePath1.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "CustomerMasterMaint.xlsm", DFilePath1.Text)
                System.Threading.Thread.Sleep(3 * 1000)
            End If
            '
            '開啟 顧客主檔維護
            Dim Code As Integer
            Code = 999
            Do Until Code = 0
                If File.Exists(DFilePath1.Text) Then
                    Code = 0
                Else
                    System.Threading.Thread.Sleep(3 * 1000)
                End If
            Loop
            If Code = 0 Then
            Else
                uJavaScript.PopMsg(Me, "顧客MST維護程式不存在，請確認!")
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BMCustomerHistory_Click)
    '**     顧客履歷
    '**
    '*****************************************************************
    Protected Sub BMCustomerHistory_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMCustomerHistory.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('INQ_CustomerGroup.aspx?pUserID=" + DUserID.Text & "','顧客履歷','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BPIReport_Click)
    '**     PI報表
    '**
    '*****************************************************************
    Protected Sub BPIReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BPIReport.Click
        If DUserID.Text <> "" Then
            '
            '複製 PI報表 
            If Not File.Exists(DFilePath2.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "PackingInformation.xlsm", DFilePath2.Text)
                System.Threading.Thread.Sleep(3 * 1000)
            End If
            '
            '開啟 PI報表 
            Dim Code As Integer
            Code = 999
            Do Until Code = 0
                If File.Exists(DFilePath2.Text) Then
                    Code = 0
                Else
                    System.Threading.Thread.Sleep(3 * 1000)
                End If
            Loop
            If Code = 0 Then
            Else
                uJavaScript.PopMsg(Me, "製作PI報表(未PACKCP)程式不存在，請確認!")
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BOPREPORT_Click)
    '**     OP報表
    '**
    '*****************************************************************
    Protected Sub BOPReportSys_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BOPReportSys.Click
        If DUserID.Text <> "" Then
            '
            '複製 OP報表 
            If Not File.Exists(DFilePath3.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "OPReportSys.xlsm", DFilePath3.Text)
                System.Threading.Thread.Sleep(3 * 1000)
            End If
            '
            '開啟 OP報表 
            Dim Code As Integer
            Code = 999
            Do Until Code = 0
                If File.Exists(DFilePath3.Text) Then
                    Code = 0
                Else
                    System.Threading.Thread.Sleep(3 * 1000)
                End If
            Loop
            If Code = 0 Then
            Else
                uJavaScript.PopMsg(Me, "製作OP報表程式不存在，請確認!")
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BMSystem_Click)
    '**     系統參數
    '**
    '*****************************************************************
    Protected Sub BMSystem_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BMSystem.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('INQ_SystemConfig.aspx?pUserID=" + DUserID.Text & "','系統參數','status=1,toolbar=0,width=600,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BInqReport_Click)
    '**     調閱報表
    '**
    '*****************************************************************
    Protected Sub BInqReport_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BInqReport.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('INQ_ReportCustomerGroup.aspx?pUserID=" + DUserID.Text & "','調閱報表','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BInqSendOffMail_Click)
    '**     調閱郵件
    '**
    '*****************************************************************
    Protected Sub BInqSendOffMail_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BInqSendOffMail.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('INQ_MailCustomerGroup.aspx?pUserID=" + DUserID.Text & "','調閱郵件','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub

    Protected Sub BRigisterSheet_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRigisterSheet.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.3/edLink/TempFiles/顧客登錄單.xlsx','登錄單','status=1,toolbar=0,width=1200,height=600,resizable=yes,scrollbars=yes');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub

    Protected Sub BSpoolStore_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSpoolStore.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.29/SpoolStore/SpoolMgnt/ReportQuery.aspx?SystemCode=YKK&ProgId=SpoolMgnt&Side=1&Top=1');" + _
              "</script>"
        '
        Response.Write(Cmd)
    End Sub

End Class
