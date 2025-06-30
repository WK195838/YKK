Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading
Imports System.IO

Partial Class GESSMenu
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
        '
        DUserID.Style("left") = -500 & "px"
        '
        DFileName.Style("left") = -500 & "px"
        DFilePath1.Style("left") = -500 & "px"
        DFilePath2.Style("left") = -500 & "px"
        DFilePath3.Style("left") = -500 & "px"
        DFilePath4.Style("left") = -500 & "px"
        DFilePath5.Style("left") = -500 & "px"
        '
        '動作按鈕設定
        BConvert.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[EDI轉換]作業？" + "');if(!ok){return false;} else {CheckAttribute('GESSCONVERT')}"
        BModify.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[EDI修改]作業？" + "');if(!ok){return false;} else {CheckAttribute('GESSMODIFY')}"
        BMaintConvert.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[MST維護]作業？" + "');if(!ok){return false;} else {CheckAttribute('GESSMAINTMST')}"
        BInqHistory.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[查詢修改履歷]作業？" + "');if(!ok){return false;} else {CheckAttribute('GESSINQHISTORY')}"
        BOther.Attributes("onclick") = "var ok=window.confirm('" + "是否執行[其它修改]作業？" + "');if(!ok){return false;} else {CheckAttribute('GESSOTHER')}"
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
            Case "BStart"
                BStart.Visible = pShow
            Case "BConvert"
                BConvert.Visible = pShow
            Case "BModify"
                BModify.Visible = pShow
            Case "BMaintConvert"
                BMaintConvert.Visible = pShow
            Case "BInqHistory"
                BInqHistory.Visible = pShow
            Case "BOther"
                BOther.Visible = pShow
            Case Else
                BStart.Visible = True
                BConvert.Visible = False
                BModify.Visible = False
                BMaintConvert.Visible = False
                BInqHistory.Visible = False
                BOther.Visible = False
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BStart_Click)
    '**     
    '**
    '*****************************************************************
    Protected Sub BStart_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BStart.Click
        Dim sql As String
        Dim errcode As Integer = 0
        Dim msg As String
        '
        '  300	AUTHORITY-IT003
        sql = "Select Data From M_Referp "
        sql = sql & " Where Cat = '300' "
        sql = sql & "   And DKey = '" & "AUTHORITY-" & xUserID & "' "
        Dim dt_Referp As DataTable = uDataBase.GetDataTable(sql)
        If dt_Referp.Rows.Count > 0 Then
            '
            'USERID
            DUserID.Text = xUserID
            '
            'CONVERT / EDI轉換
            'MODIFY / EDI修改
            DFileName.Text = "GESS_ModDataPrepare_" & Trim(DUserID.Text) & ".xlsm"
            DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") & DFileName.Text
            DFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") & DFileName.Text
            If Not File.Exists(DFilePath1.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") & "GESS_ModDataPrepare.xlsm", DFilePath1.Text)
            End If
            '
            'MAINTMST / 維護MST
            DFileName.Text = "GESS_MaintConvertData_" & Trim(DUserID.Text) + ".xlsm"
            DFilePath3.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            If Not File.Exists(DFilePath3.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") & "GESS_MaintConvertData.xlsm", DFilePath3.Text)
            End If
            '
            'INQ / 查詢修改履歷
            DFileName.Text = "GESS_InqModHistory_" & Trim(DUserID.Text) + ".xlsm"
            DFilePath4.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
            If Not File.Exists(DFilePath4.Text) Then
                File.Copy(uCommon.GetAppSetting("DataPrepareFile") & "GESS_InqModHistory.xlsm", DFilePath4.Text)
            End If
            '
            'INQ / 查詢修改履歷
            DStart.Items.Clear()
            DStart.Items.Add(dt_Referp.Rows(0).Item("Data"))
            '
            '按鈕顯示或不顯示
            SetTaskButton("BStart", False)
            SetTaskButton("BConvert", True)
            SetTaskButton("BModify", True)
            SetTaskButton("BMaintConvert", True)
            SetTaskButton("BInqHistory", True)
            '
            If InStr(xUserID, "IT") > 0 Then
                '
                '其它修改
                DFileName.Text = "GESS_Mod_EDIData_" & Trim(DUserID.Text) + ".xlsm"
                DFilePath5.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
                If Not File.Exists(DFilePath5.Text) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") & "GESS_Mod_EDIData.xlsm", DFilePath5.Text)
                End If
                '
                '按鈕顯示或不顯示
                SetTaskButton("BOther", True)
            End If
        Else
            errcode = 9001
        End If
        '
        If errcode <> 0 Then
            If errcode = 9001 Then msg = "無權限可以使用，請確認!"
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub


End Class