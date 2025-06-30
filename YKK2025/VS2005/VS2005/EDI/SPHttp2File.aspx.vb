Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class SPHttp2File
    Inherits System.Web.UI.Page
    '
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pFun As String              'Fun
    Dim pSPNo As String             'SPNo
    Dim pItem As String             'Item
    Dim pColor As String            'Color
    Dim pKeep As String             'Keep
    Dim pSPCode As String           'SPCode
    Dim pUserID As String           'UserID

    Dim uCommon As New Utility.Common

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "SPHttp2File.aspx"    '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        pFun = Request.QueryString("pFun")                  'Fun
        pSPNo = Request.QueryString("pSPNo")                'SPNo
        pItem = Request.QueryString("pItem")                'Item
        pColor = Request.QueryString("pColor")              'Color
        pKeep = Request.QueryString("pKeep")                'Keep
        pSPCode = Request.QueryString("pSPCode")            'SPCode
        pUserID = Request.QueryString("pUserID")            'UserID
        '
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim Cmd As String
        Dim xFileName, xPathImport As String
        '
        If pFun = "SPNO" Then
            If pSPNo <> "" Then

                xFileName = "SP_" & pSPNo & "_PILFinal.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_PILFinal.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "FDM" Then
            If pItem <> "" Then

                xFileName = "SP_" & pItem & "-" & pColor & "-" & pKeep & "_BuyMonth.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_BuyMonth.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "PROD" Then
            If pItem <> "" Then

                xFileName = "SP_" & pItem & "-" & pColor & "-" & pKeep & "_ProductionStatus.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_ProductionStatus.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "WINGS" Then
            If pSPCode <> "" Then

                xFileName = "SP_" & pSPCode & "_WINGSRefresh.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_WINGSRefresh.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "PDFINAL" Then
            If pSPCode <> "" Then

                xFileName = "SP_" & pSPCode & "_PDFINAL.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_PDFINAL.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "KP" Then
            If pUserID <> "" Then
                '
                'SP_KP-WINGS_Check.xlsm --> SP_KP-WINGS_Check_MK043.xlsm
                xFileName = "SP_KP-WINGS_Check_" + UCase(pUserID) + ".xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_KP-WINGS_Check.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "PDFINALSPNO" Then
            If pSPNo <> "" Then

                xFileName = "SP_" & pSPNo & "_PDFINALSPNO.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_PDFINALSPNO.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "CHGFINALSPNO" Then
            If pSPNo <> "" Then

                xFileName = "SP_" & pSPNo & "_CHGFINALSPNO.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "SP_COMMON_CHGFINALSPNO.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If
            End If
        End If
        '
        If pFun = "SRSHORT" Then
            '
            If pUserID <> "" Then
                'IRW_NO-ID_SpecialShort6-2.xlsm --> IRW_yyyyMMdd-IT003_SpecialShort6-2.xlsm
                xFileName = "IRW_" & pUserID & "_SpecialShort6-2.xlsm"
                xPathImport = uCommon.GetAppSetting("DataPrepareFile") + xFileName
                If Not File.Exists(xPathImport) Then
                    File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "IRW_ID_SpecialShort6-2.xlsm", xPathImport)
                    System.Threading.Thread.Sleep(5 * 1000)
                End If
                '
                If File.Exists(xPathImport) Then
                    Cmd = "<script>" + _
                                "window.open('file://10.245.0.192/program$/EDI-2011/DataPrepare/" & xFileName & "' " & ",'',''); " + _
                          "</script>"
                    '
                    Response.Write(Cmd)
                End If

            End If
        End If
        '
        '限IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

    End Sub

End Class
