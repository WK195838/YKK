Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Partial Class EOESMenu
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
            main()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        '動作按鈕設定
        BTemplate.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟範本檔設定？" + "');if(!ok){return false;} else {CheckAttribute('Template')}"
        BItem.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟ITEM轉換設定？" + "');if(!ok){return false;} else {CheckAttribute('Item')}"
        BColor.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟COLOR轉換設定？" + "');if(!ok){return false;} else {CheckAttribute('Color')}"
    End Sub
    '---------------------------------------------------------------------------------------------------
    Sub main()
        DFileName.Text = "TemplateMaint-" + xUserID + ".xlsm"
        DFilePath1.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
        '
        If Not File.Exists(DFilePath1.Text) Then
            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "TemplateMaint.xlsm", DFilePath1.Text)
        End If

        DFileName.Text = "ItemMaint-" + xUserID + ".xlsm"
        DFilePath2.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
        '
        If Not File.Exists(DFilePath2.Text) Then
            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ItemMaint.xlsm", DFilePath2.Text)
        End If

        DFileName.Text = "ColorMaint-" + xUserID + ".xlsm"
        DFilePath3.Text = uCommon.GetAppSetting("DataPrepareFile") + DFileName.Text
        '
        If Not File.Exists(DFilePath3.Text) Then
            File.Copy(uCommon.GetAppSetting("DataPrepareFile") + "ColorMaint.xlsm", DFilePath3.Text)
        End If
    End Sub

End Class
