Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

Partial Class CloseSystem
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim uJavaScript As New Utility.JScript
    Dim uEDICommon As New EDI2011.CommonService
    Dim NowDateTime As String               ' 現在日時
    Dim xBuyer As String                    ' Buyer
    Dim xUserID As String                   ' 使用者ID
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        '
        If Not IsPostBack Then
            If uEDICommon.UpdateControlRecord(xBuyer, xUserID, "Reset", 0, "Reset", 0) = 0 Then
            Else
                uJavaScript.PopMsg(Me, "無法更新客戶控制檔(Flag)，請連絡系統人員!")
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xBuyer = Request.QueryString("pBuyer")
        xUserID = Request.QueryString("pUserID")
    End Sub

End Class
