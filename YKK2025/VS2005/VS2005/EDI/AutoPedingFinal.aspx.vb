Imports System.Data

Partial Class AutoPedingFinal
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
    '
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            BatchApproveProc()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     自動簽核處理
    '**
    '*****************************************************************
    Sub BatchApproveProc()
        Dim LogID As String = Now.ToString("yyyyMMddHHmmss")

        uFASCommon.SPAutoFinalPlan(LogID, "SHA-AD", "IT003", "000071", "XXXXXXXXXX")

        'IE8 可以自行關閉網頁
        Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
        'IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

    End Sub

End Class
