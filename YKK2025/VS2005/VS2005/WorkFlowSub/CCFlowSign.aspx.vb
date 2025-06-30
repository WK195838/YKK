Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class CCFlowSign
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wOPT As String
    Dim wUserID As String
    Dim wURL As String
    Dim wOPURL As String
    Dim wIP As String
    Dim NowDateTime As String       '現在日期時間
    '
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '
        '設定共用參數
        SetParameter()
        '
        If Not Me.IsPostBack Then
            URLRedirection()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        '
        wOPT = Request.QueryString("pOPT")
        wUserID = Request.QueryString("pUserID")
        wURL = Request.QueryString("pURL")
        wOPURL = Request.QueryString("pOPURL")
        wIP = Request.ServerVariables("REMOTE_ADDR")
        '
        '調整KEY
        '
        '
    End Sub

    '*****************************************************************
    '**
    '**     URLRedirection
    '**
    '*****************************************************************
    Sub URLRedirection()
        Dim sql As String
        '
        '--------------------------------------
        'URL解析
        Dim wFormNo, wDepo As String
        Dim wFormSno, wStep, wSeqNo As Integer
        Dim xUrl, str As String
        '
        If wOPT = "1" Then
            xUrl = wURL
        Else
            xUrl = wOPURL
        End If
        'FormNo
        str = Mid(xUrl, InStr(xUrl, "pFormNo=") + 8)
        wFormNo = Mid(str, 1, InStr(str, "@") - 1)
        'FormSno
        str = Mid(xUrl, InStr(xUrl, "pFormSno=") + 9)
        wFormSno = CDbl(Mid(str, 1, InStr(str, "@") - 1))
        'Step
        str = Mid(xUrl, InStr(xUrl, "pStep=") + 6)
        wStep = CDbl(Mid(str, 1, InStr(str, "@") - 1))
        'SeqNo
        str = Mid(xUrl, InStr(xUrl, "pSeqNo=") + 7)
        wSeqNo = CDbl(Mid(str, 1, InStr(str, "@") - 1))
        'Depo
        wDepo = oCommon.GetCalendarGroup(wUserID)
        '
        '--------------------------------------
        '流程處理
        oFlow.UpdateFlow(wFormNo, wFormSno, wStep, wSeqNo, wDepo, wUserID)
        '
        '--------------------------------------
        '儲存閱讀記錄
        sql = "Insert into L_CCFlowReadInf "
        sql &= "Select "
        sql &= "'" & wFormNo & "', "
        sql &= " " & CStr(wFormSno) & ", "
        sql &= " " & CStr(wStep) & ", "
        sql &= " " & CStr(wSeqNo) & ", "
        sql &= "'" & NowDateTime & "', "
        sql &= "'" & wUserID & "', "
        sql &= "'" & wIP & "', "
        sql &= "'" & NowDateTime & "' "
        sql &= " "
        uDataBase.ExecuteNonQuery(sql)
        '
        '--------------------------------------
        '轉址
        str = Replace(xUrl, "@", "&")
        Response.Redirect(str)
        '
        '--------------------------------------
        '關閉網頁
        'IE8 可以自行關閉網頁
        'Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
        'IE11
        'Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

    End Sub
End Class
