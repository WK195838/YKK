Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class CCFlowList
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
    Dim wFormNo As String
    Dim wFormSno As Integer
    Dim wStep As Integer
    Dim wFlowType As Integer
    Dim wUserID As String
    Dim NowDateTime As String       '現在日期時間
    '
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
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
            DataList()
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
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")
        wStep = Request.QueryString("pStep")
        wFlowType = Request.QueryString("pFlowType")
        wUserID = Request.QueryString("pUserID")
        '
    End Sub
    '*****************************************************************
    '**
    '**     DataList
    '**
    '*****************************************************************
    Sub DataList()
        Dim sql As String
        '
        '取得DATA
        sql = "SELECT"
        sql = sql + " FormNo, FormSno, Step, SeqNo,Division,No,substring(stepnamedesc,9,len(stepnamedesc)-1)stepnamedesc ,"
        sql = sql + " StsDesc, FormName, FlowTypeDesc, ApplyName, AgentName,"
        sql = sql + " '履歷' AS OP, URL, "
        sql = sql + "'http://10.245.1.10/WorkFlow/BefOPList.aspx?' + "
        sql = sql + "'pFormNo='   + FormNo + "
        sql = sql + "'&pFormSno=' + str(FormSno,Len(FormSno)) + "
        sql = sql + "'&pStep='    + str(Step,Len(Step)) + "
        sql = sql + "'&pSeqNo='   + str(SeqNo,Len(SeqNo)) + "
        sql = sql + "'&pApplyID=' + ApplyID "
        sql = sql + " As OPURL "
        sql = sql + " FROM V_WaitHandle_01"
        '
        sql = sql + " Where FormNo = '" & wFormNo & "' "
        sql = sql + " And FormSno = " & wFormSno & " "
        sql = sql + " And Step = " & wStep & " "
        sql = sql + " And FlowType = " & wFlowType & " "
        sql = sql + " And DecideID = '" & wUserID & "' "
        '
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then

            Dim h1, h2 As New HyperLink
            Dim xURL, xOPURL As String
            xURL = Replace(e.Row.Cells(5).Text, "&amp;", "&")
            xOPURL = Replace(e.Row.Cells(6).Text, "&amp;", "&")
            '
            xURL = Replace(xURL, "&", "@")
            xOPURL = Replace(xOPURL, "&", "@")
            '
            h1.Text = e.Row.Cells(1).Text
            h1.NavigateUrl = "http://10.245.1.6/WorkflowSub/CCFlowSign.aspx?pOPT=1" & "&pUserID=" & wUserID & "&pURL=" & xURL
            e.Row.Cells(1).Text = ""
            e.Row.Cells(1).Controls.Add(h1)
            '
            h2.Text = e.Row.Cells(4).Text
            h2.NavigateUrl = "http://10.245.1.6/WorkflowSub/CCFlowSign.aspx?pOPT=2" & "&pUserID=" & wUserID & "&pOPURL=" & xOPURL
            e.Row.Cells(4).Text = ""
            e.Row.Cells(4).Controls.Add(h2)
        End If
        '隱藏
        e.Row.Cells(5).Visible = False
        e.Row.Cells(6).Visible = False
    End Sub
    '
End Class
