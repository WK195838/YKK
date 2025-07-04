Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class EApprovalRDScheduleAll
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
    Dim pFormNo As String
    Dim wStep, wDelay, wDepo As String
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()          '設定共用參數
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
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        wStep = Request.QueryString("pStep")
        wDelay = Request.QueryString("pDelay")
        wDepo = Request.QueryString("pDepo")
    End Sub



    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs)







    End Sub


    Sub DataList()
        Dim StepStr As String
        StepStr = " and step = '" + wStep + "'"
        Dim DelayStr As String
        If wDelay <> "" Then
            DelayStr = " where delaycase = '" + wDelay + "'"
        Else
            DelayStr = ""
        End If

        Dim sql As String
        sql = "  Select * from ( SELECT"
        sql = sql + " FormNo, FormSno, Step, SeqNo,Division,No,substring(stepnamedesc,6,len(stepnamedesc)-1)stepnamedesc ,"
        sql = sql + " StsDesc, FormName, FlowTypeDesc, ApplyID, ApplyName, AgentName,"
        sql = sql + "'申請時間：[' + Convert(VarChar, ApplyTime, 20) + '], '+"
        sql = sql + "'收件時間：[' + Convert(VarChar, ReceiptTime, 20) + '], ' +"
        sql = sql + "'閱讀期限：[' + Convert(VarChar, ReadTimeLimit, 20) + '], ' +"
        sql = sql + "'首次閱讀：[' + FirstReadTimeDesc + '], ' +"
        sql = sql + "'最後閱讀：[' + LastReadTimeDesc  + '], ' +"
        sql = sql + "'預定開始：[' + Convert(VarChar, BStartTime, 20) + '], ' +"
        sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '], ' +"
        sql = sql + "'實際開始：[' + Convert(VarChar, AStartTime, 20) + '] ' +"
        sql = sql + "'預定完成：[' + Convert(VarChar, BEndTime, 20) + '] '  "
        sql = sql + " As Description, URL ,Delaysts "
        sql = sql + " ,case when delaysts ='正常' then  0 else 1 end as delaycase "
        sql = sql + " FROM V_WaitHandle_01"
        sql = sql + " Where Active = '1'"
        sql = sql + " and flowtype =1 and formno ='007003'  " + StepStr
        sql = sql + " )a " + DelayStr


        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim h1 As New HyperLink


            Dim wFormsno As String
            wFormsno = e.Row.Cells(9).Text
            h1.Text = e.Row.Cells(2).Text
            ' 連結到待處理LIST
            ' h1.NavigateUrl = "http://10.245.1.10/N2W/ComplaintOutScheduleAll.aspx?&pStep=" + wStep
            'h1.NavigateUrl = "FundingSheet_02.aspx?&pFormno=003110&pFormsno=" + wFormsno
            If e.Row.Cells(10).Text = Request.QueryString("pUserID") Then
                h1.NavigateUrl = "EApprovalRDSheet_02.aspx?&pFormno=007003&pFormsno=" + wFormsno

            End If

            e.Row.Cells(2).Controls.Add(h1)

        End If



        e.Row.Cells(9).Visible = False
        e.Row.Cells(10).Visible = False

        '連結
    End Sub

    Protected Sub BExcel_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=EApprovalRDScheduleAll.xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同


    End Sub


End Class

