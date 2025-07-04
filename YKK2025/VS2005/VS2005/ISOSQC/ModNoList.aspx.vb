Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class ModNoList
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
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim pModNo As String





    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        pModNo = Request.QueryString("pModNo")
        SetParameter()          '設定共用參數
        DataList()

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
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")




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

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click


        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=BusinessTripCommission_ist.xls")     '程式別不同
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
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


    Sub DataList()
        Dim SQL As String
      
        SQL = "  select no as 單號,completedtime 日期,name 修改人,formno,formsno,ModReason AS 修改理由,ModContent AS 修改內容,CASE WHEN sts = 0 THEN '核定中' WHEN sts = 1 THEN '完成' ELSE '取消' END AS 狀態 from F_QAModSheet "
        SQL = SQL + " where modno ='" + pModNo + "' order by 日期 desc "

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)
        GridView1.PageIndex = e.NewPageIndex
        DataList()
    End Sub



    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound



        If (e.Row.RowType = DataControlRowType.DataRow) <> (e.Row.RowType = DataControlRowType.Footer) Then
            Dim h1 As New HyperLink
            Dim formno As String
            Dim formsno As String
            h1.Target = "_blank"
            h1.Text = e.Row.Cells(0).Text
            formno = e.Row.Cells(3).Text
            formsno = e.Row.Cells(4).Text
            ' 連結到客訴表單   
            h1.NavigateUrl = "QAModSheet_02.aspx?pFormNo=" + formno + "&pFormSno=" + formsno
            ' e.Row.Cells(3).Text = ""
            e.Row.Cells(0).Controls.Add(h1)
        End If



        e.Row.Cells(3).Visible = False
        e.Row.Cells(4).Visible = False




    End Sub
End Class

