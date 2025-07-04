Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO



 


Partial Class RateList
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
    Dim wUserID As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ItemCount As Integer = 5 '預先定義欄位數量
    Dim message As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Me.IsPostBack Then
            SetParameter()          '設定共用參數
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
        wUserID = Request.QueryString("pUserID")      'UserID


        

        '幣別
        Dim SQL As String
        SQL = "  select* from ("
        SQL = SQL + "  select distinct currency+'('+country+'/'+currencycode+')' as  Data1,country from F_exchangerate  "
        SQL = SQL + "  )a order by country"

        Dim dtReferp4 As DataTable = uDataBase.GetDataTable(SQL)
        DCurrency.Items.Clear()
        DCurrency.Items.Add("全部(all)")
        Dim i As Integer

        For i = 0 To dtReferp4.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp4.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp4.Rows(i).Item("Data1")
            DCurrency.Items.Add(ListItem1)

        Next
        dtReferp4.Clear()


        '年

        SQL = "  select distinct Syear Data1 from F_exchangerate order by syear desc  "

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DSYear.Items.Clear()

        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data1")
            DSYear.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()


        '月

        SQL = "  select distinct SMonth Data1 from F_exchangerate order by  SMonth   "

        Dim dtReferp2 As DataTable = uDataBase.GetDataTable(SQL)
        DSMonth.Items.Clear()
        DSMonth.Items.Add("全部(all)")

        For i = 0 To dtReferp2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp2.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp2.Rows(i).Item("Data1")
            DSMonth.Items.Add(ListItem1)
        Next
        dtReferp2.Clear()


        '日

        SQL = "  select distinct SDay Data1 from F_exchangerate order by SDay   "

        Dim dtReferp3 As DataTable = uDataBase.GetDataTable(SQL)
        DSDay.Items.Clear()
        DSDay.Items.Add("全部(all)")

        For i = 0 To dtReferp3.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp3.Rows(i).Item("Data1")
            ListItem1.Value = dtReferp3.Rows(i).Item("Data1")
            DSDay.Items.Add(ListItem1)
        Next
        dtReferp3.Clear()



    End Sub

    Sub DataList()


 

        '年
        Dim Syear As String = ""
        If DSYear.SelectedValue <> "" Then
            Syear = " and Syear='" + DSYear.SelectedValue + "'"

        End If

        '月
        Dim SMonth As String = ""
        If DSMonth.SelectedValue <> "全部(all)" Then
            SMonth = " and Smonth = '" + DSMonth.SelectedValue + "'"
        End If

        '年
        Dim SDay As String = ""
        If DSDay.SelectedValue <> "全部(all)" Then
            SDay = " and SDay='" + DSDay.SelectedValue + "'"
        End If

        Dim Rate As String = ""
        '幣別
        If DCurrency.SelectedValue <> "全部(all)" Then
            ' Rate = " and   EndD >= convert(char,dateadd(year,-1,getdate()),111) and EndD <= convert(char,getdate(),111) "
            Rate = Rate + " and  Data1='" + DCurrency.SelectedValue + "'"
        End If


        Dim a As String
        a = SDay


        Dim SQL As String


        SQL = " select seqno ,country  ,currency  ,currencycode ,syear ,smonth ,sday ,averagerate ,StrD,EndD, Data1 from ("
        SQL = SQL + " select seqno ,country  ,currency  ,currencycode ,syear,smonth ,sday  ,averagerate,"
        SQL = SQL + " syear+'/'+smonth+'/'+substring(ltrim(sday),1,2) as StrD,syear+'/'+smonth+'/'+substring(ltrim(sday),4,2) as EndD,currency+'('+country+'/'+currencycode+')' as  Data1  from  F_exchangerate"
        SQL = SQL + " )a where 1=1 " + Syear + SMonth + SDay + Rate
        SQL = SQL + " order by StrD desc  "


        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            GridView1.DataSource = DBAdapter1
            GridView1.DataBind()

        End If


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
    Sub Excel()
        Response.AppendHeader("Content-Disposition", "attachment;filename=RateList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub



    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound


        e.Row.Cells(8).Visible = False
        e.Row.Cells(9).Visible = False
        e.Row.Cells(10).Visible = False


        e.Row.Cells(4).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Row.Cells(5).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
        e.Row.Cells(6).Attributes.Add("style", "vnd.ms-excel.numberformat:@")


        'End If


        '選擇列時變色 
        If e.Row.RowIndex > -1 Then
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='#FBE5D6'")
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c;")
        End If

    End Sub

  
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataList()

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Excel()
    End Sub

End Class

