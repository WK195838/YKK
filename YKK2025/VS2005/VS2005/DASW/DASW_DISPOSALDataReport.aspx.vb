Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO



 


Partial Class DASW_DISPOSALDataReport
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
       

        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            CheckAuthority()
            GetDisposalYM()
            'DataList()

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

    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim srule As String = ""
        If DDISPOSALRULE.SelectedValue <> "" Then
            srule = " and a.DISPOSALRULE ='" + DDISPOSALRULE.SelectedValue + "'"
        End If


        Dim appname As String = ""
        If DAppName.SelectedValue <> "" Then
            appname = " and a.appname = '" + DAppName.SelectedValue + "'"
        End If



        Dim SQL As String


        SQL = "  Select  sMonth '月份',a.no '單號',convert(char(10),a.appdate,111) '申請日期',appdepo '申請部門',a.appname '申請人',"
        SQL = SQL + " CODE,itemname1 as [ITEM NAME 1],itemname2 as [ITEM NAME 2], Case when LENGTH = '' then '0' else LENGTH END AS LENGTH,U,COLOR,"
        SQL = SQL + " LOCATION, ACTUAL, FREE, SIZE, CHAIN, CLS, UNIT,unitweight AS [UNIT WEIGHT],WEIGHTKG AS [WEIGHT(KG)], "
        SQL = SQL + "a.disposalrule AS '報廢準則',pnstock AS [PN 倉位],costa AS [COST A], costb AS [COST B],"
        SQL = SQL + " actualamount AS [ACTUAL AMOUNT], UT2, LASTIN, LASTOUT, a.deponame AS '部門',"
        SQL = SQL + " a.sales AS  '營業擔當', a.disposalreason AS '原因', oneyear AS '１年使用量',BUYER,BUYERNAME,STYPE AS 報廢形式 "
        SQL = SQL + " from f_disposaldata a,f_disposalsheet b where a.no =b.no and smonth  between '" & DDisposalYM1.SelectedValue & "'" & " and '" & DDisposalYM2.SelectedValue & "'" + srule + appname


        'SQL = " select smonth '月份',a.no '單號',convert(char(10),appdate,111) '申請日期',appdepo '申請部門',appname '申請人',code as CODE,itemname1 as [ITEM NAME 1],itemname2 as [ITEM NAME 2],"
        'SQL = SQL + " Case when LENGTH = '' then '0' else LENGTH END AS LENGTH,"
        'SQL = SQL + " U,COLOR,"
        'SQL = SQL + " LOCATION, ACTUAL, FREE, SIZE, CHAIN, CLS, UNIT,unitweight AS [UNIT WEIGHT],WEIGHTKG AS [WEIGHT(KG)], disposalrule AS '報廢準則', pnstock AS [PN 倉位],costa AS [COST A], costb AS [COST B],"
        'SQL = SQL + " actualamount AS [ACTUAL AMOUNT], UT2, LASTIN, LASTOUT, deponame AS '部門', sales AS  '營業擔當', disposalreason AS '原因', oneyear AS '１年使用量'"
        'SQL = SQL + "  from f_disposalData "
        'SQL = SQL + "  a,( SELECT no from F_disposalsheet where sts <>2) b  where a.no=b.no  "
        'SQL = SQL + "  and  smonth between '" + DDisposalYM1.SelectedValue + "' and '" + DDisposalYM2.SelectedValue + "'" + srule + appname
        'SQL = SQL + " order by smonth,a.no,seqno"

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()

    End Sub

    Sub DataList1()
        Dim srule As String = ""
        If DDISPOSALRULE.SelectedValue <> "" Then
            srule = " and DISPOSALRULE ='" + DDISPOSALRULE.SelectedValue + "'"
        End If


        Dim appname As String = ""
        If DAppName.SelectedValue <> "" Then
            appname = " and appname = '" + DAppName.SelectedValue + "'"
        End If



        Dim SQL As String


        SQL = "  Select sMonth '月份',a.no '單號',convert(char(10),a.appdate,111) '申請日期',appdepo '申請部門',a.appname '申請人',"
        SQL = SQL + " CODE,itemname1 as [ITEM NAME 1],itemname2 as [ITEM NAME 2], Case when LENGTH = '' then '0' else LENGTH END AS LENGTH,U,COLOR,"
        SQL = SQL + " LOCATION, ACTUAL, FREE, SIZE, CHAIN, CLS, UNIT,unitweight AS [UNIT WEIGHT],WEIGHTKG AS [WEIGHT(KG)], "
        SQL = SQL + "a.disposalrule AS '報廢準則',pnstock AS [PN 倉位],costa AS [COST A], costb AS [COST B],"
        SQL = SQL + " actualamount AS [ACTUAL AMOUNT], UT2, LASTIN, LASTOUT, a.deponame AS '部門',"
        SQL = SQL + " a.sales AS  '營業擔當', a.disposalreason AS '原因', oneyear AS '１年使用量',BUYER,BUYERNAME,STYPE AS 報廢形式 "
        SQL = SQL + " from f_disposaldata a,f_disposalsheet b where a.no =b.no and smonth  between '" & DDisposalYM1.SelectedValue & "'" & " and '" & DDisposalYM2.SelectedValue & "'" + srule + appname

        'SQL = " select smonth '月份',a.no '單號',convert(char(10),appdate,111) '申請日期',appdepo '申請部門',appname '申請人',code as CODE,itemname1 as [ITEM NAME 1],itemname2 as [ITEM NAME 2],"
        'SQL = SQL + " Case when LENGTH = '' then '0' else LENGTH END AS LENGTH,"
        'SQL = SQL + " U,COLOR,"
        'SQL = SQL + " LOCATION, ACTUAL, FREE, SIZE, CHAIN, CLS, UNIT,unitweight AS [UNIT WEIGHT],WEIGHTKG AS [WEIGHT(KG)], disposalrule AS '報廢準則', pnstock AS [PN 倉位],costa AS [COST A], costb AS [COST B],"
        'SQL = SQL + " actualamount AS [ACTUAL AMOUNT], UT2, LASTIN, LASTOUT, deponame AS '部門', sales AS  '營業擔當', disposalreason AS '原因', oneyear AS '１年使用量'"
        'SQL = SQL + "  from f_disposalData "
        'SQL = SQL + "  a,( SELECT no from F_disposalsheet where sts <>2) b  where a.no=b.no  "
        'SQL = SQL + "  and  smonth between '" + DDisposalYM1.SelectedValue + "' and '" + DDisposalYM2.SelectedValue + "'" + srule + appname
        'SQL = SQL + " order by smonth,a.no,seqno"



        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)

        GridView2.DataSource = DBAdapter2
        GridView2.DataBind()


    End Sub




    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        DataList()
        Excel()
    End Sub

    Sub GetDisposalYM()
        Dim SQL As String
        Dim i As Integer
        SQL = " select distinct smonth from F_DISPOSALData where smonth>='2018/12' order by smonth desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DDisposalYM1.Items.Clear()
        DDisposalYM2.Items.Clear()
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("smonth")
            ListItem1.Value = DBAdapter1.Rows(i).Item("smonth")
            DDisposalYM1.Items.Add(ListItem1)
            DDisposalYM2.Items.Add(ListItem1)
        Next

        DDisposalYM1.SelectedValue = Now.ToString("yyyy/MM")
        DDisposalYM2.SelectedValue = Now.ToString("yyyy/MM")

        DBAdapter1.Clear()

        '報廢準則

        SQL = "  select data from M_referp  "
        SQL = SQL & " where cat =6001 and dkey ='RULE'  order by unique_id "

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALRULE.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("data")
            ListItem1.Value = dtReferp.Rows(i).Item("data")
            DDISPOSALRULE.Items.Add(ListItem1)
        Next
        dtReferp.Clear()

        '報廢準則

        SQL = " select  distinct appname  from F_DISPOSALData order by appname    "

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DAppName.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("appname")
            ListItem1.Value = dtReferp1.Rows(i).Item("appname")
            DAppName.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()


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
        Response.AppendHeader("Content-Disposition", "attachment;filename=DisposalDataReport_Division.xls")     '程式別不同
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

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        DataList1()
        Excel()
        'Response.AppendHeader("Content-Disposition", "attachment;filename=DisposalDataReport_Division.xls")     '程式別不同
        'Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        ''Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        'Response.ContentType = "application/vnd.ms-excel"
        'Response.Charset = ""
        'Me.EnableViewState = False

        'Dim tw As New System.IO.StringWriter
        'Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        'GridView2.RenderControl(hw)
        'Response.Write(tw.ToString())
        'Response.End()
    
    End Sub

 
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound



        e.Row.Cells(0).Visible = False
        e.Row.Cells(1).Visible = True
        'Dim mydate As DateTime = Convert.ToDateTime(e.Row.Cells(2).ToString)

        Dim i As Integer
        For i = 0 To 33

            If i <> 8 And i <> 9 And i <> 12 And i <> 13 And i <> 14 And i <> 18 And i <> 19 And i <> 22 And i <> 23 And i <> 24 And i <> 31 Then
                e.Row.Cells(i).Attributes.Add("style", "vnd.ms-excel.numberformat:@")
            End If




        Next


        ' e.Row.Cells(2).Text = mydate.ToString("yyyy/MM/dd")
    End Sub



End Class

