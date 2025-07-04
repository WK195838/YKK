Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System
Imports System.IO


Imports System.ComponentModel

Imports System.Text



 


Partial Class PNoCommList
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
    Dim pCust As String
    Dim pCode As String
    Dim pNo As String






    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "PNoCommList.aspx"
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
   
        pCust = Request.QueryString("pCust")
        pCode = Request.QueryString("pCode")
        pNo = Request.QueryString("pNo")
        ' pCust = "M52952"
        ' pCode = "6286198"
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


    Sub DataList()



        Dim i As Integer = 0
        Dim SQL As String = ""


        If pNo = "" Then
            SQL = " select * from("
            SQL = SQL + "   select   confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT  "
            SQL = SQL + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=0  "
            SQL = SQL + "  and cust ='" + pCust + "' and item ='" + pCode + "'"

            SQL = SQL + " UNION ALL"
            SQL = SQL + "  select   confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT  "
            SQL = SQL + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=1  "
            If pCode = "" Then
                SQL = SQL + " and cust ='" + pCust + "'"
            Else
                SQL = SQL + " and cust ='" + pCust + "' and item ='" + pCode + "'"

            End If

            SQL = SQL + " UNION ALL"
            SQL = SQL + " SELECT  confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT1   FROM ("
            SQL = SQL + " select  *,SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
            SQL = SQL + " WHERE  CHARINDEX('客訴單號',comment )>0"
            SQL = SQL + " and cust ='" + pCust + "' and item ='" + pCode + "'"
            SQL = SQL + " )A  "
            SQL = SQL + " )a"

        Else
            SQL = " select  CONVERT(CHAR(10),ORDDATE,112)  as 'ORD DATE',ORDERNO AS 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,ITEMNAME,COLOR,QTY,AMOUNT,COMMENT from f_nocommsheetdt "
            SQL = SQL + " where no = '" + pNo + "'"
        End If

        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()

        'SQL = " select ITEMNAME from F_nocommlist  "

        'Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        'GridView2.DataSource = DBAdapter2
        'GridView2.DataBind()


    End Sub

    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=NoComm_ist.xls")     '程式別不同
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
End Class

