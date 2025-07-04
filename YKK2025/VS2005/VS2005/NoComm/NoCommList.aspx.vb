Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System
Imports System.IO
Imports System.ComponentModel
Imports System.Text


Partial Class NoCommList
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
    Dim pUser As String
    Dim sales As String
    Dim cust As String
    Dim Item As String
    Dim Supplier As String
    Dim Qty As Integer
    Dim Comment As String




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "NoCommList.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetFieldData()
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
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
        pUser = Request.QueryString("pUserID")
        ' pUser = "it013"
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






    End Sub


    Sub DataList()



        Dim i As Integer = 0
        Dim SQL As String = ""

        'SQL = " select  sales,name,cust,custname,Item,itemname1,sum(qty)qty,sum(amount)amount,'' as Supplier,'' as COMMENT "
        'SQL = SQL + "  from F_nocommlist"
        'SQL = SQL + " Group by          sales, name, cust, custname, Item, itemname1  "
        SQL = " SELECT * FROM ("
        SQL = SQL + " SELECT A.*,USERID FROM ("
        SQL = SQL + " SELECT ROW_NUMBER() OVER (ORDER BY sales ASC) as Unique_ID,sales,name,cust,custname,CASE WHEN ITEM ='' THEN '連結' ELSE ITEM END AS ITEM,itemname1,supplier,sample,no_comm,COMMENT,QTY,AMOUNT,"
        SQL = SQL + " 'PNoCommList.aspx?pCust='+CUST+'&pcode='+ITEM  as ViewITEM,ITEM AS ITEM1 "
        SQL = SQL + " from ("
        SQL = SQL + "   select  sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,'' AS COMMENT,sum(qty)qty,sum(amount)amount "
        SQL = SQL + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0  and sample <>1 and EXCLUDE =0"

        SQL = SQL + " Group by          sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm "
        SQL = SQL + " UNION ALL "
        SQL = SQL + "  select sales,name,cust,custname,'' Item,'' itemname1,supplier,sample,no_comm, '' AS COMMENT,sum(qty)qty,sum(amount)amount "
        SQL = SQL + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample =1 and EXCLUDE =0 "

        SQL = SQL + " Group by          sales,name,cust,custname,supplier,sample,no_comm "
        SQL = SQL + " UNION ALL"
        SQL = SQL + " SELECT sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,COMMENT1,sum(qty)qty,sum(amount)amount  FROM ("
        SQL = SQL + " select  *, SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
        SQL = SQL + " WHERE  CHARINDEX('客訴單號',comment )>0 and EXCLUDE =0 "

        SQL = SQL + " )A GROUP BY sales,name,cust,custname,Item,itemname1,supplier,sample,no_comm,COMMENT1"
        SQL = SQL + " )A  )A LEFT JOIN  ("
        SQL = SQL + " SELECT * FROM M_RELATED"
        SQL = SQL + " WHERE RELATEDID = 'A'"
        SQL = SQL + " ) B "
        SQL = SQL + " ON  A.NAME =B.USERNAME )A  ORDER BY Unique_ID  "


        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()

        'SQL = " select ITEMNAME from F_nocommlist  "

        'Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        'GridView2.DataSource = DBAdapter2
        'GridView2.DataBind()


    End Sub
 





    '*****************************************************************
    '**(ShowSheetField)
    '**     建置下拉式選單初始值
    '**
    '*****************************************************************
    Sub SetFieldData()


    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        sales = Trim(GridView1.SelectedRow.Cells(1).Text)
        cust = Trim(GridView1.SelectedRow.Cells(3).Text)
        Item = Trim(GridView1.SelectedRow.Cells(14).Text)
        Supplier = Trim(GridView1.SelectedRow.Cells(10).Text)
        Qty = Trim(GridView1.SelectedRow.Cells(8).Text)
        Comment = Trim(GridView1.SelectedRow.Cells(11).Text)
        Dim Userid As String
        Userid = Trim(GridView1.SelectedRow.Cells(15).Text)
        If Userid = "&nbsp;" Then
            Userid = ""
        End If

        If Item = "&nbsp;" Then
            Item = ""
        End If
        If Comment = "&nbsp;" Then
            Comment = ""
        End If
        Dim URL As String = ""


        If Userid = "" Then

            uJavaScript.PopMsg(Me, "此USER ID對應不到關係人A,請自行選取報告者")
            URL = "NoCommSheet_01.aspx?pFormNo=001172" & "&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=" & pUser & "&pUserID=" & pUser & "&pSales=" & sales & "&pCust=" & cust & "&pItem=" & Item & "&pSupplier=" & Supplier & "&pComment=" & Comment & "&pNoAuto=1"
            Response.Write("<script>window.open('" & URL & "','_blank')</script>")
        Else
            If Supplier = "1" Then
                If Qty > 600 Then
                    '如果起單過就把OR 放到  M_NewExcludeOrder
                    InsertExcludeOrder()
                    URL = "NoCommSheet_01.aspx?pFormNo=001172" & "&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=" & pUser & "&pUserID=" & pUser & "&pSales=" & sales & "&pCust=" & cust & "&pItem=" & Item & "&pSupplier=" & Supplier & "&pComment=" & Comment
                    Response.Write("<script>window.open('" & URL & "','_blank')</script>")

                Else
                    uJavaScript.PopMsg(Me, "外注廠樣品訂單未超過600pcs,不需報告")
                End If

            ElseIf Userid = "" Then
                uJavaScript.PopMsg(Me, "此員工對應不到 USER ID,請連絡IT")
                URL = "NoCommSheet_01.aspx?pFormNo=001172" & "&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=" & pUser & "&pUserID=" & pUser & "&pSales=" & sales & "&pCust=" & cust & "&pItem=" & Item & "&pSupplier=" & Supplier & "&pComment=" & Comment & "&pNoAuto=1"
                Response.Write("<script>window.open('" & URL & "','_blank')</script>")
            Else


                '如果起單過就把OR 放到  M_NewExcludeOrder
                InsertExcludeOrder()
                URL = "NoCommSheet_01.aspx?pFormNo=001172" & "&pFormSno=0&pStep=1&pSeqNo=1&pApplyID=" & pUser & "&pUserID=" & pUser & "&pSales=" & sales & "&pCust=" & cust & "&pItem=" & Item & "&pSupplier=" & Supplier & "&pComment=" & Comment
                Response.Write("<script>window.open('" & URL & "','_blank')</script>")





            End If


        End If
       

    End Sub



    Protected Sub BExclude_Click(ByVal sender As Object, ByVal e As System.EventArgs)



        Dim btn As Button = CType(sender, Button)

        ' 取得按下按鈕的那一列
        Dim gvr As GridViewRow = CType(btn.NamingContainer, GridViewRow)

        Dim i As Integer
        i = gvr.DataItemIndex

        'Dim i As Integer
        Dim id As String = GridView1.DataKeys(i).Value.ToString()
 

        cust = GridView1.Rows(id - 1).Cells(3).Text
        Item = GridView1.Rows(id - 1).Cells(14).Text
        If Item = "&nbsp;" Then
            Item = ""
        End If

        InsertExcludeOrder()

       
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        e.Row.Cells(12).Visible = False
        e.Row.Cells(14).Visible = False
        e.Row.Cells(15).Visible = False
    End Sub

    '如果起單過就把OR 放到  M_NewExcludeOrder

    Sub InsertExcludeOrder()

        '如果起單過就把OR 放到  M_NewExcludeOrder

        Dim DBDataSet1 As New DataSet

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("TP_Conn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Dim i As Integer
        Dim Sql As String
        'insert 排除OR  M_NewExcludeOrder


        Sql = " select distinct ORDER1,ITEM  from("
        Sql = Sql + "   select   confirm1 as 'ORD DATE',order1 ,SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT  "
        Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=0  "
        Sql = Sql + "  and cust ='" + cust + "' and item = '" + Item + "'"

        Sql = Sql + " UNION ALL"
        Sql = Sql + "  select   confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT  "
        Sql = Sql + " from F_nocommlist where CHARINDEX('客訴單號', comment) =0 and sample=1  "
        If Item = "" Then
            Sql = Sql + " and cust ='" + cust + "'"
        Else
            Sql = Sql + " and cust ='" + cust + "' and item = '" + Item + "'"

        End If

        Sql = Sql + " UNION ALL"
        Sql = Sql + " SELECT  confirm1 as 'ORD DATE',order1 as 'ORDER',SAMPLE,BUYER,BUYERNAME,ITEM,itemname1 as 'ITEMNAME',COLOR,QTY,amount as 'AMOUNT',COMMENT1   FROM ("
        Sql = Sql + " select  *,SUBSTRING(COMMENT, CHARINDEX( '客訴單號',comment )+5,10)COMMENT1 from F_nocommlist"
        Sql = Sql + " WHERE  CHARINDEX('客訴單號',comment )>0"
        Sql = Sql + " and cust ='" + cust + "' and item = '" + Item + "'"
        Sql = Sql + " )A  "
        Sql = Sql + " )a"


        Dim DBList As DataTable = uDataBase.GetDataTable(Sql)
        Dim Order1 As String = ""
        If DBList.Rows.Count > 0 Then
            For i = 0 To DBList.Rows.Count - 1
                Order1 = DBList.Rows(i).Item("ORDER1")
                Sql = "Insert into M_NewExcludeOrder (Cat,OrderNo,createtime)"
                Sql = Sql + " values('NOCOM','" + DBList.Rows(0).Item("ORDER1") + "',convert(char(10),getdate(),112))"
                OleDBCommand1.Connection = OleDbConnection1
                OleDBCommand1.CommandText = Sql
                OleDbConnection1.Open()
                OleDBCommand1.ExecuteNonQuery()
                OleDbConnection1.Close()




                '更新成已排除
                Sql = "     UPDATE  F_nocommlist "
                Sql = Sql + " SET EXCLUDE =1 "
                Sql = Sql + " ,createdate =getdate()"
                Sql = Sql + " where Order1 = '" + Order1 + "'"
                If Item <> "" Then
                    Sql = Sql + " AND ITEM ='" + Item + "'"
                End If

                uDataBase.ExecuteNonQuery(Sql)

            Next
        End If


        DataList()

    End Sub
End Class

