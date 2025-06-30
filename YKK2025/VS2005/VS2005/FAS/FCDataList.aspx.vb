Imports System.Data
Imports System.Data.OleDb
Partial Class FCDataList
    Inherits System.Web.UI.Page
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetData()
        End If
    End Sub

    Sub GetData()
        Dim SQL As String
        Dim UserID As String = Request.QueryString("userid")
        Dim wformsno As String = Request.QueryString("pFormSno")
        Dim wSts As String = Request.QueryString("pSts")
        Dim wTable As String

        If wformsno = 0 Then
            wformsno = " and formsno = 0 and createuser  ='" + UserID + "'"
        Else
            wformsno = " and  formsno ='" + wformsno + "'"
        End If

        If wSts = 1 Then
            wTable = " f_fcdataHistory"
        Else
            wTable = " f_fcdata"
        End If



        SQL = " select  Scheck,CustomerCode,CustomerName,' '+BuyerCode as BuyerCode,BuyerName,REQDATE,KEEPCODE,BUYMONTH,CORDERNO,SeqNo,AtStock"
        SQL = SQL + ",''+ITEMCODE1 as ITEMCODE1,ITEMNAME1,COLOR1,LENGTH1,LENGTHU1,ORDERQTYP1	"
        SQL = SQL + ",''+ITEMCODE2 as ITEMCODE2,ITEMNAME2,COLOR2,LENGTH2,LENGTHU2,ORDERQTYP2,convert(int,ORDERQTYP1*UnitPrice)UnitPrice,INVQTY,CONVERT(INT,ROUND(INVAMT,0))INVAMT"
        SQL = SQL + " from  " + wTable
        SQL = SQL + " where convert(date,buymonth+'/1') >=convert(date,Rtrim(convert(char,year(getdate())))+'/'+convert(char,month(getdate()))+'/1' )"
        SQL = SQL + wformsno
        SQL = SQL + " and scheck not in ('C','P')"
        SQL = SQL + " order by CustomerCode, BuyerCode,REQDATE,BUYMONTH"
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            GridView1.DataSource = dt
            GridView1.DataBind()
        End If


    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
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

        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)
        Response.AppendHeader("Content-Disposition", "attachment; filename=FCdataList.xls")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        GridView1.RenderControl(hw)
        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()

    End Sub

  
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        e.Row.Cells(3).Attributes.Add("class", "text")
        e.Row.Cells(7).Attributes.Add("class", "text")
        e.Row.Cells(11).Attributes.Add("class", "text")
        e.Row.Cells(17).Attributes.Add("class", "text")
  

    End Sub
End Class
