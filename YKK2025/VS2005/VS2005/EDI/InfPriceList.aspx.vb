Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfPriceList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim wBuyer As String            'Buyer
    Dim wCustomer As String         'Customer
    Dim wUserID As String           'UserID
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then                      'PostBack
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "InfPriceList.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyy/MM/dd")   '現在日期時間
        wBuyer = Request.QueryString("pBuyer")
        wCustomer = Request.QueryString("pBuyer")
        wUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql, xGRBuyer As String
        '
        DBuyer.Text = ""
        sql = "Select RTrim(Name) As NameA, BuyerGroup From M_ControlRecord "
        sql &= "Where Buyer = '" & wBuyer & "' "
        sql &= "Order by Name "
        Dim dt_ControlRecord As DataTable = uDataBase.GetDataTable(sql)
        If dt_ControlRecord.Rows.Count > 0 Then
            DBuyer.Text = dt_ControlRecord.Rows(0).Item("NameA").ToString + " (" + wBuyer + ") "
            xGRBuyer = dt_ControlRecord.Rows(0).Item("BuyerGroup").ToString
        End If
        '
        '<asp:BoundField DataField="PO" HeaderText="PO" ></asp:BoundField>
        '<asp:BoundField DataField="SeqNo" HeaderText="SeqNo" ></asp:BoundField>
        '<asp:BoundField DataField="OrderInf" HeaderText="Order Information" ></asp:BoundField>

        '<asp:BoundField DataField="CustBuyer" HeaderText="Customer/Buyer" ></asp:BoundField>
        '<asp:BoundField DataField="PriceVersion" HeaderText="單價版本" ></asp:BoundField>
        '<asp:BoundField DataField="OrderNo" HeaderText="OrderNo" ></asp:BoundField>
        '<asp:BoundField DataField="Item" HeaderText="Item" ></asp:BoundField>
        '<asp:BoundField DataField="Color" HeaderText="Color" ></asp:BoundField>

        '<asp:BoundField DataField="PriceList" HeaderText="單價類型" ></asp:BoundField>
        '<asp:BoundField DataField="SalesPrice" HeaderText="單價" ></asp:BoundField>
        '<asp:BoundField DataField="RegisterTime" HeaderText="最終更新時間" ></asp:BoundField>
        '<asp:BoundField DataField="PriceListRemark" HeaderText="備註" ></asp:BoundField>
        '


        If fpObj.GetFunctionCode(xGRBuyer, 2) = "P" Then        ' GRBuyer(2)=P (使用自動PO) 
            sql = "SELECT "
            sql &= "a.A1 + ' / ' + a.B1 + ' / ' + a.C1 + ' / ' + a.D1 + ' / ' + a.E1 + ' / ' + "
            sql &= "a.F1 + ' / ' + a.G1 + ' / ' + a.H1 + ' / ' + a.I1 + ' / ' + a.J1 + ' / ' + "
            sql &= "a.K1 + ' / ' + a.L1 + ' / ' + a.M1 + ' / ' + a.N1 + ' / ' + a.O1 + ' / ' "
            sql &= "as OrderInf, "
            '
            sql &= "b.PO as PO, b.Seqno as Seqno, "
            sql &= "b.Customer + '-' + b.Buyer as CustBuyer, b.PriceVersion as PriceVersion, b.OrderNo as OrderNo, b.Item as Item, b.Color as Color, "
            sql &= "b.PriceList as PriceList, b.SalesPrice as SalesPrice, b.RegisterTime as RegisterTime, b.PriceListRemark as PriceListRemark "
            '
            sql &= "From E_InputSheet a, W_SalesPrice b "
            sql &= "Where a.BM1=b.PO "
            sql &= "  And Convert(int,a.BN1)=b.Seqno "
            sql &= "  And a.Buyer = '" & wBuyer & "' "
            sql &= "Order by a.BM1, Convert(int,a.BN1) "
        Else
            sql = "SELECT "
            sql &= "a.A1 + ' / ' + a.B1 + ' / ' + a.C1 + ' / ' + a.D1 + ' / ' + a.E1 + ' / ' + "
            sql &= "a.F1 + ' / ' + a.G1 + ' / ' + a.H1 + ' / ' + a.I1 + ' / ' + a.J1 + ' / ' + "
            sql &= "a.K1 + ' / ' + a.L1 + ' / ' + a.M1 + ' / ' + a.N1 + ' / ' + a.O1 + ' / ' "
            sql &= "as OrderInf, "
            '
            sql &= "b.PO as PO, b.Seqno as Seqno, "
            sql &= "b.Customer + '-' + b.Buyer as CustBuyer, b.PriceVersion as PriceVersion, b.OrderNo as OrderNo, b.Item as Item, b.Color as Color, "
            sql &= "b.PriceList as PriceList, b.SalesPrice as SalesPrice, b.RegisterTime as RegisterTime, b.PriceListRemark as PriceListRemark "
            '
            sql &= "From E_InputSheet a, W_SalesPrice b "
            sql &= "Where a.E1=b.PO "
            sql &= "  And Convert(int,a.D1)=b.Seqno "
            sql &= "  And a.Buyer = '" & wBuyer & "' "
            sql &= "Order by a.E1, Convert(int,a.D1) "
        End If

        '
        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click

        Response.AppendHeader("Content-Disposition", "attachment;filename=PriceListInf.xls")     '程式別不同

        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        For i As Integer = 0 To 8
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub

End Class
