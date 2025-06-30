Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class OrderBuyerPriceCheckList
    Inherits System.Web.UI.Page
    '
    Dim fpObj As New ForProject
    Dim uJavaScript As New Utility.JScript
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim xBuyer As String
    '---------------------------------------------------------------------------------------------------
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            SetParameter()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    Sub SetParameter()
        xBuyer = Request.QueryString("pBuyer")
        If xBuyer = "" Then
            DCustomer.Text = "CUSTOMER"
            DBuyer.Text = "BUYER"
        Else
            DCustomer.Text = Mid(xBuyer, 1, InStr(xBuyer, "-") - 1)
            DBuyer.Text = Mid(xBuyer, InStr(xBuyer, "-") + 1)
        End If
        DStartDate.Text = DateAdd(DateInterval.Day, -7, Now).ToString("yyyyMMdd")
        DEndDate.Text = Now.ToString("yyyyMMdd")
    End Sub
    '---------------------------------------------------------------------------------------------------
    Protected Sub Go_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Go.Click
        Dim err As Boolean = False
        Dim msg As String = ""
        '
        If DCustomer.Text = "" Then
            err = True
            msg = "CUSTOMER 不可空白 ! "
        End If
        If DBuyer.Text = "" Then
            err = True
            msg = "BUYER 不可空白 ! "
        End If
        If DStartDate.Text = "" Then
            err = True
            msg = "起迄日期不可空白 ! "
        End If
        If DEndDate.Text = "" Then
            err = True
            msg = "起迄日期不可空白 ! "
        End If
        If Not err Then
            If Not IsNumeric(DStartDate.Text) Then
                err = True
                msg = "起迄日期輸入不正確 ! (正確例:20121105) "
            End If
            If Not IsNumeric(DEndDate.Text) Then
                err = True
                msg = "起迄日期輸入不正確 ! (正確例:20121105) "
            End If
        End If
        If Not err Then
            Dim xDays As Integer = DateDiff(DateInterval.Day, CDate(Mid(DStartDate.Text, 1, 4) + "/" + Mid(DStartDate.Text, 5, 2) + "/" + Mid(DStartDate.Text, 7, 2)), CDate(Mid(DEndDate.Text, 1, 4) + "/" + Mid(DEndDate.Text, 5, 2) + "/" + Mid(DEndDate.Text, 7, 2)))
            If xDays > 7 Then
                err = True
                msg = "指定的起迄日期天數共 " + CStr(xDays) + " 天已經超過限制，請調整至 7 天內 ! "
            End If
        End If
        '
        If Not err Then
            ShowData()
        Else
            uJavaScript.PopMsg(Me, msg)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    Sub ShowData()
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ds, ds1, ds2, ds3 As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '
        cn.ConnectionString = ConnectString

        SQL = "SELECT "
        SQL &= "ORDN5C AS ORDERNO, "    '0
        SQL &= "OSBN5E AS SEQ, "        '1
        SQL &= "CSTC5C AS CUST, "       '2
        SQL &= "BYRC5C AS BUYER, "      '3
        SQL &= "PLVC5C AS VERSION, "    '4
        SQL &= "OINU5C AS ORDERDATE, "  '5
        SQL &= "ITMC5E AS CODE, "       '6
        SQL &= "LNGV5E AS LENGTH, "     '7
        SQL &= "LUNC5E AS U, "          '8
        SQL &= "CLRC5E AS COLOR, "      '9
        SQL &= "ORRQ5E AS QUANTITY, "   '10
        SQL &= "OQUC5E AS UNIT, "       '11
        SQL &= "SUNP5E AS SALESPRICE, " '12
        SQL &= "TCRC5C AS CURRENCY, "   '13
        SQL &= "''     AS PRICEA, "     '14
        SQL &= "''     AS PRICEB, "     '15
        SQL &= "''     AS REGISTER, "   '16
        SQL &= "''     AS UPDATE, "     '17
        SQL &= "''     AS ITEMNAME1, "  '18
        SQL &= "''     AS ITEMNAME2, "  '19
        SQL &= "FL1I39 AS NAME, "       '20
        SQL &= "''     AS CUSTITEM, "   '21
        SQL &= "''     AS MARK "        '22
        SQL &= "FROM WAVEDLIB.S5C00 LEFT JOIN WAVEDLIB.S5E00 ON ORDN5C = ORDN5E "
        SQL &= "                    LEFT JOIN WAVEDLIB.S3900 ON CSTC5C = CLNC39 "
        SQL &= "WHERE CSTC5C = '" & DCustomer.Text & "' "
        SQL &= "  AND BYRC5C = '" & DBuyer.Text & "' "
        SQL &= "  AND OINU5C >= " & DStartDate.Text & " "
        SQL &= "  AND OINU5C <= " & DEndDate.Text & " "
        SQL &= "ORDER BY OINU5C,ORDN5C,OSBN5E "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "ORDER")
        If ds.Tables("ORDER").Rows.Count > 0 Then
            '
            For i As Integer = 0 To ds.Tables("ORDER").Rows.Count - 1
                'Get Customer Buyer Price
                ds1.Clear()
                SQL = "SELECT PFAP3M,PFBP3M,RADU3M,RUPU3M "
                SQL &= "FROM WAVEDLIB.S3M00 "
                SQL &= "WHERE CSTC3M = '" & ds.Tables("ORDER").Rows(i)(2) & "' "
                SQL &= "  AND BYRC3M = '" & ds.Tables("ORDER").Rows(i)(3) & "' "
                SQL &= "  AND PLVC3M = '" & ds.Tables("ORDER").Rows(i)(4) & "' "
                SQL &= "  AND ITMC3M = '" & ds.Tables("ORDER").Rows(i)(6) & "' "
                SQL &= "  AND CRRC3M = '" & ds.Tables("ORDER").Rows(i)(13) & "' "
                SQL &= "  AND TTRC3M = '" & "F" & "' "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                DBAdapter2.Fill(ds1, "PRICE")
                If ds1.Tables("PRICE").Rows.Count > 0 Then
                    ds.Tables("ORDER").Rows(i)(14) = ds1.Tables("PRICE").Rows(0)(0)
                    ds.Tables("ORDER").Rows(i)(15) = ds1.Tables("PRICE").Rows(0)(1)
                    ds.Tables("ORDER").Rows(i)(16) = ds1.Tables("PRICE").Rows(0)(2)
                    ds.Tables("ORDER").Rows(i)(17) = ds1.Tables("PRICE").Rows(0)(3)
                Else
                    ds.Tables("ORDER").Rows(i)(22) = "★ Price Error"
                End If
                'Get ItemName
                ds2.Clear()
                SQL = "SELECT TRIM(IT1IA0),TRIM(IT2IA0) "
                SQL &= "FROM WAVEDLIB.FA000 "
                SQL &= "WHERE ITMCA0 = '" & ds.Tables("ORDER").Rows(i)(6) & "' "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                DBAdapter3.Fill(ds2, "ITEM")
                If ds2.Tables("ITEM").Rows.Count > 0 Then
                    ds.Tables("ORDER").Rows(i)(18) = ds2.Tables("ITEM").Rows(0)(0)
                    ds.Tables("ORDER").Rows(i)(19) = ds2.Tables("ITEM").Rows(0)(1)
                End If
                'Get CUST ITEM
                ds3.Clear()
                SQL = "SELECT NINOW2 "
                SQL &= "FROM WAVEDLIB.SC750W2 "
                SQL &= "WHERE ORDNW2 = '" & ds.Tables("ORDER").Rows(i)(0) & "' "
                SQL &= "  AND OSBNW2 = '" & ds.Tables("ORDER").Rows(i)(1) & "' "
                Dim DBAdapter4 As New OleDbDataAdapter(SQL, cn)
                DBAdapter4.Fill(ds3, "VDP")
                If ds3.Tables("VDP").Rows.Count > 0 Then
                    ds.Tables("ORDER").Rows(i)(21) = ds3.Tables("VDP").Rows(0)(0)
                End If
            Next
            '
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim C_Blank As TableCell = New TableCell
            Dim C_Name As TableCell = New TableCell
            Dim C_CUSTITEM As TableCell = New TableCell
            Dim C_ItemName1 As TableCell = New TableCell
            Dim C_ItemName2 As TableCell = New TableCell
            '
            C_Blank.BackColor = Color.Silver
            C_Blank.ColumnSpan = 2
            C_Blank.Text = ""
            row.Cells.Add(C_Blank)

            C_Name.Font.Bold = True
            C_Name.BackColor = Color.Silver
            C_Name.HorizontalAlign = HorizontalAlign.Center
            C_Name.ColumnSpan = 5
            C_Name.Text = "CUST NAME"
            row.Cells.Add(C_Name)

            C_CUSTITEM.Font.Bold = True
            C_CUSTITEM.BackColor = Color.Silver
            C_CUSTITEM.HorizontalAlign = HorizontalAlign.Center
            C_CUSTITEM.Text = "CUST ITEM"
            row.Cells.Add(C_CUSTITEM)

            C_ItemName1.Font.Bold = True
            C_ItemName1.BackColor = Color.Silver
            C_ItemName1.HorizontalAlign = HorizontalAlign.Center
            C_ItemName1.ColumnSpan = 5
            C_ItemName1.Text = "ITEM NAME-1"
            row.Cells.Add(C_ItemName1)

            C_ItemName2.Font.Bold = True
            C_ItemName2.BackColor = Color.Silver
            C_ItemName2.HorizontalAlign = HorizontalAlign.Center
            C_ItemName2.ColumnSpan = 5
            C_ItemName2.Text = "ITEM NAME-2"
            row.Cells.Add(C_ItemName2)

            e.Row.Parent.Controls.Add(row)
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Normal)
            Dim C_Blank As TableCell = New TableCell
            Dim C_Name As TableCell = New TableCell
            Dim C_CUSTITEM As TableCell = New TableCell
            Dim C_ItemName1 As TableCell = New TableCell
            Dim C_ItemName2 As TableCell = New TableCell
            Dim xColor As System.Drawing.Color = Color.White
            '
            If DataBinder.Eval(e.Row.DataItem, "MARK") <> "" Then
                xColor = Color.LightPink
                e.Row.Cells(0).BackColor = xColor
                e.Row.Cells(1).BackColor = xColor
                e.Row.Cells(2).BackColor = xColor
                e.Row.Cells(3).BackColor = xColor
                e.Row.Cells(4).BackColor = xColor
                e.Row.Cells(5).BackColor = xColor
                e.Row.Cells(6).BackColor = xColor
                e.Row.Cells(7).BackColor = xColor
                e.Row.Cells(8).BackColor = xColor
                e.Row.Cells(9).BackColor = xColor
                '
                e.Row.Cells(10).BackColor = xColor
                e.Row.Cells(11).BackColor = xColor
                e.Row.Cells(12).BackColor = xColor
                e.Row.Cells(13).BackColor = xColor
                e.Row.Cells(14).BackColor = xColor
                e.Row.Cells(15).BackColor = xColor
                e.Row.Cells(16).BackColor = xColor
                e.Row.Cells(17).BackColor = xColor
            End If
            '
            C_Blank.ColumnSpan = 2
            C_Blank.BackColor = xColor
            C_Blank.ForeColor = Color.Red
            C_Blank.Text = DataBinder.Eval(e.Row.DataItem, "MARK")
            row.Cells.Add(C_Blank)

            C_Name.ColumnSpan = 5
            C_Name.BackColor = xColor
            C_Name.Text = DataBinder.Eval(e.Row.DataItem, "NAME")
            row.Cells.Add(C_Name)

            C_CUSTITEM.BackColor = xColor
            C_CUSTITEM.Text = DataBinder.Eval(e.Row.DataItem, "CUSTITEM")
            row.Cells.Add(C_CUSTITEM)

            C_ItemName1.ColumnSpan = 5
            C_ItemName1.BackColor = xColor
            C_ItemName1.Text = DataBinder.Eval(e.Row.DataItem, "ITEMNAME1")
            row.Cells.Add(C_ItemName1)
            e.Row.Parent.Controls.Add(row)

            C_ItemName2.ColumnSpan = 5
            C_ItemName2.BackColor = xColor
            C_ItemName2.Text = DataBinder.Eval(e.Row.DataItem, "ITEMNAME2")
            row.Cells.Add(C_ItemName2)
            e.Row.Parent.Controls.Add(row)
        End If
        '
        For i As Integer = 0 To 17
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=OrderBuyerCheckList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
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
End Class
