Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PC_INQWingsOrder
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pOrderNo As String
    Dim pPuller As String

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            SetParameter()                              '設定參數
            GetData()
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
        Server.ScriptTimeout = 900                                  '設定逾時時間
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pOrderNo = Request.QueryString("pOrderNo")
        pPuller = Request.QueryString("pPuller")
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds, ds1, ds2 As New DataSet
        Dim wTop As Integer
        Dim xDatoFound As Boolean
        '
        wTop = 40
        '
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()
        '
        Sql = "SELECT "
        Sql = Sql + "OrderNo AS ORDN5E, OrderSubNo AS OSBN5E, SalesDate AS OCNU5C, "
        Sql = Sql + "Custmer AS CSTC5C, [CustmerName] AS FL1I39, Buyer AS BYRC5C, [BuyerName] AS BYRI35, Salesman AS SPRC5C,[SalesName] AS FEMI05, "
        Sql = Sql + "Sample AS SMPF5C, NC AS NCMF5C, ITEM AS ITMC5E, [ITEMNAME] AS ITEMNAME, Length AS LNGV5E, L_Unit as LUNC5E, Color AS CLRC5E, Quantity AS ORRQ5E "
        Sql = Sql + "FROM W_SALES_DATA_10Y "
        '
        Sql = Sql & "WHERE OrderNo <> 'ORZZZZZZZZ' "
        '
        If pOrderNo <> "" Then
            Sql = Sql & "AND OrderNo = '" & Trim(pOrderNo) & "' "
        Else
            Sql = Sql & "AND OrderNo = '" & "ORZZZZZZZZ" & "' "
        End If
        Sql = Sql & "ORDER BY OrderNo, OrderSubNo "
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, OleDbConnection1)
        DBAdapter1.Fill(ds1, "SQLORDER")
        If ds1.Tables("SQLORDER").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds1
            GridView1.DataBind()
            '
            GridView2.Style("top") = (ds1.Tables("SQLORDER").Rows.Count + 4) * 25 + wTop & "px"
            wTop = (ds1.Tables("SQLORDER").Rows.Count + 4) * 25 + wTop
            '
            If Trim(pPuller) <> "" Then xDatoFound = True
            '
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
        OleDbConnection1.Close()
        '---------------
        '
        'cn.ConnectionString = ConnectString

        'Sql = "SELECT "
        'Sql = Sql & "A.ORDN5E, A.OSBN5E, B.OCNU5C, B.CSTC5C, C.FL1I39, B.BYRC5C, D.BYRI35, B.SPRC5C, F.FEMI05, B.SMPF5C, B.NCMF5C, "
        'Sql = Sql & "A.ITMC5E, E.IT1IA0 || ' ' || E.IT2IA0 || ' ' || E.IT3IA0 as ITEMNAME, A.LNGV5E, A.LUNC5E, A.CLRC5E, A.ORRQ5E "
        ''
        'Sql = Sql & "FROM WAVEDLIB.S5E00 A "
        'Sql = Sql & "LEFT JOIN WAVEDLIB.S5C00 B ON A.ORDN5E=B.ORDN5C "
        'Sql = Sql & "LEFT JOIN WAVEDLIB.S3900 C ON B.CSTC5C=C.CLNC39 "
        'Sql = Sql & "LEFT JOIN WAVEDLIB.S3500 D ON B.BYRC5C=D.BYRC35 "
        'Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 E ON A.ITMC5E=E.ITMCA0 "
        'Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 F ON B.SPRC5C=F.EMPC05 "
        ''
        ''Sql = Sql & "WHERE A.ORDN5E <> 'ORZZZZZZZZ' "
        'Sql = Sql & "WHERE A.ORDN5E = 'ORZZZZZZZZ' "
        ''
        'If pOrderNo <> "" Then
        '    Sql = Sql & "AND A.ORDN5E = '" & Trim(pOrderNo) & "' "
        'Else
        '    Sql = Sql & "AND A.ORDN5E = '" & "ORZZZZZZZZ" & "' "
        'End If
        'Sql = Sql & "ORDER BY A.ORDN5E, A.OSBN5E "
        ''
        'Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        'DBAdapter1.Fill(ds, "ORDER")
        'If ds.Tables("ORDER").Rows.Count > 0 Then
        '    GridView1.Visible = True
        '    GridView1.DataSource = ds
        '    GridView1.DataBind()
        '    '
        '    GridView2.Style("top") = (ds.Tables("ORDER").Rows.Count + 4) * 25 + wTop & "px"
        '    wTop = (ds.Tables("ORDER").Rows.Count + 4) * 25 + wTop
        '    xDatoFound = True
        'Else
        '    '
        '    Dim OleDbConnection1 As New OleDbConnection
        '    OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '    OleDbConnection1.Open()
        '    '
        '    Sql = "SELECT "
        '    Sql = Sql + "OrderNo AS ORDN5E, OrderSubNo AS OSBN5E, SalesDate AS OCNU5C, "
        '    Sql = Sql + "Custmer AS CSTC5C, [CustmerName] AS FL1I39, Buyer AS BYRC5C, [BuyerName] AS BYRI35, Salesman AS SPRC5C,[SalesName] AS FEMI05, "
        '    Sql = Sql + "Sample AS SMPF5C, NC AS NCMF5C, ITEM AS ITMC5E, [ITEMNAME] AS ITEMNAME, Length AS LNGV5E, L_Unit as LUNC5E, Color AS CLRC5E, Quantity AS ORRQ5E "
        '    Sql = Sql + "FROM W_SALES_DATA_10Y "
        '    '
        '    Sql = Sql & "WHERE OrderNo <> 'ORZZZZZZZZ' "
        '    '
        '    If pOrderNo <> "" Then
        '        Sql = Sql & "AND OrderNo = '" & Trim(pOrderNo) & "' "
        '    Else
        '        Sql = Sql & "AND OrderNo = '" & "ORZZZZZZZZ" & "' "
        '    End If
        '    Sql = Sql & "ORDER BY OrderNo, OrderSubNo "
        '    Dim DBAdapter2 As New OleDbDataAdapter(Sql, OleDbConnection1)
        '    DBAdapter2.Fill(ds1, "SQLORDER")
        '    If ds1.Tables("SQLORDER").Rows.Count > 0 Then
        '        GridView1.Visible = True
        '        GridView1.DataSource = ds1
        '        GridView1.DataBind()
        '        '
        '        GridView2.Style("top") = (ds1.Tables("SQLORDER").Rows.Count + 4) * 25 + wTop & "px"
        '        wTop = (ds1.Tables("SQLORDER").Rows.Count + 4) * 25 + wTop
        '        xDatoFound = True
        '    Else
        '        uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        '    End If
        '    OleDbConnection1.Close()
        'End If

        'cn.Close()
        '---------------
        '
        ' Sales Qty & Amount
        '
        If xDatoFound = True Then
            '
            OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()
            '
            Sql = "SELECT "
            Sql = Sql + "[puller], [Color] "
            Sql = Sql + ",[YY],[QTY],[AMT],[YY1Y],[QTY1Y],[AMT1Y] "
            Sql = Sql + ",[YY2Y],[QTY2Y],[AMT2Y],[YY3Y],[QTY3Y],[AMT3Y] "
            Sql = Sql + ",[YY4Y],[QTY4Y],[AMT4Y],[YY5Y],[QTY5Y],[AMT5Y] "
            Sql = Sql + "FROM W_Puller10YAmount "
            Sql = Sql & "WHERE puller <> '' "
            Sql = Sql & "AND Puller+Color = '" & Trim(pPuller) & "' "
            Sql = Sql & "ORDER BY Puller+Color  "
            Dim DBAdapter2 As New OleDbDataAdapter(Sql, OleDbConnection1)
            DBAdapter2.Fill(ds2, "SALESAMOUNT")
            If ds2.Tables("SALESAMOUNT").Rows.Count > 0 Then
                GridView2.Visible = True
                GridView2.DataSource = ds2
                GridView2.DataBind()
                '
            End If
            OleDbConnection1.Close()
        End If
        '
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i = 0 To 16
                Select Case i
                    Case 3, 5
                        e.Row.Cells(i).Attributes.Add("class", "text")
                    Case Else
                End Select

            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            ' detail
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.Black
            'BUYER PULLER
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Puller" & "<BR>" & "Code"
            tcl(1).BackColor = Color.Black
            'Puller Code

            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Color" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Black
            'Color Code	

            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "QTY" & "<BR>" & ""
            tcl(3).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "AMT" & "<BR>" & ""
            tcl(4).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "QTY" & "<BR>" & ""
            tcl(5).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "AMT" & "<BR>" & ""
            tcl(6).BackColor = Color.Blue

            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "QTY" & "<BR>" & ""
            tcl(7).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "AMT" & "<BR>" & ""
            tcl(8).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "QTY" & "<BR>" & ""
            tcl(9).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "AMT" & "<BR>" & ""
            tcl(10).BackColor = Color.Blue

            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "QTY" & "<BR>" & ""
            tcl(11).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(12).Text = "AMT" & "<BR>" & ""
            tcl(12).BackColor = Color.Green

            tcl.Add(New TableHeaderCell())
            tcl(13).Text = "QTY" & "<BR>" & ""
            tcl(13).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(14).Text = "AMT" & "<BR>" & ""
            tcl(14).BackColor = Color.Blue
            '
            gv.Controls(0).Controls.AddAt(0, H3row)

            '-----------------------------------------
            ' 表頭-N-1行
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "過去5年販賣實績"
            H3tc1.ColumnSpan = 3
            H3tc1.BackColor = Color.Black
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "N" & "<BR>" & "This Year"
            H3tc2.ColumnSpan = 2
            H3tc2.BackColor = Color.Green
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "N-1" & "<BR>" & "Year"
            H3tc3.ColumnSpan = 2
            H3tc3.BackColor = Color.Blue
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "N-2" & "<BR>" & "Year"
            H3tc4.ColumnSpan = 2
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "N-3" & "<BR>" & "Year"
            H3tc5.ColumnSpan = 2
            H3tc5.BackColor = Color.Blue
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = "N-4" & "<BR>" & "Year"
            H3tc6.ColumnSpan = 2
            H3tc6.BackColor = Color.Green
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = "N-5" & "<BR>" & "Year"
            H3tc7.ColumnSpan = 2
            H3tc7.BackColor = Color.Blue
            H3row.Cells.Add(H3tc7)
            '
            gv.Controls(0).Controls.AddAt(0, H3row)

        End If
        '
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 14
                Select Case i
                    Case 0
                        e.Row.Cells(0).Text = ""
                        e.Row.Cells(0).Enabled = False
                        'e.Row.Cells(0).Visible = False
                    Case 1, 2
                        e.Row.Cells(i).ForeColor = Color.Red
                    Case 3, 4, 7, 8, 11, 12
                        e.Row.Cells(i).ForeColor = Color.Green
                    Case Else
                        e.Row.Cells(i).ForeColor = Color.Blue

                End Select
                e.Row.Cells(i).Font.Bold = True
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
End Class
