Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_StockProdList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim Buyer As String             'Buyer
    Dim UserID As String            'UserID
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件

        If Not IsPostBack Then                      'PostBack
            SetDefaultValue()                       '設定初值
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
        Response.Cookies("PGM").Value = "InfF_StockProdList.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If

        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBuyer.ReadOnly = True
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        Dim sql As String
        '
        DBuyer.Text = Buyer

        DKeepCode.Items.Clear()
        DKeepCode.Items.Add("ALL")

        sql = "SELECT Data From M_Referp "
        sql &= "Where Cat  = '153' "
        sql &= "  And DKey = '" & Buyer & "' "
        sql &= "Order by Data "
        Dim dt_KeepCode As DataTable = uDataBase.GetDataTable(sql)
        For i As Integer = 0 To dt_KeepCode.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt_KeepCode.Rows(i).Item("Data")
            ListItem1.Value = dt_KeepCode.Rows(i).Item("Data")
            ListItem1.Selected = False
            DKeepCode.Items.Add(ListItem1)
        Next

        DItem.Text = ""
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        Dim sql As String
        '
        Dim xKeepString As String = "('" + DKeepCode.Text + "')"
        If DKeepCode.Text = "ALL" Then
            sql = "SELECT Data From M_Referp "
            sql &= "Where Cat  = '153' "
            sql &= "  And DKey = '" & Buyer & "' "
            sql &= "Order by Data "
            Dim dt_KeepCode As DataTable = uDataBase.GetDataTable(sql)
            If dt_KeepCode.Rows.Count > 0 Then
                xKeepString = "("
                For i As Integer = 0 To dt_KeepCode.Rows.Count - 1
                    If i = 0 Then
                        xKeepString = xKeepString + "'" + dt_KeepCode.Rows(i).Item("Data") + "'"
                    Else
                        xKeepString = xKeepString + ",'" + dt_KeepCode.Rows(i).Item("Data") + "'"
                    End If
                Next
                xKeepString = xKeepString + ")"
            End If
        End If
        '
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        cn.ConnectionString = ConnectString
        '
        sql = "SELECT Rtrim(ITMCDB) As Item, Rtrim(CLRCDB) As Color, Rtrim(KEPCDB) As KeepCode,  "
        sql &= "'' As ItemName, "
        sql &= "0 As KeepQty, '' As KeepQtyDescr, "
        sql &= "0 As ProdQty, '' As ProdQtyDescr1, '' As ProdQtyDescr2, "
        sql &= "'' As DataTime "
        sql &= "FROM WAVEDLIB.TDB00 "
        sql &= "WHERE DPTCDB = '" & "01" & "' "
        sql &= "  AND KEPCDB IN " & xKeepString & " "
        If DItem.Text <> "" Then
            sql &= "  AND ITMCDB = '" & DItem.Text & "' "
        End If
        sql &= "GROUP BY ITMCDB, CLRCDB, KEPCDB "
        sql &= "ORDER BY ITMCDB, CLRCDB, KEPCDB  "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
        DBAdapter1.Fill(ds, "TDB00")
        If ds.Tables("TDB00").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds.Tables("TDB00")
            GridView1.DataBind()
        End If
        '
        cn.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim sql As String
            Dim cn As New OleDbConnection
            Dim ds, ds1, ds2 As New DataSet
            cn.ConnectionString = ConnectString
            '
            ' 取得 ItemName
            Dim xItemName As String = ""
            oWaves.GetItemName(DataBinder.Eval(e.Row.DataItem, "Item"), xItemName)
            e.Row.Cells(1).Text = xItemName
            '
            ' 取得 KeepQty
            Dim xQty As String = "0"
            oWaves.GetKeepCodeInventory("01", _
                                        DataBinder.Eval(e.Row.DataItem, "Item"), _
                                        DataBinder.Eval(e.Row.DataItem, "Color"), _
                                        DataBinder.Eval(e.Row.DataItem, "KeepCode"), _
                                        xQty)
            If CDbl(xQty) > 0 Then
                e.Row.Cells(4).Text = Format(CDbl(xQty) / 10000000, "###,###,##0.00")
            Else
                e.Row.Cells(4).Text = "0.00"
            End If
            '
            ' 取得 KeepQty說明-STOCK
            Dim xDescr As String = ""
            sql = "SELECT Ltrim(Rtrim(SLCCDB)) AS LOC, OKSQDB AS QTY, LUSUDB AS LASTTIME FROM WAVEDLIB.TDB00 "
            sql &= "Where DPTCDB = '" & "01" & "' "
            sql &= "  And ITMCDB = '" & DataBinder.Eval(e.Row.DataItem, "Item") & "' "
            sql &= "  And CLRCDB = '" & DataBinder.Eval(e.Row.DataItem, "Color") & "' "
            sql &= "  And KEPCDB = '" & DataBinder.Eval(e.Row.DataItem, "KeepCode") & "' "
            sql &= "Order By SLCCDB "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "KeepQty")
            For i As Integer = 0 To ds.Tables("KeepQty").Rows.Count - 1
                If i = 0 Then
                    xDescr = "[" + ds.Tables("KeepQty").Rows(i).Item("LOC") + "/" + Format(ds.Tables("KeepQty").Rows(i).Item("QTY"), "###,###,##0.00") + "/" + ds.Tables("KeepQty").Rows(i).Item("LASTTIME").ToString
                Else
                    xDescr = xDescr + "," + "[" + ds.Tables("KeepQty").Rows(i).Item("LOC") + "/" + Format(ds.Tables("KeepQty").Rows(i).Item("QTY"), "###,###,##0.00") + "/" + ds.Tables("KeepQty").Rows(i).Item("LASTTIME").ToString
                End If
                '
                sql = "SELECT Ltrim(Rtrim(SIONDF)) AS ORNO "
                sql &= "FROM WAVEDLIB.TDF00 "
                sql &= "Where DPTCDF = '" & "01" & "' "
                sql &= "  And ITMCDF = '" & DataBinder.Eval(e.Row.DataItem, "Item") & "' "
                sql &= "  And CLRCDF = '" & DataBinder.Eval(e.Row.DataItem, "Color") & "' "
                sql &= "  And KEPCDF = '" & DataBinder.Eval(e.Row.DataItem, "KeepCode") & "' "
                sql &= "  And SLCCDF = '" & ds.Tables("KeepQty").Rows(i).Item("LOC") & "' "
                sql &= "  And SHSCDF = '" & "06" & "' "
                sql &= "Order By SIOUDF DESC, SIOTDF DESC "
                '
                Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
                DBAdapter2.Fill(ds1, "LASTORDER")
                For j As Integer = 0 To ds1.Tables("LASTORDER").Rows.Count - 1
                    '
                    sql = "SELECT CORN5C FROM WAVEDLIB.S5C00 "
                    sql &= "Where ORDN5C = '" & ds1.Tables("LASTORDER").Rows(j).Item("ORNO") & "' "

                    Dim DBAdapter3 As New OleDbDataAdapter(sql, cn)
                    DBAdapter3.Fill(ds2, "S5C00")

                    If ds2.Tables("S5C00").Rows.Count > 0 Then
                        If Not IsDBNull(ds2.Tables("S5C00").Rows(0).Item("CORN5C")) Then
                            If ds2.Tables("S5C00").Rows(0).Item("CORN5C") <> "" Then
                                xDescr = xDescr + "/" + ds1.Tables("LASTORDER").Rows(j).Item("ORNO")
                                xDescr = xDescr + "/" + ds2.Tables("S5C00").Rows(0).Item("CORN5C")
                                Exit For
                            End If
                        End If
                    End If
                Next
                '
                If xDescr <> "" Then
                    xDescr = xDescr + "]"
                End If
            Next
            e.Row.Cells(5).Text = xDescr
            '
            ' 取得 ProdQty
            Dim xScheProdQty As Double = 0
            Dim xOnProdQty As Double = 0
            Dim xProd(6) As String
            xQty = "0"
            '   1.Sche Prod Qty
            oWaves.GetProductionQty("01", _
                                    DataBinder.Eval(e.Row.DataItem, "Item"), _
                                    DataBinder.Eval(e.Row.DataItem, "Color"), _
                                    DataBinder.Eval(e.Row.DataItem, "KeepCode"), _
                                    0, _
                                    xQty, _
                                    xProd)
            If CDbl(xQty) > 0 Then
                xScheProdQty = CDbl(xQty) / 10000000
            End If
            '   2.On Prod Qty
            oWaves.GetProductionQty("01", _
                                    DataBinder.Eval(e.Row.DataItem, "Item"), _
                                    DataBinder.Eval(e.Row.DataItem, "Color"), _
                                    DataBinder.Eval(e.Row.DataItem, "KeepCode"), _
                                    1, _
                                    xQty, _
                                    xProd)
            If CDbl(xQty) > 0 Then
                xOnProdQty = CDbl(xQty) / 10000000
            End If
            e.Row.Cells(6).Text = Format(xScheProdQty + xOnProdQty, "###,###,##0.00")
            '
            ' 取得 ProdQty說明-1
            e.Row.Cells(7).Text = "[" + "SCHE" + "/" + Format(xScheProdQty, "###,###,##0.00") + "]," + _
                                  "[" + "ON" + "/" + Format(xOnProdQty, "###,###,##0.00") + "]"
            '
            ' 取得 ProdQty說明-2
            '   1.Sche Prod Qty
            Dim xDescr1 As String = ""
            ds.Clear()
            sql = "SELECT PSCN9F, RLTN9F, ALMQ9F, KEPC9F, RLON9F, UAVU9F FROM WAVEDLIB.F9F00 "
            sql &= "Where DPTC9F = '" & "01" & "' "
            sql &= "  And ITMC9F = '" & DataBinder.Eval(e.Row.DataItem, "Item") & "' "
            sql &= "  And CLRC9F = '" & DataBinder.Eval(e.Row.DataItem, "Color") & "' "
            sql &= "  And KEPC9F = '" & DataBinder.Eval(e.Row.DataItem, "KeepCode") & "' "
            sql &= "  And PSHN9F = '" & "" & "' "           ' SCHE PROD
            sql &= "  And PCPU9F = '" & "0" & "' "          ' COMPLETE DATE (PRODUCTION) 
            sql &= "  And PSCN9F <> RLTN9F "                ' PRODUCTION SHEDULE NO. <> RELATION NO.
            sql &= "  And ALMQ9F > 0 "                      ' ALLOCATE QUANTITY
            '
            Dim DBAdapter4 As New OleDbDataAdapter(sql, cn)
            DBAdapter4.Fill(ds, "ScheProdQty")
            If ds.Tables("ScheProdQty").Rows.Count > 0 Then
                For i As Integer = 0 To ds.Tables("ScheProdQty").Rows.Count - 1
                    If i = 0 Then
                        xDescr1 = "["
                    Else
                        xDescr1 = xDescr1 + ",["
                    End If
                    If InStr(ds.Tables("ScheProdQty").Rows(i).Item("RLTN9F"), "OR") > 0 Or InStr(ds.Tables("ScheProdQty").Rows(i).Item("RLON9F"), "OR") > 0 Then
                        If InStr(ds.Tables("ScheProdQty").Rows(i).Item("RLTN9F"), "OR") > 0 Then
                            xDescr1 = xDescr1 + ds.Tables("ScheProdQty").Rows(i).Item("RLTN9F") + "/"
                        Else
                            xDescr1 = xDescr1 + ds.Tables("ScheProdQty").Rows(i).Item("RLON9F") + "/"
                        End If
                        xDescr1 = xDescr1 + Format(ds.Tables("ScheProdQty").Rows(i).Item("ALMQ9F"), "###,###,##0.00") + "/" + ds.Tables("ScheProdQty").Rows(i).Item("UAVU9F").ToString
                    End If
                    xDescr1 = xDescr1 + "]"
                Next
            End If
            '   2.On Prod Qty
            Dim xDescr2 As String = ""
            ds.Clear()
            sql = "SELECT PSCN9F, RLTN9F, ALMQ9F, KEPC9F, RLON9F, UAVU9F FROM WAVEDLIB.F9F00 "
            sql &= "Where DPTC9F = '" & "01" & "' "
            sql &= "  And ITMC9F = '" & DataBinder.Eval(e.Row.DataItem, "Item") & "' "
            sql &= "  And CLRC9F = '" & DataBinder.Eval(e.Row.DataItem, "Color") & "' "
            sql &= "  And KEPC9F = '" & DataBinder.Eval(e.Row.DataItem, "KeepCode") & "' "
            sql &= "  And PSHN9F <> '" & "" & "' "          ' ON PROD
            sql &= "  And PCPU9F = '" & "0" & "' "          ' COMPLETE DATE (PRODUCTION) 
            sql &= "  And PSCN9F <> RLTN9F "                ' PRODUCTION SHEDULE NO. <> RELATION NO.
            sql &= "  And ALMQ9F > 0 "                      ' ALLOCATE QUANTITY
            '
            Dim DBAdapter5 As New OleDbDataAdapter(sql, cn)
            DBAdapter5.Fill(ds, "OnProdQty")
            If ds.Tables("OnProdQty").Rows.Count > 0 Then
                For i As Integer = 0 To ds.Tables("OnProdQty").Rows.Count - 1
                    If i = 0 Then
                        xDescr2 = "["
                    Else
                        xDescr2 = xDescr2 + ",["
                    End If
                    If InStr(ds.Tables("OnProdQty").Rows(i).Item("RLTN9F"), "OR") > 0 Or InStr(ds.Tables("OnProdQty").Rows(i).Item("RLON9F"), "OR") > 0 Then
                        If InStr(ds.Tables("OnProdQty").Rows(i).Item("RLTN9F"), "OR") > 0 Then
                            xDescr2 = xDescr2 + ds.Tables("OnProdQty").Rows(i).Item("RLTN9F") + "/"
                        Else
                            xDescr2 = xDescr2 + ds.Tables("OnProdQty").Rows(i).Item("RLON9F") + "/"
                        End If
                        xDescr2 = xDescr2 + Format(ds.Tables("OnProdQty").Rows(i).Item("ALMQ9F"), "###,###,##0.00") + "/" + ds.Tables("OnProdQty").Rows(i).Item("UAVU9F").ToString
                    End If
                    xDescr2 = xDescr2 + "]"
                Next
            End If
            e.Row.Cells(8).Text = ""
            If xDescr1 <> "" Or xDescr2 <> "" Then
                If xDescr1 <> "" Then
                    e.Row.Cells(8).Text = xDescr1
                End If
                If xDescr2 <> "" Then
                    If e.Row.Cells(8).Text <> "" Then
                        e.Row.Cells(8).Text = e.Row.Cells(8).Text + "," + xDescr2
                    Else
                        e.Row.Cells(8).Text = xDescr2
                    End If
                End If
            End If
            '
            ' 取得 資料時點
            e.Row.Cells(9).Text = Now.ToString("yyyy/MM/dd HH:mm:ss")
            '
            cn.Close()
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub

End Class
