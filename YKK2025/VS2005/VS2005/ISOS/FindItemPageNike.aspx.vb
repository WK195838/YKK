Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class FindItemPageNike
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String = ""      '申請者
    Dim wBuyer As String = ""       '申請者
    Dim wBuyerItem As String = ""   '申請者
    Dim wItemname As String = ""    'ItemName
    Dim wYCode As String = ""    'ItemName

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESDB")
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間

        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            SetDefaultValue()                   '設定初值
            If DCode.Text <> "" Or DName1.Text <> "" Then
                DataList()
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wUserID = Request.QueryString("pUserID")
        wBuyer = Request.QueryString("pBuyer")
        wBuyerItem = Request.QueryString("pBuyerItem")
        wItemname = Request.QueryString("pItemName")
        wYCode = Request.QueryString("pYCode")
        Response.Cookies("PGM").Value = "FindItemPage.aspx"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        Dim i, j, k As Integer
        Dim str, key As String
        Dim xItemName As String()
        '
        DCode.Text = ""
        DName1.Text = ""
        DName2.Text = ""
        DName3.Text = ""
        DName4.Text = ""
        '
        'ItemName(1)
        str = wItemname & " "
        xItemName = str.Split(" ")
        i = 0
        str = ""

        If (wBuyer <> "000013") And (InStr(xItemName(1) & xItemName(2), "*") > 0 Or InStr(xItemName(1) & xItemName(2), "+") > 0) Then
            '
            ' BUYER ITEM (VSOL-39 DSNF211* X6 P12 KENSIN N-ANTI NEWKOB4 OP-NW2P)
            '     VSOL-39 DSNF211* X6 P12 KENSIN N-ANTI NEWKOB4 OP-NW2P
            ' --> VSOL-39 DSNF211 / X6 P12 / KENSIN / N-ANTI  
            For Each str In xItemName
                Select Case i
                    Case 1, 2
                        If InStr(str, "*") > 0 Or InStr(str, "+") > 0 Then
                            For j = 0 To i
                                If DName1.Text = "" Then
                                    DName1.Text = xItemName(j)
                                    DName1.Text = Replace(DName1.Text, "*", "")
                                    DName1.Text = Replace(DName1.Text, "+", "")
                                Else
                                    DName1.Text = DName1.Text & " " & xItemName(j)
                                    DName1.Text = Replace(DName1.Text, "*", "")
                                    DName1.Text = Replace(DName1.Text, "+", "")
                                End If
                            Next
                            Exit For
                        End If
                    Case Else
                End Select
                i = i + 1
            Next
            '
            'ItemName(2)
            i = i + 1
            For j = i To xItemName.Length - 1
                If Mid(xItemName(j), 1, 1) = "P" Then
                    For k = i To j
                        If DName2.Text = "" Then
                            DName2.Text = xItemName(k)
                        Else
                            DName2.Text = DName2.Text & " " & xItemName(k)
                        End If
                    Next
                    Exit For
                End If
            Next
            '
            'ItemName(3)(4)
            For j = i + 1 To xItemName.Length - 1
                If j = i + 2 Then DName3.Text = xItemName(j)
                If j = i + 3 Then DName4.Text = xItemName(j)
            Next
        Else
            '
            ' 一般
            '     VSOL-39 DSNF211A X6 P12 KENSIN N-ANTI NEWKOB4 OP-NW2P)
            ' --> VSOL-39 DSNF211A X6 P12 / KENSIN / N-ANTI   
            '?pUserID=it003&pBuyer=000021&pBuyerItem=HMQ5&pItemName=05 RC DFL EF
            '
            For Each str In xItemName
                Select Case i
                    Case 3, 4, 5
                        If Mid(str, 1, 1) = "P" Then
                            For j = 0 To i - 1
                                If DName1.Text = "" Then
                                    DName1.Text = xItemName(j)
                                Else
                                    DName1.Text = DName1.Text & " " & xItemName(j)
                                End If
                            Next
                            Exit For
                        End If
                    Case Else
                End Select
                i = i + 1
            Next
            '
            If DName1.Text = "" Then
                i = 0
                DName1.Text = xItemName(i)
            End If
            '
            'ItemName(2)(3)(4)
            For j = i + 1 To xItemName.Length - 1
                If j = i + 1 Then
                    'DName2.Text = xItemName(j)
                    DName2.Text = "PBR"
                End If
                If j = i + 2 Then DName3.Text = xItemName(j)
                If j = i + 3 Then DName4.Text = xItemName(j)
            Next
        End If
    End Sub
    '------------------------------------------------------
    Protected Sub BFindItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindItem.Click
        If DCode.Text <> "" Or DName1.Text <> "" Then
            DataList()
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        DataList()
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataBind()
    End Sub

    Sub DataList()
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        Sql = "SELECT "
        Sql = Sql + "LTRIM(RTRIM(ITEM)) As Code, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME1)) As Name1, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME2)) As Name2, "
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME3)) + '　NoDisp[' + NoDisp + ']' As Name3, "
        '
        Sql = Sql + "'' As Season, "
        Sql = Sql + "'' As Color, "
        Sql = Sql + "'' As ColorName, "
        Sql = Sql + "'材料' as ISOS2FAS, "
        Sql = Sql + "'http://10.245.0.3/ISOS/ISOS2FAS.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pItem=' + LTRIM(RTRIM(ITEM)) + '&pItemName1=' + LTRIM(RTRIM(ITEM_NAME1)) + '&pItemName2=' + LTRIM(RTRIM(ITEM_NAME2)) + '&pItemName3=' + LTRIM(RTRIM(ITEM_NAME3)) as URL, "
        Sql = Sql + "LTRIM(RTRIM(SLIDER)) As Slider, "
        '
        Sql = Sql + "'SPC' as SPC, "
        If wBuyer = "000001" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        If wBuyer = "000013" Then
            Sql = Sql + "'SPC' as SPC, "
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        '
        Sql = Sql + "'SPP' as SPP, "
        Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPP" & "' as SPPURL, "
        '
        Sql = Sql + "'BC' as BC, "
        Sql = Sql + "'' as BCURL, "
        Sql = Sql + "'IRW' as IRW, "
        Sql = Sql + "'' as IRWURL "
        '
        Sql = Sql + "From MST_ITEM "
        '
        If DCode.Text <> "" Then
            Sql = Sql + "Where ITEM LIKE '%" & DCode.Text.ToUpper & "%' "
        Else
            Sql = Sql + "Where LTRIM(RTRIM(ITEM_NAME1)) + LTRIM(RTRIM(ITEM_NAME2)) + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName1.Text.ToUpper & "%' "
            If DName2.Text <> "" Then
                Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + LTRIM(RTRIM(ITEM_NAME2)) + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName2.Text.ToUpper & "%' "
            End If
            If DName3.Text <> "" Then
                Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + LTRIM(RTRIM(ITEM_NAME2)) + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName3.Text.ToUpper & "%' "
            End If
            If DName4.Text <> "" Then
                Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + LTRIM(RTRIM(ITEM_NAME2)) + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName4.Text.ToUpper & "%' "
            End If
        End If

        'Sql = Sql + "And NoDisp <> '1' "

        Sql = Sql + "ORDER BY ITEM DESC, ITEM_NAME1, ITEM_NAME2, ITEM_NAME3 "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "FA000")
        If ds.Tables("FA000").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        Dim xCode As String = GridView1.SelectedRow.Cells(2).Text
        cn.ConnectionString = EDLConnect
        '
        If wBuyer <> "" And wUserID <> "" And xCode <> "" Then
            sql = "SELECT Item "
            sql = sql + "From T_Quotation "
            sql = sql + "Where Buyer = '" & wBuyer & "' "
            sql = sql + "And   Item = '" & xCode & "' "
            sql = sql + "And   Active = 0 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "Quotation")
            If ds.Tables("Quotation").Rows.Count <= 0 Then
                '
                sql = " insert into T_Quotation (Active, Buyer, BuyerItem, Item, CreateUser, CreateTime) "
                sql = sql + "values(0, '" + wBuyer + "','" + wBuyerItem + "','" + xCode + "','" + wUserID + "', getdate()" + ") "
                '
                dc.Connection = cn
                dc.CommandText = sql
                cn.Open()
                dc.ExecuteNonQuery()
                cn.Close()
                '
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        Dim i As Integer
        '
        cn.ConnectionString = EDLConnect
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            '=================================================================
            '共通
            '=================================================================
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = ""
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = ""
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Item Code"
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Item Name-1"
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Item Name-2"
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Item Name-3"
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Slider Code"
            '=================================================================
            'TNF
            '=================================================================
            If wBuyer = "000021" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Season"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "Color Name"
            End If
            '=================================================================
            'ADIDAS
            '=================================================================
            If wBuyer = "000001" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = ""
            End If
            '=================================================================
            'NIKE
            '=================================================================
            If wBuyer = "000013" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = ""
            End If
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            For i = 0 To 11
                If e.Row.Cells(2).Text = wYCode Then
                    e.Row.Cells(2).ForeColor = Color.Red
                End If
                '
                If i = 8 Then
                    '=================================================================
                    'TNF
                    '=================================================================
                    If wBuyer = "000021" Then
                        sql = "SELECT TOP 1 B1, C1, D1 "
                        sql &= "From M_PSCommonData "
                        sql &= "Where Buyer = '" & wBuyer & "' "
                        sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        sql &= "And Active = '0' "
                        '
                        sql &= "And A1 = 'PULLER' "
                        sql &= "And C1 <> '' "
                        sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + C1 + '%' "
                        sql &= "Order by  LEN(C1) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(9).Text = ds.Tables("Color").Rows(0).Item("D1")
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.0.3/ISOS/PS_INQ_Color01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBColor=" & Mid(ds.Tables("Color").Rows(0).Item("C1"), 1, Len(ds.Tables("Color").Rows(0).Item("C1")) - 1)
                        End If
                    End If
                    '=================================================================
                    'ADIDAS
                    '=================================================================
                    If wBuyer = "000001" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(C1)) AS C1, F1, H1 "
                        sql &= "From M_PSCommonData "
                        sql &= "Where Buyer = '" & wBuyer & "' "
                        sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        sql &= "And Active = '0' "
                        '
                        sql &= "And A1 = 'PULLER' "
                        sql &= "And LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) <> '' "
                        sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '%' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("F1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.0.3/ISOS/PS_INQ_Color01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBColor=" & ds.Tables("Color").Rows(0).Item("C1")
                        End If
                    End If
                    '=================================================================
                    'NIKE
                    '=================================================================
                    If wBuyer = "000013" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, D1, H1 "
                        sql &= "From M_PSCommonData "
                        sql &= "Where Buyer = '" & wBuyer & "' "
                        sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        sql &= "And Active = '0' "
                        '
                        sql &= "And A1 = 'PULLER' "
                        sql &= "And LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) <> '' "
                        sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '%' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.0.3/ISOS/PS_INQ_Color01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBColor=" & ds.Tables("Color").Rows(0).Item("B1")
                        End If
                    End If
                End If
                '
                If i = 11 Then
                    Dim h1 As HyperLink = e.Row.Cells(11).Controls(0)
                    h1.NavigateUrl = "http://10.245.1.6/IRW/ItemRegisterSheet_01.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & wUserID & "&pApplyID=" & wUserID & "&pCode=" & e.Row.Cells(2).Text
                End If

            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
End Class
