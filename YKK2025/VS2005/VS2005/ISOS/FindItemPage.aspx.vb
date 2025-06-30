Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class FindItemPage
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


    '外部Object
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESDB")
    Dim ConnectStringCL As String = uCommon.GetAppSetting("WAVESDBCL")
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
        If wBuyer = "000098" Or wBuyer = "000068" Then
            DCode.Text = wYCode
        End If
        DName1.Text = ""
        DName2.Text = ""
        DName3.Text = ""
        DName4.Text = ""
        '
        'ItemName(1)
        '
        wItemname = Trim(Replace(Replace(wItemname, "[1]", ""), "[]", ""))
        str = wItemname & " "
        xItemName = str.Split(" ")
        i = 0
        str = ""

        '雙拉頭*保留，PULLER後*去掉
        If (wBuyer = "000055") Then
            xItemName(1) = Replace(xItemName(1), "*", "")
        End If


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
                            For j = 0 To i
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
                If j = i + 1 Then DName2.Text = xItemName(j)
                If j = i + 2 Then DName3.Text = xItemName(j)
                If j = i + 3 Then DName4.Text = xItemName(j)
            Next
        End If
        If (wBuyer = "TW5068") Then
            If UCase(wYCode) = "SLIDER" And InStr(wItemname, "PARTS") <= 0 Then
                DName1.Text = ""
                DName2.Text = ""
                DName3.Text = ""
                DName4.Text = ""
                '
                For j = 0 To xItemName.Length - 1
                    If DName1.Text = "" Then
                        DName1.Text = xItemName(j)
                    Else
                        DName1.Text = DName1.Text & " " & xItemName(j)
                    End If
                Next
                DName4.Text = "BOX"
            End If
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
        Sql = Sql + "LTRIM(RTRIM(ITEM_NAME3)) As Name3, "
        '
        Sql = Sql + "'' As Season, "
        Sql = Sql + "'' As Color, "
        Sql = Sql + "'' As ColorName, "
        Sql = Sql + "LTRIM(RTRIM(SLIDER)) As Slider, "
        '
        '材料
        Sql = Sql + "'材料' as ISOS2FAS, "
        If wBuyer = "TW1741" Or wBuyer = "000098" Or wBuyer = "000068" Or wBuyer = "000151" Or wBuyer = "TW5068" Or wBuyer = "000055" Or wBuyer = "TW0020" Then
            Sql = Sql + "'' as URL, "
        Else
            Sql = Sql + "'http://10.245.0.3/ISOS/ISOS2FAS.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pItem=' + LTRIM(RTRIM(ITEM)) + '&pItemName1=' + LTRIM(RTRIM(ITEM_NAME1)) + '&pItemName2=' + LTRIM(RTRIM(ITEM_NAME2)) + '&pItemName3=' + LTRIM(RTRIM(ITEM_NAME3)) as URL, "
        End If
        '
        '**BCP
        Sql = Sql + "'BCP' as BCP, "
        Sql = Sql + "'' as BCPURL, "
        '
        '**BCT
        Sql = Sql + "'BCT' as BCT, "
        If wBuyer = "000013" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Color01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBColor=' + SUBSTRING(LTrim(RTrim(ChainType)),1,LEN(LTrim(RTrim(ChainType)))-1) + '&pFun=BCT" & "' as BCTURL, "
        Else
            Sql = Sql + "'' as BCTURL, "
        End If
        '
        '**IRW
        Sql = Sql + "'IRW' as IRW, "
        Sql = Sql + "'' as IRWURL, "
        '
        '**SPC
        Sql = Sql + "'SPC' as SPC, "
        If wBuyer = "000021" Or wBuyer = "000098" Or wBuyer = "000068" Or wBuyer = "000151" Or wBuyer = "TW5068" Or wBuyer = "TW0655" Or wBuyer = "000141" Or wBuyer = "000053" Or wBuyer = "000055" Or wBuyer = "TW0020" Then
            Sql = Sql + "'' as SPCURL, "
        End If
        If wBuyer = "000001" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        If wBuyer = "000013" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        If wBuyer = "000003" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        If wBuyer = "TW1741" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPC" & "' as SPCURL, "
        End If
        If wBuyer = "TW0371" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & "" & "&pFun=SPC" & "' as SPCURL, "
        End If
        '
        '**SPP
        Sql = Sql + "'SPP' as SPP, "
        If wBuyer = "000013" Then
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_Special01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=SPP" & "' as SPPURL, "
        Else
            Sql = Sql + "'' as SPPURL, "
        End If
        '
        '**IMG
        Sql = Sql + "'IMG' as IMG, "
        Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_OpenFile.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=' + LTrim(RTrim(ITEM)) + '&pFun=IMG" & "' as IMGURL, "
        '
        '**COMBI
        Sql = Sql + "'COMBI' as COMBI, "
        If wBuyer = "TW1741" Or wBuyer = "000098" Or wBuyer = "000068" Or wBuyer = "000151" Or wBuyer = "TW5068" Or wBuyer = "000141" Or wBuyer = "000053" Or wBuyer = "000055" Or wBuyer = "TW0020" Then
            Sql = Sql + "'' as COMBIURL, "
        Else
            Sql = Sql + "'http://10.245.0.3/ISOS/PS_INQ_COMBI01.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pBuyerItem=" & wBuyerItem & "&pFun=COMBI" & "' as COMBIURL, "
        End If
        '
        '**估價
        Sql = Sql + "'估價' as VAL, "
        Sql = Sql + "'http://10.245.1.6/IRW/ItemValuationSheet_01.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & wUserID & "&pApplyID=" & wUserID & "&pCode=" & "' + LTrim(RTrim(ITEM)) as VALURL, "
        '
        Sql = Sql + "'' as XXX "
        '
        Sql = Sql + "From MST_ITEM "
        '
        If DCode.Text <> "" Then
            Sql = Sql + "Where ITEM LIKE '%" & DCode.Text.ToUpper & "%' "
        Else
            Sql = Sql + "Where REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName1.Text.ToUpper, " ", "") & "%' "
            If DName2.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName2.Text.ToUpper, " ", "") & "%' "
            End If
            If DName3.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName3.Text.ToUpper, " ", "") & "%' "
            End If
            If DName4.Text <> "" Then
                Sql = Sql + "And REPLACE(ITEM_NAME1+ITEM_NAME2+ITEM_NAME3,' ','') LIKE '%" & Replace(DName4.Text.ToUpper, " ", "") & "%' "
            End If
            '
            'Sql = Sql + "Where LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName1.Text.ToUpper & "%' "
            'If DName2.Text <> "" Then
            '    Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName2.Text.ToUpper & "%' "
            'End If
            'If DName3.Text <> "" Then
            '    Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName3.Text.ToUpper & "%' "
            'End If
            'If DName4.Text <> "" Then
            '    Sql = Sql + "And LTRIM(RTRIM(ITEM_NAME1)) + ' ' + LTRIM(RTRIM(ITEM_NAME2)) + ' ' + LTRIM(RTRIM(ITEM_NAME3)) LIKE '%" & DName4.Text.ToUpper & "%' "
            'End If
        End If
        Sql = Sql + "And NoDisp <> '1' "
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
        Dim cn, cn1 As New OleDbConnection
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
                tcl(9).Text = "Color Name" & "<BR>" & "[PFAS]"
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
                tcl(9).Text = "PFAS"
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
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'COLUMBIA
            '=================================================================
            If wBuyer = "000003" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'LULULEMON
            '=================================================================
            If wBuyer = "TW1741" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'UA
            '=================================================================
            If wBuyer = "TW0371" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'HELLY HANSEN
            '=================================================================
            If wBuyer = "000098" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'BURTON
            '=================================================================
            If wBuyer = "000068" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Code"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'PUMA
            '=================================================================
            If wBuyer = "000151" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'HERSCHEL
            '=================================================================
            If wBuyer = "TW5068" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'VERA BRADLEY
            '=================================================================
            If wBuyer = "TW0655" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'PATAGONIA
            '=================================================================
            If wBuyer = "000141" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'ARCTERYX
            '=================================================================
            If wBuyer = "000053" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Name"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'SALOMON
            '=================================================================
            If wBuyer = "000055" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "Puller Name"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "Color Name"
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If
            '=================================================================
            'TUMI
            '=================================================================
            If wBuyer = "TW0020" Then
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = ""
                tcl.Add(New TableHeaderCell())
                tcl(9).Text = "PFAS"
            End If

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            For i = 0 To 17
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
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + C1 + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + C1 + '' "
                        End If
                        'MOD-END
                        sql &= "And C1 <> '' "
                        sql &= "Order by  LEN(C1) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(9).Text = ds.Tables("Color").Rows(0).Item("D1")
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & Mid(ds.Tables("Color").Rows(0).Item("C1"), 1, Len(ds.Tables("Color").Rows(0).Item("C1")) - 1) & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'ADIDAS
                    '=================================================================
                    If wBuyer = "000001" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(C1)) AS C1, F1, H1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("F1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("C1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'NIKE
                    '=================================================================
                    If wBuyer = "000013" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, D1, H1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("B1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'COLUMBIA
                    '=================================================================
                    If wBuyer = "000003" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, C1, D1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("B1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'LULULEMON
                    '=================================================================
                    If wBuyer = "TW1741" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(E1)) AS E1, F1, D1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(E1))+LTRIM(RTRIM(F1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(E1))+LTRIM(RTRIM(F1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(E1))+LTRIM(RTRIM(F1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(E1))+LTRIM(RTRIM(F1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("E1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("E1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'UA
                    '=================================================================
                    If wBuyer = "TW0371" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, C1, D1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(B1))+LTRIM(RTRIM(C1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("B1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'HELLY HANSEN
                    '=================================================================
                    If wBuyer = "000098" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(C1)) AS C1, D1, E1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("E1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("C1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'BURTON
                    '=================================================================
                    If wBuyer = "000068" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(C1)) AS C1, D1, E1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("E1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("C1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'PUMA
                    '=================================================================
                    If wBuyer = "000151" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(C1)) AS C1, D1, E1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(C1)) + '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(C1)) + '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(C1))+LTRIM(RTRIM(D1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("C1")
                            e.Row.Cells(8).Text = ""
                            'e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("E1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("C1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'HERSCHEL
                    '=================================================================
                    If wBuyer = "TW5068" Then
                        'NO DATA
                    End If
                    '=================================================================
                    'VERA BRADLEY
                    '=================================================================
                    If wBuyer = "TW0655" Then
                        'NO DATA
                    End If
                    '=================================================================
                    'PATAGONIA
                    '=================================================================
                    If wBuyer = "000141" Then
                        'NO DATA
                    End If
                    '=================================================================
                    'ARCTERYX
                    '=================================================================
                    If wBuyer = "000053" Then
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, D1, E1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(e.Row.Cells(6).Text, 1, InStr(e.Row.Cells(6).Text, "-") - 1) & "' like '%' + LTRIM(RTRIM(D1)) + LTRIM(RTRIM(E1))+ '' "
                        Else
                            sql &= "And '" & e.Row.Cells(6).Text & "' like '%' + LTRIM(RTRIM(D1)) + LTRIM(RTRIM(E1))+ '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(D1))+LTRIM(RTRIM(E1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(D1))+LTRIM(RTRIM(E1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("D1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'SALOMON
                    '=================================================================
                    If wBuyer = "000055" Then
                        'puller補0'
                        Dim puller As String
                        puller = Replace(e.Row.Cells(6).Text, "SLN24", "SLN024")
                        puller = Replace(puller, "SLN28", "SLN028")
                        puller = Replace(puller, "SLN30", "SLN030")
                        sql = "SELECT TOP 1 LTRIM(RTRIM(B1)) AS B1, D1, E1 "
                        'MOD-START BY JOY 240503
                        sql &= "From V_ISIPPullerColor "
                        'sql &= "From M_PSCommonData "
                        sql &= "Where Buyer like '%" & wBuyer & "|%' "
                        'sql &= "Where Buyer = '" & wBuyer & "' "
                        'sql &= "And DataType = '" & "BUYERCOLOR" & "' "
                        'sql &= "And Active = '0' "
                        'sql &= "And A1 = 'PULLER' "
                        If InStr(e.Row.Cells(6).Text, "-") > 0 Then
                            sql &= "And '" & Mid(puller, 1, InStr(puller, "-") - 1) & "' like '%' + LTRIM(RTRIM(D1)) + LTRIM(RTRIM(E1))+ '' "
                        Else
                            sql &= "And '" & puller & "' like '%' + LTRIM(RTRIM(D1)) + LTRIM(RTRIM(E1))+ '' "
                        End If
                        'MOD-END
                        sql &= "And LTRIM(RTRIM(D1))+LTRIM(RTRIM(E1)) <> '' "
                        sql &= "Order by  LEN(LTRIM(RTRIM(D1))+LTRIM(RTRIM(E1))) DESC "
                        '
                        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                        DBAdapter1.Fill(ds, "Color")
                        If ds.Tables("Color").Rows.Count > 0 Then
                            e.Row.Cells(7).Text = ds.Tables("Color").Rows(0).Item("D1")
                            e.Row.Cells(8).Text = ds.Tables("Color").Rows(0).Item("B1")
                            e.Row.Cells(9).Text = ""
                            Dim h As HyperLink = e.Row.Cells(10).Controls(0)
                            h.NavigateUrl = "http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=" & wUserID & "&pBuyer=" & wBuyer & "&pSlider=" & ds.Tables("Color").Rows(0).Item("D1") & "&pSource=ISOS"
                        End If
                    End If
                    '=================================================================
                    'TUMI
                    '=================================================================
                    If wBuyer = "TW0020" Then
                        'NO DATA
                    End If
                End If
                '
                'JOY
                If i = 9 Then
                    cn1.ConnectionString = ConnectStringCL
                    ds.Clear()
                    '
                    sql = "SELECT ISNULL(IF8CA0,'') AS PFAS From MST_ITEM "
                    sql &= "Where ITEM = '" & Trim(e.Row.Cells(2).Text) & "' "
                    Dim DBAdapter2 As New OleDbDataAdapter(sql, cn1)
                    DBAdapter2.Fill(ds, "ITEM")
                    If ds.Tables("ITEM").Rows.Count > 0 Then
                        If ds.Tables("ITEM").Rows(0).Item("PFAS") <> "" Then
                            e.Row.Cells(9).ForeColor = Color.Red
                            e.Row.Cells(9).BackColor = Color.Yellow
                            e.Row.Cells(9).Text = e.Row.Cells(9).Text & "[" & Trim(ds.Tables("ITEM").Rows(0).Item("PFAS")) & "]"
                        End If
                    End If
                    '
                    cn1.Close()
                End If
                '
                If i = 11 Then
                    Dim h1 As HyperLink = e.Row.Cells(12).Controls(0)
                    h1.NavigateUrl = "http://10.245.1.6/IRW/ItemRegisterSheet_03.aspx?pFormNo=001151&pFormSno=0&pStep=1&pSeqNo=0&pUserID=" & wUserID & "&pApplyID=" & wUserID & "&pCode=" & e.Row.Cells(2).Text
                End If
                '
                'If i = 17 Then
                '    If InStr("SL034/MK045/IT003", UCase(wUserID)) > 0 Then
                '    Else
                '        e.Row.Cells(i).Visible = False
                '    End If
                'End If
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=FindItemPage.xls")     '程式別不同

        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        DataList()
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
