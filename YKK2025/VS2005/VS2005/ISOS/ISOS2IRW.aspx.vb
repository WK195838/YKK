Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class ISOS2IRW
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
    Dim pBuyer As String            'Buyer
    Dim pSlider1 As String          'Slider1
    Dim pSlider2 As String          'Slider2
    Dim UserID As String            'UserID
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
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
        Response.Cookies("PGM").Value = "ISOS2IRW.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        pSlider1 = Request.QueryString("pSlider1")
        pSlider2 = Request.QueryString("pSlider2")
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        GridView3.Visible = False
        DBUYER.ReadOnly = True
        DSlider1.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        HBuyerColor.Style("left") = -500 & "px"
        HIRW.Style("left") = -500 & "px"
        HBuyerItem.Style("left") = -500 & "px"
        HSales.Style("left") = -500 & "px"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBUYER.Text = ""
        HBuyerCode.Text = ""
        DSlider1.Text = pSlider1
        '
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBUYER.Text = "TNF"
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "000003" Then DBUYER.Text = "COLUMBIA"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            HBuyerCode.Text = "FALL-" & pBuyer
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '
        '-- BUYER COLOR (ISOS)
        'TNF
        If pBuyer = "000021" Then
            sql = "SELECT TOP 300 "
            sql &= "B1 AS A1, "
            sql &= "C1 AS B1, "
            sql &= "D1 AS C1, "
            sql &= "'' AS D1, "
            sql &= "I1 AS E1, "
            sql &= "'開發' AS Z1 "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType = '" & "BUYERCOLOR" & "' "
            sql &= "And Active = '0' "
            sql &= "And A1 = '" & "PULLER" & "' "
            '
            sql &= "And C1 = '" & pSlider1 & "' "
            '
            sql &= " Order by Unique_ID Desc "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "Puller")
            If ds.Tables("Puller").Rows.Count > 0 Then
                GridView4.Visible = True
                GridView4.DataSource = ds
                GridView4.DataBind()
            Else
                HBuyerColor.Style("left") = 24 & "px"
            End If
        End If
        'ADIDAS
        If pBuyer = "000001" Then
            sql = "SELECT TOP 300 "
            sql &= "'ALL' AS A1, "
            sql &= "C1+D1 AS B1, "
            sql &= "E1 AS C1, "
            sql &= "G1 AS D1, "
            sql &= "B1 AS E1, "
            sql &= "'開發' AS Z1 "
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType = '" & "BUYERCOLOR" & "' "
            sql &= "And Active = '0' "
            sql &= "And A1 = '" & "PULLER" & "' "
            '
            sql &= "And C1+D1 = '" & pSlider1 & "' "
            '
            sql &= " Order by Unique_ID Desc "
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "Puller")
            If ds.Tables("Puller").Rows.Count > 0 Then
                GridView4.Visible = True
                GridView4.DataSource = ds
                GridView4.DataBind()
            Else
                HBuyerColor.Style("left") = 24 & "px"
            End If
        End If
        '
        '--ITEM REGISTER (IRW)
        sql = "SELECT top 300 "
        sql &= "ItemName1 + ' ' + ItemName2 + ' ' + ItemName3 AS A1, "
        sql &= "Code AS B1, "
        sql &= "Size AS C1, "
        sql &= "Family AS D1, "
        sql &= "Slider1 AS E1, "
        sql &= "Finish1 AS F1, "
        sql &= "No AS G1, Name AS H1, Sts AS I1, ApplyDate AS J1, "
        '
        sql &= "'Ｆ' as FormMark, "
        sql &= "FURL AS FormURL, "
        sql &= "'ITEM申請' AS Z1 "
        '
        sql &= "From V_Develop_Digital_IRW "
        '
        sql &= "Where BuyerCode = '" & pBuyer & "' "
        sql &= "And  SUBSTRING(SYS,1,3) = 'IRW' "
        'sql &= "And  sts = '完成' "
        '
        If DSlider1.Text <> "" Then sql &= " And ItemName1+ItemName2+ItemName3 Like '%" & DSlider1.Text & "%' "
        '
        '*
        sql &= " Order by ApplyDate DESC,Code,ItemName1,ItemName2 "
        '
        Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
        DBAdapter2.Fill(ds1, "IRW")
        If ds1.Tables("IRW").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds1
            GridView1.DataBind()
        Else
            HIRW.Style("left") = 24 & "px"
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GRIDVIEW1)
    '**     展開-取得販賣 & BUYER ITEM
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds, ds1 As New DataSet
        Dim dc As New OleDbCommand
        Dim sql, xSlider, oSlider, xFinish, xSize, xSize1, xFamily, str As String
        Dim i As Integer
        '
        cn.ConnectionString = EDLConnect
        '
        i = Len(GridView1.SelectedRow.Cells(6).Text)
        oSlider = Replace(GridView1.SelectedRow.Cells(6).Text, " ", "")
        str = Replace(GridView1.SelectedRow.Cells(6).Text, " ", "")
        Do Until i < 0
            If Mid(str, i, 1) >= "0" And Mid(str, i, 1) <= "9" Then
                xSlider = Mid(str, 1, i)
                Exit Do
            End If
            i = i - 1
        Loop
        xSize = GridView1.SelectedRow.Cells(4).Text
        xFamily = GridView1.SelectedRow.Cells(5).Text
        xSlider = xSlider & "A"
        xFinish = GridView1.SelectedRow.Cells(7).Text
        '        
        ShowData()
        '
        xSize1 = "-" & xSize
        If InStr(xSize1, "-0") > 0 Then xSize1 = Replace(xSize1, "-0", "-")
        '
        '--BUYER ITEM
        If pBuyer = "000021" Then
            sql = "SELECT TOP 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(P1,'+','*') + '&pYCode=' + O1 as URL, "
            sql &= "P1 AS ZZ1, A1 AS ZZ2, B1 AS ZZ3, C1 AS ZZ4, D1 AS ZZ5, O1 AS ZZ6, "
            sql &= "'' AS ZZ "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "@@BUYERITEM-B" & "') "
            sql &= "And Active = '0' "
            '
            'sql &= "And LEN(O1) = 7 And P1 LIKE '%[ ]%' "
            'sql &= " And P1 Like '%" & xSlider & " " & xFinish & "%' "
            sql &= " And P1 Like '%" & xSize1 & "%' "
            sql &= " And P1 Like '%" & xSlider & "%' "
            sql &= " And P1 Like '%" & xFinish & "%' "
            '
            sql &= " Order by P1, A1,B1,C1,D1,O1 "
        End If
        If pBuyer = "000001" Then
            sql = "SELECT TOP 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(N1,'+','*') + '&pYCode=' + M1 as URL, "
            sql &= "N1 AS ZZ1, B1 AS ZZ2, D1 AS ZZ3, F1 AS ZZ4, G1 AS ZZ5, M1 AS ZZ6, "
            sql &= "'' AS ZZ "
            '
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & pBuyer & "' "
            sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "@@BUYERITEM-B" & "') "
            sql &= "And Active = '0' "
            '
            sql &= " And N1 Like '%" & xSize1 & "%' "
            sql &= " And N1 Like '%" & xSlider & "%' "
            sql &= " And N1 Like '%" & xFinish & "%' "
            '
            sql &= " Order by N1, B1,D1,F1,G1,M1 "
        End If
        '
        Dim DBAdapter2 As New OleDbDataAdapter(sql, cn)
        DBAdapter2.Fill(ds1, "item")
        If ds1.Tables("item").Rows.Count > 0 Then
            GridView2.Visible = True
            GridView2.DataSource = ds1
            GridView2.DataBind()
        Else
            HBuyerItem.Style("left") = 24 & "px"
        End If
        '
        '--SALES
        sql = "SELECT top 300 "
        sql &= "'IRW' as Mark, "
        sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + '&pItemName=' + Replace(ITEMNAME,'+','*') as URL, "

        sql &= "A.ItemName AS A1, "
        '
        sql &= "SUBSTRING( "
        sql &= "  (  select '/' + B.CUSTOMERNAME + '(' + RTRIM(B.CUSTOMER) + ')' from [EDLINK].[dbo].[V_SalesData_Digital_IRW] B "
        sql &= "     WHERE B.ITEMNAME=A.ITEMNAME "
        '
        If DCust.Text <> "" Then
            sql &= " And ( B.CUSTOMERNAME Like '%" & DCust.Text & "%' OR B.CUSTOMER Like '%" & DCust.Text & "%' ) "
        End If
        '
        sql &= "     GROUP BY B.CUSTOMERNAME, B.CUSTOMER "
        sql &= "     FOR XML PATH('') "
        sql &= "  ),2 ,999 "
        sql &= ") AS B1, "
        '
        sql &= "SUM(A.[QTY]) AS C1, "
        sql &= "'' AS ZZ "
        '
        sql &= "From V_SalesData_Digital_IRW A "
        sql &= "Where A.Buyer = '" & pBuyer & "' "
        '
        If xSlider <> "" Then
            sql &= " And ( A.ITEMNAME Like '%" & xSlider & " " & xFinish & "%' OR A.ITEMNAME Like '%" & oSlider & " " & xFinish & "%' ) "
            sql &= " And ( A.SIZE = '" & xSize & "' ) "
            sql &= " And ( A.FMLCA0 = '" & xFamily & "' ) "
        End If
        '
        If DCust.Text <> "" Then
            sql &= " And ( A.CUSTOMERNAME Like '%" & DCust.Text & "%' OR A.CUSTOMER Like '%" & DCust.Text & "%' ) "
        End If
        '
        sql &= "GROUP BY A.ITEMNAME "
        sql &= "Order by SUM(A.[QTY]) DESC "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
        DBAdapter1.Fill(ds, "SALES")
        If ds.Tables("SALES").Rows.Count > 0 Then
            GridView3.Visible = True
            GridView3.DataSource = ds
            GridView3.DataBind()
        Else
            HSales.Style("left") = 24 & "px"
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(GRIDVIEW1)
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = ""
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = ""
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "ItemName"
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Code"
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Size"
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Family"
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Slider"
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "Finish"
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "IRW NO."
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "申請者"
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "Status"
            tcl.Add(New TableHeaderCell())
            tcl(11).Text = "申請日"
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(GRIDVIEW2)
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        'TNF
        If pBuyer = "000021" Then
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = ""
                tcl(0).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "Article" & "<BR>" & "(Puller:Tnf Black僅供參考)"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "Season"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "Cat."
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "Type"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "Web Code"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "YKK/Web Code"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                If Len(e.Row.Cells(6).Text) = 7 Then
                    e.Row.Cells(0).Text = "[BUYER]+[FCT]"
                Else
                    e.Row.Cells(0).Text = "BUYER"
                End If
                e.Row.Cells(1).Text = Replace(e.Row.Cells(1).Text, "[]", "")
                e.Row.Cells(7).Visible = False
            End If
        End If
        'ADIDAS
        If pBuyer = "000001" Then
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = ""
                tcl(0).BackColor = Color.Blue
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "YKK ITEM" & "<BR>" & "(Puller:095A僅供參考)"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "Season"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "Iten Status"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "Kids safe"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "PLM#"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "YKK/PLM"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                If Len(e.Row.Cells(6).Text) = 7 Then
                    e.Row.Cells(0).Text = "[BUYER]+[FCT]"
                Else
                    e.Row.Cells(0).Text = "BUYER"
                End If
                e.Row.Cells(1).Text = Replace(e.Row.Cells(1).Text, "[]", "")
                e.Row.Cells(7).Visible = False
            End If
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(GRIDVIEW3)
    '**
    '*****************************************************************
    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "販賣"
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Item Name"
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Customer"
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Qty"
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).Text = CStr(e.Row.RowIndex + 1)
            e.Row.Cells(0).ForeColor = Color.Red
            e.Row.Cells(0).HorizontalAlign = HorizontalAlign.Center
            '
            e.Row.Cells(3).Text = String.Format("{0:N0}", CDbl(e.Row.Cells(3).Text))
            e.Row.Cells(3).HorizontalAlign = HorizontalAlign.Right
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料(GRIDVIEW1)
    '**
    '*****************************************************************
    Protected Sub GridView4_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = ""
            tcl(0).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Season"
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Puller"
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "ColorName"
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "核可樣送核日"
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Statsu"
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------

End Class

