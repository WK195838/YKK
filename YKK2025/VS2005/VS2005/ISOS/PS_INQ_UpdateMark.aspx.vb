Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_UpdateMark
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
    Dim pBuyer As String             'Buyer
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
        Response.Cookies("PGM").Value = "PS_INQ_UpdateMark.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        '動作按鈕設定
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
        DKEY1.Text = ""
        DKEY2.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBUYER.Text = "TNF"
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "000003" Then DBUYER.Text = "COLUMBIA"
            If pBuyer = "TW1741" Then DBUYER.Text = "LULULEMON"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            If pBuyer = "000098" Then DBUYER.Text = "HELLY HANSEN"
            '
            If pBuyer = "TW0371" Then
                DDataType.Items.Clear()
                '
                Dim ListItem1 As New ListItem
                ListItem1.Text = "ITEM(APP)"
                ListItem1.Value = "BUYERITEM-APP"
                ListItem1.Selected = True
                DDataType.Items.Add(ListItem1)
                '
                Dim ListItem2 As New ListItem
                ListItem2.Text = "ITEM(BAG)"
                ListItem2.Value = "BUYERITEM-BAG"
                DDataType.Items.Add(ListItem2)
                '
                Dim ListItem3 As New ListItem
                ListItem3.Text = "COLOR(TAPE)"
                ListItem3.Value = "BUYERCOLOR-TAPE"
                DDataType.Items.Add(ListItem3)
                '
                Dim ListItem4 As New ListItem
                ListItem4.Text = "COLOR(PULLER)"
                ListItem4.Value = "BUYERCOLOR-PULLER"
                DDataType.Items.Add(ListItem4)
                '
                Dim ListItem5 As New ListItem
                ListItem5.Text = "PRICE(APP)"
                ListItem5.Value = "BUYERPRICE-APP"
                DDataType.Items.Add(ListItem5)
                '
                Dim ListItem6 As New ListItem
                ListItem6.Text = "PRICE(BAG)"
                ListItem6.Value = "BUYERPRICE-BAG"
                DDataType.Items.Add(ListItem6)
            End If
            '
            HBuyerCode.Text = pBuyer
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        'On Error GoTo LBL_Error

        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '
        If pBuyer = "000021" Or pBuyer = "000001" Or pBuyer = "000013" Or pBuyer = "000003" Or pBuyer = "TW1741" Or pBuyer = "TW0371" Or pBuyer = "000098" Then
            sql = "SELECT TOP 300 "
            sql &= "CASE WHEN ACTIVE=0 THEN 'N' ELSE 'O' END AS CAT,  "
            sql &= "ID,A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1,Active, "
            sql &= "WA1,WB1,WC1,WD1,WE1,WF1,WG1,WH1,WI1,WJ1,WK1,WL1,WM1,WN1,WO1,WP1,WQ1,WR1,WS1,WT1,WU1,WV1,WW1,WX1,WY1,WZ1, "
            sql &= "WA1,WB1,WC1,WD1,WE1,WF1,WG1,WH1,WI1,WJ1,WK1,WL1,WM1,WN1,WO1,WP1,WQ1,WR1,WS1,WT1,WU1,WV1,WW1,WX1,WY1,WZ1, "
            '
            sql &= "'IRW' AS IRW, "
            ' TNF
            If HBuyerCode.Text = "000021" Then sql &= "'http://10.245.0.3/ISOS/ISOS2IRW.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pSlider1=' + LTRIM(RTRIM(C1)) as IRWURL "
            ' ADIDAS
            If HBuyerCode.Text = "000001" Then sql &= "'http://10.245.0.3/ISOS/ISOS2IRW.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pSlider1=' + LTRIM(RTRIM(C1)) + LTRIM(RTRIM(D1)) as IRWURL "
            '
            ' <>
            If HBuyerCode.Text <> "000021" And HBuyerCode.Text <> "000001" Then sql &= "'' as IRWURL "
            '
            sql &= "From V_PSCInfUpdateHistory "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            '
            ' NIKE
            If HBuyerCode.Text = "000013" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= "And ( "
                sql &= "     DataType IN ('" & DDataType.SelectedValue & "') "
                sql &= "  OR ColorDataType IN ('BUYERCOLOR-TAPE','BUYERCOLOR-TEETH','BUYERCOLOR-SLDFINISH') "
                sql &= "    ) "
            Else
                sql &= "And ( "
                sql &= "     DataType IN ('" & DDataType.SelectedValue & "') "
                sql &= "  OR ColorDataType IN ('" & DDataType.SelectedValue & "') "
                sql &= "    ) "
            End If
            '
            If DKEY1.Text <> "" Then sql &= " And ID+A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And ID+A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by REPLACE(Buyer,' ','') + '/' + REPLACE(DataType,' ','') + '/' + "
            ' TNF
            If HBuyerCode.Text = "000021" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/' "
            End If
            ' ADIDASA
            If HBuyerCode.Text = "000001" And (DDataType.SelectedValue = "BUYERPRICE" Or DDataType.SelectedValue = "BUYERITEM") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(G1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "000001" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            If HBuyerCode.Text = "000001" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            ' NIKE
            If HBuyerCode.Text = "000013" And (DDataType.SelectedValue = "BUYERPRICE" Or DDataType.SelectedValue = "BUYERITEM") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "000013" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            If HBuyerCode.Text = "000013" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            ' COLUMBIA
            If HBuyerCode.Text = "000003" And (DDataType.SelectedValue = "BUYERPRICE" Or DDataType.SelectedValue = "BUYERITEM") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(E1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "000003" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            If HBuyerCode.Text = "000003" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            ' LULULEMON
            If HBuyerCode.Text = "TW1741" And (DDataType.SelectedValue = "BUYERPRICE" Or DDataType.SelectedValue = "BUYERITEM") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(D1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW1741" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(D1,' ','') + '/'"
            End If
            If HBuyerCode.Text = "TW1741" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(D1,' ','') + '/' + REPLACE(E1,' ','') + '/'"
            End If
            ' UA
            If HBuyerCode.Text = "TW0371" And (DDataType.SelectedValue = "BUYERPRICE-APP" Or DDataType.SelectedValue = "BUYERITEM-APP") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW0371" And (DDataType.SelectedValue = "BUYERPRICE-BAG" Or DDataType.SelectedValue = "BUYERITEM-BAG") Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' + REPLACE(C1,' ','') + '/' + REPLACE(D1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW0371" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(B1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW0371" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(A1,' ','') + '/' + REPLACE(D1,' ','') + '/' + REPLACE(B1,' ','') + '/' "
            End If
            ' HELLY HANSEN
            If HBuyerCode.Text = "000098" Then
                sql &= " '' "
            End If
            'TUMI
            If HBuyerCode.Text = "TW0020" And (DDataType.SelectedValue = "BUYERPRICE" Or DDataType.SelectedValue = "BUYERITEM") Then
                sql &= " REPLACE(F1,' ','') + '/' + REPLACE(G1,' ','') + '/' + REPLACE(H1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW0020" And DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                sql &= " REPLACE(F1,' ','') + '/' + REPLACE(G1,' ','') + '/' + REPLACE(H1,' ','') + '/' "
            End If
            If HBuyerCode.Text = "TW0020" And DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                sql &= " REPLACE(F1,' ','') + '/' + REPLACE(G1,' ','') + '/' + REPLACE(H1,' ','') + '/' "
            End If
            '
            sql &= " ,Active "

            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
            '
            GoTo LBL_End
        End If
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
    '**     Get Search Field
    '**
    '*****************************************************************
    Public Function GetSearchField(ByVal pCmd As String, ByVal pFieldStr As String, ByVal pDataStr As String) As String
        Dim RtnString As String = ""
        Dim str As String
        Dim i As Integer
        Dim FieldNames(), DataNames() As String
        '
        Try
            str = UCase(pCmd)
            FieldNames = pFieldStr.Split("/")
            DataNames = pDataStr.Split("/")
            For i = 0 To DataNames.Length - 1
                If FieldNames(i) <> DataNames(i) Then
                    str = str.Replace(FieldNames(i), DataNames(i))
                End If
            Next
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)
        Dim style As String = "<style> .text { mso-number-format:\@; } </style> " '文字樣式字串
        Response.Clear()
        Response.ContentEncoding = System.Text.Encoding.UTF8 '指定編碼
        Response.AppendHeader("Content-Disposition", "attachment; filename=DiffInformation.xls") '指定匯出檔案名稱
        Response.ContentType = "application/vnd.ms-excel" '指定檔案類型
        '
        ShowData()
        '
        GridView1.RenderControl(hw) '將Gridview Render出來
        Response.Write(style)
        Response.Write(sw.ToString()) '將檔案輸出
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        '=========================================================================
        '共通
        '=========================================================================
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(1).Attributes.Add("class", "text")
        End If
        '
        '=========================================================================
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Base length (in inch, for zipper only)"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Price for basic length (USD/100pcs)"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Price for 1inch up ( USD/100pcs)"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Requested by"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Production place"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Remark1"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Remark2"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "Remark3"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "YKK/Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "Article" & "<BR>" & "(Puller:Tnf Black僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 18 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Or DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Tape/Puller"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "ColorName"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "[T]TWN Color"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "[T]Statsu(P)"
                    tcl(7).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "[T]PBR Color"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "[T]Statsu(PBR)"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "[P]Statsu"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 11 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                        If e.Row.Cells(28).Text = "0" And InStr(e.Row.Cells(10).Text, "APPROVED") > 0 Then
                            If InStr("SL034/MK045/IT003", UCase(UserID)) > 0 Then
                                e.Row.Cells(55).Visible = True
                            End If
                        End If
                    End If
                End If
            End If
            '
            '
            If DDataType.SelectedValue = "BUYERPRICE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Base length (in inch, for zipper only)"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Price for basic length (USD/100pcs)"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Price for 1inch up ( USD/100pcs)"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Requested by"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Production place"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Remark1"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Remark2"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "Remark3"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "YKK/Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "Article" & "<BR>" & "(Puller:Tnf Black僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 18 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Item Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "End Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Kids safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "PLM#"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "YKK ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Base length"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Base (US$/PC)"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "add.1 inch (US$/PC)"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Sample L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "FAS ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "FAS ITEMNAME"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, Chr(10), "<br>")
                    e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, Chr(10), "<br>")

                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 17 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKKColor"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "PBR*D"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "DT*/VT*/CNT*/CFT*" & "<BR>" & "(with P tape)"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "DT*/VT*/CZT*/D*" & "<BR>" & "(with Recycled tape)"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "NM T.Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "YKK S.Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "YKK S.Color " & "<BR>" & "for EAA/EFA/EX Neon color" & "<BR>" & "(SLS-*** / SLS*** / SLS#****)"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "P Tape(VSNT)"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "PBR*D Tape(VSNT)"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Clear Teeth(VSNT)"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "SLS-***(VSNT)"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "Remark(for MKT)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 20 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Puller Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Color Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "核可樣送核日"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 10 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    If e.Row.Cells(28).Text = "0" And InStr(e.Row.Cells(3).Text, "APPROVAL") > 0 Then
                        If InStr("SL034/MK045/IT003", UCase(UserID)) > 0 Then
                            e.Row.Cells(55).Visible = True
                        End If
                    End If
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERPRICE" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Item Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "End Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Kids safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "PLM#"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "YKK ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Base length"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Base (US$/PC)"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "add.1 inch (US$/PC)"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "FAS ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "FAS ITEMNAME"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, Chr(10), "<br>")
                    e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, Chr(10), "<br>")

                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 16 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "PCX"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "IM Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Description(s)"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Size(inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Add 1 inch"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "YPGOLD" & "<BR>" & "(OQG)"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "YKKSHB" & "<BR>" & "(VKL)"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "SHINY ICESIL" & "<BR>" & "(C5V)"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "YKKH6N" & "<BR>" & "(H6N)"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "01M" & "<BR>" & "(H1)"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "OP-O(95K)" & "<BR>" & "OP-H6(95H)"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Sample L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(21).Text = "YKK/IM Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(22).Text = "Item Name" & "<BR>" & "(Puller:00A僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 23 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Color Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Labdip Vendor Code" & "<BR>" & "(YKK code)"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "PBR2D/PBR6D"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "NM(transparent tape)" & "<BR>" & "Labdip Vendor Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "SHA Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "DLN Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Japan Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Teeth"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 13 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Puller"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Slider Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "核可樣送核日"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "核可狀態"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 9 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERPRICE" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "PCX"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "IM Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Description(s)"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Size(inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Add 1 inch"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "YPGOLD" & "<BR>" & "(OQG)"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "YKKSHB" & "<BR>" & "(VKL)"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "SHINY ICESIL" & "<BR>" & "(C5V)"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "YKKH6N" & "<BR>" & "(H6N)"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "01M" & "<BR>" & "(H1)"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "OP-O(95K)" & "<BR>" & "OP-H6(95H)"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Sample L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(21).Text = "YKK/IM Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(22).Text = "Item Name" & "<BR>" & "(Puller:00A僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 23 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKK Dsc"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Base Length"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "SPL w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "SPL w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "PROD w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "PROD w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Unit"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "Currency"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "PROD Loc."
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Add 1 inch"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "MOQ" & "<BR>" & "SPL/PROD"
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(21).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(22).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 23 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "State"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "LabDipCode"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "PBR*D"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "CNT8/10"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "VT8/10"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 11 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Puller"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Slider Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "核可狀態"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 9 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERPRICE" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKK Dsc"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Base Length"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "SPL w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "SPL w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "PROD w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "PROD w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Unit"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "Currency"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "PROD Loc."
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Add 1 inch"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "MOQ" & "<BR>" & "SPL/PROD"
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(21).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(22).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 28
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 23 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Material Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Supplier Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "LT Show to Sales"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "FOB Price in USD"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "FOB Price UOM"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Total Leadtime"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Season Developed For"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    'For i = 0 To 28
                    '    e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    'Next
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 12 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "TWN Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "PBR Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "HK Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "SHA Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "SZN Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "VNM Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "JP Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Korea Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Dalian Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "Soft Clear Zipper NM"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 17 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "核可狀態"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Puller Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Puller Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 9 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
        End If
        '
        '=========================================================================
        'UA
        '=========================================================================
        If pBuyer = "TW0371" Then
            If DDataType.SelectedValue = "BUYERITEM-APP" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKK" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "CPSC"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "SMS" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "BULK" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Length" & "<BR>" & "(Inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Price"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "1(Inch)" & "<BR>" & "Upcharge"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 17 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERITEM-BAG" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "#"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Rebranding" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Rebraning" & "<BR>" & "YKK Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Previous" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Previous" & "<BR>" & "YKK Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "SMS" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "BULK" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Length" & "<BR>" & "(Inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Price"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "1(Inch)" & "<BR>" & "Upcharge"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 19 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-TAPE" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Color Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "SMS ONLY"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "E/EF/EAA/EFA"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "EFR"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "NEON/GENERAL" & "<BR>" & "(BULK-Y COLOR)"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "PBR"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Remarks"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 11 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERCOLOR-PULLER" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Color Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Puller Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Rubber Puller Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Rubber Puller Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 8 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERPRICE-APP" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKK" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "CPSC"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "SMS" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "BULK" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Length" & "<BR>" & "(Inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Price"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "1(Inch)" & "<BR>" & "Upcharge"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 17 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '
            If DDataType.SelectedValue = "BUYERPRICE-BAG" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "#"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Rebranding" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Rebraning" & "<BR>" & "YKK Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Previous" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Previous" & "<BR>" & "YKK Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "SMS" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "BULK" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Length" & "<BR>" & "(Inch)"
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "Price"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "1(Inch)" & "<BR>" & "Upcharge"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "YKK/PDM"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 19 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If

            '
        End If

        '=========================================================================
        'TUMI
        '=========================================================================
        If pBuyer = "TW0020" Then
            If DDataType.SelectedValue = "BUYERITEM" Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Series"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Slider/Chain Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "TUMI Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Item Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Item Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "TUMI Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "SMS LeadTime"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "BULK LeadTime"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    'e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, Chr(10), "<br>")
                    'e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, Chr(10), "<br>")

                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 12 To 16
                        e.Row.Cells(i).Visible = False
                    Next

                    For i = 18 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
            '

            If DDataType.SelectedValue = "BUYERPRICE" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = ""
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Upload Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Series"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Slider/Chain Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "TUMI Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Item Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Item Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "TUMI Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Unit"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "NT$"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "US$"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    'e.Row.Cells(9).Text = Replace(e.Row.Cells(9).Text, Chr(10), "<br>")
                    'e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, Chr(10), "<br>")

                    If e.Row.Cells(28).Text = "0" Then
                        e.Row.Cells(0).BackColor = Color.SpringGreen
                    End If
                    For i = 10 To 11
                        e.Row.Cells(i).Visible = False
                    Next

                    e.Row.Cells(15).Visible = False

                    For i = 17 To 28
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 29 To 55
                        If e.Row.Cells(i).Text <> "0" Then
                            If e.Row.Cells(28).Text = "0" Then
                                If e.Row.Cells(i - 27).Text = "&nbsp;" Then e.Row.Cells(i - 27).Text = "空白"
                                e.Row.Cells(i - 27).ForeColor = Color.Red
                            End If
                        End If
                    Next
                    For i = 29 To 55
                        e.Row.Cells(i).Visible = False
                    Next
                End If

            End If

        End If
        '


        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
