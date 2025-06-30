Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_Special01
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
    Dim pBuyerItem As String        'BuyerItem
    Dim UserID As String            'UserID
    Dim pFun As String              'Fun
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
        Response.Cookies("PGM").Value = "PS_INQ_ListPrice01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        pBuyerItem = Request.QueryString("pBuyerItem")
        pFun = Request.QueryString("pFun")
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        LMCU.Style("left") = -500 & "px"
        '
        If pBuyer = "TW1741" Then LMCU.Style("left") = 550 & "px"
        '動作按鈕設定
        '
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
        DKEY3.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "TW1741" Then DBUYER.Text = "LULULEMON"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            '
            HBuyerCode.Text = pBuyer
        End If
        '
        If pBuyerItem <> "" Then DKEY1.Text = pBuyerItem
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
        On Error GoTo LBL_Error

        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '篩選資料
        '
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            sql = "SELECT top 300 "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '*
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If pFun = "SPC" Then
                sql &= "And DataType IN ('BUYERSPCHAIN') "
            End If
            If pFun = "SPP" Then
                sql &= "And DataType IN ('BUYERSPPULLER') "
            End If
            '
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            If DKEY3.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY3.Text & "%' "
            '*
            sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            sql = "SELECT top 300 "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '*
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If pFun = "SPC" Then
                sql &= "And DataType IN ('BUYERSPCHAIN') "
            End If
            If pFun = "SPP" Then
                sql &= "And DataType IN ('BUYERSPPULLER') "
            End If
            '
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            If DKEY3.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY3.Text & "%' "
            '*
            sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            '
            sql = "SELECT top 300 "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '*
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If pFun = "SPC" Then
                sql &= "And DataType IN ('BUYERSPCHAIN') "
            End If
            If pFun = "SPP" Then
                sql &= "And DataType IN ('BUYERSPPULLER') "
            End If
            '
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            If DKEY3.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY3.Text & "%' "
            '*
            sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '=========================================================================
        'UNDER ARMOUR
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            sql = "SELECT top 300 "
            sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '*
            sql &= "From M_PSCommonData "
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            sql &= "And Active = 0 "
            '
            If pFun = "SPC" Then
                sql &= "And DataType IN ('BUYERSPCHAIN') "
            End If
            If pFun = "SPP" Then
                sql &= "And DataType IN ('BUYERSPPULLER') "
            End If
            '
            If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
            If DKEY3.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY3.Text & "%' "
            '*
            sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
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
        '
        Response.AppendHeader("Content-Disposition", "attachment;filename=BUYERSPCHAIN.xls")     '程式別不同
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            If pFun = "SPC" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "PLM#"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Chain Code"

                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Tape" & "<BR>" & "Color"
                    tcl(3).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Tape" & "<BR>" & "YKK色號"
                    tcl(4).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Element" & "<BR>" & "Color"
                    tcl(5).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Element" & "<BR>" & "YKK色號"
                    tcl(6).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Under Yarn" & "<BR>" & "Left Stitching Color"
                    tcl(7).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Under Yarn" & "<BR>" & "Right Stitching Color"
                    tcl(8).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Upper Yarn" & "<BR>" & "Left Stitching Color"
                    tcl(9).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "Upper Yarn" & "<BR>" & "Right Stitching Color"
                    tcl(10).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "Stripe" & "<BR>" & "Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "Printed Logo" & "<BR>" & "Color"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(2).ForeColor = Color.Red
                    For i = 0 To 12
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    For i = 13 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            If pFun = "SPP" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "核可樣送核日"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "核可狀態"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "IM#"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "GCW"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Color Code Base"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Color Code Logo"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Slider Finish"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Puller"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Slider Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(2).ForeColor = Color.Red
                    For i = 0 To 9
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    For i = 10 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '=========================================================================
        'LULULEMON
        '=========================================================================
        If pBuyer = "TW1741" Then
            If pFun = "SPC" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = " Zipper Code"
                    tcl(0).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "#"
                    tcl(1).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Size"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Colour 1" & "<BR>" & "Thread (Upper/lower)"
                    tcl(3).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Corresponding" & "<BR>" & "lulu colour"
                    tcl(4).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "MCU code"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "Colour 2" & "<BR>" & "Tape Colour"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "YKK Color"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Colour Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "Supplier" & "<BR>" & "Colour Reference"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(5).ForeColor = Color.Red
                    e.Row.Cells(7).ForeColor = Color.Red
                    For i = 0 To 9
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    For i = 10 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If

        '=========================================================================
        'UNDER ARMOUR
        '=========================================================================
        If pBuyer = "TW0371" Then
            If pFun = "SPC" Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "Size"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "Reverse/Normal"

                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "Tape Color"
                    tcl(2).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "Upper Thread"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "Teeth Color"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "Under Thread"
                    tcl(5).BackColor = Color.Green

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "PK Code"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "Teeth Code"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(6).ForeColor = Color.Red
                    e.Row.Cells(7).ForeColor = Color.Red
                    For i = 0 To 8
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    For i = 9 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                End If
            End If
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
