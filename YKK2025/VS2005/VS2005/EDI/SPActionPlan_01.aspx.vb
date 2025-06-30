Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPActionPlan_01
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pItem As String
    Dim pColor As String
    Dim pKeep As String

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uEDIMapping As New EDI2011.MappingService
    Dim uEDICommon As New EDI2011.CommonService
    Dim uWFSCommon As New WFS.CommonService
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
        Response.Cookies("PGM").Value = "SPActionPlan_01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pItem = Request.QueryString("pItem")
        pColor = Request.QueryString("pColor")
        pKeep = Request.QueryString("pKeep")
        '
        '動作按鈕設定
        '
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        GridView1.Visible = False
        '
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "select "
        SQL = SQL & "ItemName1+' '+ItemName2 as ItemName, "
        SQL = SQL & "[Keep],[Item],[Color], "
        SQL = SQL & "[SPNo],[SPSubNo],[C_Code],[C_Color],[C_Special], "
        SQL = SQL & "[C_KeepCode],[C_Season],[C_Cust], "
        SQL = SQL & "[Minimum],[SchePQ],[OnPQ],[FreeQ],[KeepQ],[TotalQ], "
        SQL = SQL & "[N_FC],[N_Action],[N1_FC],[N1_Action],[N2_FC],[N2_Action],[N3_FC],[N3_Action],[N4_FC],[N4_Action], "
        '
        SQL = SQL & "'....' as PIL, '\\10.245.0.205\www$\EDI\Document\SP\PIL\Shopping_SHA_23Q4_Nov.xlsx' as PILUrl "
        '
        SQL = SQL & "from W_SPSActionPlan "
        ' code
        If pItem <> "" Then
            SQL = SQL & "where item = '" & pItem & "' "
            '
            ' keep
            If pKeep <> "" Then
                SQL = SQL & "and keep = '" & pKeep & "' "
            End If
            ' Color
            If pColor <> "" Then
                SQL = SQL & "and ltrim(rtrim(Color)) = '" & pColor & "' "
            End If
        Else
            SQL = SQL & "where Item = 'zzzzzzzz' "
        End If
        '
        SQL = SQL & "Order by [SPNo],[SPSubNo] "
        '
        Dim dt As DataTable = uDataBase.GetDataTable(SQL)
        If dt.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            '** 4 line *****************
            i = 0
            'Item
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'Name
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Name"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'Color
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Color"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            'keep
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Keep"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            '-----
            'Cust Inf.
            '[SPNo]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP No."
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[SPSubNo]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Sub."
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_Code]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Code"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_Color]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Color"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_Special]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Other"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_KeepCode]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "C.Keep"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_Season]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Season"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '[C_Cust]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "C.Code"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '-----
            '[SchePQ]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Sche P."
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '[OnPQ]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "On P."
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '[FreeQ]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Free"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '[KeepQ]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Keep"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '[TotalQ]
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Total"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '-----
            'FC-0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "FC"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            'FC-1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "FC"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            'FC-2
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "FC"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            'FC-3
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "FC"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            'FC-4
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "FC"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '-----
            'PIL
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "PIL"
            tcl(i).BackColor = Color.White
            tcl(i).ForeColor = Color.Black
            i = i + 1
            '
            '** 3 line *****************
            Dim H3tc1 As TableCell = New TableCell
            H3tc1.Text = "Material"
            H3tc1.ColumnSpan = 4
            H3tc1.BackColor = Color.Blue
            H3row.Cells.Add(H3tc1)

            Dim H3tc2 As TableCell = New TableCell
            H3tc2.Text = "Customer Inf."
            H3tc2.ColumnSpan = 8
            H3tc2.BackColor = Color.Black
            H3row.Cells.Add(H3tc2)

            Dim H3tc3 As TableCell = New TableCell
            H3tc3.Text = "Stock Inf."
            H3tc3.ColumnSpan = 5
            H3tc3.BackColor = Color.Purple
            H3row.Cells.Add(H3tc3)

            Dim H3tc4 As TableCell = New TableCell
            H3tc4.Text = "N"
            H3tc4.ColumnSpan = 2
            H3tc4.BackColor = Color.Green
            H3row.Cells.Add(H3tc4)

            Dim H3tc5 As TableCell = New TableCell
            H3tc5.Text = "N+1"
            H3tc5.ColumnSpan = 2
            H3tc5.BackColor = Color.Green
            H3row.Cells.Add(H3tc5)

            Dim H3tc6 As TableCell = New TableCell
            H3tc6.Text = "N+2"
            H3tc6.ColumnSpan = 2
            H3tc6.BackColor = Color.Green
            H3row.Cells.Add(H3tc6)

            Dim H3tc7 As TableCell = New TableCell
            H3tc7.Text = "N+3"
            H3tc7.ColumnSpan = 2
            H3tc7.BackColor = Color.Green
            H3row.Cells.Add(H3tc7)

            Dim H3tc8 As TableCell = New TableCell
            H3tc8.Text = "N+4"
            H3tc8.ColumnSpan = 2
            H3tc8.BackColor = Color.Green
            H3row.Cells.Add(H3tc8)

            Dim H3tc9 As TableCell = New TableCell
            H3tc9.Text = ""
            H3tc9.ColumnSpan = 1
            H3tc9.BackColor = Color.White
            H3row.Cells.Add(H3tc9)

            '
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 顏色+格式
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            '
            For i = 0 To 27
                Select Case i
                    Case 0 To 3
                        If i = 0 Then
                            e.Row.Cells(i).BackColor = Color.Blue
                            e.Row.Cells(i).ForeColor = Color.White
                        End If
                    Case 4 To 11
                        If i = 4 Then
                            e.Row.Cells(i).BackColor = Color.Black
                            e.Row.Cells(i).ForeColor = Color.White
                        End If
                    Case 12 To 16
                        If i = 12 Then
                            e.Row.Cells(i).BackColor = Color.Purple
                            e.Row.Cells(i).ForeColor = Color.White
                        End If
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                    Case 17, 19, 21, 23, 25
                        If i = 17 Then
                            e.Row.Cells(i).BackColor = Color.Green
                            e.Row.Cells(i).ForeColor = Color.White
                        End If
                        e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                    Case Else
                End Select
            Next
            '
            'e.Row.Cells(i).Attributes.Add("style", "border:1px solid black ")
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub


End Class
