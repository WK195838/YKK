Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class ManySupplierList
    Inherits System.Web.UI.Page
    '
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
    Dim oWaves As New WAVES.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pPuller As String
    Dim pColor As String
    Dim xPuller As String
    Dim xColor As String
    Dim pFun As String

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
        Response.Cookies("PGM").Value = "ManySupplierList.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pPuller = Request.QueryString("pPuller")
        pColor = Request.QueryString("pColor")
        pFun = Request.QueryString("pFun")
        '
        xPuller = ""
        xColor = ""
        If InStr("/5N18/6N18/7N18/", "/" & pPuller & "/") > 0 Then
            xPuller = "5N18"
        End If
        '
        If InStr("/N5N18/N6N18/N7N18/", "/" & pPuller & "/") > 0 Then
            xPuller = "N5N18"
        End If
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
        '
    End Sub

    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim xEDXHave As Boolean
        '
        xEDXHave = False
        SQL = "SELECT "
        SQL = SQL & "Cat "
        SQL = SQL & "FROM V_ManySupplierList "
        If pFun = "QC5Y" Then
            SQL = SQL & "where CAT IN ('QC5Y') AND puller+COLOR = '" & pPuller + pColor & "' "
        Else
            SQL = SQL & "where CAT IN ('EDX') AND puller+COLOR = '" & pPuller + pColor & "' "
        End If
        Dim dt_EDXCount As DataTable = uDataBase.GetDataTable(SQL)
        If dt_EDXCount.Rows.Count > 0 Then
            xEDXHave = True
        End If
        '
        If xEDXHave = True Then
            If pFun = "QC5Y" Then
                SQL = "SELECT "
                SQL = SQL & "Cat, MMSDESC,Puller, Color, SUPPLIER, Sts, NO, NOUrl, AppDate, RDiMAGES, EDXIMAGES,EDXIMGUrl, "
                SQL = SQL & "(select top 1 Remark from M_PullerColor where puller+color = V_ManySupplierList.puller+V_ManySupplierList.color) As Remark "
                SQL = SQL & "FROM V_ManySupplierList "
                SQL = SQL & "where (CAT='R&D' AND ( puller = '" & pPuller & "' OR PULLER = '" & xPuller & "') ) "
                'SQL = SQL & "where (CAT='R&D' AND puller = '" & pPuller & "') "
                SQL = SQL & "or (CAT IN ('QC5Y') AND puller+COLOR = '" & pPuller + pColor & "') "
                SQL = SQL & "or (CAT='-') "
                SQL = SQL & "ORDER BY Seqno, AppDate DESC "
            Else
                SQL = "SELECT "
                SQL = SQL & "Cat, MMSDESC,Puller, Color, SUPPLIER, Sts, NO, NOUrl, AppDate, RDiMAGES, EDXIMAGES,EDXIMGUrl, "
                SQL = SQL & "(select top 1 Remark from M_PullerColor where puller+color = V_ManySupplierList.puller+V_ManySupplierList.color) As Remark "
                SQL = SQL & "FROM V_ManySupplierList "
                'SQL = SQL & "where (CAT='R&D' AND puller = '" & pPuller & "') "
                SQL = SQL & "where (CAT='R&D' AND ( puller = '" & pPuller & "' OR PULLER = '" & xPuller & "') ) "
                SQL = SQL & "or (CAT IN ('EDX') AND puller+COLOR = '" & pPuller + pColor & "') "
                SQL = SQL & "or (CAT='-') "
                SQL = SQL & "ORDER BY Seqno, AppDate DESC "
            End If
        Else
            SQL = "SELECT "
            SQL = SQL & "Cat, MMSDESC,Puller, Color, SUPPLIER, Sts, NO, NOUrl, AppDate, RDiMAGES, EDXIMAGES,EDXIMGUrl, "
            SQL = SQL & "(select top 1 Remark from M_PullerColor where puller+color = V_ManySupplierList.puller+V_ManySupplierList.color) As Remark "
            SQL = SQL & "FROM V_ManySupplierList "
            SQL = SQL & "where (CAT='R&D' AND ( puller = '" & pPuller & "' OR PULLER = '" & xPuller & "') ) "
            SQL = SQL & "ORDER BY Seqno, AppDate DESC "
        End If
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated

        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            '
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Cat." & "<BR>" & ""
            tcl(1).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Puller" & "<BR>" & "Code"
            tcl(2).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Color" & "<BR>" & "Code"
            tcl(3).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "Remark" & "<BR>" & ""
            tcl(4).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(5).Text = "Supplier" & "<BR>" & ""
            tcl(5).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(6).Text = "Status" & "<BR>" & ""
            tcl(6).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(7).Text = "No" & "<BR>" & ""
            tcl(7).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(8).Text = "App.d" & "<BR>" & ""
            tcl(8).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(9).Text = "R&D" & "<BR>" & "Images"
            tcl(9).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(10).Text = "EDX" & "<BR>" & "Images"
            tcl(10).BackColor = Color.Green
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If e.Row.Cells(0).Text = "模廢" Then
                e.Row.Cells(0).BackColor = Color.Black
                e.Row.Cells(0).ForeColor = Color.White
            End If
            e.Row.Cells(2).Font.Bold = True
            e.Row.Cells(3).Font.Bold = True
            e.Row.Cells(4).Font.Bold = True
            e.Row.Cells(5).ForeColor = Color.Red
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
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
            If e.Row.Cells(1).Text = "EDX" Or e.Row.Cells(1).Text = "QC5Y" Then
                e.Row.Cells(9).Text = ""
            End If
            If e.Row.Cells(1).Text = "R&amp;D" Then
                e.Row.Cells(4).Text = ""
            End If
            If e.Row.Cells(1).Text = "-" Then
                For i = 0 To 10
                    Select Case i
                        Case Is >= 1
                            e.Row.Cells(i).Text = "==========="
                        Case Else
                            e.Row.Cells(i).Text = ""
                    End Select
                    e.Row.Cells(i).Font.Size = 8
                    e.Row.Cells(i).ForeColor = Color.Black
                    e.Row.Cells(i).Font.Bold = True

                    'e.Row.Cells(i).BackColor = Color.Red
                Next
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     R&D Images
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim lnk As HyperLink = GridView1.Rows(e.RowIndex).Cells(7).Controls(0)
        '
        SQL = "select top 1 [ImagePath] as Path "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where No = '" & Replace(lnk.Text, "&nbsp;", "") & "' "
        SQL = SQL & "Order by createtime "
        Dim dt_RDImage As DataTable = uDataBase.GetDataTable(SQL)
        If dt_RDImage.Rows.Count > 0 Then
            DRDImage.ImageUrl = dt_RDImage.Rows(0).Item("Path")
            DRDImage.Visible = True
        End If
        '
        ShowData()

    End Sub

End Class
