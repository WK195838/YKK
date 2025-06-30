Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class ISOSGetColorFA300
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pBuyer As String            'Buyer
    Dim UserID As String            'UserID
    Dim pColor As String            'Color
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
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")
        pColor = Request.QueryString("pColor")
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
        '
        Dim i As Integer
        For i = Len(pColor) To 4
            pColor = " " & pColor
        Next
        'MsgBox("[" & pColor & "]")
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
        Dim ds As New DataSet
        On Error GoTo LBL_Error
        '
        If pColor <> "" Then
            cn.ConnectionString = ConnectString
            '
            Sql = "select "
            Sql &= "CTBNA3 AS A1, "
            Sql &= "CCLCA3 AS B1  "
            Sql &= "FROM WAVEDLIB.FA300 "
            '
            Sql &= "WHERE PACCA3 = '" & pColor & "' "
            Sql &= "AND   ST1CA3 = '1' AND ST6CA3 = 'P' AND SCSVA3 = '1' "
            '
            Sql &= "GROUP BY CTBNA3, CCLCA3 "
            Sql &= "ORDER BY CTBNA3, CCLCA3 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
            DBAdapter1.Fill(ds, "FA300")
            If ds.Tables("FA300").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
            End If
            '
            cn.Close()
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
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "Color Table"
            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "Slider Color"
            tcl(1).BackColor = Color.Blue
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(1).ForeColor = Color.Red
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub

End Class
