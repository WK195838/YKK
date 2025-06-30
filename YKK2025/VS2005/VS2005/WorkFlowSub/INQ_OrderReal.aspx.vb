Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_OrderReal
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     全域變數  
    '**     CALL WFS890R1    SQLRPG      ORDER REAL S8G00 JOY
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間

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
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
        GridView2.Visible = False
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
        Dim ds, ds1 As New DataSet
        '

        cn.ConnectionString = ConnectString
        Sql = "select "
        Sql &= "ODRG8G As YYMM, "
        Sql &= "ROUND(sum(DORK8G)/1000,0) as AMOUNT "
        Sql &= "FROM SONGOLIB.S8G00WK "
        Sql &= "GROUP BY ODRG8G "
        Sql &= "ORDER BY ODRG8G "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "S8G00")
        If ds.Tables("S8G00").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        '----------------------------------------------------------
        If DCat.SelectedValue = "MONTH" Then
            Sql = "select "
            Sql &= "RSMC8G As SUMCODE, "
            Sql &= "ODRG8G As YYMM, "
            Sql &= "'' As DAY, "
            Sql &= "ROUND(sum(DORK8G)/1000,0) as AMOUNT, "
            Sql &= "'Buyer' as Buyer, "
            Sql &= "'Sales' as Sales "
            Sql &= "FROM SONGOLIB.S8G00WK "
            '
            If DSum.SelectedValue <> "00" Then
                Sql &= "where RSMC8G = '" & DSum.SelectedValue & "' "
            End If
            '
            Sql &= "GROUP BY RSMC8G, ODRG8G "
            Sql &= "ORDER BY RSMC8G, ODRG8G "
        Else
            Sql = "select "
            Sql &= "RSMC8G As SUMCODE, "
            Sql &= "ODRG8G As YYMM, "
            Sql &= "OCND8G As DAY, "
            Sql &= "ROUND(sum(DORK8G)/1000,0) as AMOUNT, "
            Sql &= "'Buyer' as Buyer, "
            Sql &= "'Sales' as Sales "
            Sql &= "FROM SONGOLIB.S8G00WK "
            '
            If DSum.SelectedValue <> "00" Then
                Sql &= "where RSMC8G = '" & DSum.SelectedValue & "' "
            End If
            '
            Sql &= "GROUP BY RSMC8G, ODRG8G, OCND8G "
            Sql &= "ORDER BY OCND8G DESC "
        End If
        '
        Dim DBAdapter2 As New OleDbDataAdapter(Sql, cn)
        DBAdapter2.Fill(ds1, "S8G00A")
        If ds1.Tables("S8G00A").Rows.Count > 0 Then
            GridView2.Visible = True
            GridView2.DataSource = ds1
            GridView2.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

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
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

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
            For i = 0 To 5
                Select Case e.Row.Cells(i).Text
                    Case "Buyer"
                        Dim h1 As New HyperLink
                        h1.Target = "_blank"
                        h1.Text = e.Row.Cells(i).Text
                        h1.NavigateUrl = "http://10.245.1.6/WorkFlowSub/INQ_OrderReal01.aspx?pSUMCODE=" & e.Row.Cells(0).Text & "&pMonth=" & e.Row.Cells(1).Text & "&pDATE=" & e.Row.Cells(2).Text & "&pAction=BUYER"
                        e.Row.Cells(i).Text = ""
                        e.Row.Cells(i).Controls.Add(h1)
                    Case "Sales"
                        Dim h1 As New HyperLink
                        h1.Target = "_blank"
                        h1.Text = e.Row.Cells(i).Text
                        h1.NavigateUrl = "http://10.245.1.6/WorkFlowSub/INQ_OrderReal01.aspx?pSUMCODE=" & e.Row.Cells(0).Text & "&pMonth=" & e.Row.Cells(1).Text & "&pDATE=" & e.Row.Cells(2).Text & "&pAction=SALES"
                        e.Row.Cells(i).Text = ""
                        e.Row.Cells(i).Controls.Add(h1)
                    Case Else
                End Select
            Next
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub

    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        GetData()
    End Sub
End Class
