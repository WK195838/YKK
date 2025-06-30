Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class OR2PR2
    Inherits System.Web.UI.Page

    Dim uCommon As New Utility.Common


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        DORNO.Text = Request.QueryString("PCODE")
        GOGO()
    End Sub
    '---------------------------------------------------------------------------------------------------
    Sub SetOrderNo()
        Dim str As String = DORNO.Text
        For i As Integer = DORNO.Text.Length To 8 - 1
            str = "0" + str
        Next
        DORNO.Text = "OR" + str
    End Sub

    Sub GOGO()
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        Dim ds, ds1 As New DataSet
        '--------------------------------------------------------------------------------------
        'SetOrderNo()
        '--------------------------------------------------------------------------------------
        'F9F00
        SQL = "SELECT PSHN9F,LNGV9F,LUNC9F,CLRC9F,ALMQ9F,PCPU9F, "  '0~5
        SQL = SQL + "'' AS XORNO, "         '6
        SQL = SQL + "'' AS XF44, "          '7
        SQL = SQL + "'' AS XF44DATE, "      '8
        SQL = SQL + "'' AS XF64, "          '9
        SQL = SQL + "'' AS XF64DATE "       '10
        SQL = SQL + "FROM WAVEDLIB.F9F00 "
        SQL = SQL + "WHERE RLTN9F =  '" & DORNO.Text & "' "
        SQL = SQL + "AND   PSHN9F <> ' ' "
        SQL = SQL + "AND   RLTN9F <> PSCN9F "
        SQL = SQL + "ORDER BY PSHN9F "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "F9F00")
        If ds.Tables("F9F00").Rows.Count > 0 Then
            For i As Integer = 0 To ds.Tables("F9F00").Rows.Count - 1
                'OR-NO.(6)
                ds.Tables("F9F00").Rows(i)(6) = DORNO.Text
                'FAB00 (F44)(7~8)
                cn.Close()
                ds1.Clear()
                SQL = "SELECT PPLCAB||'-'||PPMCAB, MAX(PCPUAB) "
                SQL = SQL + "FROM WAVEDLIB.FAB00 "
                SQL = SQL + "WHERE PSHNAB = '" & ds.Tables("F9F00").Rows(i)(0) & "' "
                SQL = SQL + "AND   PPTNAB IN('F40','F44','F45','F61','F65') "
                SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                DBAdapter2.Fill(ds1, "FAB00")
                If ds1.Tables("FAB00").Rows.Count > 0 Then
                    ds.Tables("F9F00").Rows(i)(7) = ds1.Tables("FAB00").Rows(0)(0)
                    ds.Tables("F9F00").Rows(i)(8) = ds1.Tables("FAB00").Rows(0)(1)
                End If
                'FAB00 (F64)(9~10)
                cn.Close()
                ds1.Clear()
                SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB) "
                SQL = SQL + "FROM WAVEDLIB.FAB00 "
                SQL = SQL + "WHERE PSHNAB = '" & ds.Tables("F9F00").Rows(i)(0) & "' "
                SQL = SQL + "AND   PPTNAB IN('F32','F41','F64','F64B') "
                SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                DBAdapter3.Fill(ds1, "FAB00")
                If ds1.Tables("FAB00").Rows.Count > 0 Then
                    ds.Tables("F9F00").Rows(i)(9) = ds1.Tables("FAB00").Rows(0)(0)
                    ds.Tables("F9F00").Rows(i)(10) = ds1.Tables("FAB00").Rows(0)(1)
                End If
            Next
        End If
        '
        GridView1.DataSource = ds
        GridView1.DataBind()
    End Sub
        
    '---------------------------------------------------------------------------------------------------

    Protected Sub BORNO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BORNO.Click
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        Dim ds, ds1 As New DataSet
        '--------------------------------------------------------------------------------------
        'SetOrderNo()
        '--------------------------------------------------------------------------------------
        'F9F00
        SQL = "SELECT PSHN9F,LNGV9F,LUNC9F,CLRC9F,ALMQ9F,PCPU9F, "  '0~5
        SQL = SQL + "'' AS XORNO, "         '6
        SQL = SQL + "'' AS XF44, "          '7
        SQL = SQL + "'' AS XF44DATE, "      '8
        SQL = SQL + "'' AS XF64, "          '9
        SQL = SQL + "'' AS XF64DATE "       '10
        SQL = SQL + "FROM WAVEDLIB.F9F00 "
        SQL = SQL + "WHERE RLTN9F =  '" & DORNO.Text & "' "
        SQL = SQL + "AND   PSHN9F <> ' ' "
        SQL = SQL + "AND   RLTN9F <> PSCN9F "
        SQL = SQL + "ORDER BY PSHN9F "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "F9F00")
        If ds.Tables("F9F00").Rows.Count > 0 Then
            For i As Integer = 0 To ds.Tables("F9F00").Rows.Count - 1
                'OR-NO.(6)
                ds.Tables("F9F00").Rows(i)(6) = DORNO.Text
                'FAB00 (F44)(7~8)
                cn.Close()
                ds1.Clear()
                SQL = "SELECT PPLCAB||'-'||PPMCAB, MAX(PCPUAB) "
                SQL = SQL + "FROM WAVEDLIB.FAB00 "
                SQL = SQL + "WHERE PSHNAB = '" & ds.Tables("F9F00").Rows(i)(0) & "' "
                SQL = SQL + "AND   PPTNAB IN('F40','F44','F45','F61','F65') "
                SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                DBAdapter2.Fill(ds1, "FAB00")
                If ds1.Tables("FAB00").Rows.Count > 0 Then
                    ds.Tables("F9F00").Rows(i)(7) = ds1.Tables("FAB00").Rows(0)(0)
                    ds.Tables("F9F00").Rows(i)(8) = ds1.Tables("FAB00").Rows(0)(1)
                End If
                'FAB00 (F64)(9~10)
                cn.Close()
                ds1.Clear()
                SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB) "
                SQL = SQL + "FROM WAVEDLIB.FAB00 "
                SQL = SQL + "WHERE PSHNAB = '" & ds.Tables("F9F00").Rows(i)(0) & "' "
                SQL = SQL + "AND   PPTNAB IN('F32','F41','F64','F64B') "
                SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                DBAdapter3.Fill(ds1, "FAB00")
                If ds1.Tables("FAB00").Rows.Count > 0 Then
                    ds.Tables("F9F00").Rows(i)(9) = ds1.Tables("FAB00").Rows(0)(0)
                    ds.Tables("F9F00").Rows(i)(10) = ds1.Tables("FAB00").Rows(0)(1)
                End If
            Next
        End If
        '
        GridView1.DataSource = ds
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            '核定時數
            e.Row.Cells(5).Text = FormatNumber(CInt(e.Row.Cells(5).Text.ToString), 0)
        End If

    End Sub
End Class
