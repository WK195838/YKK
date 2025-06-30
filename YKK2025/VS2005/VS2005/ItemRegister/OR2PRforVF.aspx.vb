Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class OR2PRforVF
    Inherits System.Web.UI.Page

    Dim uCommon As New Utility.Common


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900      ' 設定逾時時間
        If Not IsPostBack Then
            DORNO.Text = ""
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    Sub SetOrderNo()
        Dim str As String = DORNO.Text
        For i As Integer = DORNO.Text.Length To 8 - 1
            str = "0" + str
        Next
        DORNO.Text = "OR" + str
    End Sub
    '---------------------------------------------------------------------------------------------------
    Protected Sub BORNO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BORNO.Click
        SetOrderNo()
        gogo()
    End Sub

    Sub gogo()
        Dim SQL As String
        Dim cn As New OleDbConnection
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        cn.ConnectionString = ConnectString
        Dim ds, ds1 As New DataSet
        '--------------------------------------------------------------------------------------

        '--------------------------------------------------------------------------------------
        'F9F00
        SQL = "SELECT PSCN9F,PSHN9F,LNGV9F,LUNC9F,CLRC9F,ALMQ9F,PCPU9F, "  '0~6
        SQL = SQL + "'' AS XORNO, "         '7
        SQL = SQL + "'' AS XF44, "          '8
        SQL = SQL + "'' AS XF44DATE, "      '9
        SQL = SQL + "'' AS EMPCAB1, "      '10
        SQL = SQL + "'' AS XF64, "          '11
        SQL = SQL + "'' AS XF64DATE, "       '12
        SQL = SQL + "'' AS EMPCAB2 "      '13
        SQL = SQL + "FROM WAVEDLIB.F9F00 "
        SQL = SQL + "WHERE RLTN9F =  '" & DORNO.Text & "' "
        'SQL = SQL + "AND   PSHN9F <> ' ' "
        SQL = SQL + "AND   RLTN9F <> PSCN9F "
        SQL = SQL + "ORDER BY PSHN9F "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "F9F00")
        If ds.Tables("F9F00").Rows.Count > 0 Then
            Dim PSHN9F As String
            Dim PSCN9F As String
            For i As Integer = 0 To ds.Tables("F9F00").Rows.Count - 1
                PSHN9F = Trim(ds.Tables("F9F00").Rows(i).Item("PSHN9F"))
                PSCN9F = Trim(ds.Tables("F9F00").Rows(i).Item("PSCN9F"))
                ds.Tables("F9F00").Rows(i)(7) = DORNO.Text

                If PSHN9F <> "" Then
                    'OR-NO.(6)
                    'FAB00 (F44)(7~8)
                    cn.Close()
                    ds1.Clear()
                    SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB),MAX(EMPCAB) "
                    SQL = SQL + "FROM WAVEDLIB.FAB00 "
                    SQL = SQL + "WHERE PSHNAB = '" & PSHN9F & "' "
                    SQL = SQL + "AND   PPTNAB IN('F30') "
                    SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                    DBAdapter2.Fill(ds1, "FAB00")
                    If ds1.Tables("FAB00").Rows.Count > 0 Then
                        ds.Tables("F9F00").Rows(i)(8) = ds1.Tables("FAB00").Rows(0)(0)
                        ds.Tables("F9F00").Rows(i)(9) = ds1.Tables("FAB00").Rows(0)(1)
                        ds.Tables("F9F00").Rows(i)(10) = ds1.Tables("FAB00").Rows(0)(2)
                    End If
                    'FAB00 (F64)(9~10)
                    cn.Close()
                    ds1.Clear()
                    SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB),MAX(EMPCAB) "
                    SQL = SQL + "FROM WAVEDLIB.FAB00 "
                    SQL = SQL + "WHERE PSHNAB = '" & PSHN9F & "' "
                    SQL = SQL + "AND   PPTNAB IN('F40','F61') "
                    SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB  "
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                    DBAdapter3.Fill(ds1, "FAB00")
                    If ds1.Tables("FAB00").Rows.Count > 0 Then
                        ds.Tables("F9F00").Rows(i)(11) = ds1.Tables("FAB00").Rows(0)(0)
                        ds.Tables("F9F00").Rows(i)(12) = ds1.Tables("FAB00").Rows(0)(1)
                        ds.Tables("F9F00").Rows(i)(13) = ds1.Tables("FAB00").Rows(0)(2)
                    End If
                End If
            Next
        End If


        GridView1.DataSource = ds
        GridView1.DataBind()
    End Sub
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            '核定時數
            ' e.Row.Cells(5).Text = FormatNumber(CInt(e.Row.Cells(5).Text.ToString), 0)
            e.Row.Cells(6).Text = FormatNumber(CInt(e.Row.Cells(6).Text.ToString), 0)
        End If



    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        For Each gvr As GridViewRow In GridView1.Rows
            gvr.Cells(6).HorizontalAlign = HorizontalAlign.Right
        Next
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        gogo()
    End Sub
End Class
