Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb

Partial Class OR2PR
    Inherits System.Web.UI.Page
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
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
        Dim ds, ds1, ds2 As New DataSet
        Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
        '--------------------------------------------------------------------------------------
        'F9F00
        SQL = "SELECT A.PSCN9F,A.PSHN9F,A.LNGV9F,A.LUNC9F,A.CLRC9F,A.ALMQ9F,A.PCPU9F, "  '0~6
        SQL = SQL + "'' AS XORNO, "         '7
        SQL = SQL + "'' AS XF15, "      '8
        SQL = SQL + "'' AS XF15DATA, "      '9
        SQL = SQL + "'' AS EMPCAB3, "      '10

        SQL = SQL + "'' AS XF44, "          '11
        SQL = SQL + "'' AS XF44DATE, "      '12
        SQL = SQL + "'' AS EMPCAB1, "      '13
        SQL = SQL + "'' AS XF64, "          '14
        SQL = SQL + "'' AS XF64DATE, "       '15
        SQL = SQL + "'' AS EMPCAB2,B.LN1C9D "      '16
  
        SQL = SQL + " FROM WAVEDLIB.F9F00 A, WAVEDLIB.F9D00 B"
        SQL = SQL + " WHERE A.RLTN9F =  '" & DORNO.Text & "' "
        'SQL = SQL + "AND   PSHN9F <> ' ' "
        SQL = SQL + "AND   A.RLTN9F <> PSCN9F AND A.PSCN9F=B.PSCN9D "
        SQL = SQL + "ORDER BY A.PSHN9F "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, cn)
        DBAdapter1.Fill(ds, "F9F00")
        If ds.Tables("F9F00").Rows.Count > 0 Then
            Dim PSHN9F As String
            Dim PSCN9F As String
            For i As Integer = 0 To ds.Tables("F9F00").Rows.Count - 1
                PSHN9F = Trim(ds.Tables("F9F00").Rows(i).Item("PSHN9F"))
                PSCN9F = Trim(ds.Tables("F9F00").Rows(i).Item("PSCN9F"))
                ds.Tables("F9F00").Rows(i)(7) = DORNO.Text
                Dim set1 As String = ""
                Dim set2 As String = ""
                Dim set3 As String = ""

                If Mid(Trim(ds.Tables("F9F00").Rows(i).Item("LN1C9D")), 1, 2) = "32" Then 'PF
                    set1 = " AND   PPTNAB IN('F32','F41','F64','F64B') "
                    set2 = " AND   PPTNAB IN('F40','F44','F45','F61','F65') "
                ElseIf Mid(Trim(ds.Tables("F9F00").Rows(i).Item("LN1C9D")), 1, 2) = "42" Then 'VF
                    set1 = " AND   PPTNAB IN('F30') "
                    set2 = " AND   PPTNAB IN('F40','F61')"
                    set3 = " AND   PPTNAB IN('F15','VFAI','VFAI2','VF2H')"
                ElseIf Mid(Trim(ds.Tables("F9F00").Rows(i).Item("LN1C9D")), 1, 2) = "22" Then 'MF
                    set1 = " AND   PPTNAB IN('F13','F43','7NL') "
                    set2 = " AND   PPTNAB IN('F40','F32','F47','F54','F65')"

                End If

                If PSHN9F <> "" Then
                    'OR-NO.(6)
                    'FAB00 (F44)(7~8)

                    cn.Close()
                    ds2.Clear()
                    SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB),MAX(EMPCAB) "
                    SQL = SQL + "FROM WAVEDLIB.FAB00 "
                    SQL = SQL + "WHERE PSHNAB = '" & PSHN9F & "' "
                    SQL = SQL + set3
                    SQL = SQL + " GROUP BY PPLCAB||'-'||PPMCAB  "
                    Dim DBAdapter2 As New OleDbDataAdapter(SQL, cn)
                    DBAdapter2.Fill(ds2, "FAB00")
                    'GridView3.DataSource = ds2
                    'GridView3.DataBind()
                    If ds2.Tables("FAB00").Rows.Count > 0 Then
                        ds.Tables("F9F00").Rows(i)(8) = ds2.Tables("FAB00").Rows(0)(0)
                        ds.Tables("F9F00").Rows(i)(9) = ds2.Tables("FAB00").Rows(0)(1)
                        ds.Tables("F9F00").Rows(i)(10) = ds2.Tables("FAB00").Rows(0)(2)
                    End If

                    cn.Close()
                    ds1.Clear()
                    SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB),MAX(EMPCAB) "
                    SQL = SQL + "FROM WAVEDLIB.FAB00 "
                    SQL = SQL + "WHERE PSHNAB = '" & PSHN9F & "' "
                    SQL = SQL + set1
                    SQL = SQL + "GROUP BY PPLCAB||'-'||PPMCAB "
                    Dim DBAdapter3 As New OleDbDataAdapter(SQL, cn)
                    DBAdapter3.Fill(ds1, "FAB00")
                    ' GridView2.DataSource = ds1
                    ' GridView2.DataBind()

                    If ds1.Tables("FAB00").Rows.Count > 0 Then
                        ds.Tables("F9F00").Rows(i)(11) = ds1.Tables("FAB00").Rows(0)(0)
                        ds.Tables("F9F00").Rows(i)(12) = ds1.Tables("FAB00").Rows(0)(1)
                        ds.Tables("F9F00").Rows(i)(13) = ds1.Tables("FAB00").Rows(0)(2)
                    End If
                    'FAB00 (F64)(9~10)
                    cn.Close()
                    ds2.Clear()
                    SQL = "SELECT PPLCAB||'-'||PPMCAB,MAX(PCPUAB),MAX(EMPCAB) "
                    SQL = SQL + "FROM WAVEDLIB.FAB00 "
                    SQL = SQL + "WHERE PSHNAB = '" & PSHN9F & "' "
                    SQL = SQL + set2
                    SQL = SQL + " GROUP BY PPLCAB||'-'||PPMCAB  "
                    Dim DBAdapter4 As New OleDbDataAdapter(SQL, cn)
                    DBAdapter4.Fill(ds2, "FAB00")
                    'GridView3.DataSource = ds2
                    'GridView3.DataBind()
                    If ds2.Tables("FAB00").Rows.Count > 0 Then
                        ds.Tables("F9F00").Rows(i)(14) = ds2.Tables("FAB00").Rows(0)(0)
                        ds.Tables("F9F00").Rows(i)(15) = ds2.Tables("FAB00").Rows(0)(1)
                        ds.Tables("F9F00").Rows(i)(16) = ds2.Tables("FAB00").Rows(0)(2)
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

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub
End Class
