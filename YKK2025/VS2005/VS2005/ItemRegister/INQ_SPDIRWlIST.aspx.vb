Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_SPDIRWlIST
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim pSize As String = ""
    Dim pFamily As String = ""
    Dim pSlider As String = ""
    Dim UserID As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "INQ_SPDIRWlIST.aspx"

        If Not Me.IsPostBack Then
            SetParameter()          '設定共用參數
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        pSize = Request.QueryString("pSize")
        pFamily = Request.QueryString("pFamily")
        pSlider = Request.QueryString("pSlider")
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
    End Sub

    Sub DataList()
        Dim SQL, wSpecGroup As String
        Dim xPuller As String
        Dim xShort As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet1A As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim DBDataSet3 As New DataSet
        Dim DBDataSet4 As New DataSet
        Dim xHaveData As Boolean
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        xHaveData = False
        If pSize <> "" And pFamily <> "" And pSlider <> "" Then
            OleDbConnection1.Open()
            '
            'GET IRW 
            DBDataSet3.Clear()
            '
            wSpecGroup = ""
            SQL = "SELECT "
            SQL = SQL & "top 1 '' As IRWSpec "
            SQL = SQL & "FROM M_SPDPullerSpec "
            '
            Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter3.Fill(DBDataSet3, "IRWSPEC")
            If DBDataSet3.Tables("IRWSPEC").Rows.Count > 0 Then
                GridView3.DataSource = DBDataSet3
                GridView3.DataBind()
                GridView3.Visible = True
            End If
            '
            'GET SpecGroup
            DBDataSet1.Clear()
            '
            wSpecGroup = ""
            SQL = "SELECT "
            SQL = SQL & "top 1 SpecGroup "
            SQL = SQL & "FROM M_SPDPullerSpec "
            SQL = SQL & "Where size = '" & pSize & "' "
            SQL = SQL & "And Family = '" & pFamily & "' "
            SQL = SQL & "And '" & pSlider & "' like body+'%' "
            SQL = SQL & "Order By len(body) DESC, SpecGroup "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "SPECGR")
            If DBDataSet1.Tables("SPECGR").Rows.Count > 0 Then
                wSpecGroup = DBDataSet1.Tables("SPECGR").Rows(0).Item("SpecGroup")
                '
                DBDataSet1A.Clear()
                '
                SQL = "SELECT top 1"
                SQL = SQL & "'" & wSpecGroup & " = ' + " & "(select a.spec + '/' from M_SPDPullerSpec a Where a.SpecGroup = '" & wSpecGroup & "' FOR XML PATH('') ) as SpecGr "
                SQL = SQL & "FROM M_SPDPullerSpec "
                SQL = SQL & "Where SpecGroup = '" & wSpecGroup & "' "
                '
                Dim DBAdapter1A As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter1A.Fill(DBDataSet1A, "SPEC")
                If DBDataSet1A.Tables("SPEC").Rows.Count > 0 Then
                    GridView1.DataSource = DBDataSet1A
                    GridView1.DataBind()
                    GridView1.Visible = True
                End If
            End If
            '
            If wSpecGroup <> "" Then
                'MsgBox(wSpecGroup)
                '
                'SHORT
                ' SHORT PULLER
                DBDataSet2.Clear()
                xPuller = ""
                xShort = ""
                '
                SQL = "select top 1 Puller, Short "
                SQL = SQL & "from M_PullerShortDetail "
                SQL = SQL & "where BUYER <> 'zzzzzz' "
                SQL = SQL & "AND '" & pSlider & "' like '%' + Short + '%' "
                SQL = SQL & "Order by len(short) desc "
                '
                Dim DBAdapter9 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter9.Fill(DBDataSet3, "PULLERSHORT")
                If DBDataSet3.Tables("PULLERSHORT").Rows.Count > 0 Then
                    xPuller = Trim(DBDataSet3.Tables("PULLERSHORT").Rows(0).Item("Puller").ToString)
                    xShort = Trim(DBDataSet3.Tables("PULLERSHORT").Rows(0).Item("Short").ToString)
                End If
                '
                'MsgBox(pSlider & "--" & xPuller & "--" & xShort)
                '
                DBDataSet2.Clear()
                '
                SQL = "SELECT "
                SQL = SQL & "RDNo, size+'-'+family+'-'+body as RDSpec, "
                SQL = SQL & "Puller, Color, "
                SQL = SQL & "case when formno='000003' then 'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?pFormNo=000003&pFormSno=' + convert(varchar,formsno) "
                SQL = SQL & "     when formno='000007' then 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo=000007&pFormSno=' + convert(varchar,formsno) "
                SQL = SQL & "     when formno='000011' then 'http://10.245.1.10/WorkFlow/ImportSheet_02.aspx?pFormNo=000011&pFormSno=' + convert(varchar,formsno) "
                SQL = SQL & "     else '' "
                SQL = SQL & "end as FURL "
                SQL = SQL & "FROM M_SPDIRWEDX "
                SQL = SQL & "Where PULLER <> '' "
                SQL = SQL & "And SpecGroup = '" & wSpecGroup & "' "
                '
                SQL = SQL & "And ( '" & pSlider & "' like '%'+puller+'%' "
                SQL = SQL & "    OR "
                SQL = SQL & "      '" & Replace(pSlider, xShort, xPuller) & "' like '%'+puller+'%' "
                SQL = SQL & "    ) "

                '模具報廢不顯示  Hower  2024/10/17
                SQL = SQL & "And RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) NOT IN ( "
                SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO FROM F_ManufInSheet  "
                SQL = SQL & "WHERE MMSSTS =1     "
                SQL = SQL & "UNION ALL "
                SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO  FROM F_ManufOUTSheet     "
                SQL = SQL & "WHERE MMSSTS =1   ) "

                '
                SQL = SQL & "Group By RDNo, size, family, body, Puller, Color, FormNo, FormSno "
                SQL = SQL & "Order By RDNo Desc, size, family, body, Puller, Color, FormNo, FormSno "
                '
                Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
                DBAdapter2.Fill(DBDataSet2, "SPD")
                If DBDataSet2.Tables("SPD").Rows.Count > 0 Then
                    GridView2.DataSource = DBDataSet2
                    GridView2.DataBind()
                    GridView2.Visible = True
                    xHaveData = True
                Else
                    'Puller放寬 DSBZIE08A
                    '
                    DBDataSet2.Clear()
                    '
                    SQL = "SELECT top 10  "
                    SQL = SQL & "RDNo, size+'-'+family+'-'+body as RDSpec, "
                    SQL = SQL & "Puller, Color, "
                    SQL = SQL & "case when formno='000003' then 'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?pFormNo=000003&pFormSno=' + convert(varchar,formsno) "
                    SQL = SQL & "     when formno='000007' then 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo=000007&pFormSno=' + convert(varchar,formsno) "
                    SQL = SQL & "     when formno='000011' then 'http://10.245.1.10/WorkFlow/ImportSheet_02.aspx?pFormNo=000011&pFormSno=' + convert(varchar,formsno) "
                    SQL = SQL & "     else '' "
                    SQL = SQL & "end as FURL "
                    SQL = SQL & "FROM M_SPDIRWEDX "
                    SQL = SQL & "Where PULLER <> '' "
                    'SQL = SQL & "And SpecGroup = '" & wSpecGroup & "' "
                    SQL = SQL & "And CAT = 'SPD' "
                    SQL = SQL & "And ( "
                    SQL = SQL & "     '" & pSlider & "' like '%'+ substring(puller,1,len(puller)-0) + '%' "
                    SQL = SQL & "    ) "

                    '模具報廢不顯示  Hower  2024/10/17
                    SQL = SQL & "And RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) NOT IN ( "
                    SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO FROM F_ManufInSheet  "
                    SQL = SQL & "WHERE MMSSTS =1     "
                    SQL = SQL & "UNION ALL "
                    SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO  FROM F_ManufOUTSheet     "
                    SQL = SQL & "WHERE MMSSTS =1   ) "

                    SQL = SQL & "Group By RDNo, size, family, body, Puller, Color, FormNo, FormSno "
                    SQL = SQL & "Order By len(Puller) desc, puller desc, RDNo Desc "
                    '
                    Dim DBAdapter2A As New OleDbDataAdapter(SQL, OleDbConnection1)
                    DBAdapter2A.Fill(DBDataSet2, "SPD")
                    If DBDataSet2.Tables("SPD").Rows.Count > 0 Then
                        GridView2.DataSource = DBDataSet2
                        GridView2.DataBind()
                        GridView2.Visible = True
                        xHaveData = True
                    Else
                        'Puller放寬 DSBZIE08A
                        '
                        DBDataSet2.Clear()
                        '
                        SQL = "SELECT top 10  "
                        SQL = SQL & "RDNo, size+'-'+family+'-'+body as RDSpec, "
                        SQL = SQL & "Puller, Color, "
                        SQL = SQL & "case when formno='000003' then 'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?pFormNo=000003&pFormSno=' + convert(varchar,formsno) "
                        SQL = SQL & "     when formno='000007' then 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo=000007&pFormSno=' + convert(varchar,formsno) "
                        SQL = SQL & "     when formno='000011' then 'http://10.245.1.10/WorkFlow/ImportSheet_02.aspx?pFormNo=000011&pFormSno=' + convert(varchar,formsno) "
                        SQL = SQL & "     else '' "
                        SQL = SQL & "end as FURL "
                        SQL = SQL & "FROM M_SPDIRWEDX "
                        SQL = SQL & "Where PULLER <> '' "
                        'SQL = SQL & "And SpecGroup = '" & wSpecGroup & "' "
                        SQL = SQL & "And CAT = 'SPD' "
                        SQL = SQL & "And ( "
                        SQL = SQL & "     '" & pSlider & "' like '%'+ substring(puller,1,len(puller)-1) + '%' "
                        SQL = SQL & "  OR '" & pSlider & "' like '%'+ substring(puller,1,len(puller)-2) + '%' "
                        SQL = SQL & "    ) "

                        '模具報廢不顯示  Hower  2024/10/17
                        SQL = SQL & "And RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) NOT IN ( "
                        SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO FROM F_ManufInSheet  "
                        SQL = SQL & "WHERE MMSSTS =1     "
                        SQL = SQL & "UNION ALL "
                        SQL = SQL & "SELECT RTRIM(CONVERT(CHAR,FORMNO))+CONVERT(CHAR,FORMSNO) AS MMSSNO  FROM F_ManufOUTSheet     "
                        SQL = SQL & "WHERE MMSSTS =1   ) "

                        SQL = SQL & "Group By RDNo, size, family, body, Puller, Color, FormNo, FormSno "
                        SQL = SQL & "Order By len(Puller) desc, puller desc, RDNo Desc "
                        '
                        Dim DBAdapter3A As New OleDbDataAdapter(SQL, OleDbConnection1)
                        DBAdapter3A.Fill(DBDataSet2, "SPD")
                        If DBDataSet2.Tables("SPD").Rows.Count > 0 Then
                            GridView2.DataSource = DBDataSet2
                            GridView2.DataBind()
                            GridView2.Visible = True
                            xHaveData = True
                        Else
                            xHaveData = False
                        End If
                    End If
                End If
            Else
                xHaveData = False
            End If
            '
            'GET WINGS
            DBDataSet4.Clear()
            '
            wSpecGroup = ""
            SQL = "SELECT TOP 1 "
            SQL = SQL & "CASE WHEN ORDERNO<>'' THEN ORDERNO + '/' + ITEM + ' ' + ITEMNAME "
            SQL = SQL & "     ELSE 'NO ORDER(3Y)/' + ITEM + ' ' + ITEMNAME "
            SQL = SQL & "END AS WINGS "
            SQL = SQL & "FROM M_WINGSEDX "
            SQL = SQL & "Where SLIDER <> '' "
            SQL = SQL & "And SLIDER = '" & pSlider & "' "
            SQL = SQL & "Order BY SALESDATE DESC, ITEM DESC "
            '
            Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter4.Fill(DBDataSet4, "WINGS")
            If DBDataSet4.Tables("WINGS").Rows.Count > 0 Then
                GridView4.DataSource = DBDataSet4
                GridView4.DataBind()
                GridView4.Visible = True
                '
                xHaveData = True
            End If
            '
            OleDbConnection1.Close()
        End If
        '
        If xHaveData = False Then
            uJavaScript.PopMsg(Me, "無資料!")
        End If
        '
    End Sub
    '

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).Text = Replace(e.Row.Cells(0).Text, "=", "<br>")
        End If

    End Sub

    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).Text = pSize & "-" & pFamily & "-" & pSlider
        End If
    End Sub

    Protected Sub GridView4_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView4.RowDataBound
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).Text = Replace(e.Row.Cells(0).Text, "/", "<br>")
        End If
    End Sub
End Class
