Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SPWingsOrder
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
    Dim pType As String
    Dim pItem As String
    Dim pColor As String
    Dim pKeep As String
    Dim pPlanNo As String
    Dim pOrderNo As String
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
        UserID = Request.QueryString("pUserID")             'UserID
        pType = Request.QueryString("pType")
        pItem = Request.QueryString("pItem")
        pColor = Request.QueryString("pColor")
        pKeep = Request.QueryString("pKeep")
        pPlanNo = Request.QueryString("pPlanNo")
        pOrderNo = Request.QueryString("pOrderNo")
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
        GridView2.Visible = False
        'JOYJOY
        If pType = "ST" Or InStr("IT003////", UCase(UserID)) <= 0 Then
            DDetail.Style("left") = -500 & "px"
        Else
            DDetail.Style("left") = 24 & "px"
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim i As Integer
        Dim Sql, str As String
        Dim xPPList As String()
        Dim cn As New OleDbConnection
        Dim ds, ds1, ds2 As New DataSet
        '
        cn.ConnectionString = ConnectString

        Select Case pType
            '
            Case "ORI"
                Sql = "SELECT "
                Sql = Sql & "A.ORDN5E, A.OSBN5E, B.CORN5C, B.OCNU5C, B.RDLU5C, B.CSTC5C, C.FL1I39, B.BYRC5C, D.BYRI35, B.SPRC5C, F.FEMI05, B.SMPF5C, B.NCMF5C, "
                Sql = Sql & "A.ITMC5E, E.IT1IA0 || ' ' || E.IT2IA0 || ' ' || E.IT3IA0 as ITEMNAME, A.LNGV5E, A.LUNC5E, A.CLRC5E, A.ORRQ5E "
                '
                Sql = Sql & "FROM WAVEDLIB.S5E00 A "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S5C00 B ON A.ORDN5E=B.ORDN5C "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S3900 C ON B.CSTC5C=C.CLNC39 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S3500 D ON B.BYRC5C=D.BYRC35 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 E ON A.ITMC5E=E.ITMCA0 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 F ON B.SPRC5C=F.EMPC05 "
                Sql = Sql & "WHERE A.ORDN5E <> 'ORZZZZZZZZ' "
                '
                If pOrderNo <> "" Then
                    Sql = Sql & "AND A.ORDN5E = '" & Trim(pOrderNo) & "' "
                Else
                    Sql = Sql & "AND A.ORDN5E = '" & "ORZZZZZZZZ" & "' "
                End If
                Sql = Sql & "ORDER BY A.ORDN5E, A.OSBN5E "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
                DBAdapter1.Fill(ds, "ORDER")
                If ds.Tables("ORDER").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                End If
                '
                ' PROD INFORMATION
                str = ""
                xPPList = pPlanNo.Trim.Split("/")
                For i = 0 To xPPList.Length - 1
                    If xPPList(i) <> "" Then
                        If str = "" Then
                            str = "'" & xPPList(i) & "'"
                        Else
                            str = str & ", '" & xPPList(i) & "'"
                        End If
                    End If
                Next
                '
                If InStr("IT003////", UCase(UserID)) > 0 Then
                    '
                    Sql = "SELECT "
                    Sql = Sql & "A.PSCN9F AS ORDN5E, A.ITMC9F AS OSBN5E, B.IT1IA0 || ' ' || B.IT2IA0 || ' ' || B.IT3IA0 AS CORN5C, "
                    Sql = Sql & "A.CLRC9F AS OCNU5C, A.KEPC9F AS RDLU5C, A.UIDC9F AS CSTC5C, C.FEMI05 AS FL1I39, A.RADU9F AS BYRC5C, "
                    Sql = Sql & "A.PSHN9F AS BYRI35, A.LCTC9F AS SPRC5C, '' AS FEMI05, '' AS SMPF5C, "
                    Sql = Sql & "'' AS NCMF5C, '' AS ITMC5E, '' as ITEMNAME, '' AS LNGV5E, 0 AS LUNC5E, A.ALMQ9F AS CLRC5E, A.PCPQ9F AS ORRQ5E "
                    '
                    Sql = Sql & "FROM WAVEDLIB.F9F00 A "
                    Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 B ON A.ITMC9F=B.ITMCA0 "
                    Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 C ON A.UIDC9F=C.EMPC05 "
                    Sql = Sql & "WHERE A.PCPQ9F > 0 "
                    '
                    If pItem <> "" Then
                        Sql = Sql & "AND A.ITMC9F = '" & Trim(pItem) & "' "
                    End If
                    '
                    If pColor <> "" Then
                        Sql = Sql & "AND A.CLRC9F = '" & Trim(pColor) & "' "
                    End If
                    '
                    'If pKeep <> "" Then
                    '    Sql = Sql & "AND A.AKPCDD = '" & Trim(pKeep) & "' "
                    'End If
                    ''
                    If pPlanNo <> "" Then
                        Sql = Sql & "AND A.PSCN9F in ( " & str & " ) "
                    Else
                        Sql = Sql & "AND A.PSCN9F = '" & "ORZZZZZZZZ" & "' "
                    End If
                    '
                    Sql = Sql & "ORDER BY A.PSCN9F, A.ITMC9F, A.CLRC9F, A.AKPC9F "
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(Sql, cn)
                    DBAdapter2.Fill(ds1, "F9F")
                    If ds1.Tables("F9F").Rows.Count > 0 Then
                        GridView2.Visible = True
                        GridView2.DataSource = ds1
                        GridView2.DataBind()
                    End If
                    '
                End If
                '
            Case "ORO"
                Sql = "SELECT "
                Sql = Sql & "A.ORDN5E, A.OSBN5E, B.CORN5C, B.OCNU5C, B.RDLU5C, B.CSTC5C, C.FL1I39, B.BYRC5C, D.BYRI35, B.SPRC5C, F.FEMI05, B.SMPF5C, B.NCMF5C, "
                Sql = Sql & "A.ITMC5E, E.IT1IA0 || ' ' || E.IT2IA0 || ' ' || E.IT3IA0 as ITEMNAME, A.LNGV5E, A.LUNC5E, A.CLRC5E, A.ORRQ5E "
                '
                Sql = Sql & "FROM WAVEDLIB.S5E00 A "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S5C00 B ON A.ORDN5E=B.ORDN5C "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S3900 C ON B.CSTC5C=C.CLNC39 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.S3500 D ON B.BYRC5C=D.BYRC35 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 E ON A.ITMC5E=E.ITMCA0 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 F ON B.SPRC5C=F.EMPC05 "
                Sql = Sql & "WHERE A.ORDN5E <> 'ORZZZZZZZZ' "
                '
                If pOrderNo <> "" Then
                    Sql = Sql & "AND A.ORDN5E = '" & Trim(pOrderNo) & "' "
                Else
                    Sql = Sql & "AND A.ORDN5E = '" & "ORZZZZZZZZ" & "' "
                End If
                Sql = Sql & "ORDER BY A.ORDN5E, A.OSBN5E "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
                DBAdapter1.Fill(ds, "ORDER")
                If ds.Tables("ORDER").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                End If
                '
                ' STOCK INFORMATION
                str = ""
                xPPList = pPlanNo.Trim.Split("/")
                For i = 0 To xPPList.Length - 1
                    If xPPList(i) <> "" Then
                        If str = "" Then
                            str = "'" & xPPList(i) & "'"
                        Else
                            str = str & ", '" & xPPList(i) & "'"
                        End If
                    End If
                Next
                '
                If InStr("IT003////", UCase(UserID)) > 0 Then
                    '
                    Sql = "SELECT "
                    Sql = Sql & "A.RLTNDD AS ORDN5E, A.ITMCDD AS OSBN5E, B.IT1IA0 || ' ' || B.IT2IA0 || ' ' || B.IT3IA0 AS CORN5C, "

                    Sql = Sql & "A.CLRCDD AS OCNU5C, A.AKPCDD AS RDLU5C, A.UIDCDD AS CSTC5C, C.FEMI05 AS FL1I39, A.RADUDD AS BYRC5C, "

                    Sql = Sql & "A.STONDD AS BYRI35, A.ASLCDD AS SPRC5C, A.WSSCDD AS FEMI05, '' AS SMPF5C, "
                    Sql = Sql & "'' AS NCMF5C, '' AS ITMC5E, '' as ITEMNAME, '' AS LNGV5E, A.ASRQDD AS LUNC5E, A.SALQDD AS CLRC5E, A.KSAQDD AS ORRQ5E "

                    '
                    Sql = Sql & "FROM WAVEDLIB.TDD00 A "
                    Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 B ON A.ITMCDD=B.ITMCA0 "
                    Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 C ON A.UIDCDD=C.EMPC05 "
                    Sql = Sql & "WHERE A.DECNDD = ' ' "
                    '
                    If pItem <> "" Then
                        Sql = Sql & "AND A.ITMCDD = '" & Trim(pItem) & "' "
                    End If
                    '
                    If pColor <> "" Then
                        Sql = Sql & "AND A.CLRCDD = '" & Trim(pColor) & "' "
                    End If
                    '
                    'If pKeep <> "" Then
                    '    Sql = Sql & "AND A.AKPCDD = '" & Trim(pKeep) & "' "
                    'End If
                    ''
                    If pPlanNo <> "" Then
                        Sql = Sql & "AND A.RLTNDD in ( " & str & " ) "
                    Else
                        Sql = Sql & "AND A.RLTNDD = '" & "ORZZZZZZZZ" & "' "
                    End If
                    '
                    Sql = Sql & "ORDER BY A.RLTNDD, A.ITMCDD, A.CLRCDD, A.AKPCDD "
                    '
                    Dim DBAdapter2 As New OleDbDataAdapter(Sql, cn)
                    DBAdapter2.Fill(ds1, "STOCK")
                    If ds1.Tables("STOCK").Rows.Count > 0 Then
                        GridView2.Visible = True
                        GridView2.DataSource = ds1
                        GridView2.DataBind()
                    End If
                    '
                End If
                '
            Case "ST"
                Sql = "SELECT "
                Sql = Sql & "A.RADUXX AS ORDN5E, A.HCDCXX AS OSBN5E, A.ITMCDB AS CORN5C, "
                Sql = Sql & "B.IT1IA0 || ' ' || B.IT2IA0 || ' ' || B.IT3IA0  AS OCNU5C, A.CLRCDB AS RDLU5C, A.KEPCDB AS CSTC5C, A.SLCCDB AS FL1I39, A.UIDCXX AS BYRC5C, "
                Sql = Sql & "C.FEMI05 AS BYRI35, '' AS SPRC5C, '' AS FEMI05, '' AS SMPF5C, '' AS NCMF5C, "
                Sql = Sql & "'' AS ITMC5E, '' as ITEMNAME, '' AS LNGV5E, '' AS LUNC5E, '' AS CLRC5E, OKSQDB AS ORRQ5E "
                '
                Sql = Sql & "FROM WAVEDLIB.TDB80 A "
                Sql = Sql & "LEFT JOIN WAVEDLIB.FA000 B ON A.ITMCDB=B.ITMCA0 "
                Sql = Sql & "LEFT JOIN WAVEDLIB.C0500 C ON A.UIDCXX=C.EMPC05 "
                Sql = Sql & "WHERE A.HCDCXX <> '' "
                '
                If pItem <> "" Then
                    Sql = Sql & "AND A.ITMCDB = '" & Trim(pItem) & "' "
                End If
                '
                If pColor <> "" Then
                    Sql = Sql & "AND A.CLRCDB = '" & Trim(pColor) & "' "
                End If
                '
                If pKeep <> "" Then
                    Sql = Sql & "AND A.KEPCDB = '" & Trim(pKeep) & "' "
                End If
                '
                Sql = Sql & "ORDER BY A.RADUXX, A.RADTXX, A.HCDCXX DESC, A.ITMCDB, A.CLRCDB, A.KEPCDB "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
                DBAdapter1.Fill(ds, "HISTORY")
                If ds.Tables("HISTORY").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                End If
                '
            Case Else
                '
        End Select
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料 TITLE-1
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i, j As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Select Case pType
                '
                Case "ORI", "ORO"
                    DTitle.Text = "Order Inf."
                    '
                    Dim gv As GridView = sender
                    Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
                    '
                    ' 清除
                    e.Row.Cells.Clear()
                    '
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '
                    j = 0
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "OrderNo"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "SubNo"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "C.OrderNo"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Cfm.D"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Req.D"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Customer"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Cust. Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Buyer"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "By Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Sales"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Sample"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "No.C"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "ItemN ame"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Color"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Length"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Length.U"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    gv.Controls(0).Controls.AddAt(0, H3row)
                    '
                Case Else
                    'STOCK TRANSFER
                    DTitle.Text = "WINGS Stock Transfer Inf."
                    '
                    Dim gv As GridView = sender
                    Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
                    '
                    ' 清除
                    e.Row.Cells.Clear()
                    '
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '
                    j = 0
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Date"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Action"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Color"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Keep"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Location"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "User"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Name"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    gv.Controls(0).Controls.AddAt(0, H3row)
                    '
            End Select
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料-1
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Select Case pType
                Case "ORI", "ORO"
                    '
                Case Else
                    For i = 0 To 18
                        Select Case i
                            Case 9, 10, 11, 12, 13, 14, 15, 16, 17
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
                    '
            End Select
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            Select Case pType
                Case "ORI", "ORO"
                    For i = 0 To 18
                    Next
                    '
                Case Else
                    For i = 0 To 18
                        Select Case i
                            Case 9, 10, 11, 12, 13, 14, 15, 16, 17
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
            End Select

        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料 TITLE-2
    '**
    '*****************************************************************
    Protected Sub GridView2_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowCreated
        Dim i, j As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '

            Select Case pType
                Case "ORI"
                    'Production
                    DDetail.Text = "Production Relation Inf."
                    '
                    j = 0
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Relation"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Color"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Keep"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Allocate Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Complete Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Date"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Production No"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Location"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "User"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Name"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    gv.Controls(0).Controls.AddAt(0, H3row)
                    '
                Case Else
                    'Production
                    DDetail.Text = "Stock Relation Inf."
                    '
                    j = 0
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Relation"
                    tcl(j).BackColor = Color.Blue
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Item Name"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Color"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Keep"
                    tcl(j).BackColor = Color.Purple
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Request Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Stock Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Keep Qty"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Date"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Stock Out No"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "From"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "To"
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = ""
                    tcl(j).BackColor = Color.Green
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "User"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(j).Text = "Name"
                    tcl(j).BackColor = Color.Black
                    j = j + 1
                    '
                    gv.Controls(0).Controls.AddAt(0, H3row)
                    '
            End Select

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料-2
    '**
    '*****************************************************************
    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then
            Select Case pType
                Case "ORI"
                    For i = 0 To 18
                        Select Case i
                            Case 5, 11, 12, 13, 14, 15, 16
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
                    '
                Case Else
                    For i = 0 To 18
                        Select Case i
                            Case 12, 13, 14, 15, 16
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
                    '
            End Select
        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Select Case pType
                Case "ORI"
                    For i = 0 To 18
                        Select Case i
                            Case 5, 11, 12, 13, 14, 15, 16
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
                    '
                Case Else
                    For i = 0 To 18
                        Select Case i
                            Case 12, 13, 14, 15, 16
                                e.Row.Cells(i).Visible = False
                            Case Else
                        End Select
                    Next
                    '
            End Select
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
    '
End Class
