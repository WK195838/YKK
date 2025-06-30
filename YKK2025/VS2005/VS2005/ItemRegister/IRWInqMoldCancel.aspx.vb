Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class IRWInqMoldCancel
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim uJavaScript As New Utility.JScript
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String
    Dim UserID As String            'UserID

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IRWInqMoldCancel.aspx"

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
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DSearchKey.Text = ""
        '
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        '
        'SQL = SQL & "CASE WHEN FORMNO='001151' THEN '營業申請單' "
        'SQL = SQL & "     WHEN FORMNO='001152' THEN 'ZIP申請單'  "
        'SQL = SQL & "     WHEN FORMNO='001153' THEN 'SLD申請單'  "
        'SQL = SQL & "     ELSE 'CH申請單' END AS FORM, "

        SQL = "SELECT "
        SQL = SQL & "CancelDate,Item,ItemName,Size,Chain,Slider,Tape_,Finish "
        SQL = SQL & "FROM M_IRWMoldCancel "
        SQL = SQL & "WHERE Active = 1 "
        If DSearchKey.Text <> "" Then
            SQL = SQL & "and [CancelDate]+[Item]+[ItemName]+[Size]+[Chain]+[Slider]+[Tape_]+[Finish] like '%" & DSearchKey.Text & "%' "
        End If
        SQL = SQL & "Order By canceldate desc "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "ITEM")
        If DBDataSet1.Tables("ITEM").Rows.Count > 0 Then
            GridView1.DataSource = DBDataSet1
            GridView1.DataBind()
            GridView1.Visible = True
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        OleDbConnection1.Close()
        '
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            ' 清除
            e.Row.Cells.Clear()

            '-----------------------------------------
            ' 表頭-8行
            Dim H4tc0 As TableCell = New TableCell
            H4tc0.Text = "報廢日"
            H4row.Cells.Add(H4tc0)

            Dim H4tc2 As TableCell = New TableCell
            H4tc2.Text = "Item"
            H4row.Cells.Add(H4tc2)
            H4tc2.BackColor = Color.Green

            Dim H4tc3 As TableCell = New TableCell
            H4tc3.Text = "Item Name"
            H4row.Cells.Add(H4tc3)
            H4tc3.BackColor = Color.Green

            Dim H4tc4 As TableCell = New TableCell
            H4tc4.Text = "Size"
            H4row.Cells.Add(H4tc4)
            H4tc4.BackColor = Color.Blue

            Dim H4tc5 As TableCell = New TableCell
            H4tc5.Text = "Chain"
            H4row.Cells.Add(H4tc5)

            Dim H4tc6 As TableCell = New TableCell
            H4tc6.Text = "Slider"
            H4tc6.BackColor = Color.Blue
            H4row.Cells.Add(H4tc6)

            Dim H4tc7 As TableCell = New TableCell
            H4tc7.Text = "Tape"
            H4row.Cells.Add(H4tc7)

            Dim H4tc8 As TableCell = New TableCell
            H4tc8.Text = "Finish"
            H4row.Cells.Add(H4tc8)

            gv.Controls(0).Controls.AddAt(0, H4row)
        End If

    End Sub

    '
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            For i As Integer = 0 To 7
                If i = 3 Or i = 5 Then
                    e.Row.Cells(i).ForeColor = Color.Red
                End If
            Next
        End If
        '
    End Sub

    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        DataList()
    End Sub

End Class
