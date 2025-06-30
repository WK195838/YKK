Imports System.Data
Imports System.Data.OleDb

Public Class MapPicker_01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BKey As System.Web.UI.WebControls.Button
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pKey As String     'Search Key

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '不是PostBack
            If Request.Cookies("MapNo").Value <> "" Then
                DKey.Text = Request.Cookies("MapNo").Value
                pKey = Request.Cookies("MapNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            MapData()  '取得資料
        End If
    End Sub


    Sub MapData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim SQL As String

        If Request.QueryString("field") = "Ori" Then
            SQL = "SELECT "   '取得DB資料
            SQL = SQL & "MapNo, 'MapSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_MapSheet "
        Else
            SQL = "SELECT "   '取得DB資料
            SQL = SQL & "MapNo, 'MapModSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_MapModSheet "
        End If
        SQL = SQL + "Where Sts = '1' "
        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And MapNo Like '%" & pKey & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Map")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        MapData()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '點選選取檢查
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '所選取的Map No
            Response.Cookies("MapNo").Value = Key

            Dim wLevel As String = ""
            Dim wFormNo As String = ""
            Dim wFormSno As Integer = 0

            Dim SQL As String
            Dim DBDataSet1 As New DataSet

            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            If Request.QueryString("field") = "Ori" Then
                SQL = "SELECT "   '取得DB資料
                SQL = SQL & "FormNo, FormSno, Level "
                SQL = SQL & "FROM F_MapSheet "
            Else
                SQL = "SELECT "   '取得DB資料
                SQL = SQL & "FormNo, FormSno, Level "
                SQL = SQL & "FROM F_MapModSheet "
            End If
            SQL = SQL & " Where MapNo =  '" & Key & "'"
            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_MapSheet")
            If DBDataSet1.Tables("F_MapSheet").Rows.Count > 0 Then
                wLevel = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("Level")
                wFormNo = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FormNo")
                wFormSno = DBDataSet1.Tables("F_MapSheet").Rows(0).Item("FormSno")
            End If
            'DB連結關閉
            OleDbConnection1.Close()

            If Request.QueryString("pFormNo") = "000004" Or Request.QueryString("pFormNo") = "000008" Then
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();</script>", "Form1.DMapNo", Key, "Form1.DLevel", wLevel))
            End If

            If Request.QueryString("pFormNo") = "000005" Or Request.QueryString("pFormNo") = "000009" Then
                Response.Write(String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.close();</script>", "Form1.DMapNo", Key))
            End If

        End If

    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid跳上下頁
        pKey = DKey.Text
        MapData()
    End Sub

End Class
