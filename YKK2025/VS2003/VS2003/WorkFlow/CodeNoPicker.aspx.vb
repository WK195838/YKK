Imports System.Data
Imports System.Data.OleDb

Public Class CodeNoPicker
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DKey As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents BKey As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK共通涵式

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pKey As String     'Search Key

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then   '不是PostBack
            If Request.Cookies("CodeNo").Value <> "" Then
                DKey.Text = Request.Cookies("CodeNo").Value
                pKey = Request.Cookies("CodeNo").Value
            Else
                DKey.Text = ""
                pKey = "ALL"
            End If
            CodeData()  '取得資料
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandSource.CommandName = "Select" Then  '點選選取檢查
            Dim Key As String = DataGrid1.DataKeys(e.Item.ItemIndex)  '所選取的Code

            Dim wNo As String = ""
            Dim wFormNo As String = ""
            Dim wFormSno As String = ""
            Dim wBuyer As String = ""
            Dim wSellVendor As String = ""

            Dim SQL As String
            Dim DBDataSet1 As New DataSet

            'DB連結設定
            Dim OleDbConnection1 As New OleDbConnection
            OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
            OleDbConnection1.Open()

            If Request.QueryString("field") = "Suface" Then
                SQL = "Select Code, No, FormNo, FormSno, Buyer, SellVendor From F_SufaceSheet "
            End If

            SQL = SQL & " Where Sts =  '1' "
            If pKey <> "ALL" Then
                SQL = SQL & "   And  Code Like '%" & Key & "%'"
            End If

            Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
            DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
            If DBDataSet1.Tables("F_SufaceSheet").Rows.Count > 0 Then
                wNo = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("No")
                wFormNo = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FormNo")
                wFormSno = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("FormSno")
                wBuyer = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("Buyer")
                wSellVendor = DBDataSet1.Tables("F_SufaceSheet").Rows(0).Item("SellVendor")
            End If
            'DB連結關閉
            OleDbConnection1.Close()

            '所選取的Map No回父視窗並關閉
            Dim Cmd As String
            Response.Cookies("CodeNo").Value = Key

            If Request.QueryString("pFormNo") = "000015" Then
                Cmd = String.Format("<script>window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.opener.document.{4}.value = '{5:d}'; window.opener.document.{6}.value = '{7:d}'; window.opener.document.{8}.value = '{9:d}'; window.opener.document.{10}.value = '{11:d}'; window.close();</script>", "Form1.DCode", Key, "Form1.DONo", wNo, "Form1.DOFormNo", wFormNo, "Form1.DOFormSno", wFormSno, "Form1.DBuyer", wBuyer, "Form1.DSellVendor", wSellVendor)
            End If

            Response.Write(Cmd)
        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        'DataGrid1.DataBind()
        CodeData()
    End Sub

    Sub CodeData()
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        Dim SQL As String

        SQL = "SELECT "   '取得DB資料
        SQL = SQL & "Code, No, "

        If Request.QueryString("field") = "Suface" Then
            SQL = SQL & "'SufaceSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
            SQL = SQL & "FROM F_SufaceSheet "
        End If
        SQL = SQL & " Where Sts =  '1' "

        If Request.QueryString("pFormNo") = "000015" Then
            SQL = SQL + "And   UPDSts = '0' "
        End If

        If (pKey <> "ALL") And (pKey <> "") Then
            SQL = SQL + "And Code Like '%" & pKey & "%' "
        End If
        SQL = SQL + "Order by CompletedTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_SufaceSheet")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Private Sub BKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BKey.Click
        DataGrid1.CurrentPageIndex = 0  'DataGrid跳上下頁
        pKey = DKey.Text
        CodeData()
    End Sub

End Class
