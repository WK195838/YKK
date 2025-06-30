Imports System.Data
Imports System.Data.OleDb

Partial Class CheckBankInf
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "CheckBankInf.aspx"

        If Not Me.IsPostBack Then
            CheckAuthority()
            SetSearchItem()
            DataList()
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub SetSearchItem()
    End Sub

    Sub DataList()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        SQL = "SELECT CustCode + '-' + CustName As Cust "
        SQL = SQL + "FROM M_CustControl "
        SQL = SQL + "Where IntCust <> '1' "
        SQL = SQL + "  And 'R' + CustCode Not In (Select CustCode From M_Bank) "
        SQL = SQL + "Order by CustCode "

        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_CustInf")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

End Class
