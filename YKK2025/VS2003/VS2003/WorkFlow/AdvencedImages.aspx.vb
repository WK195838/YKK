Imports System.Data
Imports System.Data.OleDb

Public Class AdvencedImages
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BFind As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKStart As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox11 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DKEnd As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataList1 As System.Web.UI.WebControls.DataList

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************

    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not Me.IsPostBack Then
        End If
    End Sub

    Private Sub BFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BFind.Click

        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        'MsgBox(DKStart.Text & "~" & DKEnd.Text)

        SQL = "select TOP 10 "
        SQL = SQL & "case when formno='000003' then [SliderGRCode] + '  內製' "
        SQL = SQL & "     else [SliderGRCode] + '  外注' "
        SQL = SQL & "end as Code, "
        SQL = SQL & "No As No, '' as NoUrl, "
        SQL = SQL & "ImagePath As ImagePath "
        SQL = SQL & "from M_RDPullerImage "
        SQL = SQL & "where formno in ('000003','000007') "
        ' Start
        If DKStart.Text <> "" Then
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKStart.Text) & ") >= '" & DKStart.Text & "' "
        End If
        ' End
        If DKEnd.Text <> "" Then
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKEnd.Text) & ") >= '" & DKEnd.Text & "' "
        End If
        '
        SQL = SQL & "Order by substring([SliderGRCode],1," & Len(DKEnd.Text) & ") "
        '
        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "IMAGE")
        '
        DataList1.DataSource = DBDataSet1.Tables("IMAGE")
        DataList1.DataBind()

        OleDbConnection1.Close()
    End Sub

End Class
