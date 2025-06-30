Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class InquiryDocument
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DMonth As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DClass As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Go As System.Web.UI.WebControls.Button

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

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "InquiryDocument.aspx"


        If Not Me.IsPostBack Then   '不是PostBack
            CheckAuthority()
            SetInitData()   '設定畫面初始值
            DataList()    '篩選資料
        End If
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定畫面初始值
    '**
    '*****************************************************************
    Sub SetInitData()
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim SQL As String
        Dim i As Integer
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '文件類別
        SQL = "Select * From M_Referp Where Cat='015' and DKey='CLASS' Order by Data "
        DClass.Items.Clear()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Referp")
        DBTable1 = DBDataSet1.Tables("M_Referp")
        For i = 0 To DBTable1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBTable1.Rows(i).Item("Data")
            ListItem1.Value = DBTable1.Rows(i).Item("Data")
            DClass.Items.Add(ListItem1)
        Next
        'DB連結關閉
        OleDbConnection1.Close()

        '年
        DYear.Items.Clear()
        For i = 2004 To 2020
            Dim ListItem1 As New ListItem
            ListItem1.Text = CStr(i)
            ListItem1.Value = CStr(i)
            If ListItem1.Value = CStr(DateTime.Now.Year) Then ListItem1.Selected = True
            DYear.Items.Add(ListItem1)
        Next

        '月
        DMonth.Items.Clear()
        DMonth.Items.Add("ALL")
        For i = 1 To 12
            Dim ListItem1 As New ListItem
            If i < 10 Then
                ListItem1.Text = "0" & CStr(i)
                ListItem1.Value = "0" & CStr(i)
            Else
                ListItem1.Text = CStr(i)
                ListItem1.Value = CStr(i)
            End If
            If CStr(i) = CStr(DateTime.Now.Month - 1) Then ListItem1.Selected = True
            DMonth.Items.Add(ListItem1)
        Next
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選資料
    '**
    '*****************************************************************
    Sub DataList()
        Dim SQL As String
        Dim OleDbConnection1 As New OleDbConnection
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("DocFilePath")
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定

        SQL = "SELECT "   '取得DB資料
        SQL = SQL & "Class, Type, Year, Month, Description, Maker, "
        SQL = SQL & "str(Year(MakerTime)) + '/' + str(Month(MakerTime)) + '/' + str(Day(MakerTime)) As MakerTimeDesc, "
        SQL = SQL & "Case Type When '0' then '月報' else '年報' end As TypeDesc, "
        SQL = SQL & "'" & Path & "' + DocFileName As URL "
        SQL = SQL & "FROM F_UPDocument "
        SQL = SQL + "Where Sts = '1' "
        SQL = SQL + "And Class = '" & DClass.SelectedValue & "' "
        SQL = SQL + "And Type  = '" & DType.SelectedValue & "' "
        If DType.SelectedValue = "0" Then
            SQL = SQL + "And Year  = '" & DYear.SelectedValue & "' "
        End If
        If (DMonth.SelectedValue <> "ALL") And (DType.SelectedValue <> "1") Then
            SQL = SQL + "And Month  = '" & DMonth.SelectedValue & "' "
        End If
        SQL = SQL + "Order by MakerTime Desc "

        Dim DBDataSet1 As New DataSet
        OleDbConnection1.Open()
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_UPDocument")
        DataGrid1.DataSource = DBDataSet1
        DataGrid1.DataBind()
        OleDbConnection1.Close()
    End Sub

    Private Sub Go_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Go.Click
        DataGrid1.CurrentPageIndex = 0
        DataList()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex   'DataGrid跳上下頁
        DataList()
    End Sub
End Class
