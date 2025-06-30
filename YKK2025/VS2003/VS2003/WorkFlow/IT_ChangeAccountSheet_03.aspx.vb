Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class IT_ChangeAccountSheet_03
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BFinishDate As System.Web.UI.WebControls.Button
    Protected WithEvents DFinishDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSystem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFun4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFun3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DFun2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DAccount4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAccount3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAccount2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAccount1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFun1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DRemark As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEngineer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DChangeAccountSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DataGrid9 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents DHistoryLabel As System.Web.UI.WebControls.Label

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
    Dim wNo As String               'No
    Dim NowDateTime As String       '現在日期時間

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IT_ChangeAccountSheet_03.aspx"

        SetParameter()          '設定共用參數
        TopPosition()           '按鈕及RequestedField的Top位置
        If Not Me.IsPostBack Then   '不是PostBack
            ShowFormData()      '顯示表單資料
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時
        wNo = Request.QueryString("pNo")            'No
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '*****************************************************************
    '**
    '**     顯示表單資料
    '**
    '*****************************************************************
    Sub ShowFormData()
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("ChangeAccountFilePath")
        Dim SQL, Str, wFormNo As String
        Dim i, wFormSno As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet9 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        DFun1.Items.Clear()
        DFun2.Items.Clear()
        DFun3.Items.Clear()
        DFun4.Items.Clear()
        DEngineer.Items.Clear()
        DSystem.Items.Clear()

        SQL = "Select * From F_ChangeAccountSheet "
        SQL = SQL & " Where No =  '" & wNo & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_ChangeAccountSheet")
        If DBDataSet1.Tables("F_ChangeAccountSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("No")
            DDate.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Date")
            DName.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Name")
            DEmpID.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("JobTitle")
            DJobCode.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("JobCode")
            DDepoName.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("DepoName")
            DDepoCode.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("DepoCode")
            DDivision.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Division")
            DDivisionCode.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("DivisionCode")

            DFinishDate.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("FinishDate")

            DAccount1.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Account1")
            DAccount2.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Account2")
            DAccount3.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Account3")
            DAccount4.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Account4")

            If DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("BStartDate") = "1900/1/1" Then
                DBStartDate.Text = ""
            Else
                DBStartDate.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("BStartDate")
            End If

            If DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("BEndDate") = "1900/1/1" Then
                DBEndDate.Text = ""
            Else
                DBEndDate.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("BEndDate")
            End If

            DBDays.Text = CStr(DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("BDays"))

            Dim ListItem1 As New ListItem
            ListItem1.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Engineer")
            ListItem1.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Engineer")
            DEngineer.Items.Add(ListItem1)

            Dim ListItem2 As New ListItem
            ListItem2.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("System")
            ListItem2.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("System")
            DSystem.Items.Add(ListItem2)

            Dim ListItem3 As New ListItem
            ListItem3.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun1")
            ListItem3.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun1")
            DFun1.Items.Add(ListItem3)

            Dim ListItem4 As New ListItem
            ListItem4.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun2")
            ListItem4.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun2")
            DFun2.Items.Add(ListItem4)

            Dim ListItem5 As New ListItem
            ListItem5.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun3")
            ListItem5.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun3")
            DFun3.Items.Add(ListItem5)

            Dim ListItem6 As New ListItem
            ListItem6.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun4")
            ListItem6.Value = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Fun4")
            DFun4.Items.Add(ListItem6)

            DRemark.Text = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("Remark")
            'Key
            wFormNo = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("FormNo")
            wFormSno = DBDataSet1.Tables("F_ChangeAccountSheet").Rows(0).Item("FormSno")
        End If
        '核定履歷資料
        SQL = "SELECT "
        SQL = SQL + "FormNo, FormSno, Step, SeqNo, "
        SQL = SQL + "StsDesc, FormName, FlowTypeDesc, DecideName, DelaySts, StepNameDesc, DecideDescA, AgentName, "
        SQL = SQL + "AEndTimeDesc As Description, "
        SQL = SQL + "URL "
        SQL = SQL + "FROM V_WaitHandle_01 "
        SQL = SQL + "Where FormNo  = '" & wFormNo & "' "
        SQL = SQL + "  And FormSno = '" & CStr(wFormSno) & "' "
        SQL = SQL + "Order by Unique_ID Desc "
        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet9, "DecideHistory")
        DataGrid9.DataSource = DBDataSet9
        DataGrid9.DataBind()
        'DB連結關閉
        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(TopPosition)
    '**     按鈕及RequestedField的Top位置
    '**
    '*****************************************************************
    Sub TopPosition()
    End Sub

End Class
