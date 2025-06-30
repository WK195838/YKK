Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class IT_NewSystemSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DSystem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DITEffect As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDevelopItem As System.Web.UI.WebControls.DropDownList
    Protected WithEvents LRefFile As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DEffect As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTarget As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFinishDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEngineer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNewSystemSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox

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
    Dim FieldName(50) As String     '各欄位
    Dim Attribute(50) As Integer    '各欄位屬性    
    Dim Top As Integer              '動態元件的Top位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim NowDateTime As String       '現在日期時間

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Page Load Event
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Response.Cookies("PGM").Value = "IT_NewSystemSheet_02.aspx"

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
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wStep = 999                                 '工程代碼
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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("NewSystemFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        DEngineer.Items.Clear()
        DDevelopItem.Items.Clear()
        DSystem.Items.Clear()

        SQL = "Select * From F_NewSystemSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_NewSystemSheet")
        If DBDataSet1.Tables("F_NewSystemSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("No")
            DDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Date")
            DName.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Name")
            DEmpID.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("EmpID")
            DJobTitle.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("JobTitle")
            DJobCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("JobCode")
            DDepoName.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DepoName")
            DDepoCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DepoCode")
            DDivision.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Division")
            DDivisionCode.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DivisionCode")
            DFinishDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("FinishDate")
            DTarget.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Target")
            DEffect.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Effect")

            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("RefFile") <> "" Then
                LRefFile.NavigateUrl = Path & DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("RefFile")
            Else
                LRefFile.Visible = False
            End If
            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BStartDate") = "1900/1/1" Then
                DBStartDate.Text = ""
            Else
                DBStartDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BStartDate")
            End If

            If DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BEndDate") = "1900/1/1" Then
                DBEndDate.Text = ""
            Else
                DBEndDate.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BEndDate")
            End If

            DBDays.Text = CStr(DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("BDays"))

            Dim ListItem1 As New ListItem
            ListItem1.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Engineer")
            ListItem1.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("Engineer")
            DEngineer.Items.Add(ListItem1)

            Dim ListItem2 As New ListItem
            ListItem2.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DevelopItem")
            ListItem2.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("DevelopItem")
            DDevelopItem.Items.Add(ListItem2)

            Dim ListItem3 As New ListItem
            ListItem3.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("System")
            ListItem3.Value = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("System")
            DSystem.Items.Add(ListItem3)

            DITEffect.Text = DBDataSet1.Tables("F_NewSystemSheet").Rows(0).Item("ITEffect")
        End If

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
        Top = 432
    End Sub

End Class
