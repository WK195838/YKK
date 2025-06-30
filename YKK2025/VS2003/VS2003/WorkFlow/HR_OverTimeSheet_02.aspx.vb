Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_OverTimeSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDateType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFCH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAH1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFAM1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOverTimeSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DOverTimeDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFood As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTraffic As System.Web.UI.WebControls.TextBox
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DCVacation As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "HR_OverTimeSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("OverTimeFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_OverTimeSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_OverTimeSheet")
        If DBDataSet1.Tables("F_OverTimeSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("No")                     'No
            DOverTimeDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("OverTimeDate")   '加班日期
            DDateType.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DateType")           '加班別
            DSalaryYM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("SalaryYM")           '所屬年月
            DDate.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Date")                   '申請日期
            DName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Name")                   '姓名
            DEmpID.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobTitle")           '職稱
            DJobCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("JobCode")             '職稱代碼
            DDepoName.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoName")           '公司
            DDepoCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DepoCode")           '公司代碼
            DDivision.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Division")           '部門
            DDivisionCode.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("DivisionCode")   '部門代碼
            DCVacation.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("CVacation")         '調假
            DFood.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Food")                   '伙食
            DTraffic.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("Traffic")             '交通

            DBStartH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartH").ToString    '預定開始-時
            DBStartM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BStartM").ToString    '預定開始-分
            DBEndH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndH").ToString        '預定終止-時
            DBEndM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BEndM").ToString        '預定終止-分
            DBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BH").ToString              '計算結果-時
            DBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("BM").ToString              '計算結果-分

            DAStartH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartH").ToString    '實際開始-時
            DAStartM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AStartM").ToString    '實際開始-分
            DAEndH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndH").ToString        '實際終止-時
            DAEndM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AEndM").ToString        '實際終止-分
            DAH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AH").ToString              '計算結果-時
            DAM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("AM").ToString              '計算結果-分

            DFAH1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH1").ToString          '核定平日2內-時
            DFAM1.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM1").ToString          '核定平日2內-分
            DFAH2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAH2").ToString          '核定平日2外-時
            DFAM2.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FAM2").ToString          '核定平日2外-分

            DFBH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBH").ToString            '核定假日-時
            DFBM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FBM").ToString            '核定假日-分
            DFCH.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCH").ToString            '核定國定假日-時
            DFCM.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FCM").ToString            '核定國定假日-分

            DFReason.Text = DBDataSet1.Tables("F_OverTimeSheet").Rows(0).Item("FReason")             '加班理由

            BCardTime.Attributes("onclick") = "ShowCardTime();"
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
