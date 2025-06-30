Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration

Public Class HR_TimeOffSheet_02
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents BVARecord As System.Web.UI.WebControls.Button
    Protected WithEvents DAEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents LOTNo4 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo2 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo5 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo3 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents LOTNo1 As System.Web.UI.WebControls.HyperLink
    Protected WithEvents DOTHours5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents DOTHours As System.Web.UI.WebControls.TextBox
    Protected WithEvents DADays As System.Web.UI.WebControls.TextBox
    Protected WithEvents BCardTime As System.Web.UI.WebControls.Button
    Protected WithEvents DVDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalary As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEvidence As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTimeOffAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobAgent As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDepoName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DSalaryYM As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBDays As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivisionCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDivision As System.Web.UI.WebControls.TextBox
    Protected WithEvents DJobTitle As System.Web.UI.WebControls.TextBox
    Protected WithEvents DEmpID As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents DFReason As System.Web.UI.WebControls.TextBox
    Protected WithEvents DTimeOffSheet1 As System.Web.UI.WebControls.Image
    Protected WithEvents DNo As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAfter As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVacation As System.Web.UI.WebControls.TextBox
    Protected WithEvents DDieType As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAStartH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DBEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DAEndH As System.Web.UI.WebControls.TextBox
    Protected WithEvents DName As System.Web.UI.WebControls.TextBox
    Protected WithEvents DVDaysBlank As System.Web.UI.WebControls.TextBox

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
        Response.Cookies("PGM").Value = "HR_TimeOffSheet_02.aspx"

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
        Dim Path As String = ConfigurationSettings.AppSettings.Get("Http") & ConfigurationSettings.AppSettings.Get("TimeOffFilePath")
        Dim SQL, Str As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        'DB連結設定
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = ConfigurationSettings.AppSettings.Get("SqlConn")  'SQL連結設定
        OleDbConnection1.Open()

        SQL = "Select * From F_TimeOffSheet "
        SQL = SQL & " Where FormNo =  '" & wFormNo & "'"
        SQL = SQL & "   And FormSno =  '" & CStr(wFormSno) & "'"
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_TimeOffSheet")
        If DBDataSet1.Tables("F_TimeOffSheet").Rows.Count > 0 Then
            '表單資料
            DNo.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("No")                       'No
            DDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Date")                   '申請日期
            DSalaryYM.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("SalaryYM")           '所屬年月

            DName.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Name")                   '姓名
            DEmpID.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("EmpID")                 'EMPID
            DJobTitle.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobTitle")           '職稱
            DJobCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobCode")             '職稱代碼
            DDepoName.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoName")           '公司別
            DDepoCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DepoCode")           '公司別代碼
            DDivision.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Division")           '部門
            DDivisionCode.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DivisionCode")   '部門代碼

            DAfter.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("After")                 '事前,事後
            DJobAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("JobAgent")           '職務代理人
            DTimeOffAgent.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("TimeOffAgent")   '代請假人
            DVacation.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VacationCode") + _
                           ":" + _
                           DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Vacation")             '假別

            DEvidence.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Evidence")           '憑證
            DSalary.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("Salary")               '薪水
            DDieType.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("DieType")             '喪假別
            DVDays.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("VDays").ToString        '可請天數

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1") <> "" Then                 '加班No1 
                LOTNo1.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
                LOTNo1.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo1")
            Else
                LOTNo1.Visible = False
                DOTHours1.Visible = False
            End If
            DOTHours1.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours1"), 1)  '加班No1-時數

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2") <> "" Then                 '加班No2 
                LOTNo2.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
                LOTNo2.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo2")
            Else
                LOTNo2.Visible = False
                DOTHours2.Visible = False
            End If
            DOTHours2.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours2"), 1)  '加班No2-時數

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3") <> "" Then                 '加班No3
                LOTNo3.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
                LOTNo3.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo3")
            Else
                LOTNo3.Visible = False
                DOTHours3.Visible = False
            End If
            DOTHours3.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours3"), 1)  '加班No3-時數

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4") <> "" Then                         '加班No4
                LOTNo4.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
                LOTNo4.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo4")
            Else
                LOTNo4.Visible = False
                DOTHours4.Visible = False
            End If
            DOTHours4.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours4"), 1)  '加班No4-時數

            If DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5") <> "" Then                         '加班No5
                LOTNo5.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
                LOTNo5.NavigateUrl = "HR_OverTimeSheet_03.aspx?pNo=" & DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTNo5")
            Else
                LOTNo5.Visible = False
                DOTHours5.Visible = False
            End If
            DOTHours5.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours5"), 1)  '加班No5-時數

            DOTHours.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("OTHours"), 1)    '加班總時數

            DBStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartDate")       '預定開始日期
            DBStartH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BStartH").ToString    '預定開始時
            DBEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndDate")           '預定結束日期
            DBEndH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BEndH").ToString        '預定結束時
            DBDays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("BDays"), 1)    '預定天

            DAStartDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartDate")       '實際開始日期
            DAStartH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AStartH").ToString    '實際開始時
            DAEndDate.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndDate")           '實際結束日期
            DAEndH.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("AEndH").ToString        '實際結束時
            DADays.Text = FormatNumber(DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("ADays"), 1)    '實際天

            DFReason.Text = DBDataSet1.Tables("F_TimeOffSheet").Rows(0).Item("FReason")             '請假理由

            BCardTime.Attributes("onclick") = "ShowCardTime();"
            BVARecord.Attributes("onclick") = "ShowVacation();"    '請假記錄
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
