Imports System.Data
Imports System.Data.OleDb

Partial Class HR_VacationList01
    Inherits System.Web.UI.Page

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
    Dim wYear As String           '工作年
    Dim wEmpID As String          'Emp-ID
    Dim wNo As String             '委託No
    Dim wVacationCode As String   '假別代碼
    Dim wDieTypeCode As String    '喪假類別

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetParameter()
        If Not Me.IsPostBack Then
            VacationData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wYear = Left(Request.QueryString("pMonth"), 4)  '工作年
        wEmpID = Request.QueryString("pEmpID")          'Emp-ID
        wNo = Request.QueryString("pNo")                '委託No
        wVacationCode = Left(Request.QueryString("pVacation"), 1)  '假別代碼
        wDieTypeCode = Left(Request.QueryString("pDieType"), 1)    '喪假類別
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     篩選請假資料
    '**
    '*****************************************************************
    Sub VacationData()
        Dim wStartDate As String = ""
        Dim wEndDate As String = ""

        If CInt(Now.Date.Month.ToString) < 4 Then
            wStartDate = CStr(CInt(Now.Date.Year.ToString) - 1) + "/4/1"
            wEndDate = Now.Date.Year.ToString + "/3/31"
        Else
            wStartDate = Now.Date.Year.ToString + "/4/1"
            wEndDate = CStr(CInt(Now.Date.Year.ToString) + 1) + "/3/31"
        End If

        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBDataSet2 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        '--------------------------------------------------------------------------------------------------------------------
        '上線前資料
        'SQL = "SELECT "
        'SQL = SQL + "'完成' As StsDesc, "
        'SQL = SQL + "Name + '('+ EmpID +')' As NameDesc, "
        'If wVacationCode = "X" Then
        'SQL = SQL + "CodeName + '('+ Code1 + '.' + CodeName1 +')' As VacationDesc, "
        'Else
        'SQL = SQL + "CodeName As VacationDesc, "
        'End If
        'SQL = SQL + "'2008/1/1 ~ HRW上線前' As AVacationDate, "
        'SQL = SQL + "CodeValue As DaysDesc "
        'SQL = SQL + "FROM HR_BeforeVacation "

        'SQL = SQL + "Where EmpID  =  '" & wEmpID & "' "
        '事假/家顧
        'If wVacationCode = "I" Or wVacationCode = "Y" Then
        'SQL = SQL + "  And ( Code  = 'I' or Code  = 'Y' ) "
        'Else
        '病假/生理
        'If wVacationCode = "S" Or wVacationCode = "Z" Then
        'SQL = SQL + "  And ( Code  = 'S' or Code  = 'Z' ) "
        'Else
        '其他
        'SQL = SQL + "  And Code  = '" & wVacationCode & "' "
        '喪假
        'If wVacationCode = "X" Then
        'SQL = SQL + "  And Code1  Like '" & wDieTypeCode & "%' "
        'End If
        'End If
        'End If
        'SQL = SQL + "Order by Code Desc "

        'Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter2.Fill(DBDataSet2, "Before")
        'GridView2.DataSource = DBDataSet2
        'GridView2.DataBind()
        '--------------------------------------------------------------------------------------------------------------------
        '系統資料
        SQL = "SELECT "
        SQL = SQL + "case sts when 0 then '核定中' when 1 then '完成' else '取消' end As StsDesc, "
        SQL = SQL + "Name + '('+ EmpID +')' As NameDesc, "
        If wVacationCode = "X" Then
            SQL = SQL + "Vacation + '('+ DieType +')' As VacationDesc, "
        Else
            SQL = SQL + "Vacation As VacationDesc, "
        End If
        SQL = SQL + "Convert(VARCHAR(10), AStartDate, 111) + ' ' + str(AStartH) + ':00 ~ ' + "
        SQL = SQL + "Convert(VARCHAR(10), AEndDate, 111)   + ' ' + str(AEndH) + ':00'  As AVacationDate, "
        SQL = SQL + "ADays As DaysDesc "
        SQL = SQL + "FROM F_TimeOffSheet "

        SQL = SQL + "Where EmpID  =  '" & wEmpID & "' "
        SQL = SQL + "  And No     <> '" & wNo & "' "
        SQL = SQL + "  And AEndDate  >= '" & wStartDate & "' "
        SQL = SQL + "  And AEndDate  <= '" & wEndDate & "' "
        '事假/家顧
        If wVacationCode = "I" Or wVacationCode = "Y" Then
            SQL = SQL + "  And ( VacationCode  = 'I' or VacationCode  = 'Y' ) "
        Else
            '病假/生理
            If wVacationCode = "S" Or wVacationCode = "Z" Then
                SQL = SQL + "  And ( VacationCode  = 'S' or VacationCode  = 'Z' ) "
            Else
                '其他
                SQL = SQL + "  And VacationCode  = '" & wVacationCode & "' "
                '喪假
                If wVacationCode = "X" Then
                    SQL = SQL + "  And DieType  Like '" & wDieTypeCode & "%' "
                End If
            End If
        End If
        SQL = SQL + "Order by VacationCode, AStartDate Desc, AEndDate Desc "

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "Vacation")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()

        OleDbConnection1.Close()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(4).Text = FormatNumber(CDbl(e.Row.Cells(4).Text.ToString), 1)
        End If
    End Sub

    Protected Sub GridView2_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView2.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(4).Text = FormatNumber(CDbl(e.Row.Cells(4).Text.ToString), 1)
        End If
    End Sub
End Class
