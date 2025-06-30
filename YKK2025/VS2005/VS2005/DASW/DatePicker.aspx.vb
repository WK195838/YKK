Imports System
Imports System.Data
Imports System.Data.SqlClient

Partial Class DatePicker
    Inherits System.Web.UI.Page

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()

    Dim holidays(13, 32) As String
    Protected dsHolidays As DataSet

    Protected Sub Page_Load(ByVal sender As Object, _
            ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Calendar1.VisibleDate = DateTime.Today
            FillHolidayDataset()

        End If

    End Sub

    Protected Sub FillHolidayDataset()
        Dim firstDate As New DateTime(Calendar1.VisibleDate.Year, _
             Calendar1.VisibleDate.Month, 1)
        Dim lastDate As DateTime = GetFirstDayOfNextMonth()

        dsHolidays = GetCurrentMonthData(firstDate, lastDate)
    End Sub

    Protected Function GetFirstDayOfNextMonth() As DateTime
        Dim monthNumber, yearNumber As Integer
        If Calendar1.VisibleDate.Month = 12 Then
            monthNumber = 1
            yearNumber = Calendar1.VisibleDate.Year + 1
        Else
            monthNumber = Calendar1.VisibleDate.Month + 1
            yearNumber = Calendar1.VisibleDate.Year
        End If
        Dim lastDate As New DateTime(yearNumber, monthNumber, 1)
        Return lastDate
    End Function

    Protected Sub Calendar1_VisibleMonthChanged(ByVal sender As Object, _
            ByVal e As System.Web.UI.WebControls.MonthChangedEventArgs) _
            Handles Calendar1.VisibleMonthChanged
        FillHolidayDataset()
    End Sub

    Function GetCurrentMonthData(ByVal firstDate As DateTime, ByVal lastDate As DateTime) As DataSet
        '抓資料庫休假的資訊
        Dim dsMonth As New DataSet
        Dim cs As ConnectionStringSettings
        cs = ConfigurationManager.ConnectionStrings("WF_Con")
        Dim connString As String = cs.ConnectionString
        Dim dbConnection As New SqlConnection(connString)
        Dim query As String
        query = "select ymd as HolidayDate from m_vacation where depo='" + Mid(Request.QueryString("pDepo"), 1, 3) + "' and active='1' "

        Dim dbCommand As New SqlCommand(query, dbConnection)
        'dbCommand.Parameters.Add(New SqlParameter("@firstDate", firstDate))
        'dbCommand.Parameters.Add(New SqlParameter("@lastDate", lastDate))


        Dim sqlDataAdapter As New SqlDataAdapter(dbCommand)

        Try
            sqlDataAdapter.Fill(dsMonth)

        Catch
        End Try
        Return dsMonth
    End Function

    Protected Sub Calendar1_DayRender(ByVal sender As Object, _
            ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) _
            Handles Calendar1.DayRender
        '設定假日的字體及顏色

        Dim nextDate As DateTime
        If Not dsHolidays Is Nothing Then
            For Each dr As DataRow In dsHolidays.Tables(0).Rows
                nextDate = CType(dr("HolidayDate"), DateTime)
                If nextDate = e.Day.Date Then
                    e.Cell.BackColor = System.Drawing.Color.NavajoWhite
                    e.Cell.ForeColor = System.Drawing.Color.Red
                End If
            Next
        End If
        If e.Day.IsOtherMonth Then
            e.Cell.Controls.Clear()
        Else
            Dim aDate As Date = e.Day.Date
            Dim aHoliday As String = holidays(aDate.Month, aDate.Day)
            If (Not aHoliday Is Nothing) Then
                Dim aLabel As Label = New Label()
                aLabel.Text = "<br>" & aHoliday
                e.Cell.Controls.Add(aLabel)
            End If
        End If

        '限定只能選不含依賴日大於4個工作天
        ' If e.Day.Date < Today.AddDays(4) Then
        'e.Cell.Controls.Clear()
        'e.Cell.Text = e.Day.Date.Day.ToString()
        'End If




    End Sub

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged

        FillHolidayDataset()
        '
        '傳回父視窗三個值 , 日期 , 日期別 , 薪資所屬年月
        Dim param As String = "JavaScript:"
        param &= "window.opener.document.getElementById('" & Request.QueryString("field1") & "').value = '" & Me.Calendar1.SelectedDate.ToString("yyyy/MM/dd") & "'; "
        'param &= "window.opener.HideGridview();"
        param &= "window.close(); "

        Dim lJScript As New System.Text.StringBuilder("")
        lJScript.Append(param)
        ClientScript.RegisterStartupScript(Me.GetType(), "ReturnValue", lJScript.ToString(), True)
    End Sub


End Class
