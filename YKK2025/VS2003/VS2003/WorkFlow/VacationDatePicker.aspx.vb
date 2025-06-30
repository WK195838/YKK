Imports System.Data
Imports System.Data.OleDb
Imports System.Text

Public Class VacationDatePicker
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Calendar1 As System.Web.UI.WebControls.Calendar

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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Nothing to do, here
    End Sub

    Private Sub Calendar1_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles Calendar1.DayRender
        '// Clear the link from this day
        e.Cell.Controls.Clear()

        Dim DBDataSet1 As New DataSet
        Dim SQL As String
        Dim SalaryYearMonth As String = ""
        'Salary Year Month
        If e.Day.Date.Month < 10 Then
            SalaryYearMonth = CStr(e.Day.Date.Year) + "/0" + CStr(e.Day.Date.Month)
        Else
            SalaryYearMonth = CStr(e.Day.Date.Year) + "/" + CStr(e.Day.Date.Month)
        End If
        '// Add the custom link
        Dim Link As System.Web.UI.HtmlControls.HtmlGenericControl
        Link = New System.Web.UI.HtmlControls.HtmlGenericControl
        Link.TagName = "a"
        Link.InnerText = e.Day.DayNumberText

        If Request.QueryString("field3") <> "" Then
            Link.Attributes.Add("href", String.Format("JavaScript:window.opener.document.{0}.value = '{1:d}'; window.opener.document.{2}.value = '{3:d}'; window.close();", Request.QueryString("field2"), e.Day.Date, Request.QueryString("field3"), SalaryYearMonth))
        Else
            Link.Attributes.Add("href", String.Format("JavaScript:window.opener.document.{0}.value = '{1:d}'; window.close();", Request.QueryString("field2"), e.Day.Date))
        End If

        '// By default, this will highlight today's date.
        If e.Day.IsSelected Then
            Link.Attributes.Add("style", Me.Calendar1.SelectedDayStyle.ToString())
        End If


        '// Now add our custom link to the page
        e.Cell.Controls.Add(Link)

    End Sub

End Class
