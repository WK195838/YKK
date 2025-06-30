Imports System.Data
Imports System.Data.OleDb
Imports System.Text

Public Class VacationDatePicker
    Inherits System.Web.UI.Page

#Region " Web Form �]�p�u�㲣�ͪ��{���X "

    '���� Web Form �]�p�u��һݪ��I�s�C
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Calendar1 As System.Web.UI.WebControls.Calendar

    '�`�N: �U�C�w�d��m�ŧi�O Web Form �]�p�u��ݭn�����ءC
    '�ФŧR���β��ʥ��C
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: ���� Web Form �]�p�u��һݪ���k�I�s
        '�ФŨϥε{���X�s�边�i��ק�C
        InitializeComponent()
    End Sub

#End Region

    Public YKK As New YKK_SPDClass   'YKK�@�q�[��

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
