Imports System.Web.UI.WebControls
Imports System.Data
Imports System.Data.SqlClient

Partial Class Xml2Excel
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '
    Dim NowDateTime As String       '現在日期時間
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            SetSearchItem()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetSearchItem)
    '**     設定搜尋欄位
    '**
    '*****************************************************************
    Sub SetSearchItem()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(匯入)
    '**     按鈕(匯入)
    '**
    '*****************************************************************
    Protected Sub BImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BImport.Click
        Dim Path As String = Server.MapPath(uCommon.GetAppSetting("XMLFileUpload"))
        Dim UploadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                             CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)
        Dim FileName As String = ""
        '
        If DSourcefile.PostedFile.FileName <> "" Then  '判斷有檔案上傳
            FileName = UploadDateTime & "-" & IO.Path.GetFileName(DSourcefile.PostedFile.FileName)
            Try    '上傳圖檔
                DSourcefile.PostedFile.SaveAs(Path & FileName)
                '
                Dim AuthorsDataSet As New System.Data.DataSet()
                AuthorsDataSet.ReadXml(Path & FileName)
                GridView1.DataSource = AuthorsDataSet
                GridView1.DataBind()
                '
            Catch ex As Exception
            End Try
        End If
        '
    End Sub
    '*****************************************************************
    '**
    '**     按鈕(轉Excel)
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BExcel.Click
        Dim DownloadDateTime = CStr(DateTime.Now.Year) + CStr(DateTime.Now.Month) + CStr(DateTime.Now.Day) + _
                               CStr(DateTime.Now.Hour) + CStr(DateTime.Now.Minute) + CStr(DateTime.Now.Second)
        '
        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "

        Response.AppendHeader("Content-Disposition", "attachment; filename=" & DownloadDateTime & "_Xml2Excel.xls")
        Response.ContentType = "application/vnd.ms-excel"

        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        For i As Integer = 0 To 37
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub
End Class
