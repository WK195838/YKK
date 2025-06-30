Imports System.Web.UI.WebControls
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO

Partial Class AutoXml2Excel
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
    Dim xBuyer As String            'Buyer
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not IsPostBack Then
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
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Protected Sub BImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BImport.Click
        Xml2Excel()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Xml2Excel)
    '**      
    '**
    '*****************************************************************
    Sub Xml2Excel()
        Dim Http As String = uCommon.GetAppSetting("Http")
        Dim EDIPath As String = uCommon.GetAppSetting("EDIFile")
        '
        Dim XmlPath As String = Server.MapPath(uCommon.GetAppSetting("XMLFile"))
        Dim TNFPath As String = Server.MapPath(uCommon.GetAppSetting("TNFFile"))
        Dim ADIDASPath As String = Server.MapPath(uCommon.GetAppSetting("ADIDASFile"))
        Dim REEBOKPath As String = Server.MapPath(uCommon.GetAppSetting("REEBOKFile"))
        '
        Dim source As DirectoryInfo = New DirectoryInfo(EDIPath)
        Dim file1 As FileInfo
        '
        Dim xPath, xExcel, sql, xDate, xTime, str, xBuyer, xFileName As String
        Dim i As Integer = 0
        '
        '讀取該目錄下所有檔案清單()
        For Each file1 In source.GetFiles
            '
            ' Create TextWriter
            Dim sw As New System.IO.StringWriter
            Dim hw As New System.Web.UI.HtmlTextWriter(sw)
            Dim sw1 As New System.IO.StringWriter
            Dim hw1 As New System.Web.UI.HtmlTextWriter(sw1)
            '
            ' XML to GridView
            Dim AuthorsDataSet As New System.Data.DataSet()
            AuthorsDataSet.ReadXml(file1.FullName)
            GridView1.DataSource = AuthorsDataSet
            GridView1.DataBind()
            '
            ' 產生Buyer別Excel
            '品牌代號(品牌)
            'A(adidas)
            'B(Reebok)
            'T(TNF)
            'U(UNIQLO)
            xBuyer = ""
            '
            ' ADIDAS
            If UCase(Trim(AuthorsDataSet.Tables(0).Rows(0)("BRAND_NO").ToString)) = "A" Then
                Dim ADIDASDataSet As New System.Data.DataSet()
                ADIDASDataSet.ReadXml(file1.FullName)
                ADIDASGridView.DataSource = ADIDASDataSet
                ADIDASGridView.DataBind()
                '
                xPath = ADIDASPath
                xBuyer = "ADIDAS"
                xFileName = Http + uCommon.GetAppSetting("ADIDASFile")
                '
                ADIDASGridView.RenderControl(hw1)
                xExcel = xPath + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + ".xls"
                System.IO.File.WriteAllText(xExcel, sw1.ToString())
            End If
            '
            ' REEBOK
            If UCase(Trim(AuthorsDataSet.Tables(0).Rows(0)("BRAND_NO").ToString)) = "B" Then
                Dim REEBOKDataSet As New System.Data.DataSet()
                REEBOKDataSet.ReadXml(file1.FullName)
                REEBOKGridView.DataSource = REEBOKDataSet
                REEBOKGridView.DataBind()
                '
                xPath = REEBOKPath
                xBuyer = "REEBOK"
                xFileName = Http + uCommon.GetAppSetting("REEBOKFile")
                '
                REEBOKGridView.RenderControl(hw1)
                xExcel = xPath + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + ".xls"
                System.IO.File.WriteAllText(xExcel, sw1.ToString())
            End If
            '
            ' TNF
            If UCase(Trim(AuthorsDataSet.Tables(0).Rows(0)("BRAND_NO").ToString)) = "T" Then
                Dim TNFDataSet As New System.Data.DataSet()
                TNFDataSet.ReadXml(file1.FullName)
                TNFGridView.DataSource = TNFDataSet
                TNFGridView.DataBind()
                '
                xPath = TNFPath
                xBuyer = "TNF"
                xFileName = Http + uCommon.GetAppSetting("TNFFile")
                '
                TNFGridView.RenderControl(hw1)
                xExcel = xPath + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + ".xls"
                System.IO.File.WriteAllText(xExcel, sw1.ToString())
            End If
            '
            If xBuyer <> "" Then
                '
                '產生Excel File
                GridView1.RenderControl(hw)
                xExcel = xPath + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + "@.xls"
                System.IO.File.WriteAllText(xExcel, sw.ToString())
                '
                '移動XML
                If System.IO.File.Exists(file1.FullName) Then
                    System.IO.File.Move(file1.FullName, XmlPath + file1.Name)    ' Move(原始路徑,目標路徑)
                End If
                '
                ' Write Data to DB
                'SU00029_20120717_144000.xml
                str = Mid(file1.Name, InStr(file1.Name, "_") + 1)
                '20120717_144000.xml
                xDate = Mid(str, 1, InStr(str, "_") - 1)
                xDate = Mid(xDate, 1, 4) + "/" + Mid(xDate, 5, 2) + "/" + Mid(xDate, 7, 2)
                str = Mid(str, InStr(file1.Name, "_") + 2)
                '144000.xml
                xTime = Mid(str, 1, InStr(str, ".") - 1)
                xTime = Mid(xTime, 1, 2) + ":" + Mid(xTime, 3, 2) + ":" + Mid(xTime, 5, 2)
                '
                sql = "INSERT INTO W_EDIData (Customer, Buyer, DataDate, DataTime, EDI, EDIAll, EDIOri, CreateTime) " & _
                "VALUES("
                sql &= "'DINSEN', "
                sql &= "'" & xBuyer & "' ,"
                sql &= "'" & xDate & "' ,"
                sql &= "'" & xTime & "' ,"
                sql &= "'" & xFileName + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + ".xls" & "', "
                sql &= "'" & xFileName + Mid(file1.Name, 1, InStr(file1.Name, ".") - 1) + "@.xls" & "', "
                sql &= "'" & Http + uCommon.GetAppSetting("XMLFile") + file1.Name & "', "
                sql &= "'" & NowDateTime & "' )"
                uDataBase.ExecuteNonQuery(sql)
            End If
            '
            ' Close TextWriter
            hw.Close()
            sw.Close()
            hw1.Close()
            sw1.Close()
        Next
        '
        GridView1.Visible = False
        ADIDASGridView.Visible = False
        REEBOKGridView.Visible = False
        TNFGridView.Visible = False
        '
        DMsg.Text = "接收完成請確認資料！"
        '限IE7使用
        'Response.Write("<script>window.opener=null;window.open(' ','_self');window.close();</script>")
    End Sub
    '
    '*****************************************************************
    '**
    '**     GridView轉文字
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        For i As Integer = 0 To 46
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub

    Protected Sub ADIDASGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ADIDASGridView.RowDataBound
        For i As Integer = 0 To 15
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub

    Protected Sub REEBOKGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles ADIDASGridView.RowDataBound
        For i As Integer = 0 To 15
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub

    Protected Sub TNFGridView_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles TNFGridView.RowDataBound
        For i As Integer = 0 To 15
            e.Row.Cells(i).Attributes.Add("class", "text")
        Next
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

End Class
