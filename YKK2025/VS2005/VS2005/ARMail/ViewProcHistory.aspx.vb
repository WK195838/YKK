Imports System.Data
Partial Class ViewProcHistory
    Inherits PageBase

    
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        If Not IsPostBack Then
            Dim uDatabase As New Utility.DataBase
            uDatabase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

            TextBox1.Text = uDatabase.SelectQuery("Select CustName + '(' + Custcode + ')' from T_ProcHistory where CustCode='" & Request.QueryString("pCust") & "'")

            GV_Bind()
        End If
    End Sub

    Sub GV_Bind()

        Dim strSql As String = ""
        strSql &= "SELECT top 500 * from T_ProcHistory where custcode='" & Request.QueryString("pCust") & "'"


        strSql &= " order by ProcPeriod DESC"

        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

        GridView1.DataSource = uDataBase.GetDataTable(strSql)
        GridView1.DataBind()

    End Sub

    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GV_Bind()
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub


    Function ConvertDate(ByVal obj As Object) As String
        Dim uCommon As New Utility.Common
        If uCommon.ReplaceDBnull(obj, "") = "" Then

            Return ""

        Else

            Return CDate(obj).ToString("yyyy/MM/dd HH:mm:ss")
        End If

    End Function

    '自訂標題
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If e.Row.RowType = DataControlRowType.Header Then
            Dim header As TableCellCollection = e.Row.Cells
            header.Clear()
            header.Add(New TableHeaderCell)
            header(0).Attributes.Add("rowspan", "2")
            header(0).Text = "期間"
            header.Add(New TableHeaderCell)
            header(1).Attributes.Add("colspan", "4")
            header(1).Text = "對帳單"
            header.Add(New TableHeaderCell)
            header(2).Attributes.Add("colspan", "5")
            header(2).Text = "郵件</TR><TR  align='center' valign='middle' style='color:#FFFFCC;background-color:#CC6600;'>"

            header.Add(New TableHeaderCell)
            header(3).Text = "ID"

            header.Add(New TableHeaderCell)
            header(4).Text = "製作時間"

            header.Add(New TableHeaderCell)
            header(5).Text = "檔案名"

            header.Add(New TableHeaderCell)
            header(6).Text = "附件"

            header.Add(New TableHeaderCell)
            header(7).Text = "傳送時間"

            header.Add(New TableHeaderCell)

            header(8).Text = "傳送者"

            header.Add(New TableHeaderCell)

            header(9).Text = "收件者"

            header.Add(New TableHeaderCell)

            header(10).Text = "CC者"

            header.Add(New TableHeaderCell)

            header(11).Text = "主旨</tr>"


        End If
    End Sub

    '開啟PDF
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim imgbtn As ImageButton = e.Row.FindControl("ImageButton1")
            Dim img As Image = e.Row.FindControl("Image1")
            Dim strMail_ID As String = imgbtn.CommandArgument.ToString.Split("|")(1)
            Dim strPDF_File As String = imgbtn.CommandArgument.ToString.Split("|")(0)

            Dim strFolder As String = ""
            If String.IsNullOrEmpty(strMail_ID) Then
                strFolder = "NoSend"
            Else
                strFolder = "SendOff"
            End If

            Dim uCommon As New Utility.Common
            img.Attributes.Add("Onclick", "OpenFile('" & uCommon.GetAppSetting("PDF_Folder") & strFolder & "/" & strPDF_File & "')")

            
        End If
    End Sub
End Class
