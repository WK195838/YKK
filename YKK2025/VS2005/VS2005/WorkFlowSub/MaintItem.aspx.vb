Imports System.Data
Imports System.Data.SqlClient
Partial Class _MaintItem
    Inherits System.Web.UI.Page
    Dim Userid As String
    Dim xCat As String
    Dim NowDateTime As String       '現在日期時間

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim msg As String = ""

        NowDateTime = CStr(DateTime.Today) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時

        Userid = Request.QueryString("pUserID")
        xCat = Request.QueryString("pCat")

        If Not IsPostBack Then

            If Request.QueryString("pfun") = "MOD" Then
                Dim UniqueID As String = Request.QueryString("pid")
                Dim sqlStr As String = "select Unique_ID,cat ,dkey,data  from m_referp where Unique_ID = @UniqueID"
                Dim dbcon As New SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WorkFlow_Con").ToString)
                Dim dbcmd As New SqlCommand(sqlStr, dbcon)
                dbcmd.Parameters.AddWithValue("@UniqueID", UniqueID)
                Dim dbAdpter As New SqlDataAdapter(dbcmd)
                Dim ds As New Data.DataSet
                dbAdpter.Fill(ds)
                Dim dt As DataTable = ds.Tables(0)
                BID.Text = dt.Rows(0)("Unique_ID").ToString
                BCAT.Text = dt.Rows(0)("CAT").ToString
                BDKEY.Text = dt.Rows(0)("DKEY").ToString
                BDATA.Text = dt.Rows(0)("DATA").ToString

                BID.ReadOnly = True
                BID.BackColor = Drawing.Color.LightGray
                BCAT.ReadOnly = True
                BCAT.BackColor = Drawing.Color.LightGray
                msg = "修改"
            Else
                BID.ReadOnly = True
                BID.BackColor = Drawing.Color.LightGray
                BCAT.Text = Request.QueryString("pcat")
                BCAT.ReadOnly = True
                BCAT.BackColor = Drawing.Color.LightGray
                msg = "新增"
            End If
            '帶入已寫好的ｊｓ檔
            BSave.Attributes("onclick") = "IsConfirm('" & msg & "');"

        End If

    End Sub


    Private Sub BSave_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSave.ServerClick
        Dim sqlStr As String = ""
        Dim msg As String = ""
        Dim dbcon As New SqlConnection(System.Web.Configuration.WebConfigurationManager.ConnectionStrings("WorkFlow_Con").ToString)
        Dim dbcmd As SqlCommand

        '判斷資料是否存在
        sqlStr = " select count(*)  from M_referp where Cat =@Cat and dkey =@dkey "
        dbcmd = New SqlCommand(sqlStr, dbcon)
        dbcmd.Parameters.AddWithValue("@Unique_ID", BID.Text.Trim)
        dbcmd.Parameters.AddWithValue("@CAT", BCAT.Text)
        dbcmd.Parameters.AddWithValue("@DKEY", UCase(BDKEY.Text))
        dbcmd.Parameters.AddWithValue("@DATA", BDATA.Text)
        dbcmd.Connection.Open()

        '判斷要新增還是修改
        If Request.QueryString("pFun") = "MOD" Then
            If dbcmd.ExecuteScalar = 1 Then
                sqlStr = "Update M_Referp set CAT =@CAT,Dkey=@Dkey,Data=@Data,ModifyUser=@User,ModifyTime=@Time where Unique_ID =@Unique_ID"
                msg = "已變更"
            Else
                msg = "此筆資料不存在"
            End If
        Else
            If dbcmd.ExecuteScalar = 0 Then
                sqlStr = "Insert into M_Referp (Cat,Dkey,Data,CreateUser,CreateTime) values (@Cat,@Dkey,@Data,@User,@Time)"
                msg = "已新增"
            Else
                msg = "此筆資料已存在"
            End If
        End If

        dbcmd.Connection.Close()
        'show 提醒視窗
        Dim jsStr As String = "alert('" & msg & "');"

        jsStr &= "location.href('MaintReferp.aspx?pUserID=" & Userid & "&pCat=" & xCat & "');"
        dbcmd = New SqlCommand(sqlStr, dbcon)

        dbcmd.Parameters.AddWithValue("@Unique_ID", BID.Text.Trim)
        dbcmd.Parameters.AddWithValue("@CAT", BCAT.Text)
        dbcmd.Parameters.AddWithValue("@DKEY", UCase(BDKEY.Text))
        dbcmd.Parameters.AddWithValue("@DATA", BDATA.Text)

        dbcmd.Parameters.AddWithValue("@User", Userid)
        dbcmd.Parameters.AddWithValue("@Time", NowDateTime)

        dbcmd.Connection.Open()

        dbcmd.ExecuteNonQuery()

        dbcmd.Connection.Close()
        '註冊java script
        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "abc", jsStr, True)
        'Response.Redirect("Default.aspx")

    End Sub


    Protected Sub BBack_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BBack.ServerClick
        Response.Redirect("MaintReferp.aspx?pUserID=" & Userid & "&pCat=" & xCat)

    End Sub
End Class
