
Imports System.Data

Partial Class MaintCustomer
    Inherits PageBase

    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Button1.Attributes.Add("onclick", "history.go(-1);return false;")
        Button2.Attributes.Add("onclick", "if (!confirm('" & ForProject.strSaveAlertMessage & "')){return false;}")

        If Not IsPostBack Then
            Dim strStatus = Request.QueryString("pFun")

            If strStatus = "MOD" Then

                Dim uDataBase As New Utility.DataBase
                uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())

                Dim strSql As String = "Select * from M_CustControl where Unique_ID=" & Request.QueryString("pID")
                Dim dt As DataTable = uDataBase.GetDataTable(strSql)
                Dim uCommon As New Utility.Common

                DActive.SelectedValue = dt.Rows(0)("Active")
                DUniqueID.Text = dt.Rows(0)("Unique_ID")

                DCustCode.Text = dt.Rows(0)("CustCode")
                DCustCode.ReadOnly = True
                DCustName.Text = dt.Rows(0)("CustName")
                DIntCust.SelectedValue = dt.Rows(0)("IntCust")

                DMailToName1.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToName1"), "")
                TextBox3.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToAddress1"), "")
                DMailToPosition1.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToPosition1"), "")
                DMailToName2.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToName2"), "")
                TextBox2.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToAddress2"), "")
                DMailToPosition2.Text = uCommon.ReplaceDBnull(dt.Rows(0)("MailToPosition2"), "")

                DMailCCList.SelectedValue = dt.Rows(0)("MailCCList")
                DSalesMan.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SalesMan"), "")
                DSalesCode.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SalesCode"), "")
                Textbox1.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SalesMailAddress"), "")

                DPDFCreate.SelectedValue = dt.Rows(0)("PDF_Create")
                DPDFRecreate.SelectedValue = dt.Rows(0)("PDF_Recreate")
                DPDFPeriod.Text = uCommon.ReplaceDBnull(dt.Rows(0)("PDF_Period"), "")
                DPDFFirstTime.Text = uCommon.ReplaceDBnull(dt.Rows(0)("PDF_FirstTime"), "")
                DPDFLastTime.Text = uCommon.ReplaceDBnull(dt.Rows(0)("PDF_LastTime"), "")

                DSMTPSend.SelectedValue = dt.Rows(0)("SMTP_Send")
                DSMTPResend.SelectedValue = dt.Rows(0)("SMTP_ReSend")
                DSMTPPeriod.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SMTP_Period"), "")
                DSMTPFirstTime.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SMTP_FirstTime"), "")
                DSMTPLastTime.Text = uCommon.ReplaceDBnull(dt.Rows(0)("SMTP_LastTime"), "")

            End If
        End If


    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())
        Dim uJScript As New Utility.JScript

        If Not String.IsNullOrEmpty(TextBox3.Text.Trim) And String.IsNullOrEmpty(DMailToName1.Text.Trim) Then
            uJScript.PopMsg(Me, "請輸入收件者名稱")
        Else
            Dim strStatus = Request.QueryString("pFun")
            Dim strSql As String = "Select Count(CustCode) from M_CustControl where CustCode='" & DCustCode.Text.Trim & "'"
            Dim strResult As String = uDataBase.SelectQuery(strSql)

            'insert history
            Dim strValue As String = DCustCode.Text.Trim & "-" & _
                                     DCustName.Text.Trim & "-" & _
                                     DMailToName1.Text.Trim & "-" & _
                                     TextBox3.Text.Trim & "-" & _
                                     DMailToPosition1.Text.Trim & "-" & _
                                     DMailToName2.Text.Trim & "-" & _
                                     TextBox2.Text.Trim & "-" & _
                                     DMailToPosition2.Text.Trim & "-" & _
                                     DMailCCList.Text.Trim & "-" & _
                                     DSalesCode.Text.Trim & "-" & _
                                     DSalesMan.Text.Trim & "-" & _
                                     Textbox1.Text.Trim & "-" & _
                                     DPDFCreate.SelectedValue & "-" & _
                                     DSMTPSend.SelectedValue

            Dim cmdSql As SqlClient.SqlCommand

            If strStatus = "MOD" Then
                cmdSql = New SqlClient.SqlCommand(strSql)

                If strResult = 0 Then
                    uJScript.PopMsg(Me, ForProject.strRecordNotExist)
                Else
                    strSql = "Select * from M_CustControl where Unique_ID=" & Request.QueryString("pID")
                    Dim dt As DataTable = uDataBase.GetDataTable(strSql)
                    Dim strBefore As String = dt.Rows(0)("CustCode") & "-" & _
                                              dt.Rows(0)("CustName") & "-" & _
                                              dt.Rows(0)("MailToName1") & "-" & _
                                              dt.Rows(0)("MailToAddress1") & "-" & _
                                              dt.Rows(0)("MailToPosition1") & "-" & _
                                              dt.Rows(0)("MailToName2") & "-" & _
                                              dt.Rows(0)("MailToAddress2") & "-" & _
                                              dt.Rows(0)("MailToPosition2") & "-" & _
                                              dt.Rows(0)("MailCCList") & "-" & _
                                              dt.Rows(0)("SalesCode") & "-" & _
                                              dt.Rows(0)("SalesMan") & "-" & _
                                              dt.Rows(0)("SalesMailAddress") & "-" & _
                                              dt.Rows(0)("PDF_Create") & "-" & _
                                              dt.Rows(0)("SMTP_Send")
                    ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintCustControl", "MOD", strBefore, strValue)

                    strSql = "Update M_CustControl Set Active=@Active,CustCode=@CustCode,CustName=@CustName,IntCust=@IntCust,MailToName1=@MailToName1,MailToAddress1=@MailToAddress1," & _
                                    "MailToPosition1=@MailToPosition1,MailToName2=@MailToName2,MailToAddress2=@MailToAddress2,MailToPosition2=@MailToPosition2,MailCCList=@MailCCList,SalesMan=@SalesMan,SalesCode=@SalesCode,SalesMailAddress=@SalesMailAddress," & _
                                    "PDF_Create=@PDF_Create,PDF_Recreate=@PDF_Recreate,SMTP_Send=@SMTP_Send,SMTP_ReSend=@SMTP_ReSend,ModifyUser=@ModifyUser,ModifyTime=@ModifyTime where Unique_ID=@Unique_ID"

                    cmdSql = New SqlClient.SqlCommand(strSql)
                    'ADD
                    cmdSql.Parameters.AddWithValue("@Active", DActive.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@CustCode", DCustCode.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@CustName", DCustName.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@IntCust", DIntCust.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@MailToName1", DMailToName1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToAddress1", TextBox3.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToPosition1", DMailToPosition1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToName2", DMailToName2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToAddress2", TextBox2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToPosition2", DMailToPosition2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailCCList", DMailCCList.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesMan", DSalesMan.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesCode", DSalesCode.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesMailAddress", Textbox1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@PDF_Create", DPDFCreate.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@SMTP_Send", DSMTPSend.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@PDF_Recreate", DPDFRecreate.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@SMTP_ReSend", DSMTPResend.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@ModifyUser", Request.QueryString("pUserID"))
                    cmdSql.Parameters.AddWithValue("@ModifyTime", Now)
                    cmdSql.Parameters.AddWithValue("@Unique_ID", Request.QueryString("pID"))

                    uDataBase.ExecuteNonQuery(cmdSql)
                    Response.Redirect("MaintCustControl.aspx?pUserID=" & Request.QueryString("pUserID"))

                End If

            Else
                '判斷是否已存在
                If strResult <> "0" Then
                    uJScript.PopMsg(Me, ForProject.strRecordExist)
                Else


                    ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintCustControl", "ADD", "", strValue)

                    strSql = "Insert into M_CustControl (Active,CustCode,CustName,IntCust,MailToName1,MailToAddress1," & _
                    "MailToPosition1,MailToName2,MailToAddress2,MailToPosition2,MailCCList,SalesMan,SalesCode,SalesMailAddress," & _
                    "PDF_Create,PDF_Recreate,SMTP_Send,SMTP_ReSend,CreateUser,CreateTime,ModifyUser,ModifyTime) values (@Active,@CustCode,@CustName,@IntCust,@MailToName1,@MailToAddress1,@MailToPosition1,@MailToName2," & _
                    "@MailToAddress2,@MailToPosition2,@MailCCList,@SalesMan,@SalesCode,@SalesMailAddress,@PDF_Create,@PDF_Recreate,@SMTP_Send,@SMTP_ReSend,@CreateUser,@CreateTime,@ModifyUser,@ModifyTime)"

                    cmdSql = New SqlClient.SqlCommand(strSql)
                    'ADD
                    cmdSql.Parameters.AddWithValue("@Active", DActive.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@CustCode", DCustCode.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@CustName", DCustName.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@IntCust", DIntCust.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@MailToName1", DMailToName1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToAddress1", TextBox3.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToPosition1", DMailToPosition1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToName2", DMailToName2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToAddress2", TextBox2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailToPosition2", DMailToPosition2.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@MailCCList", DMailCCList.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesMan", DSalesMan.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesCode", DSalesCode.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@SalesMailAddress", Textbox1.Text.Trim)
                    cmdSql.Parameters.AddWithValue("@PDF_Create", DPDFCreate.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@PDF_Recreate", DPDFRecreate.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@SMTP_Send", DSMTPSend.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@SMTP_ReSend", DSMTPResend.SelectedValue)
                    cmdSql.Parameters.AddWithValue("@CreateUser", Request.QueryString("pUserID"))
                    cmdSql.Parameters.AddWithValue("@CreateTime", Now)

                    cmdSql.Parameters.AddWithValue("@ModifyUser", Request.QueryString("pUserID"))
                    cmdSql.Parameters.AddWithValue("@ModifyTime", Now)

                    uDataBase.ExecuteNonQuery(cmdSql)
                    Response.Redirect("MaintCustControl.aspx?pUserID=" & Request.QueryString("pUserID"))

                End If


            End If
        End If
        
        


    End Sub



End Class
