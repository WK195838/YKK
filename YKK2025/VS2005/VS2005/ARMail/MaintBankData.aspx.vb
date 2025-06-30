Imports System.Data

Partial Class MaintBankData
    Inherits PageBase
    Dim uDataBase As New Utility.DataBase
    Dim uCommon As New Utility.Common
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Button1.Attributes.Add("onclick", "if (!confirm('" & ForProject.strSaveAlertMessage & "')){return false;}")

        If Not IsPostBack Then
            Dim strStatus = Request.QueryString("pFun")
            If strStatus = "MOD" Then
                uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())
                Dim sql As String = "Select CustCode,CustName,SalesCode,BankName,BankAddress,BankACNo,BankACName,Swift , ID from M_Bank where CustCode='" & Request.QueryString("pID") & "'"
                Dim dt As DataTable = uDataBase.GetDataTable(sql)

                DCustCode.Text = uCommon.ReplaceDBnull(dt.Rows(0)(0).ToString, "")
                DCustName.Text = uCommon.ReplaceDBnull(dt.Rows(0)(1).ToString, "")
                DSalesCode.Text = uCommon.ReplaceDBnull(dt.Rows(0)(2).ToString, "")
                DBankName.Text = uCommon.ReplaceDBnull(dt.Rows(0)(3).ToString, "")
                DBankAddress.Text = uCommon.ReplaceDBnull(dt.Rows(0)(4).ToString, "")
                DBankACNo.Text = uCommon.ReplaceDBnull(dt.Rows(0)(5).ToString, "")
                DBankACName.Text = uCommon.ReplaceDBnull(dt.Rows(0)(6).ToString, "")
                DSwift.Text = uCommon.ReplaceDBnull(dt.Rows(0)(7).ToString, "")

                DID.Value = uCommon.ReplaceDBnull(dt.Rows(0)(8).ToString, "")

            End If
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim uDataBase As New Utility.DataBase
        uDataBase.DefaultConnection = New SqlClient.SqlConnection(ForProject.GetConnectionString())
        Dim uJScript As New Utility.JScript

        Dim strStatus = Request.QueryString("pFun")
        Dim checkSql As String = "Select Count(CustCode) from M_Bank Where CustCode='" & DCustCode.Text.Trim & "'"
        Dim checkResult As Integer = 0
        Dim strValue As String = DCustCode.Text.Trim & "-" & _
                                DCustName.Text.Trim & "-" & _
                                DSalesCode.Text.Trim & "-" & _
                                DBankName.Text.Trim & "-" & _
                                DBankAddress.Text.Trim & "-" & _
                                DBankACNo.Text.Trim & "-" & _
                                DBankACName.Text.Trim & "-" & _
                                DSwift.Text.Trim
                                            
        Dim sql As String = ""
        If strStatus = "MOD" Then
            checkSql &= " and ID <> " & DID.Value
            checkResult = uDataBase.SelectQuery(checkSql)
            If checkResult = 1 Then
                uJScript.PopMsg(Me, ForProject.strRecordExist)

            Else
                sql = "Select * from M_Bank Where ID=" & DID.Value
                Dim dt As DataTable = uDataBase.GetDataTable(sql)
                Dim strBefore As String = dt.Rows(0)("CustCode") & "-" & _
                                        dt.Rows(0)("CustName") & "-" & _
                                        dt.Rows(0)("SalesCode") & "-" & _
                                        dt.Rows(0)("BankName") & "-" & _
                                        dt.Rows(0)("BankAddress") & "-" & _
                                        dt.Rows(0)("BankACNo") & "-" & _
                                        dt.Rows(0)("BankACName") & "-" & _
                                     dt.Rows(0)("Swift")
                                                              
                ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintBank", "MOD", strBefore, strValue)

                sql = "Update M_Bank Set "
                sql &= "CustCode = '" & DCustCode.Text.Trim & "',"
                sql &= "CustName = '" & DCustName.Text.Trim & "',"
                sql &= "SalesCode = '" & DSalesCode.Text.Trim & "',"
                sql &= "BankName = '" & DBankName.Text.Trim & "',"
                sql &= "BankAddress = '" & DBankAddress.Text.Trim & "',"
                sql &= "BankACNo = '" & DBankACNo.Text.Trim & "',"
                sql &= "BankACName = '" & DBankACName.Text.Trim & "',"
                sql &= "Swift = '" & DSwift.Text.Trim & "' Where ID=" & DID.Value

                uDataBase.ExecuteNonQuery(sql)
                Response.Redirect("MaintBank.aspx?pUserID=" & Request.QueryString("pUserID"))
            End If


        Else
            checkResult = uDataBase.SelectQuery(checkSql)
            If checkResult = 0 Then

                ForProject.InsertMaintHistory(Request.QueryString("pUserID"), "MaintBank", "ADD", "", strValue)
                sql = "Insert into M_Bank (CustCode,CustName,SalesCode,BankName,BankAddress,BankACNO,BankACName,Swift) Values ("
                sql &= "'" & DCustCode.Text.Trim & "',"
                sql &= "'" & DCustName.Text.Trim & "',"
                sql &= "'" & DSalesCode.Text.Trim & "',"
                sql &= "'" & DBankName.Text.Trim & "',"
                sql &= "'" & DBankAddress.Text.Trim & "',"
                sql &= "'" & DBankACNo.Text.Trim & "',"
                sql &= "'" & DBankACName.Text.Trim & "',"
                sql &= "'" & DSwift.Text.Trim & "')"

                uDataBase.ExecuteNonQuery(sql)

                Response.Redirect("MaintBank.aspx?pUserID=" & Request.QueryString("pUserID"))

            Else
                uJScript.PopMsg(Me, ForProject.strRecordExist)


            End If


        End If
    End Sub
End Class
