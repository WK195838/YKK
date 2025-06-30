Imports System
Imports System.Data
Imports System.Data.SqlClient

Partial Class StsPicker
    Inherits System.Web.UI.Page

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()

    Dim holidays(13, 32) As String
    Protected dsHolidays As DataSet

    Protected Sub Page_Load(ByVal sender As Object, _
            ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub BGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGo.Click
        Dim I As Integer = 0
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        For I = 0 To ListBox1.Items.Count - 1
            If ListBox1.Items(I).Selected Then
                If Str1 = "" Then
                    Str1 = ListBox1.Items(I).Text
                    If ListBox1.Items(I).Text = "完成" Then
                        Str2 = "1"
                    ElseIf ListBox1.Items(I).Text = "取消" Then
                        Str2 = "2"
                    ElseIf ListBox1.Items(I).Text = "核定中" Then
                        Str2 = "0"
                    ElseIf ListBox1.Items(I).Text = "全部" Then
                        Str2 = "0,1,2"
                    End If
                    If ListBox1.Items(I).Text = "全部" Then
                        Str1 = "全部"
                        Exit For
                    End If
                Else
                    Str1 = Str1 + "," + ListBox1.Items(I).Text
                    If Str2 <> "0,1,2" Then
                        If ListBox1.Items(I).Text = "完成" Then
                            Str2 = Str2 + "," + "1"
                        ElseIf ListBox1.Items(I).Text = "取消" Then
                            Str2 = Str2 + "," + "2"
                        ElseIf ListBox1.Items(I).Text = "核定中" Then
                            Str2 = Str2 + "," + "0"
                        ElseIf ListBox1.Items(I).Text = "全部" Then
                            Str2 = "0,1,2"
                        End If
                        If ListBox1.Items(I).Text = "全部" Then
                            Str1 = "全部"
                            Exit For
                        End If
                    End If

                End If
            End If
        Next
        TextBox1.Text = Str1
        TextBox2.Text = Str2


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DSts1;"
        js &= "obj.value = '" & Str1 & "';"
        js &= "var obj1 = window.opener.document.all.DSts2;"
        js &= "obj1.value = '" & Str2 & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub
End Class
