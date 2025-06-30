Imports System
Imports System.Data
Imports System.Data.SqlClient

Partial Class DISPOSALTYPE
    Inherits System.Web.UI.Page

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Me.IsPostBack Then
            GetFormList()
        End If
    End Sub

    Protected Sub BGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGo.Click
        Dim I As Integer = 0
        Dim Str1 As String = ""
        Dim Str2 As String = ""
        Dim AllFormNo As String = ""
        Dim sql As String = ""

        sql = "  Select  * from M_referp"
        sql = sql & " where  cat = '6001'"
        sql = sql & " and dkey = 'DISPOSALTYPE'"
        sql = sql & " order by unique_id"

        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        'ListBox3.Items.Clear()

        For I = 0 To dt.Rows.Count - 1
            If I = 0 Then
                AllFormNo = dt.Rows(I).Item("Data")
            Else
                AllFormNo = AllFormNo + "," + dt.Rows(I).Item("Data")
            End If
        Next

        For I = 0 To ListBox1.Items.Count - 1

            If ListBox1.Items(I).Selected Then

                If Str1 = "" Then
                    Str1 = ListBox1.Items(I).Text
                    Str2 = ListBox1.Items(I).Value
                    If ListBox1.Items(I).Text = "全部" Then
                        Str2 = AllFormNo
                        Exit For
                    End If
                Else
                    Str1 = Str1 + "," + ListBox1.Items(I).Text
                    Str2 = Str2 + "," + ListBox1.Items(I).Value
                    If ListBox1.Items(I).Text = "全部" Then
                        Str2 = AllFormNo
                        Exit For
                    End If
                End If


            End If
        Next
        TextBox1.Text = Str1
        TextBox2.Text = Str2


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DDISPOSALTYPE;"
        js &= "obj.value = '" & Str1 & "';"
        ' js &= "var obj1 = window.opener.document.all.DFormNo2;"
        'js &= "obj1.value = '" & Str2 & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub

    Sub GetFormList()
        Dim sql As String
        Dim i As Integer
        sql = "  Select  * from M_referp"
        sql = sql & " where  cat = '6001'"
        sql = sql & " and dkey = 'DISPOSALTYPE'"
        sql = sql & " order by unique_id"
        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        'ListBox3.Items.Clear()
        ' ListBox1.Items.Add("全部")
        For i = 0 To dt.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt.Rows(i).Item("Data")
            ListItem1.Value = dt.Rows(i).Item("Data")
            ListBox1.Items.Add(ListItem1)
        Next


    End Sub


End Class
