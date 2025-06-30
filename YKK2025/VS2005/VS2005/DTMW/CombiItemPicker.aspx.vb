Imports System
Imports System.Data
Imports System.Data.SqlClient

Partial Class CombiItemPicker
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
        Dim AllItem As String = ""
        Dim sql As String = ""

        sql = " select * from M_referp "
        sql = sql + " where cat = 5001 and dkey = 'COMBIItem'"
        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        'ListBox3.Items.Clear()

        For I = 0 To dt.Rows.Count - 1
            If I = 0 Then
                AllItem = dt.Rows(I).Item("Data")
            Else
                AllItem = AllItem + "," + dt.Rows(I).Item("Data")
            End If
        Next

        For I = 0 To ListBox1.Items.Count - 1

            If ListBox1.Items(I).Selected Then

                If Str1 = "" Then
                    Str1 = ListBox1.Items(I).Text
                    Str2 = ListBox1.Items(I).Value
                    If ListBox1.Items(I).Text = "全部" Then
                        Str2 = AllItem
                        Exit For
                    End If
                Else
                    Str1 = Str1 + "," + ListBox1.Items(I).Text
                    Str2 = Str2 + "," + ListBox1.Items(I).Value
                    If ListBox1.Items(I).Text = "全部" Then
                        Str2 = AllItem
                        Exit For
                    End If
                End If


            End If
        Next
        TextBox1.Text = Str1
        TextBox2.Text = Str2


        Dim js As String = ""
        js &= "var obj = window.opener.document.all.DCombiItem1;"
        js &= "obj.value = '" & Str1 & "';"
        js &= "var obj1 = window.opener.document.all.DCombiItem2;"
        js &= "obj1.value = '" & Str2 & "';"
        js &= "window.opener.document.forms[0].submit();"
        js &= "window.close();"


        Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

    End Sub

    Sub GetFormList()
        Dim sql As String
        Dim i As Integer
        sql = " select * from M_referp "
        sql = sql + " where cat = 5001 and dkey = 'COMBIItem'"
        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        'ListBox3.Items.Clear()
        ListBox1.Items.Add("全部")
        For i = 0 To dt.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dt.Rows(i).Item("Data")
            ListItem1.Value = dt.Rows(i).Item("Data")
            ListBox1.Items.Add(ListItem1)
        Next


    End Sub


End Class
