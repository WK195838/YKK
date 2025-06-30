Imports System.Data
Imports System.Data.OleDb
Partial Class PostCodeList
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Dim InputData As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetCountry1()

        End If
    End Sub

    Sub GetData()
        Dim SQL As String
       
        Dim county1, county2, address As String

        county1 = ""
        county2 = ""
        address = ""

        If DCounty1.SelectedValue <> "" And DCounty1.SelectedValue <> "-請選擇-" Then
            county1 = " and County1='" + DCounty1.SelectedValue + "'"
        Else
            county1 = ""
        End If
        If DCounty2.SelectedValue <> "" And DCounty2.SelectedValue <> "-請選擇-" Then
            county2 = " and County2='" + DCounty2.SelectedValue + "'"
        Else
            county2 = ""
        End If
        If DAddress2.SelectedValue <> "" And DAddress2.SelectedValue <> "-請選擇-" Then
            address = " and Address='" + DAddress2.SelectedValue + "'"
        Else
            address = ""
        End If

        SQL = " select *  from f_postcode"
        SQL = SQL + " where 1=1 " + county1 + county2 + address
        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
  
        GridView1.DataSource = dtReferp
        GridView1.DataBind()
 
    End Sub


    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        GridView1.PageIndex = e.NewPageIndex
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

        Dim fpObj As New ForProject
        Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()



        Dim D1 As String = DAddress3.Text

        Dim D2 As String = GridView1.SelectedRow.Cells(1).Text

        Dim D3 As String = GridView1.SelectedRow.Cells(6).Text


        'Dim js As String = ""
        'js &= "var obj = window.opener.document.all.D1;"
        'js &= "'" & D2 & "';"
        ' ''js &= "obj.value = '" & D1 & "';"
        ' ''param &= "window.opener.document.getElementById('" & Request.QueryString("field1") & "').value = '" & Me.Calendar1.SelectedDate.ToString("yyyy/MM/dd") & "'; "

        ' ''js &= "var obj = window.opener.document.getElementById('" & Request.QueryString("field1") & "').value = "
        ' ''js &= "obj.value = '" & D2 & "';"
        'js &= "window.opener.document.forms[0].submit();"
        'js &= "window.close();"


        'Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)

        '
        '傳回父視窗
        Dim param As String = "JavaScript:"
        param &= "window.opener.document.getElementById('" & Request.QueryString("field") & "').value = '" & D1 & "'; "
        If Request.QueryString("field") = "DAddCH" Then
            param &= "window.opener.document.getElementById('DPostCode').value = '" & D2 & "'; "
            param &= "window.opener.document.getElementById('DTJCode').value = '" & D3 & "'; "
        End If

        'param &= "window.opener.HideGridview();"
        param &= "window.close(); "

        Dim lJScript As New System.Text.StringBuilder("")
        lJScript.Append(param)
        ClientScript.RegisterStartupScript(Me.GetType(), "ReturnValue", lJScript.ToString(), True)


    End Sub


    Protected Sub GridView1_PageIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PageIndexChanged
        GetData()
    End Sub

    Sub GetCountry1()
        Dim SQL As String
        Dim i As Integer
        SQL = "  select distinct county1 as Data from f_postcode "
        SQL = SQL & " order by county1 "
        Dim dtReferp As DataTable = uDataBase.GetDataTable(Sql)
        DCounty1.Items.Add("-請選擇-")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DCounty1.Items.Add(ListItem1)
        Next
        dtReferp.Clear()
    End Sub

    Protected Sub DCounty1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCounty1.SelectedIndexChanged
        Dim SQL As String
        Dim i As Integer
        SQL = "  select distinct county2 as Data from f_postcode "
        SQL = SQL + " where county1='" + DCounty1.SelectedValue + "'"
        SQL = SQL & " order by county2 "
        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DCounty2.Items.Clear()
        DCounty2.Items.Add("-請選擇-")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DCounty2.Items.Add(ListItem1)
        Next
        dtReferp.Clear()
        DAddress3.Text = DCounty1.SelectedValue
    End Sub

 
    Protected Sub DCounty2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DCounty2.SelectedIndexChanged
        Dim SQL As String
        Dim i As Integer
        SQL = "  select distinct Address as Data from f_postcode "
        SQL = SQL + " where county2='" + DCounty2.SelectedValue + "'"
        SQL = SQL & " order by Address "
        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DAddress2.Items.Clear()
        DAddress2.Items.Add("-請選擇-")

        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DAddress2.Items.Add(ListItem1)
        Next
        dtReferp.Clear()
        DAddress3.Text = DCounty1.SelectedValue + DCounty2.SelectedValue
    End Sub

    Sub CheckAddress()
        If DAddress1.Text <> "" Then
            Dim i As Integer
            For i = 0 To DAddress2.Items.Count - 1
             
                '方法二
                If DAddress2.Items(i).Text Like DAddress1.Text & "*" Then
                    DAddress2.SelectedIndex = i


                    Exit For
                End If
            Next
        End If
    End Sub

    Protected Sub DAddress1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAddress1.TextChanged
        CheckAddress()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        DAddress3.Text = DCounty1.SelectedValue + DCounty2.SelectedValue + DAddress2.SelectedValue
        GetData()
    End Sub

    Protected Sub DAddress2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DAddress2.SelectedIndexChanged
        DAddress3.Text = DCounty1.SelectedValue + DCounty2.SelectedValue + DAddress2.SelectedValue
    End Sub
End Class
