Imports System.Data
Imports System.Data.OleDb

Partial Class MPMMarqueeList
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    '*****************************************************************
    '**
    '**     全域變數ds
    '**
    '*****************************************************************
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService


    Dim NowDateTime As String    '現在日期時間
    Dim wName As String = ""
    Dim wType As String = ""     '類別
    Dim wSystem As String = ""   '系統
    Dim wUserID As String = ""   'User ID

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()
        If Not Me.IsPostBack Then   '不是PostBack
            SetComboboxList()
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)     '現在日時

    
        wType = Request.QueryString("pType")        '類別
        wSystem = Request.QueryString("pSystem")    'System
        wUserID = Request.QueryString("pUserID")    'User ID
    End Sub
    
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        SQL = "SELECT "
        SQL = SQL + "Case Action When 1 then '啟用' else '關閉' end as Action, "
        SQL = SQL + "Dep_Name,Subject,convert(char(10),startdate,111) Startdate,convert(char(10),enddate,111) EndDate,FontSize,Color, "
        SQL = SQL + "convert(char(10),Appdate,111) AppDate,AppUser,Unique_ID "
        SQL = SQL + "FROM F_Marquee Where 1=1 "
        '姓名
        If DName.SelectedValue <> "ALL" Then
            SQL = SQL + " And AppUser = '" + DName.SelectedValue + "'"
        End If
        SQL = SQL + "ORDER BY Unique_id "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_Marquee")
        DBTable1 = DBDataSet1.Tables("F_Marquee")
        GridView1.DataSource = DBDataSet1
        GridView1.DataBind()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     查詢(Go)
    '**
    '*****************************************************************
    Protected Sub BGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BGo.Click
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     新開發
    '**
    '*****************************************************************
    Protected Sub BNewItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BNewItem.Click
        Dim URL As String = "MPMMarquee.aspx?pFun=ADD" & "&pUserID=" & wUserID & _
                                                            "&pType=" & DType.SelectedValue & _
                                                            "&pSystem=" & DSystem.SelectedValue
        Response.Redirect(URL)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--欄位隱藏處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        e.Row.Cells(7).Visible = False   'ID
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--刪除
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        Dim SQL As String
        Dim Key As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Key = GridView1.DataKeys(e.RowIndex).Value

        OleDbConnection1.Open()
        SQL = "Select * From F_Marquee "
        SQL = SQL + "Where Unique_ID = '" + CStr(Key) + "' "
        SQL = SQL + "  And ( "
        SQL = SQL + "       CreateUser = '" + wUserID + "' or CreateUser = 'admin' "
        SQL = SQL + "      ) "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_Marquee")
        DBTable1 = DBDataSet1.Tables("F_Marquee")
        If DBTable1.Rows.Count > 0 Then
            SQL = "Delete From  F_Marquee Where Unique_ID = '" + CStr(Key) + "' "
            OleDBCommand1.Connection = OleDbConnection1
            OleDBCommand1.CommandText = SQL
            OleDBCommand1.ExecuteNonQuery()
        Else
            ShowMessage("無權利刪除資料")
        End If
        OleDbConnection1.Close()

        ShowData()              '顯示資料 
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--更新
    '**
    '*****************************************************************
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        Dim SQL, URL As String
        Dim Key As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable

        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定
        Dim OleDBCommand1 As New OleDbCommand

        Key = GridView1.DataKeys(e.NewEditIndex).Value

        OleDbConnection1.Open()
        SQL = "Select * From F_Marquee "
        SQL = SQL + "Where Unique_ID = '" + CStr(Key) + "' "
        SQL = SQL + "  And ( "
        SQL = SQL + "       CreateUser = '" + wUserID + "' or CreateUser = 'admin' "
        SQL = SQL + "      ) "
        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "F_Marquee")
        DBTable1 = DBDataSet1.Tables("F_Marquee")
        If DBTable1.Rows.Count > 0 Then

            URL = "MPMMarquee.aspx?pFun=MOD&pUserID=" & wUserID & "&pID=" & CStr(Key)
            Response.Redirect(URL)
        Else
            ShowMessage("無權利更新資料")
        End If

        OleDbConnection1.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView--刪除鈕按下時提示的確認訊息
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            '-----刪除鈕按下時提示的確認訊息。-------------------------------------
            Dim oButton As Button
            oButton = FindCommandButton(e.Row.Cells(1), "Delete")
            oButton.OnClientClick = "if (confirm('您確定要刪除嗎?')==false) {return false;}"
        End If
    End Sub

    Function FindCommandButton(ByVal Control As Control, ByVal CommandName As String) As Button
        Dim oChildCtrl As Control

        For Each oChildCtrl In Control.Controls
            If TypeOf (oChildCtrl) Is Button Then
                If String.Compare(CType(oChildCtrl, Button).CommandName, CommandName, True) = 0 Then
                    Return oChildCtrl
                End If
            End If
        Next

        Return Nothing
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     彈出警告訊息
    '**
    '*****************************************************************
    Sub ShowMessage(ByVal Msg As String)
        Dim myMsg As New Literal()
        myMsg.Text = "<script>alert('" & Msg & "')</script><br>"
        Me.Page.Controls.Add(myMsg)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     回工作週報
    '**
    '*****************************************************************
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim URL As String = "WeekReport.aspx?pUserID=" & wUserID
        Response.Redirect(URL)
    End Sub

    Sub SetComboboxList()
        '設定姓名
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim i As Integer

        SQL = "Select UserName, DivName From M_Users Where UserID = '" + wUserID + "' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            wName = DBAdapter1.Rows(0).Item("UserName")
        End If

        '設定姓名
        DBDataSet1.Clear()
        DName.Items.Clear()
        DName.Items.Add("ALL")
        SQL = "Select distinct AppUser From F_Marquee "
        Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
        
        For i = 0 To DBAdapter2.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter2.Rows(i).Item("AppUser")
            ListItem1.Value = DBAdapter2.Rows(i).Item("AppUser")
            If ListItem1.Value = wName Then ListItem1.Selected = True
            DName.Items.Add(ListItem1)
        Next
        wName = DName.SelectedValue


    End Sub
End Class
