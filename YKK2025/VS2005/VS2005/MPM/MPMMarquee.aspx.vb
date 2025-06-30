Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class MPMMarquee
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    ''' <summary>
    ''' 以下為共用函式的宣告
    ''' </summary>
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
  
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Dim NowDateTime As String       '現在日期時間
    Dim wFun As String              '新增(ADD)/修改(MOD)
    Dim wUserID As String           'User ID
    Dim wID As String               'Unique_ID
    Dim wEMPID As String            'EMP-ID
    Dim wDivision As String         '部門
    Dim wYear As String             '年
    Dim wMonth As String            '月
    Dim wName As String             '姓名
    Dim wType As String             '類別
    Dim wSystem As String           '系統

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetCommonParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            DisplayFieldFormat()      '設定表單欄位屬性
            Dim ff As FontFamily
            '載入系統安裝的字形到ComboBox1   
            DFontStyle.Items.Clear()
            DFontStyle.Text = "標楷體"
            For Each ff In System.Drawing.FontFamily.Families
                DFontStyle.Items.Add(ff.Name)
            Next
            If wFun = "MOD" Or wFun = "COPY" Then
                ShowData()          '顯示資料(修改日誌)
            Else
                SetFormatData()     '設定初值(新日誌)
               
            End If
          

        End If


      
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetCommonParameter()
        '基本資料設定
        Response.Cookies("PGM").Value = "MPMMarquee.aspx"
        NowDateTime = CStr(DateTime.Now.Date) + " " + _
                      CStr(DateTime.Now.Hour) + ":" + _
                      CStr(DateTime.Now.Minute) + ":" + _
                      CStr(DateTime.Now.Second)       '現在日時
        '附帶參數
        wFun = Request.QueryString("pFun")            '新增/變更
        wUserID = Request.QueryString("pUserID")      'User ID
        wID = Request.QueryString("pID")              'Unique_ID
        wYear = Request.QueryString("pYear")          '年
        wMonth = Request.QueryString("pMonth")        '月
        wType = Request.QueryString("pType")          '類別
        wSystem = Request.QueryString("pSystem")      '系統
        '取得個人資訊
       
        Dim SQL As String
       
        SQL = "Select UserName, DivName, EmpID From M_Users Where UserID = '" + wUserID + "' "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

    
        If DBAdapter1.Rows.Count > 0 Then
            wDivision = DBAdapter1.Rows(0).Item("DivName")
            wName = DBAdapter1.Rows(0).Item("UserName")
            wEMPID = DBAdapter1.Rows(0).Item("EMPID")
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示資料(修改行程)
    '**
    '*****************************************************************
    Sub ShowData()
      
        Dim SQL As String
     
        SQL = "Select * From F_Marquee Where Unique_ID = '" + wID + "'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        
        If DBAdapter1.Rows.Count > 0 Then
            '行程日期

            '啟用/關閉
            Dim ListItem1 As New ListItem
            If DBAdapter1.Rows(0).Item("Action") = 0 Then
                ListItem1.Text = "關閉"
            Else
                ListItem1.Text = "啟用"
            End If
            ListItem1.Value = DBAdapter1.Rows(0).Item("Action").ToString
            DAction.SelectedIndex = DAction.Items.IndexOf(ListItem1)
            DAppUser.Text = DBAdapter1.Rows(0).Item("AppUser").ToString
            DAppDate.Text = DBAdapter1.Rows(0).Item("AppDate")
            DSubject.Text = DBAdapter1.Rows(0).Item("Subject").ToString
            DFontSize.Text = DBAdapter1.Rows(0).Item("FontSize").ToString
            DFontStyle.Text = DBAdapter1.Rows(0).Item("FontStyle").ToString
            DColor.Text = DBAdapter1.Rows(0).Item("Color").ToString
            DFontColor.Style.Add("color", DColor.Text)
            If DStartDate.Value = "1900/01/01" Then
                DStartDate.Value = ""
            Else
                DStartDate.Value = DBAdapter1.Rows(0).Item("StartDate")
            End If

            If DStartDate.Value = "1900/1/1" Then
                DStartDate.Value = ""
            Else
                DStartDate.Value = DBAdapter1.Rows(0).Item("StartDate")
            End If

            If DEndDate.Value = "1900/1/1" Then
                DEndDate.Value = ""
            Else
                DEndDate.Value = DBAdapter1.Rows(0).Item("EndDate")
            End If

 

        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定初值(新日誌)
    '**
    '*****************************************************************
    Sub SetFormatData()
        DAppUser.Text = wName
        DAppDate.Text = DateTime.Now.Date
        DFontSize.Text = "7"
        DColor.Text = "#FF0000"
        DFontColor.Style.Add("color", "red")
 

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     儲存
    '**
    '*****************************************************************
    Protected Sub BSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSave.Click
        Dim SQL As String
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As New DataTable
        Dim OleDbConnection1 As New OleDbConnection
        Dim OleDBCommand1 As New OleDbCommand
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        '新增 or 修改
        If wFun = "ADD" Or wFun = "COPY" Then
            '***新增作業
        
            '***新增作業
            SQL = "Insert into F_Marquee"
            SQL = SQL + "( "
            SQL = SQL + "Action, Dep_Name, Subject, StartDate, EndDate, "        '1~5
            SQL = SQL + "FontSize,FontStyle, Color, AppDate,AppUser, "           '6~10
            SQL = SQL + "CreateUser, CreateTime, ModifyUser, ModifyTime "      '11~15
            SQL = SQL + ")  "
            SQL = SQL + "VALUES( "
            SQL = SQL + " '" + DAction.SelectedValue + "', "               '啟用
            SQL = SQL + " '" + DDep_Name.SelectedValue + "', "            '部門
            SQL = SQL + " N'" + DSubject.Text + "', "          '內容

            If DStartDate.Value = "" Then '開始日期
                SQL = SQL + " '', "
            Else
                SQL = SQL + " '" + DStartDate.Value + "', "
            End If

            If DEndDate.Value = "" Then '開始日期
                SQL = SQL + " '', "
            Else
                SQL = SQL + " '" + DEndDate.Value + "', "
            End If

            SQL = SQL + " '" + DFontSize.Text + "', "                      '字體大小
            SQL = SQL + " '" + DFontStyle.Text + "', "                      '字體大小
            SQL = SQL + " '" + DColor.Text + "', "              '字體顏色



            SQL = SQL + " '" + DAppDate.Text + "', "              '建檔日

            SQL = SQL + " '" + DAppUser.Text + "', "              '建檔日
            SQL = SQL + " '" + wUserID + "', "                  '作成者ID
            SQL = SQL + " '" + NowDateTime + "', "              '作成時間
            SQL = SQL + " '" + "" + "', "                       '修改者
            SQL = SQL + " '" + NowDateTime + "' "               '修改時間
            SQL = SQL + " ) "
           uDataBase.ExecuteNonQuery(SQl)
        Else
            '***修改作業

            SQL = "Update  F_Marquee  Set "
            SQL = SQL + " Action = '" + DAction.SelectedValue + "',"
            SQL = SQL + "Dep_Name = '" + DDep_Name.SelectedValue + "',"
            SQL = SQL + " Subject = '" + Trim(DSubject.Text) + "',"

            If DStartDate.Value = "" Then '開始日期
                SQL = SQL + "StartDate = ' ',"
            Else
                SQL = SQL + "StartDate = '" + DStartDate.Value + "',"
            End If

            If DEndDate.Value = "" Then '結束日期
                SQL = SQL + " EndDate = ' ',"
            Else
                SQL = SQL + " EndDate = '" + DEndDate.Value + "',"
            End If

            SQL = SQL + " FontSize = '" + DFontSize.Text + "',"
            SQL = SQL + " FontStyle = '" + DFontStyle.Text + "',"
            SQL = SQL + " Color = '" + DColor.Text + "',"
            SQL = SQL + " ModifyUser = '" + wUserID + "', "
            SQL = SQL + " ModifyTime = '" + NowDateTime + "' "
            SQL = SQL + " Where Unique_ID =  '" + wID + "'"
            uDataBase.ExecuteNonQuery(SQL)
        End If

        Dim URL As String
        URL = "MPMMarqueeList.aspx?&pUserID=" & wUserID & _
                                "&pYear=" & wYear & _
                                "&pMonth=" & wMonth & _
                                "&pType=" & wType & _
                                "&pSystem=" & wSystem
        Response.Redirect(URL)

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     回行程首頁
    '**
    '*****************************************************************
    Protected Sub BBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BBack.Click
        Dim URL As String
        URL = "MPMMarqueeList.aspx?&pUserID=" & wUserID & _
                                "&pYear=" & wYear & _
                                "&pMonth=" & wMonth & _
                                "&pType=" & wType & _
                                "&pSystem=" & wSystem
        Response.Redirect(URL)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     顯示欄位初值處理
    '**
    '*****************************************************************
    Sub DisplayFieldFormat()
        BStartDate.Attributes.Add("onClick", "calendarPicker('DStartDate')")
        BEndDate.Attributes.Add("onClick", "calendarPicker('DEndDate')")
    End Sub
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

    Sub Marquee()


        ' obj.Text = "<MARQUEE onmouseover=stop(); onmouseout=start();  text-align:left scrollAmount=3 scrollDeay=100  height=20 width=900 scrolltop=0 scrollleft=0>  "
        obj.Text = " <marquee bgcolor='#000000' border='0' align='left' scrollamount='5' direction='left' scrolldelay='90'   width='900' height='50' style='font-family:Arial;color:#9370DB;font-size:18;text-align:center'> "
        'obj.Text += "<table width='100%' bordr='1' cellspacing='6' cellpaddig='1'>"
        '    obj.Text = " <marquee direction=up scrollamount=1 scrolldelay=100 onmouseover='this.stop()' onmouseout='this.start()' direction=up  height=400 width=230 >"
        ' obj.Text = "<marquee  onmouseover=stop(); onmouseout=start(); direction=left scrollDelay=50 scrollAmount=2>"

        obj.Text += "<tr><td width='900'  height=50  align='left'>"
        '     obj.Text += "<img src='/MinsuHome/images/a00_new-top06_d.gif'  border='0'>"
        '      obj.Text += "</td><td style='FONT-SIZE: 11px'>"
        '  obj.Text += "<a href='/MinsuHome/Activity/ViewActivity.aspx?ID=" & dtr1("formsno").ToString & "'>"
        obj.Text += " <font Size=" + DFontSize.Text + "' color='" + DColor.Text + "' face ='" + DFontStyle.Text + "'>" & DSubject.Text & "</font>"
        obj.Text += "</a>"
        obj.Text += "</td></tr>"



        obj.Text += "</table>"
        obj.Text += "</marquee>"
    End Sub


    Protected Sub BShow_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BShow.Click
        Marquee()
        DFontColor.Style.Add("color", DColor.Text)
    End Sub
 
End Class
