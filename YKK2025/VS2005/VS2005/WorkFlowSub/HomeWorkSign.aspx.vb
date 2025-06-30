Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class HomeWorkSign
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wOPT As String
    Dim wUserID As String
    Dim wURL As String
    Dim wOPURL As String
    Dim wIP As String
    Dim NowDateTime As String       '現在日期時間
    '
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    '
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '
        '設定共用參數
        SetParameter()
        '
        If Not Me.IsPostBack Then
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        '
        wUserID = Request.QueryString("pUserID")
        '
        wOPT = Request.QueryString("pOPT")
        wUserID = Request.QueryString("pUserID")
        wURL = Request.QueryString("pURL")
        wOPURL = Request.QueryString("pOPURL")
        wIP = Request.ServerVariables("REMOTE_ADDR")
        '
    End Sub
    '*****************************************************************
    '**
    '**     DataList
    '**
    '*****************************************************************
    Sub DataList()
        Dim sql As String
        Dim xName, xEmpID, xDivision, xDivCode As String
        '
        xName = ""
        xEmpID = ""
        xDivision = ""
        xDivCode = ""
        '
        '取得DATA
        sql = "SELECT"
        sql = sql + " UserName, EmpID, DivName, DivID "
        sql = sql + " FROM M_Users "
        sql = sql + " Where UserID = '" & wUserID & "' "
        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            xName = dt.Rows(0)("UserName").ToString
            xEmpID = dt.Rows(0)("EmpID").ToString
            xDivision = dt.Rows(0)("DivName").ToString
            xDivCode = dt.Rows(0)("DivID").ToString
        End If
        '
        sql = "Insert into L_COVID19HomeWorkInf "
        sql &= "Select "
        sql &= "'" & wUserID & "', "
        sql &= "'" & xName & "', "
        sql &= "'" & xEmpID & "', "
        sql &= "'" & xDivision & "', "
        sql &= "'" & xDivCode & "', "
        sql &= "'" & NowDateTime & "', "
        sql &= "'" & wIP & "', "
        sql &= "'" & NowDateTime & "' "
        uDataBase.ExecuteNonQuery(sql)
        '
        '取得DATA
        sql = "SELECT"
        sql = sql + " UserName, [Division], ClickTime, IPAddress "
        sql = sql + " FROM L_COVID19HomeWorkInf "
        sql = sql + " Where UserID = '" & wUserID & "' "
        sql = sql + " And DATEDIFF ( day , ClickTime , getdate()) <= 1 "
        '
        sql = sql + " Order by  ClickTime desc "
        '
        Dim dtData As DataTable = uDataBase.GetDataTable(sql)
        GridView1.DataSource = dtData
        GridView1.DataBind()
    End Sub

End Class
