Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Data.OleDb

Partial Class ITMNewCommission
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()        '設定共用參數
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
        Response.Cookies("PGM").Value = "NewCommission.aspx"
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")

        Dim SQL As String
        SQL = " select dkey,data  from M_referp "
        SQL = SQL + " where cat= '013'"
        SQL = SQL + " and dkey in ( 'DataChangeURL','VBRCSURL','SystemChangeURL')"

        Dim dtForm As DataTable = uDataBase.GetDataTable(SQL)
        If dtForm.Rows.Count > 0 Then
            For i As Integer = 0 To dtForm.Rows.Count - 1
                If dtForm.Rows(i).Item("Dkey") = "DataChangeURL" Then
                    HLink_01.NavigateUrl = dtForm.Rows(i).Item("Data")
                ElseIf dtForm.Rows(i).Item("Dkey") = "VBRCSURL" Then
                    HLink_02.NavigateUrl = dtForm.Rows(i).Item("Data")
                Else
                    HLink_03.NavigateUrl = dtForm.Rows(i).Item("Data")
                End If

            Next

        End If





    End Sub


End Class

