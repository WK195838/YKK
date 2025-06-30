Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb


Partial Class CopyEdxFile
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
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SMonth As String = ""
        Dim SYear As String = ""
        Dim SQL As String

        Dim i As Integer
        Dim File(6) As String

        '找檔案路徑
        SQL = " select substring(data,3,len(data)-1)Data from  M_referp"
        SQL = SQL + " where cat = 8002  "
        SQL = SQL + " and dkey = 'file' and left(data,1) in ('3','6','7') "

        Dim DBAdapter3 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter3.Rows.Count > 0 Then

            For i = 0 To DBAdapter3.Rows.Count - 1

                File(i + 1) = DBAdapter3.Rows(i).Item("Data")

            Next


        End If



        'copy edx 檢測報告
        Dim Source As String = File(1)  '來源
        Dim Dest As String = "\\10.245.1.6\wfs$\ISOSQC\008002\EDX"  '目的

        SMonth = Now.ToString("yyyyMM") '
        My.Computer.FileSystem.CopyDirectory(Source + "\" + SMonth, Dest, True)


        'copy edx 外測報告
        Source = File(2) '來源
        Dest = "\\10.245.1.6\wfs$\ISOSQC\008002\OutTest"  '目的

        SYear = Now.ToString("yyyy") '
        My.Computer.FileSystem.CopyDirectory(Source + "\" + SYear, Dest, True)


        'copy 色粉配方
        Source = File(3) '來源
        Dest = "\\10.245.1.6\wfs$\ISOSQC\008002\Formula"  '目的

        SYear = Now.ToString("yyyy") '
        My.Computer.FileSystem.CopyDirectory(Source, Dest, True)





        ' 關閉網頁
        Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
        'IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")
    End Sub
End Class
