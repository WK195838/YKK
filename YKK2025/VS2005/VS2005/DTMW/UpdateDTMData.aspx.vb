Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Image
Imports System


Partial Class UpdateDTMData
    Inherits System.Web.UI.Page
    Dim InputData As String
    ''' <remarks></remarks>
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim Color1, Color2 As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        UpdateVFColor()

        'IE8 可以自行關閉網頁
        Response.Write("<script   language=javascript>  window.opener=null;    window.open('','_self');  window.close();</script> ")
        'IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")



    End Sub

    Sub UpdateVFColor()
        Dim SQL As String
        Dim Table As String
        Dim i As Integer
        Dim wFormNo, wFormSno As String
        '檢查今天80工程-SLD核定後的單
        SQL = "select * from T_WaitHandle "
        SQL = SQL & " where formno between  '005001' and '005099'"
        SQL = SQL & " and formno <> '005007' and step =90  "
        SQL = SQL & " and active =1"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            For i = 0 To DBAdapter1.Rows.Count - 1
                '找到表單TABLE 
                wFormNo = CStr(DBAdapter1.Rows(i).Item("formno"))
                wFormSno = CStr(DBAdapter1.Rows(i).Item("formsno"))
                SQL = " select tableName1 from M_form "
                SQL = SQL + "  where formno ='" + wFormNo + "'"
                Dim DBAdapter2 As DataTable = uDataBase.GetDataTable(SQL)
                If DBAdapter2.Rows.Count > 0 Then
                    Table = "F_" + DBAdapter2.Rows(0).Item("TableName1")
                    ' 的四個BUYER決議：BUYER的新色依賴表單，『兼用色VF上/下止色號』 直接帶入『塑鋼牙齒色號
                    If wFormNo = "005011" Or wFormNo = "005014" Or wFormNo = "005015" Or wFormNo = "005016" Then
                        '更新VFCOLOR 
                        SQL = " UPDATE " + Table
                        SQL = SQL + " SET YKKColorCodeVF = YKKColorCode"
                        SQL = SQL + " where formsno = '" + wFormSno + "'"
                        SQL = SQL + " AND  YKKColorCode <>'' and YKKColorCodeVF='' "
                        uDataBase.ExecuteNonQuery(SQL)
                        '更新VFCOLOR 
                        SQL = " UPDATE " + Table
                        SQL = SQL + " SET VFCOLOR = SLDColor"
                        SQL = SQL + " where formsno = '" + wFormSno + "'"
                        SQL = SQL + " and SLDColor <>'' and VFCOLOR=''  "
                        uDataBase.ExecuteNonQuery(SQL)
                    Else
                        '更新VFCOLOR 
                        SQL = " UPDATE " + Table
                        SQL = SQL + " SET VFCOLOR = SLDCOLOR"
                        SQL = SQL + " where formsno = '" + wFormSno + "'"
                        SQL = SQL + " and SLDCOLOR <>''  and VFCOLOR ='' "
                        uDataBase.ExecuteNonQuery(SQL)
                    End If
                 
                End If
            Next

        End If
    End Sub



End Class
