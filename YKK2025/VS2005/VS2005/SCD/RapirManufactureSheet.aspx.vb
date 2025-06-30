Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class RapirManufactureSheet
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '群組行事曆
    Dim wCalendar As String = ""
    Dim NowDateTime As String       '現在日期時間
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService

    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim SQL As String
        Dim i As Integer
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
        '-----------------------------------------------------------------
        '-- op3
        '-----------------------------------------------------------------
        SQL = "Select "
        SQL &= "case "
        SQL &= "  when substring(op3btime,1,2)>'05' then  '2011/'+substring(op3btime,1,11)+':00' "
        SQL &= "  else '2012/'+substring(op3btime,1,11)+':00' "
        SQL &= "end as BStartTime, "
        SQL &= "case "
        SQL &= "  when substring(op3atime,13,2)>'05' then  '2011/'+substring(op3atime,13,11)+':00' "
        SQL &= "  else '2012/'+substring(op3atime,13,11)+':00' "
        SQL &= "end as AEndTime, * "
        SQL &= "from FS_ManufactureSheet "
        SQL &= "Where substring(op3btime,1,11) <> substring(op3atime,1,11) "
        SQL &= "  And op3atime<>'' "
        Dim dtTable1 As DataTable = uDataBase.GetDataTable(SQL)
        For i = 0 To dtTable1.Rows.Count - 1
            Dim wDevelopTime As Integer = 0
            '
            oSchedule.GetDevelopTime(dtTable1.Rows(i).Item("BStartTime").ToString, dtTable1.Rows(i).Item("AEndTime").ToString, wDevelopTime, "CL2")
            '
            If wDevelopTime < 0 Then wDevelopTime = 0
            '
            SQL = "Update FS_ManufactureSheet Set "
            SQL = SQL + " op3atime = '" & Mid(dtTable1.Rows(i).Item("op3btime").ToString, 1, 11) + Mid(dtTable1.Rows(i).Item("op3atime").ToString, 12) & "', "
            SQL = SQL + " op3AHOURS = '" & CStr(wDevelopTime) & "', "
            SQL = SQL + " ModifyUser = '" & "Joy" & "',"
            SQL = SQL + " ModifyTime = '" & NowDateTime & "' "
            SQL = SQL + " Where Unique_ID = '" & dtTable1.Rows(i).Item("Unique_ID").ToString & "'"
            uDataBase.ExecuteNonQuery(SQL)
            '
        Next











    End Sub
End Class
