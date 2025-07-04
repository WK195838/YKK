Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO
Imports System
Imports System.Web.UI
Imports System.Text
Imports System.Web.Configuration
Imports System.Data.Common
Imports System.Web.Security
Imports System.Web.UI.HtmlControls
Imports ThoughtWorks.QRCode.Codec
Imports ThoughtWorks.QRCode.Codec.Data


Partial Class StockInSheet_05
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim FieldName(60) As String     '各欄位
    Dim Attribute(60) As Integer    '各欄位屬性    
    Dim Top As String               '預設的控制項位置
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wStep As Integer            '工程代碼
    Dim wbFormSno As Integer        '連續起單-表單流水號
    Dim wbStep As Integer           '連續起單-工程代碼
    Dim wSeqNo As Integer           '序號
    Dim wApplyID As String          '申請者ID
    Dim wAgentID As String          '被代理人ID
    Dim NowDateTime As String       '現在日期時間
    Dim wNextGate As String         '下一關人
    Dim wOKSts As Integer = 1
    Dim wNGSts1 As Integer = 1
    Dim wNGSts2 As Integer = 1
    '群組行事曆
    Dim wApplyCalendar As String = ""       '申請者
    Dim wDecideCalendar As String = ""      '核定者
    Dim wNextGateCalendar As String = ""    '下一核定者
    Dim wUserName As String = ""            '姓名代理用
    Dim YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    Dim OpenDir2 As String = ""
    Dim AttachFileName As String
    Dim wStockNo As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        wStockNo = Request.QueryString("pStockNo")  '表單流水號
        ' wFormSno = 1
        getdate()

    End Sub


    Sub getdate()
        Dim SQL As String
        SQL = " Select a.date,a.no,name,depname,a.type,b.stockno From F_StockInSheet a,F_StockInSheetdt b"
        SQL = SQL + " where a.no = b.no "
        SQL = SQL + " and  b.StockNo = '" + wStockNo + "'"
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then
            DNo.Text = dtData.Rows(0).Item("No")
            DDate.Text = dtData.Rows(0).Item("Date")
            DDepName.Text = dtData.Rows(0).Item("DepName")
            DName.Text = dtData.Rows(0).Item("Name")
            DType.Text = dtData.Rows(0).Item("Type")
            TextBox_39.Text = dtData.Rows(0).Item("StockNo")
            Image1.ImageUrl = "http://10.245.1.10/BarCode/code128.aspx?num=" + TextBox_39.Text
        End If

        Dim DBDataSet1 As New DataSet
        Dim OleDbConnection1 As New OleDbConnection


    End Sub


End Class
