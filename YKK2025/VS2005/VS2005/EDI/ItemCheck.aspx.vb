Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Threading

Partial Class ItemCheck
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '
    Dim uFASMapping As New EDI2011.FMappingService
    Dim uFASCommon As New EDI2011.FCommonService
    Dim uWFSCommon As New WFS.CommonService
    '
    Dim NowDateTime As String       '現在日期時間
    Dim xUserID, xBuyer As String
    '
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            DataList()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        uFASMapping.Timeout = Timeout.Infinite
        uFASCommon.Timeout = Timeout.Infinite
        uWFSCommon.Timeout = Timeout.Infinite
        '
        NowDateTime = CStr(Now.Date) + " " + _
                      CStr(Now.Hour) + ":" + _
                      CStr(Now.Minute) + ":" + _
                      CStr(Now.Second)     '現在日時
        '
        xBuyer = Request.QueryString("pBuyer")
        xUserID = Request.QueryString("pUserID")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     資料抽出
    '**
    '*****************************************************************
    Sub DataList()
        Dim i As Integer
        Dim Sql, errItem As String
        '
        errItem = ""
        Sql = "Select AN1 From E_InputSheet "
        Sql &= "Where Buyer = '" & xBuyer & "' "
        Sql &= "Group By AN1 "
        Sql &= "Order by AN1 "
        Dim dt_InputSheet As DataTable = uDataBase.GetDataTable(Sql)
        '
        For i = 0 To dt_InputSheet.Rows.Count - 1
            If uFASCommon.CheckItemNoDisplay(dt_InputSheet.Rows(i).Item("AN1")) <> 0 Then
                If errItem = "" Then
                    errItem = "'" + dt_InputSheet.Rows(i).Item("AN1") + "'"
                Else
                    errItem = errItem + "," + "'" + dt_InputSheet.Rows(i).Item("AN1") + "'"
                End If
            End If
        Next
        '
        If errItem <> "" Then
            Sql = "Select *, '異常' AS STS From E_InputSheet "
            Sql &= "Where Buyer = '" & xBuyer & "' "
            Sql &= "  And AN1 IN (" & errItem & ") "
            Sql &= "Order by AN1, Unique_ID "
            Dim dt_ErrData As DataTable = uDataBase.GetDataTable(Sql)
            '
            GridView1.DataSource = dt_ErrData
            GridView1.DataBind()
            '
            uJavaScript.PopMsg(Me, "發現有ITEM異常！")
        Else
            uJavaScript.PopMsg(Me, "無ITEM異常！")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     顏色處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim xBackColor As Boolean = False
            '--異常--------------------------------------------
            If e.Row.Cells(1).Text.ToString = "異常" Then
                e.Row.ForeColor = Color.Blue
                e.Row.BackColor = Color.LightPink
            End If
        End If
    End Sub

End Class
