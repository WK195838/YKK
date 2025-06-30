Imports System.Data
Imports System.Drawing

Partial Class DevelopmentDelivery_Load
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
    Dim wFormNo As String           '表單號碼
    Dim wFormSno As Integer         '表單流水號
    Dim wWID As String              '工作ID
    '
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim oCommon As New Common.CommonService
    Dim oSchedule As New Schedule.ScheduleService
    '
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cookies("PGM").Value = "DevelopmentDelivery_Load.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            '負荷   待處理或預定(Active=1 or Active=9) 
            ActiveDataList()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        wFormNo = Request.QueryString("pFormNo")    '表單號碼
        wFormSno = Request.QueryString("pFormSno")  '表單流水號
        wWID = Request.QueryString("pWID")          'Work ID
        Dim wStartTime As String = Request.QueryString("pStartTime") + " 00:00:00"
        If DBStartTime.Text = "" Or DBEndTime.Text = "" Then
            DBStartTime.Text = DateAdd(DateInterval.Day, -365, CDate(wStartTime)).ToString("yyyy/MM/dd")
            DBEndTime.Text = DateAdd(DateInterval.Day, 365, CDate(wStartTime)).ToString("yyyy/MM/dd")
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     待處理或預定-篩選資料
    '**
    '*****************************************************************
    Sub ActiveDataList()
        Dim SQL As String
        Dim wBStartTime As String = DBStartTime.Text + " 00:00:00"
        Dim wBEndTime As String = DBEndTime.Text + " 23:59:59"
        '
        SQL = "Select "
        SQL &= "'' As NO, "
        SQL &= "'' As DEVNO, "
        SQL &= "'' As CODE, "
        SQL &= "'' As BUYER, "
        SQL &= "'' As SIZE, "
        SQL &= "'' As ITEM, "
        SQL &= "'' As BTIME, "
        SQL &= "'' As ATIME, "
        SQL &= "'' As OPPER, "
        SQL &= "FormNo, FormSno, BStartTime, BEndTime, AStartTime, AEndTime, DecideID, "
        SQL &= "Case Active When 9 Then 'P' When 1 Then 'W' Else '' End As STSD "
        SQL &= "From T_WaitHandle "
        SQL &= "Where FormNo = '" + wFormNo + "' "
        SQL &= "  and WorkID = '" + wWID + "'"
        SQL &= "  and BStartTime >= '" + wBStartTime + "'"
        SQL &= "  and BStartTime <= '" + wBEndTime + "'"
        SQL &= "  and (Active = '1' or Active = '9') "
        SQL &= "  and (FlowType = '1' or FlowType = '9') "
        SQL &= "Order by BStartTime, AStartTime "
        Dim dt_WaitHandle As DataTable = uDataBase.GetDataTable(SQL)
        '設定各欄位資料
        If dt_WaitHandle.Rows.Count > 0 Then
            '
            For i As Integer = 0 To dt_WaitHandle.Rows.Count - 1
                '
                SQL = "Select * From F_CommissionSheet "
                SQL &= "Where FormNo  = '" + dt_WaitHandle.Rows(i).Item("FormNo") + "' "
                SQL &= "  and FormSno = '" + CStr(dt_WaitHandle.Rows(i).Item("FormSno")) + "'"
                Dim dt_Commission As DataTable = uDataBase.GetDataTable(SQL)
                If dt_Commission.Rows.Count > 0 Then
                    '
                    SQL = "Select * From FS_ManufactureSheet "
                    SQL &= "Where No = '" + dt_Commission.Rows(0).Item("No") + "' "
                    Dim dt_Manufacture As DataTable = uDataBase.GetDataTable(SQL)
                    If dt_Manufacture.Rows.Count > 0 Then
                        '
                        ' NO(0)
                        dt_WaitHandle.Rows(i)(0) = dt_Commission.Rows(0).Item("No")
                        ' DEVNO(1)
                        dt_WaitHandle.Rows(i)(1) = dt_Manufacture.Rows(0).Item("DevNo")
                        ' CODE(2)
                        dt_WaitHandle.Rows(i)(2) = dt_Manufacture.Rows(0).Item("CodeNo")
                        ' BUYER(3)
                        dt_WaitHandle.Rows(i)(3) = dt_Commission.Rows(0).Item("APPBuyer")
                        ' SIZE(4)
                        dt_WaitHandle.Rows(i)(4) = dt_Commission.Rows(0).Item("SizeNo")
                        ' ITEM(5)
                        dt_WaitHandle.Rows(i)(5) = dt_Commission.Rows(0).Item("Item")
                        ' BTIME(6)
                        Dim xTime As Integer = 0
                        oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i).Item("BStartTime")).ToString("yyyy/MM/dd HH:mm"), _
                                                 CDate(dt_WaitHandle.Rows(i).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm"), _
                                                 xTime, _
                                                 oCommon.GetCalendarGroup(wWID))
                        dt_WaitHandle.Rows(i)(6) = CDate(dt_WaitHandle.Rows(i).Item("BStartTime")).ToString("yyyy/MM/dd HH:mm") + _
                                                   " ~ " + _
                                                   CDate(dt_WaitHandle.Rows(i).Item("BEndTime")).ToString("yyyy/MM/dd HH:mm") + _
                                                   "  (" + CStr(xTime) + ")"
                        ' ATIME(7)
                        If dt_WaitHandle.Rows(i).Item("STSD") = "" Then
                            xTime = 0
                            oSchedule.GetDevelopTime(CDate(dt_WaitHandle.Rows(i).Item("AStartTime")).ToString("yyyy/MM/dd HH:mm"), _
                                                     CDate(dt_WaitHandle.Rows(i).Item("AEndTime")).ToString("yyyy/MM/dd HH:mm"), _
                                                     xTime, _
                                                     oCommon.GetCalendarGroup(wWID))
                            dt_WaitHandle.Rows(i)(7) = CDate(dt_WaitHandle.Rows(i).Item("AStartTime")).ToString("yyyy/MM/dd HH:mm") + _
                                                       " ~ " + _
                                                       CDate(dt_WaitHandle.Rows(i).Item("AEndTime")).ToString("yyyy/MM/dd HH:mm") + _
                                                       "  (" + CStr(xTime) + ")"
                        End If
                        ' OPPER(8)
                        dt_WaitHandle.Rows(i)(8) = oCommon.GetUserName(dt_WaitHandle.Rows(i).Item("DecideID"))
                    End If
                End If
            Next
            '
        End If
        '
        'GridView1.Columns.Item(8).HeaderText = "預定"
        '
        GridView1.DataSource = dt_WaitHandle
        GridView1.DataBind()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Go Command
    '**
    '*****************************************************************
    Protected Sub BOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        '負荷   待處理或預定(Active=1 or Active=9) 
        ActiveDataList()
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_RowDataBound)
    '**     當筆處理
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            '
            If e.Row.Cells(1).Text.ToString <> "&nbsp;" Then
                Dim sql As String
                sql = "Select No From F_CommissionSheet "
                sql &= "Where FormNo  = '" + wFormNo + "' "
                sql &= "  and FormSno = '" + CStr(wFormSno) + "'"
                Dim dt_Commission As DataTable = uDataBase.GetDataTable(Sql)
                If dt_Commission.Rows.Count > 0 Then
                    If e.Row.Cells(2).Text.ToString = dt_Commission.Rows(0).Item("No") Then
                        e.Row.BackColor = Color.LightPink
                    End If
                End If
            End If
        End If
    End Sub

End Class
