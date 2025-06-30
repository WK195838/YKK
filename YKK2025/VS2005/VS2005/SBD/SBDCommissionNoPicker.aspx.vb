Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Collections.Generic

Partial Class SBDCommissionNoPicker
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim oFlow As New Flow.FlowService
    Dim oSchedule As New Schedule.ScheduleService
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        '
        If Not Me.IsPostBack Then   '不是PostBack
            DKey.Text = ""
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '-----------------------------------------------------------------
        '-- 系統參數
        '-----------------------------------------------------------------
        Server.ScriptTimeout = 900                                                                  '設定逾時時間
        Response.Cookies("PGM").Value = "SBDCommissionNoPicker.aspx"                                         '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")                                           '現在日時
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BKey_Click)
    '**     搜尋按鈕按下時
    '**
    '*****************************************************************
    Protected Sub BKey_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BKey.Click
        If DKey.Text <> "" Then
            DataList()
        Else
            uJavaScript.PopMsg(Me, "無輸入資料無法搜尋,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(DataList)
    '**     開發資料列表
    '**
    '*****************************************************************
    Protected Sub DataList()
        Dim SQL As String
        SQL = "select  No," & " FormNo +  '-' + RTrim(LTrim(str(formsno))) as NoDesc,"
        SQL &= "'SBDCommissionSheet_02.aspx?pFormNo=' + FormNo + '&pFormSno=' + str(FormSno, Len(FormSno)) As URL "
        SQL &= " from  F_SBDCommissionSheet "
        SQL &= " Where No Like '%" & DKey.Text & "%' "
        SQL &= " Order by APPDate Desc"
        Dim dtCommissionSheet As DataTable = uDataBase.GetDataTable(SQL)
        If dtCommissionSheet.Rows.Count > 0 Then
            GridView1.DataSource = dtCommissionSheet
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到開發資料,請確認!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GridView1_SelectedIndexChanged)
    '**     取得所選擇開發資料
    '**
    '*****************************************************************
    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim xNo As String = GridView1.SelectedRow.Cells(1).Text
        Dim SQL, Cmd As String
        '
        SQL = "Select * From   F_SBDCommissionSheet "
        SQL &= "Where No = '" & xNo & "' "
        Dim dtData As DataTable = uDataBase.GetDataTable(SQL)
        If dtData.Rows.Count > 0 Then


            '
            '傳回父視窗三個值 , 日期 , 日期別 , 薪資所屬年月
            Dim param As String = "JavaScript:"
            param &= "window.opener.document.getElementById('DNo').value = '" & dtData.Rows(0).Item("No") & "'; "
            param &= "window.opener.document.getElementById('DBuyer').value = '" & dtData.Rows(0).Item("Buyer") & "'; "
            param &= "window.opener.document.getElementById('DVendor').value = '" & dtData.Rows(0).Item("Vendor") & "'; "
            param &= "window.opener.document.getElementById('DSupplier').value = '" & dtData.Rows(0).Item("Supplier") & "'; "
            'param &= "window.opener.HideGridview();"
            param &= "window.close(); "

            Dim lJScript As New System.Text.StringBuilder("")
            lJScript.Append(param)
            ClientScript.RegisterStartupScript(Me.GetType(), "ReturnValue", lJScript.ToString(), True)




        End If
    End Sub
End Class
