Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_SystemConfig
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim GroupCode As String         'GroupCode

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then
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
        Server.ScriptTimeout = 900                                  '設定逾時時間
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        BCat.Attributes.Add("onclick", "window.open('categoripicker.aspx?pUserID=" & Request.QueryString("pUserID") & "&pControlID=DCat,HiddenField1','參數','scrollbars=no,status=no,width=400,height=600,top=0,left=0');")
        DGridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BSearch_Click)
    '**     查詢資料
    '**
    '*****************************************************************
    Protected Sub BSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BSearch.Click
        Dim Sql As String = ""
        If DCat.Text <> "" Then
            Sql &= "SELECT *, "
            Sql &= "(SELECT Data From M_Referp Where Cat='999' and DKey= '" & DCat.Text & "') As CatDescr "
            Sql &= "FROM M_Referp "
            Sql &= "Where Cat = '" & DCat.Text & "' "
            If DDKey.Text <> "" Then
                Sql &= "And DKey Like '%" & DDKey.Text & "%' "
            End If
            Sql &= "order by Cat, DKey, Data "

            Dim dt_Referp As DataTable = uDataBase.GetDataTable(Sql)
            If dt_Referp.Rows.Count > 0 Then
                DGridView1.Visible = True
                DGridView1.DataSource = dt_Referp
                DGridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
            End If
        End If
    End Sub

End Class
