Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class SP_ShoppingListInf
    Inherits System.Web.UI.Page
    '
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim SPName As String            'SP Name
    Dim SPNo As String              'SP NO
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            SetDefaultValue()                   '設定初值
            ShowData()
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
        Server.ScriptTimeout = 900                          '設定逾時時間
        Response.Cookies("PGM").Value = "SP_ShoppingListInf.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        SPName = Request.QueryString("pSPName")             'SP Name
        SPNo = Request.QueryString("pSPNo")                 'SP No
        '
        If SPName <> "" Then
            DKSPName.Text = SPName
        End If
        If SPNo <> "" Then
            DKOther.Text = SPNo
        End If
        '
        '動作按鈕設定
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '-----------------------------------------------------------------
        '-- Not IsPostBack
        '-----------------------------------------------------------------
        GridView1.Visible = False
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Find Puller
    '**
    '*****************************************************************
    Protected Sub BFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFind.Click
        ShowData()
    End Sub

    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "select top 150 "
        SQL = SQL & "[SPCode],[SPName],[SPTime],[SPNo],[Status], "
        SQL = SQL & "[LastAcessTime], [ChangeFinal], [UserList], "
        '
        SQL = SQL & "case when Status='Last Final' then '....' when Status='Pending Final' then '....' else '' end as Change, "
        '
        SQL = SQL & "case when Status='Last Final'    then 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=CHGFINALSPNO&pSPNo='+SPNo "
        SQL = SQL & "     when Status='Pending Final' then 'http://10.245.0.205/EDI/SPHttp2File.aspx?pFun=PDFINALSPNO&pSPNo='+SPNo "
        SQL = SQL & "     else '' "
        SQL = SQL & "end as ChangeUrl "
        '
        SQL = SQL & "from V_UserShoppingInf "
        SQL = SQL & "where SPCode <> '' "

        ' SPName
        If DKSPName.Text <> "" Then
            SQL = SQL & "and ( "
            SQL = SQL & "        SPCode = '" & DKSPName.Text & "' "
            SQL = SQL & "     or SPName = '" & DKSPName.Text & "' "
            SQL = SQL & ") "
        End If
        ' Other
        If DKOther.Text <> "" Then
            SQL = SQL & "and [SPCode]+[SPName]+[SPTime]+[SPNo]+[Status]+[LastAcessTime]+[ChangeFinal]+[UserList] like '%" & DKOther.Text & "%' "
        End If
        '
        SQL = SQL & "Order by LastAcessTime desc "
        '
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        '
        If (e.Row.RowType = DataControlRowType.Header) Then
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Name"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Shopping Time"
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP No."
            tcl(i).BackColor = Color.Purple
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Status"
            tcl(i).BackColor = Color.Green
            i = i + 1
            '
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Last Acess Time"
            tcl(i).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "ChangeFinal"
            tcl(i).BackColor = Color.Green
            i = i + 1
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Link"
            tcl(i).BackColor = Color.Black
            i = i + 1
            '
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            For i = 0 To 8
                Select Case i
                    Case 8
                        If InStr(e.Row.Cells(i).Text, UCase(UserID)) <= 0 Then
                            e.Row.Cells(i - 1).Text = ""
                        End If
                        e.Row.Cells(i).Visible = False
                        '
                        'Case 4 To 8
                        '    e.Row.Cells(i).Font.Bold = True
                        '    e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")
                        'Case 9 To 12
                        '    'e.Row.Cells(i).ForeColor = Color.Red
                        'Case 13 To 36
                        '    e.Row.Cells(i).Font.Bold = True
                        '    e.Row.Cells(i).Text = Format(CDbl(e.Row.Cells(i).Text), "###,###,##0.00")

                    Case Else
                End Select

            Next
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub
End Class
