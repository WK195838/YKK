Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class AdvencedImagesData
    Inherits System.Web.UI.Page

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
    Dim oWaves As New WAVES.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pFormNo As String           'FormNo
    Dim pFormSno As Integer         'FormSno
    Dim pCode As String             'Code
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
        Response.Cookies("PGM").Value = "AdvencedImagesData.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pFormNo = Request.QueryString("pFormNo")            'FormNo
        pFormSno = Request.QueryString("pFormSno")          'FormSno
        pCode = Request.QueryString("pCode")                'Code
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
        Dim sql As String
        sql = "select "
        sql = sql & "ImagePath "
        sql = sql & "from V_RDPullerImage "
        sql = sql & "where formno = '" & pFormNo & "' "
        sql = sql & "and   formsno = " & pFormSno & " "
        Dim dt_Map As DataTable = uDataBase.GetDataTable(sql)
        If dt_Map.Rows.Count > 0 Then
            If pFormNo = "000003" Then
                DLabel1.Text = pCode & "  " & "內製"
            Else
                DLabel1.Text = pCode & "  " & "外注"
            End If
            DImages.ImageUrl = dt_Map.Rows(0).Item("ImagePath").ToString
        End If
        '
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        '
        SQL = "select "
        SQL = SQL & "case when MMSSts='1' then '模廢' else '' end As MMS, "
        '
        SQL = SQL & "No As No, "
        SQL = SQL & "case when formno='000003' then 'http://10.245.1.10/WorkFlow/ManufInSheet_02.aspx?pFormNo=000003&pFormSno=' + ltrim(rtrim(str(formsno))) "
        SQL = SQL & "     else 'http://10.245.1.10/WorkFlow/ManufOutSheet_02.aspx?pFormNo=000007&pFormSno=' + ltrim(rtrim(str(formsno))) "
        SQL = SQL & "end as NoUrl, "
        '
        SQL = SQL & "Sts As Status, "
        '
        SQL = SQL & "Spec + Spec1 As Spec, "
        '
        SQL = SQL & "Remark As Remark "
        '
        SQL = SQL & "from V_RDPullerData "
        SQL = SQL & "where SliderGRCode = '" & pCode & "' "
        '
        SQL = SQL & "Order by SliderGRCode, FormNo, FormSno, No, Sts, Spec, Spec1, Remark "
        '
        '
        GridView1.Visible = True
        GridView1.DataSource = uDataBase.GetDataTable(SQL)
        GridView1.DataBind()
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        '
        If (e.Row.RowType = DataControlRowType.Header) Then

            Dim gv As GridView = sender
            Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            ' 清除
            e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            ' detail
            '添加新的表头第一行
            tcl.Add(New TableHeaderCell())
            tcl(0).Text = "" & "<BR>" & ""
            tcl(0).BackColor = Color.White

            tcl.Add(New TableHeaderCell())
            tcl(1).Text = "No." & "<BR>" & ""
            tcl(1).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(2).Text = "Status" & "<BR>" & ""
            tcl(2).BackColor = Color.Green
            tcl.Add(New TableHeaderCell())
            tcl(3).Text = "Size-Family-Body" & "<BR>" & ""
            tcl(3).BackColor = Color.Blue
            tcl.Add(New TableHeaderCell())
            tcl(4).Text = "R&D Remark" & "<BR>" & ""
            tcl(4).BackColor = Color.Red
'
            gv.Controls(0).Controls.AddAt(0, H3row)
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 顏色+格式
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If e.Row.Cells(0).Text = "模廢" Then
                e.Row.Cells(0).BackColor = Color.Black
                e.Row.Cells(0).ForeColor = Color.White
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If
    End Sub

End Class
