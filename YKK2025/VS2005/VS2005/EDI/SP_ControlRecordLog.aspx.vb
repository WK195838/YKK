Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SP_ControlRecordLog
    Inherits System.Web.UI.Page

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim fpObj As New ForProject             ' 操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim uFASMapping As New EDI2011.FMappingService
    Dim uFASCommon As New EDI2011.FCommonService
    Dim uWFSCommon As New WFS.CommonService
    Dim NowDateTime As String               ' 現在日時
    Dim xUserID As String                   ' 使用者ID
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        If Not IsPostBack Then
            ShowData()
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        '
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")
        xUserID = UCase(Request.QueryString("pUserID"))
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     ShowData
    '**
    '*****************************************************************
    Sub ShowData()
        '
        Dim sql As String
        '
        sql = "SELECT Top 100 "
        sql = sql + "[User],AccessTime,Cat,Active, Code, Name,Customer,Import,Demand,ActPlan,ImpActPlan,RspActPlan,KPInterface,RspWINGS,PILSheet,Final,ChgFinal,Progress "
        sql = sql + "FROM L_SPControlRecord "
        sql = sql + "Where [User] = '" & xUserID & "' "
        sql = sql + "Order by User, AccessTime desc, Cat "
        Dim dt As DataTable = uDataBase.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     GridView1 編輯資料 TITLE
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        Dim i As Integer
        ''
        If (e.Row.RowType = DataControlRowType.Header) Then

            'Dim gv As GridView = sender
            'Dim H3row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            'Dim H4row As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)
            '
            '' 清除
            'e.Row.Cells.Clear()
            '
            Dim tcl As TableCellCollection = e.Row.Cells
            tcl.Clear() '清除自动生成的表头
            '
            i = 0
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "User"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "AccessTime"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Cat"
            tcl(i).BackColor = Color.Black
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Active"
            tcl(i).BackColor = Color.Blue
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Code"
            tcl(i).BackColor = Color.Blue
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "SP Name"
            tcl(i).BackColor = Color.Blue
            i = i + 1
            ' -----
            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Customer"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Import"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Plan" & "<br>" & "Proc."
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Plan" & "<br>" & "Report"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Action"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "YOC" & "<br>" & "Confirm"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "KP"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "WINGS"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "PIL"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Final"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Change" & "<br>" & "Final"
            tcl(i).BackColor = Color.Green
            i = i + 1

            tcl.Add(New TableHeaderCell())
            tcl(i).Text = "Progress"
            tcl(i).BackColor = Color.Green
            i = i + 1

        End If
    End Sub

End Class
