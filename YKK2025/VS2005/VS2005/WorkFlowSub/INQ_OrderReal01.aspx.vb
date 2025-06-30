Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class INQ_OrderReal01
    Inherits System.Web.UI.Page


    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pMonth As String
    Dim pDate As String
    Dim pSUMCODE As String
    Dim pAction As String

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then
            GetData()
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
        pMonth = Request.QueryString("pMonth")
        pDate = Request.QueryString("pDate")
        pSUMCODE = Request.QueryString("pSUMCODE")
        pAction = Request.QueryString("pAction")
        '-----------------------------------------------------------------
        '-- 初值
        GridView1.Visible = False
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetData)
    '**     取得資料
    '**
    '*****************************************************************
    Sub GetData()
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        '
        If IsNumeric(pDate) Then
            Sql = "select "
            Sql &= "ODRG8G As YYMM, "
            Sql &= "OCND8G As DAY, "
            Sql &= "RSMC8G As SUMCODE, "
            Sql &= "ROUND(sum(DORK8G)/1000,0) as AMOUNT, "
            If pAction = "BUYER" Then
                Sql &= "BYRC8G || '(' || BUYERN || ')' AS  KEYNAME "
            Else
                Sql &= "SPRC8G || '(' || SPRCN  || ')' as KEYNAME "
            End If
            '
            Sql &= "FROM SONGOLIB.S8G00WK "
            '
            Sql &= "WHERE ODRG8G = " & pMonth & " "
            Sql &= "And RSMC8G = '" & pSUMCODE & "' "
            '
            Sql &= "And OCND8G = " & pDate & " "
            '
            If pAction = "BUYER" Then
                Sql &= "GROUP BY RSMC8G, ODRG8G, OCND8G, BYRC8G, BUYERN "
            Else
                Sql &= "GROUP BY RSMC8G, ODRG8G, OCND8G, SPRC8G, SPRCN "
            End If
        Else
            Sql = "select "
            Sql &= "ODRG8G As YYMM, "
            Sql &= "'' As DAY, "
            Sql &= "RSMC8G As SUMCODE, "
            Sql &= "ROUND(sum(DORK8G)/1000,0) as AMOUNT, "
            If pAction = "BUYER" Then
                Sql &= "BYRC8G || '(' || BUYERN || ')' AS  KEYNAME "
            Else
                Sql &= "SPRC8G || '(' || SPRCN  || ')' as KEYNAME "
            End If
            '
            Sql &= "FROM SONGOLIB.S8G00WK "
            '
            Sql &= "WHERE ODRG8G = " & pMonth & " "
            Sql &= "And RSMC8G = '" & pSUMCODE & "' "
            '
            If pAction = "BUYER" Then
                Sql &= "GROUP BY RSMC8G, ODRG8G, BYRC8G, BUYERN "
            Else
                Sql &= "GROUP BY RSMC8G, ODRG8G, SPRC8G, SPRCN "
            End If

        End If

        Sql &= "ORDER BY AMOUNT DESC "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "S8G00")
        If ds.Tables("S8G00").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        'Header
        If (e.Row.RowType = DataControlRowType.Header) Then

        End If
        '
        'DataRow
        If (e.Row.RowType = DataControlRowType.DataRow) Then
        End If
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub



End Class
