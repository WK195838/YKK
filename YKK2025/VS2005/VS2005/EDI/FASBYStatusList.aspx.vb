Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

Partial Class FASBYStatusList
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
    Dim NowDateTime As String               ' 現在日時
    Dim xBuyer As String                    ' Buyer
    Dim xUserID As String                   ' 使用者ID
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Page_Load)
    '**     Load Page 處理
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                                  ' 設定共用參數
        '
        If Not IsPostBack Then
            GridView1.Style("left") = -1000 & "px"
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
        xBuyer = UCase(Request.QueryString("pBuyer"))
        xUserID = UCase(Request.QueryString("pUserID"))
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowData()
        '
        Dim sql As String
        Dim i As Integer
        Dim tCount As Long = 0
        '
        If xBuyer = "FALL-000001" Then
            D10Reason.Text = "BRAND <> ADIDAS"
        Else
            If xBuyer = "FALL-000016" Then
                D10Reason.Text = "BRAND <> REEBOK"
            End If
        End If
        '
        sql = "Select Count(*) As RCount From W_InputSheetBY "
        sql = sql & "Where Buyer = '" & xBuyer & "' "
        Dim dt_wInputSheet As DataTable = uDataBase.GetDataTable(sql)
        If dt_wInputSheet.Rows.Count > 0 Then
            DBYCount.Text = dt_wInputSheet.Rows(0).Item("RCount").ToString
        End If
        '
        sql = "Select Count(*) As RCount From E_InputSheetBY "
        sql = sql & "Where Buyer = '" & xBuyer & "' "
        Dim dt_eInputSheet As DataTable = uDataBase.GetDataTable(sql)
        If dt_eInputSheet.Rows.Count > 0 Then
            DSUMCount.Text = dt_eInputSheet.Rows(0).Item("RCount").ToString
        End If
        '
        sql = "Select EZ1, Count(*) As RCount From E_InputSheetBY "
        sql = sql & "Where Buyer = '" & xBuyer & "' "
        sql = sql & "Group By EZ1 "
        sql = sql & "Order By EZ1 "
        Dim dt_InputSheet As DataTable = uDataBase.GetDataTable(sql)
        For i = 0 To dt_InputSheet.Rows.Count - 1
            If dt_InputSheet.Rows(i).Item("EZ1") = "00" Then D00.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "01" Then D01.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "10" Then D10.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "20" Then D20.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "30" Then D30.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "40" Then D40.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "41" Then D41.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "42" Then D42.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "50" Then D50.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "52" Then D52.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "60" Then D60.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "61" Then D61.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "70" Then D70.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "71" Then D71.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "80" Then D80.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            If dt_InputSheet.Rows(i).Item("EZ1") = "81" Then D81.Text = dt_InputSheet.Rows(i).Item("RCount").ToString
            '
            tCount = tCount + CLng(dt_InputSheet.Rows(i).Item("RCount"))
        Next
        '
        DTTL.Text = CStr(tCount)
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(BExcel)
    '**     下載Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Dim i As Integer
        Dim sql As String
        Dim tCount As Long = 0
        '
        sql = "Select EZ1, '' As StsDescr, '' As Remark, "
        '
        sql = sql & "Case EZ1 "
        sql = sql & "When '10' Then '01' "
        sql = sql & "When '20' Then '02' "
        sql = sql & "When '40' Then '03' "
        sql = sql & "When '41' Then '04' "
        sql = sql & "When '42' Then '05' "
        sql = sql & "When '50' Then '06' "
        sql = sql & "When '52' Then '07' "
        sql = sql & "When '30' Then '08' "
        sql = sql & "When '80' Then '09' "
        sql = sql & "When '81' Then '10' "
        sql = sql & "When '60' Then '11' "
        sql = sql & "When '61' Then '12' "
        sql = sql & "When '70' Then '13' "
        sql = sql & "When '71' Then '14' "
        sql = sql & "When '01' Then '15' "
        sql = sql & "When '00' Then '16' "
        sql = sql & "Else '99' "
        sql = sql & "End As CheckSeq, "
        '
        sql = sql & "Count(*) As RCount "
        '
        sql = sql & "From E_InputSheetBY "
        sql = sql & "Where Buyer = '" & xBuyer & "' "
        Sql = Sql & "Group By EZ1 "
        sql = sql & "Order By CheckSeq "
        Dim dt_InputSheet As DataTable = uDataBase.GetDataTable(Sql)
        For i = 0 To dt_InputSheet.Rows.Count - 1
            '
            If dt_InputSheet.Rows(i).Item("EZ1") = "10" Then
                dt_InputSheet.Rows(i)(1) = "非對象BUYER"
                If xBuyer = "FALL-000001" Then
                    dt_InputSheet.Rows(i)(2) = "BRAND <> ADIDAS"
                Else
                    If xBuyer = "FALL-000016" Then
                        dt_InputSheet.Rows(i)(2) = "BRAND <> REEBOK"
                    End If
                End If
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "20" Then
                dt_InputSheet.Rows(i)(1) = "非對象供應商"
                dt_InputSheet.Rows(i)(2) = "T2供應商 <> YKK GROUP"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "40" Then
                dt_InputSheet.Rows(i)(1) = "新成衣廠"
                dt_InputSheet.Rows(i)(2) = "NATIVE VENDOR MASTER不存在"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "41" Then
                dt_InputSheet.Rows(i)(1) = "姐妹社"
                dt_InputSheet.Rows(i)(2) = "NATIVE VENDOR MASTER存在但非國內客戶"
            End If
            '
            If dt_InputSheet.Rows(i).Item("EZ1") = "42" Then
                dt_InputSheet.Rows(i)(1) = "姐妹社（特殊需台灣生產）"
                dt_InputSheet.Rows(i)(2) = "ITEM DESCR.符合特殊材料 MASTER"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "50" Then
                dt_InputSheet.Rows(i)(1) = "ＣＣＳ"
                dt_InputSheet.Rows(i)(2) = "WORKING NO符合(Ａ～ OR ～Ｓ) OR CCS-ITEM"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "52" Then
                dt_InputSheet.Rows(i)(1) = "ＣＣＳ（特殊需台灣生產）"
                dt_InputSheet.Rows(i)(2) = "ITEM DESCR.符合特殊材料 MASTER"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "30" Then
                dt_InputSheet.Rows(i)(1) = "需確認的FCT DATA"
                dt_InputSheet.Rows(i)(2) = "PART WITH LEVEL2=非SOLID及MULTI(含WRONG)"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "80" Then
                dt_InputSheet.Rows(i)(1) = "(S) FCT 數量不符"
                dt_InputSheet.Rows(i)(2) = "FCT 數量 = 0"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "81" Then
                dt_InputSheet.Rows(i)(1) = "(M) FCT 數量不符"
                dt_InputSheet.Rows(i)(2) = "FCT 數量 = 0"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "60" Then
                dt_InputSheet.Rows(i)(1) = "(S) PLM OR COLOR空白"
                dt_InputSheet.Rows(i)(2) = "PLM = 空白 OR COLOR = 空白"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "61" Then
                dt_InputSheet.Rows(i)(1) = "(M) PLM OR COLOR空白"
                dt_InputSheet.Rows(i)(2) = "PLM = 空白 OR COLOR = 空白"
            End If
            '
            If dt_InputSheet.Rows(i).Item("EZ1") = "70" Then
                dt_InputSheet.Rows(i)(1) = "(M) COLOR PATTERN不符"
                dt_InputSheet.Rows(i)(2) = "ELEMENT <> COLOR PATTERN FILE"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "71" Then
                dt_InputSheet.Rows(i)(1) = "(M) COLOR 結構不符"
                dt_InputSheet.Rows(i)(2) = "不符合COLOR結構基本要素(TAPE, TEETH, PULLER)"
            End If
            '
            If dt_InputSheet.Rows(i).Item("EZ1") = "01" Then
                dt_InputSheet.Rows(i)(1) = "正常-國內進口半成品"
                dt_InputSheet.Rows(i)(2) = "備料對象"
            End If
            If dt_InputSheet.Rows(i).Item("EZ1") = "00" Then
                dt_InputSheet.Rows(i)(1) = "正常"
                dt_InputSheet.Rows(i)(2) = "備料對象"
            End If
            '
            tCount = tCount + CLng(dt_InputSheet.Rows(i).Item("RCount"))
        Next
        '
        Dim drNew As DataRow = dt_InputSheet.NewRow()
        drNew("CheckSeq") = ""
        drNew("EZ1") = ""
        drNew("StsDescr") = ""
        drNew("Remark") = "計"
        drNew("RCount") = CStr(tCount)
        dt_InputSheet.Rows.Add(drNew)
        '
        GridView1.DataSource = dt_InputSheet
        GridView1.DataBind()
        '
        '---------------------------------------------------------------------------------------------------------

        Response.AppendHeader("Content-Disposition", "attachment;filename=FASBY_StatusList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)

        GridView1.RenderControl(hw)

        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub


End Class
