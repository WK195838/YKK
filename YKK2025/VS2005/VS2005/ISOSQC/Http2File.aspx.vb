Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class Http2File
    Inherits System.Web.UI.Page
    '
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pNo As String               'No
    Dim pPuller As String           'Puller
    Dim pColor As String            'Color
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        ShowData()
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
        Response.Cookies("PGM").Value = "Http2File.aspx"    '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pNo = Request.QueryString("pNo")                    'No
        pPuller = Request.QueryString("pPuller")            'Puller
        pColor = Request.QueryString("pColor")              'Color
    End Sub

    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim Cmd As String
        '
        If pNo <> "" Then
            Cmd = "<script>" + _
                        "window.open('file://10.245.1.6/wfs$/ISOSQC/008002/" & pNo & "/核可卡','',''); " + _
                  "</script>"
        Else
            If pPuller <> "" And pColor <> "" Then
                Cmd = "<script>" + _
                            "window.open('file://10.245.1.6/www$/ISOSQC/Document/EDX/" & Trim(pPuller) & Trim(pColor) & ".jpg','EDX','height=630, width=520, top=100, left=100, toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no'); " + _
                      "</script>"
            End If
        End If
        '
        Response.Write(Cmd)
        '
        '限IE11
        Response.Write("<script>window.open('', '_self', ''); window.close();</script>")

    End Sub
End Class
