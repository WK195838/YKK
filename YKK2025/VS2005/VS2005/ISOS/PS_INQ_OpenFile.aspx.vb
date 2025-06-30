Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class PS_INQ_OpenFile
    Inherits System.Web.UI.Page

    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String = ""      '申請者
    Dim pBuyer As String = ""       '申請者
    Dim pBuyerItem As String = ""   '
    Dim pFun As String = ""         '

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim CLConnect As String = uCommon.GetAppSetting("WAVESDBCL")
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                          '設定參數
        If Not IsPostBack Then                  'PostBack
            OpenDocument()
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
        Response.Cookies("PGM").Value = "PS_INQ_OpenFile.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID
        pBuyer = Request.QueryString("pBuyer")
        pBuyerItem = Request.QueryString("pBuyerItem")
        pFun = Request.QueryString("pFun")
        '
        DPOPUPIMAGEFile.Style("left") = -500 & "px"
        DPOPUPIMAGEFile.Text = "\\10.245.0.3\Program$\IE\IV.EXE XTWN802 WTWN802 000034 " & pBuyerItem
        DPOPUPIMAGEFile1.Style("left") = -500 & "px"
        DPOPUPIMAGEFile1.Text = "\\10.245.0.3\Program$\IE\IV.EXE XTWN802 WTWN802 000034 " & pBuyerItem
        '
        '動作按鈕設定
        BGO.Style("left") = -500 & "px"
        BGO.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('POPUPIMAGE')}"        '
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(OpenFile)
    '**     
    '**
    '*****************************************************************
    Sub OpenDocument()
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql, str As String
        Dim i As Integer
        '
        Dim xFilePath, xFileName As String
        Dim Cmd As String
        '
        '*******************
        ' SPC OPEN FILE
        '*******************
        If pFun = "SPC" Then
            '
            'NIKE
            '
            If pBuyer = "000013" Then
                xFilePath = "\\10.245.0.3\www$\ISOS\Document\000013\SPC\" & Trim(pBuyerItem) & ".xls"
                xFileName = Trim(pBuyerItem) & ".xls"
                'MsgBox(xFilePath)
                If File.Exists(xFilePath) Then
                    Cmd = "<script>" & _
                                "window.open('http://10.245.0.3/ISOS/Document/000013/SPC/" & xFileName & "');" & _
                         "</script>"
                    Response.Write(Cmd)
                Else
                    xFilePath = "\\10.245.0.3\www$\ISOS\Document\000013\SPC\" & Trim(pBuyerItem) & ".xlsx"
                    xFileName = Trim(pBuyerItem) & ".xlsx"
                    If File.Exists(xFilePath) Then
                        Cmd = "<script>" & _
                                    "window.open('http://10.245.0.3/ISOS/Document/000013/SPC/" & xFileName & "');" & _
                              "</script>"
                        Response.Write(Cmd)
                    Else
                        xFilePath = "\\10.245.0.3\www$\ISOS\Document\000013\SPC\" & Trim(pBuyerItem) & ".pdf"
                        xFileName = Trim(pBuyerItem) & ".pdf"
                        If File.Exists(xFilePath) Then
                            Cmd = "<script>" & _
                                        "window.open('http://10.245.0.3/ISOS/Document/000013/SPC/" & xFileName & "');" & _
                                  "</script>"
                            Response.Write(Cmd)
                        Else
                        End If
                    End If
                End If
            End If
            '
            'COLUMBIA
            '
            If pBuyer = "000003" Then
                str = ""
                cn.ConnectionString = EDLConnect
                '
                sql = "SELECT Top 1 A1 "
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERSPCHAIN" & "') "
                sql &= "And B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & Trim(pBuyerItem) & "%' "
                sql &= "And Active = '0' "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "SPCHAIN")
                If ds.Tables("SPCHAIN").Rows.Count > 0 Then
                    str = Trim(ds.Tables(0).Rows(i).Item("A1").ToString)
                End If
                '
                If str <> "" Then
                    xFilePath = "\\10.245.0.3\www$\ISOS\Document\000003\SPC\" & str & ".xls"
                    xFileName = str & ".xls"
                Else
                    xFilePath = "\\10.245.0.3\www$\ISOS\Document\000003\SPC\" & Trim(pBuyerItem) & ".xls"
                    xFileName = Trim(pBuyerItem) & ".xls"
                End If
                'MsgBox("[" & xFilePath & "][" & xFileName & "]")
                '
                If File.Exists(xFilePath) Then
                    Cmd = "<script>" & _
                                "window.open('http://10.245.0.3/ISOS/Document/000003/SPC/" & xFileName & "');" & _
                         "</script>"
                    Response.Write(Cmd)
                Else
                    xFilePath = "\\10.245.0.3\www$\ISOS\Document\000013\SPC\" & Trim(pBuyerItem) & ".xlsx"
                    xFileName = Trim(pBuyerItem) & ".xlsx"
                    If File.Exists(xFilePath) Then
                        Cmd = "<script>" & _
                                    "window.open('http://10.245.0.3/ISOS/Document/000003/SPC/" & xFileName & "');" & _
                              "</script>"
                        Response.Write(Cmd)
                    Else
                        xFilePath = "\\10.245.0.3\www$\ISOS\Document\000013\SPC\" & Trim(pBuyerItem) & ".pdf"
                        xFileName = Trim(pBuyerItem) & ".pdf"
                        If File.Exists(xFilePath) Then
                            Cmd = "<script>" & _
                                        "window.open('http://10.245.0.3/ISOS/Document/000003/SPC/" & xFileName & "');" & _
                                  "</script>"
                            Response.Write(Cmd)
                        Else
                        End If
                    End If
                End If
            End If
        End If
        '
        '*******************
        ' IMAGE PORTAL
        '*******************
        If pFun = "IMG" Then
            cn.ConnectionString = CLConnect
            '
            sql = "SELECT CIT2XA "
            sql = sql + "From MST_STRUCTURE "
            sql = sql + "Where PIPCXA = '" & pBuyerItem & "' "
            sql = sql + "And   CLSCXA = 'PS' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "STRUCTURE")
            If ds.Tables("STRUCTURE").Rows.Count > 0 Then
                str = ""
                For i = 0 To ds.Tables("STRUCTURE").Rows.Count - 1
                    If str = "" Then
                        str = ds.Tables(0).Rows(i).Item("CIT2XA").ToString
                    Else
                        str = str & " " & ds.Tables(0).Rows(i).Item("CIT2XA").ToString
                    End If
                Next
                DPOPUPIMAGEFile.Text = "\\10.245.0.3\Program$\IE\IV.EXE XTWN802 WTWN802 000034 " & str
            End If
            '
            Dim someScript As String = ""
            someScript = "<script language='javascript'>CheckAttribute('POPUPIMAGE');</script>"
            Page.ClientScript.RegisterStartupScript(Me.GetType(), "onload", someScript)
        End If
        '
        '
        '*******************
        ' 5 SECORD CLOSE PAGE
        '*******************
        Response.Write("<script>setTimeout(function () {window.open('', '_self', ''); window.close();}, 5000);</script>")
        '
    End Sub

End Class
