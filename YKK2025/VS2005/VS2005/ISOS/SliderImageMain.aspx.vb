Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class SliderImageMain
    Inherits System.Web.UI.Page

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間VVVVV
    Dim UserID As String            'UserID
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
        Response.Cookies("PGM").Value = "SliderImageMain.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        UserID = Request.QueryString("pUserID")             'UserID

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(多Image - Slider Division )
    '**     
    '**
    '*****************************************************************
    Protected Sub BSLDImageITEM_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSLDImageITEM.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.10/sliderimage/inqslider.aspx');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(單Image - Group Portal )
    '**     
    '**
    '*****************************************************************
    Protected Sub BGRImage_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BGRImage.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.0.3/ISOS/FindItemImagePage.aspx?pUserID=" & UserID & "');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(多Image - RD & MKT Image)
    '**     
    '**
    '*****************************************************************
    Protected Sub BRDImagePuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BRDImagePuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NewAdvencedImages.aspx?pUserID=" & UserID & "&pFun=RD');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(多Image - SUPPLIER Image)
    '**     
    '**
    '*****************************************************************
    Protected Sub BSUPPLIERImagePuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSUPPLIERImagePuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NewAdvencedImages.aspx?pUserID=" & UserID & "&pFun=SUPPLIER');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(多Image - WINGS Slider Image PR)
    '**     
    '**
    '*****************************************************************
    Protected Sub BSLDImagePuller_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BSLDImagePuller.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/NewAdvencedImages.aspx?pUserID=" & UserID & "&pFun=WINGS');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Place of manufacture - R&D Slider Image)
    '**     
    '**
    '*****************************************************************
    Protected Sub BTWN2YOC_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BTWN2YOC.Click
        Dim Cmd As String
        Cmd = "<script>" + _
                    "window.open('http://10.245.1.6/ISOSQC/TWN2YOCCopyList_01.aspx?pUserID=" & UserID & "');" + _
              "</script>"
        Response.Write(Cmd)
    End Sub
End Class
