Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Partial Class NewAdvencedImages
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
    Dim oWaves As New WAVES.CommonService
    Dim oCommon As New Common.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim UserID As String            'UserID
    Dim pFun As String              'FUN
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
        Response.Cookies("PGM").Value = "AdvencedImages.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        UserID = Request.QueryString("pUserID")             'UserID
        pFun = Request.QueryString("pFun")                  'Fun
        If pFun = "RD" Then
            BFind.Style("left") = 304 & "px"
            BFindSupplier.Style("left") = -500 & "px"
            BFindWings.Style("left") = -500 & "px"
        End If
        If pFun = "SUPPLIER" Then
            BFind.Style("left") = -500 & "px"
            BFindSupplier.Style("left") = 424 & "px"
            BFindWings.Style("left") = -500 & "px"
        End If
        If pFun = "WINGS" Then
            BFind.Style("left") = -500 & "px"
            BFindSupplier.Style("left") = -500 & "px"
            BFindWings.Style("left") = 544 & "px"
        End If
        DStatus.Text = ""
        DStatus.Style("left") = -500 & "px"
        '
        DataList1.Visible = False
        DataList2.Visible = False
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
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     R&D
    '**
    '*****************************************************************
    Protected Sub BFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFind.Click
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     外注商
    '**
    '*****************************************************************
    Protected Sub BFindSupplier_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindSupplier.Click
        ShowSupplierData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Go)
    '**     Wings
    '**
    '*****************************************************************
    Protected Sub BFindWings_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindWings.Click
        ShowWingsData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim SQL As String
        Dim xRun As Boolean
        Dim RtnCode As Integer = 0
        '
        xRun = True
        DKSeries.Text = UCase(Trim(DKSeries.Text))
        '
        If xRun = True Then
            If DKSeries.Text = "" Then
                xRun = False
                uJavaScript.PopMsg(Me, "  篩選條件有誤, 請調整篩選條件 !! ")
            End If
        End If
        '
        If xRun = True Then
            DataList2.Visible = False
            '
            SQL = "select "
            SQL = SQL & "FormNo, "
            SQL = SQL & "case when formno='000003' then [SliderGRCode] + '  內製' "
            SQL = SQL & "     else [SliderGRCode] + '  外注' "
            SQL = SQL & "end as Code, "

            SQL = SQL & "case when yobi3 like '%1/%' then 'Link ..(有模廢)..' else 'Link .....' end As No, "
            SQL = SQL & "case when formno='000003' then 'http://10.245.1.6/ISOSQC/AdvencedImagesData.aspx?pFormNo=000003&pFormSno=' + ltrim(rtrim(str(formsno))) + '&pCode=' + SliderGRCode "
            SQL = SQL & "     else 'http://10.245.1.6/ISOSQC/AdvencedImagesData.aspx?pFormNo=000007&pFormSno=' + ltrim(rtrim(str(formsno))) + '&pCode=' + SliderGRCode "
            SQL = SQL & "end as NoUrl, "

            SQL = SQL & "ImagePath As ImagePath, "
            SQL = SQL & "Yobi2 As MapFile "
            '
            SQL = SQL & "from V_RDPullerImage "
            SQL = SQL & "where formno in ('000003','000007') "
            ' Series
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKSeries.Text) & ") = '" & DKSeries.Text & "' "
            '
            SQL = SQL & "Order by len([SliderGRCode]) desc, [SliderGRCode] desc "
            'SQL = SQL & "Order by substring([SliderGRCode],1," & Len(DKSeries.Text) & ") "
            '
            Dim dt_Map As DataTable = uDataBase.GetDataTable(SQL)
            For i As Integer = 0 To dt_Map.Rows.Count - 1
                '\\10.245.1.10\www$\WorkFlow\Document\000001
                '\\10.245.1.9\SyncData\inetpub\wwwroot\WorkFlow\Document\000001
                '
                If dt_Map.Rows(i)("FormNo").ToString = "000003" Then
                    RtnCode = oCommon.RecoveryArchiveFile("\\10.245.1.10\www$\", "WorkFlow\Document\000003\", dt_Map.Rows(i)("MapFile").ToString)
                Else
                    RtnCode = oCommon.RecoveryArchiveFile("\\10.245.1.10\www$\", "WorkFlow\Document\000007\", dt_Map.Rows(i)("MapFile").ToString)
                End If
                '
                DStatus.Text = "只顯示前500件" & Chr(13) & "最後1件=[" & Replace(dt_Map.Rows(i).Item("Code").ToString, " ", "") & "]" & Chr(13) & "如影像無顯示請再調整篩選條件"
            Next
            '
            DataList1.Visible = True
            DataList1.DataSource = uDataBase.GetDataTable(SQL)
            DataList1.DataBind()
        End If
        '
    End Sub '*****************************************************************
    '**(ShowSupplierData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowSupplierData()
        Dim SQL As String
        Dim xRun As Boolean
        Dim RtnCode As Integer = 0
        '
        xRun = True
        DKSeries.Text = UCase(Trim(DKSeries.Text))

        '
        If xRun = True Then
            If DKSeries.Text = "" Then
                xRun = False
                uJavaScript.PopMsg(Me, "  篩選條件有誤, 請調整篩選條件 !! ")
            End If
        End If
        ''
        If xRun = True Then
            DataList1.Visible = False
            '
            SQL = "select "
            SQL = SQL & "FormNo, "
            SQL = SQL & "case when formno='SUPPLIER' then [SliderGRCode] "
            SQL = SQL & "end as Code, "
            SQL = SQL & "'Link ISIP .....' As No, "
            SQL = SQL & "case when formno='SUPPLIER' then 'http://10.245.1.6/ISOSQC/PC_ReleaseNewcolor.aspx?pUserID=' + '" & UserID & "' + '&pSlider=' + Yobi1 + '&pBuyer=' + Yobi2 + '&pSource=IMGP' "
            SQL = SQL & "end as NoUrl, "
            SQL = SQL & "ImagePath As ImagePath, "
            SQL = SQL & "Yobi2 As MapFile, "
            '
            SQL = SQL & "'Zoom +' As IMGResize, ImagePath As IMGResizeURL  "
            '
            SQL = SQL & "from M_RDPullerImage "
            SQL = SQL & "where formno in ('SUPPLIER') "
            ' Series
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKSeries.Text) & ") = '" & DKSeries.Text & "' "
            '
            SQL = SQL & "Order by len([SliderGRCode]) desc, [SliderGRCode] desc "
            '
            DataList2.Visible = True
            DataList2.DataSource = uDataBase.GetDataTable(SQL)
            DataList2.DataBind()
        End If
        '
    End Sub
    '*****************************************************************
    '**(ShowWingsData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowWingsData()
        Dim SQL As String
        Dim xRun As Boolean
        Dim RtnCode As Integer = 0
        '
        xRun = True
        DKSeries.Text = UCase(Trim(DKSeries.Text))

        '
        If xRun = True Then
            If DKSeries.Text = "" Then
                xRun = False
                uJavaScript.PopMsg(Me, "  篩選條件有誤, 請調整篩選條件 !! ")
            End If
        End If
        '
        If xRun = True Then
            DataList2.Visible = False
            '--
            SQL = "select "
            SQL = SQL & "FormNo, "
            SQL = SQL & "[SliderGRCode] + '(' + YOBI2 + ')' as Code, "
            SQL = SQL & "Yobi1 As No, "
            SQL = SQL & "'' as NoUrl, "
            SQL = SQL & "ImagePath As ImagePath "
            '
            SQL = SQL & "from M_RDPullerImage "
            SQL = SQL & "where formno in ('WINGS')  "
            ' Series
            SQL = SQL & "and substring([SliderGRCode],1," & Len(DKSeries.Text) & ") = '" & DKSeries.Text & "' "
            '
            SQL = SQL & "Order by len([SliderGRCode]) desc, [SliderGRCode] desc "
            'SQL = SQL & "Order by substring([SliderGRCode],1," & Len(DKSeries.Text) & ") "
            '
            DataList1.Visible = True
            DataList1.DataSource = uDataBase.GetDataTable(SQL)
            DataList1.DataBind()
        End If
        '
    End Sub
    '
    '*****************************************************************
End Class
