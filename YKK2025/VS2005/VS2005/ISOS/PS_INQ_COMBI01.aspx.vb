Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing


Partial Class PS_INQ_COMBI01
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
    Dim oWaves As New Waves.CommonService

    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim pBuyer As String            'Buyer
    Dim UserID As String            'UserID
    Dim pBuyerItem As String        'BuyerItem
    Dim pFun As String              'Fun
    '
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
        Response.Cookies("PGM").Value = "PS_INQ_COMBI01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        pBuyerItem = Request.QueryString("pBuyerItem")      'BuyerItem
        pFun = Request.QueryString("pFun")                  'Fun
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        GridView2.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"        '
        '動作按鈕設定
        '
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        '
        DBUYER.Text = ""
        DKEY1.Text = ""
        DKEY2.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBUYER.Text = "TNF"
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "000003" Then DBUYER.Text = "COLUMBIA"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            If pBuyer = "TW0655" Then DBUYER.Text = "VERA BRADLEY"

            HBuyerCode.Text = "FALL-" & pBuyer
        End If
        '
        If pBuyerItem <> "" Then
            DKEY1.Text = pBuyerItem
        End If
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        ShowData()
    End Sub
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        cn.ConnectionString = EDLConnect
        '
        '--FAS
        sql = "SELECT TOP 300 "
        '
        sql &= "Code, Color1, Color2, Color3, SliderStatus, YCode, YName1 + ' ' + YName2 + ' ' + REPLACE(SUBSTRING([YName3],1,CHARINDEX('(',YNAME3)-1),'[]','') As YName "
        sql &= "From M_ItemConvert "
        If HBuyerCode.Text = "FALL-000001" Then
            sql &= "Where Buyer IN ('FALL-000001','FALL-000016') "
        Else
            sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
        End If
        '
        If DKEY1.Text <> "" Then sql &= " And CODE+COLOR1+COLOR2+COLOR3+SliderStatus+YCode+YName1+YName2+YName3 Like '%" & DKEY1.Text & "%' "
        If DKEY2.Text <> "" Then sql &= " And CODE+COLOR1+COLOR2+COLOR3+SliderStatus+YCode+YName1+YName2+YName3 Like '%" & DKEY2.Text & "%' "
        '
        sql &= " Group by Code, Color1, Color2, Color3, SliderStatus, YCode, YName1, YName2, YName3 "
        sql &= " Order by Code, Color1, Color2, Color3, SliderStatus, YCode, YName1, YName2, YName3 "

        Dim dt_Item As DataTable = uDataBase.GetDataTable(sql)
        If dt_Item.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_Item
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        On Error GoTo LBL_Error
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        Dim k_Tape As String = GridView1.SelectedRow.Cells(2).Text
        Dim k_Teeth As String = GridView1.SelectedRow.Cells(3).Text
        '
        cn.ConnectionString = EDLConnect
        '
        If pFun = "COMBI" And k_Tape <> "" And k_Teeth <> "" Then
            ShowData()
            '
            sql = "SELECT TOP 300 "
            sql &= "Season, Color1, Color2, Green, YColor "
            sql &= "From M_ColorConvert "
            If HBuyerCode.Text = "FALL-000001" Then
                sql &= "Where Buyer IN ('FALL-000001','FALL-000016') "
            Else
                sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            End If
            '
            If DColorType.SelectedValue = "TAPE" Then sql &= " And COLOR1 Like '%" & k_Tape & "%' "
            If DColorType.SelectedValue = "TAPETEETH" Then sql &= " And COLOR1 Like '%" & k_Tape & "%' And COLOR2 Like '%" & k_Teeth & "%' "
            If DColorType.SelectedValue = "ALL" Then sql &= " And ( COLOR1+'/'+COLOR2 Like '%" & k_Tape & "%' Or COLOR1+'/'+COLOR2 Like '%" & k_Teeth & "%') "
            '
            sql &= " Group by Season, COLOR1, COLOR2, Green, YColor "
            sql &= " Order by substring(Season,3,2) Desc, COLOR1, COLOR2, Green, YColor "

            Dim dt_Color As DataTable = uDataBase.GetDataTable(sql)
            If dt_Color.Rows.Count > 0 Then
                GridView2.Visible = True
                GridView2.DataSource = dt_Color
                GridView2.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub


End Class
