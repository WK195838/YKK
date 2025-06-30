Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class PS_INQ_Item01
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
    Dim pBuyer As String             'Buyer
    Dim UserID As String            'UserID
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
        Response.Cookies("PGM").Value = "PS_INQ_Item01.aspx"   '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        pBuyer = Request.QueryString("pBuyer")
        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        DBUYER.ReadOnly = True
        HBuyerCode.Style("left") = -500 & "px"
        DRemark.Style("left") = -500 & "px"
        LCPSIAPage.Style("left") = -500 & "px"
        '
        DINQBUYERITEMFile.Style("left") = -500 & "px"
        If pBuyer <> "TW0371" Then Label4.Style("left") = -500 & "px"
        If pBuyer = "TW5068" Or pBuyer = "000055" Or pBuyer = "TW0020" Then DRemark1.Style("left") = -500 & "px"

        '
        DINQBUYERITEMFile.Text = uCommon.GetAppSetting("DataPrepareFile") + "INQBUYERITEM_" & UserID & ".xlsm"
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            If pBuyer = "000098" Or pBuyer = "000068" Or pBuyer = "TW5068" Or pBuyer = "000055" Or pBuyer = "TW0020" Then
                AtBuyer.Checked = True
                AtFAS.Style("left") = -500 & "px"
                AtReplace.Style("left") = -500 & "px"
            Else
                AtBuyer.Checked = True
                'AtFAS.Checked = True
            End If
        End If
        '
        '動作按鈕設定
        BSPInq.Attributes("onclick") = "var ok=window.confirm('" + "是否開啟？" + "');if(!ok){return false;} else {CheckAttribute('INQBUYERITEMExcel')}"
        '
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DBUYER.Text = ""
        HBuyerCode.Text = ""
        DKEY1.Text = ""
        DKEY2.Text = ""
        '
        If pBuyer <> "" Then
            If pBuyer = "000021" Then DBUYER.Text = "TNF"
            If pBuyer = "000001" Then DBUYER.Text = "ADIDAS"
            If pBuyer = "000013" Then DBUYER.Text = "NIKE"
            If pBuyer = "000003" Then DBUYER.Text = "COLUMBIA"
            If pBuyer = "TW0371" Then DBUYER.Text = "UNDER ARMOUR"
            If pBuyer = "000098" Then DBUYER.Text = "HELLY HANSEN"
            If pBuyer = "000068" Then DBUYER.Text = "BURTON"
            If pBuyer = "TW5068" Then DBUYER.Text = "HERSCHEL"
            If pBuyer = "TW0655" Then DBUYER.Text = "VERA BRADLEY"
            If pBuyer = "000141" Then DBUYER.Text = "PATAGONIA"
            If pBuyer = "000053" Then DBUYER.Text = "ARCTERYX"
            If pBuyer = "000055" Then DBUYER.Text = "SALOMON"
            If pBuyer = "TW0020" Then DBUYER.Text = "TUMI"
            HBuyerCode.Text = "FALL-" & pBuyer
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(Inq)
    '**     查詢
    '**
    '*****************************************************************
    Protected Sub BInq_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BInq.Click
        If pBuyer = "000003" And AtBuyer.Checked = True Then
            LCPSIAPage.Style("left") = 40 & "px"
        Else
            LCPSIAPage.Style("left") = -500 & "px"
        End If
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
        '篩選資料
        '=========================================================================
        '共用
        '=========================================================================
        '
        '--開發(SPD)
        If AtSPD.Checked = True Then
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + Code + '&pItemName=' + Replace(ItemName1 + ' ' + ItemName2,'+','*') as URL, "
            sql &= "Code AS A1, "
            sql &= "ItemName1 AS B1, "
            sql &= "ItemName2 AS C1, "
            sql &= "ItemName3 AS D1, "
            sql &= "No AS E1, Name AS F1, Sts AS G1, ApplyDate AS H1, "
            sql &= "'' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1, "

            sql &= "'Ｆ' as FormMark, "
            sql &= "FURL AS FormURL "
            sql &= "From V_Develop_Digital "

            sql &= "Where SUBSTRING(SYS,1,3) = 'SPD' "
            
            '
            If pBuyer = "000021" Then
                sql &= "And (BuyerCode Like '%" & "000021" & "%' "
                sql &= "     or BuyerCode Like '%" & "TNF" & "%' "
                sql &= "     or BuyerCode Like '%" & "THE NORTH FACE" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000001" Then
                sql &= "And (BuyerCode Like '%" & "000001" & "%' "
                sql &= "     or BuyerCode Like '%" & "000016" & "%' "
                sql &= "     or BuyerCode Like '%" & "ADIDAS" & "%' "
                sql &= "     or BuyerCode Like '%" & "REEBOK" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000013" Then
                sql &= "And (BuyerCode Like '%" & "000013" & "%' "
                sql &= "     or BuyerCode Like '%" & "NIKE" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000003" Then
                sql &= "And (BuyerCode Like '%" & "000003" & "%' "
                sql &= "     or BuyerCode Like '%" & "COLUMBIA" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "TW0371" Then
                sql &= "And (BuyerCode Like '%" & "TW0371" & "%' "
                sql &= "     or BuyerCode Like '%" & "UA" & "%' "
                sql &= "     or BuyerCode Like '%" & "UNDER ARMOUR" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000098" Then
                sql &= "And (BuyerCode Like '%" & "000098" & "%' "
                sql &= "     or BuyerCode Like '%" & "HH" & "%' "
                sql &= "     or BuyerCode Like '%" & "HELLY" & "%' "
                sql &= "     or BuyerCode Like '%" & "HANSEN" & "%' "
                sql &= "     or BuyerCode Like '%" & "HELLY HANSEN" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000068" Then
                sql &= "And (BuyerCode Like '%" & "000068" & "%' "
                sql &= "     or BuyerCode Like '%" & "BT" & "%' "
                sql &= "     or BuyerCode Like '%" & "BUR" & "%' "
                sql &= "     or BuyerCode Like '%" & "TON" & "%' "
                sql &= "     or BuyerCode Like '%" & "BURTON" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "TW5068" Then
                sql &= "And (BuyerCode Like '%" & "TW5068" & "%' "
                sql &= "     or BuyerCode Like '%" & "HERSCHE" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "TW0655" Then
                sql &= "And (BuyerCode Like '%" & "TW0655" & "%' "
                sql &= "     or BuyerCode Like '%" & "VERA BRADLEY" & "%' "
                sql &= "     or BuyerCode Like '%" & "VB" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000141" Then
                sql &= "And (BuyerCode Like '%" & "000141" & "%' "
                sql &= "     or BuyerCode Like '%" & "PATAGONIA" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000053" Then
                sql &= "And (BuyerCode Like '%" & "000053" & "%' "
                sql &= "     or BuyerCode Like '%" & "ARC’TERYX" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "000055" Then
                sql &= "And (BuyerCode Like '%" & "000055" & "%' "
                sql &= "     or BuyerCode Like '%" & "SALOMON" & "%' "
                sql &= "    ) "
            End If
            If pBuyer = "TW0020" Then
                sql &= "And (BuyerCode Like '%" & "TW0020" & "%' "
                sql &= "     or BuyerCode Like '%" & "TUMI" & "%' "
                sql &= "    ) "
            End If

            '
            If DFKEY.Text <> "" Then
                '使用A1~G1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "Code/ItemName1/ItemName2/ItemName3/No/Name/Sts/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And Code+ItemName1+ItemName2+ItemName3+No+Name+Sts Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And Code+ItemName1+ItemName2+ItemName3+No+Name+Sts Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by No DESC,Code,ItemName1,ItemName2 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '--開發(IRW)
        If AtIRW.Checked = True Then
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + Code + '&pItemName=' + Replace(ItemName1 + ' ' + ItemName2,'+','*') as URL, "
            sql &= "Code AS A1, "
            sql &= "ItemName1 AS B1, "
            sql &= "ItemName2 AS C1, "
            sql &= "ItemName3 AS D1, "
            sql &= "No AS E1, Name AS F1, Sts AS G1, ApplyDate AS H1, "
            sql &= "'' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1, "
            '
            sql &= "'Ｆ' as FormMark, "
            sql &= "FURL AS FormURL "
            '
            sql &= "From V_Develop_Digital "
            '
            sql &= "Where BuyerCode = '" & pBuyer & "' "
            sql &= "And  SUBSTRING(SYS,1,3) = 'IRW' "
            
            '
            If DFKEY.Text <> "" Then
                '使用A1~G1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "Code/ItemName1/ItemName2/ItemName3/No/Name/Sts/ApplyDate/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And Code+ItemName1+ItemName2+ItemName3+No+Name+Sts+ApplyDate Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And Code+ItemName1+ItemName2+ItemName3+No+Name+Sts+ApplyDate Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by ApplyDate DESC,Code,ItemName1,ItemName2 "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '--FAS
        If AtFAS.Checked = True Then
            sql = "SELECT TOP 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + Code + '&pItemName=' + Replace(YName1 + ' ' + YName2,'+','*') + '&pYCode=' + YCode as URL, "
            sql &= "Code AS A1, "
            sql &= "Color1 AS B1, "
            sql &= "Color2 AS C1, "
            sql &= "Color3 AS D1, "
            sql &= "SliderStatus AS E1, "
            sql &= "YCode AS F1, "
            sql &= "YName1 AS G1, "
            sql &= "YName2 AS H1, "
            sql &= "SUBSTRING([YName3],1,CHARINDEX('(',YNAME3)-1) AS I1, "
            sql &= "'' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1, "
            '
            sql &= "'' as FormMark, "
            sql &= "'' AS FormURL "
            '
            sql &= "From M_ItemConvert "
            
            If HBuyerCode.Text = "FALL-000001" Then
                sql &= "Where Buyer IN ('FALL-000001','FALL-000016') "
            Else
                sql &= "Where Buyer = '" & HBuyerCode.Text & "' "
            End If
            '
            If DFKEY.Text <> "" Then
                '使用A1~I1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "CODE/COLOR1/COLOR2/COLOR3/SliderStatus/YCode/YName1/YName2/YName3/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And CODE+COLOR1+COLOR2+COLOR3+SliderStatus+YCode+YName1+YName2+YName3 Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And CODE+COLOR1+COLOR2+COLOR3+SliderStatus+YCode+YName1+YName2+YName3 Like '%" & DKEY2.Text & "%' "
            '
            sql &= " Group by Code, Color1, Color2, Color3, SliderStatus, YCode, YName1, YName2, YName3 "
            sql &= " Order by Code, Color1, Color2, Color3, SliderStatus, YCode "
            Dim dt_Item As DataTable = uDataBase.GetDataTable(sql)
            If dt_Item.Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = dt_Item
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '--SALES
        If AtSales.Checked = True Then
            sql = "SELECT top 300 "
            sql &= "'Ｓ' as Mark, "
            sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + BUYERITEM + '&pItemName=' + Replace(ITEMNAME,'+','*') as URL, "
            sql &= "ITEM AS A1, "
            sql &= "ITEMNAME AS B1, "
            sql &= "BUYERITEM AS C1, "
            sql &= "CUSTOMERNAME + '(' + CUSTOMER + ')' AS D1, "
            sql &= "QTY AS E1, "
            sql &= "'' AS F1, '' AS G1, '' AS H1, '' AS I1, '' AS J1, "
            sql &= "'' AS K1, '' AS L1, '' AS M1, '' AS N1, '' AS O1, '' AS P1, '' AS Q1, '' AS R1, '' AS S1, '' AS T1, "
            sql &= "'' AS U1, '' AS V1, '' AS W1, '' AS X1, '' AS Y1, '' AS Z1, "
            '
            sql &= "'' as FormMark, "
            sql &= "'' AS FormURL "
            '
            sql &= "From V_OPReportData_DigitalItem "
            sql &= "Where Buyer = '" & pBuyer & "' "

            '
            If DFKEY.Text <> "" Then
                '使用A1~D1
                sql &= "And (" & GetSearchField(DFKEY.Text, _
                                               "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                               "ITEM/ITEMNAME/BUYERITEM/CUSTOMERNAME+CUSTOMER/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
            End If
            If DKEY1.Text <> "" Then sql &= " And ITEM+ITEMNAME+BUYERITEM+CUSTOMERNAME+CUSTOMER Like '%" & DKEY1.Text & "%' "
            If DKEY2.Text <> "" Then sql &= " And ITEM+ITEMNAME+BUYERITEM+CUSTOMERNAME+CUSTOMER Like '%" & DKEY2.Text & "%' "
            '*
            sql &= " Order by QTY DESC,CUSTOMER, ITEM,ITEMNAME,BUYERITEM "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "item")
            If ds.Tables("item").Rows.Count > 0 Then
                GridView1.Visible = True
                GridView1.DataSource = ds
                GridView1.DataBind()
            Else
                uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
            End If
        End If
        '
        '=========================================================================
        'BUYER別
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(P1,'+','*') + '&pYCode=' + O1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/O1/P1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用ITEM
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + Replace(D1,'+','*') as URL, "
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACEITEM" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/O1/P1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '*
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + G1 + '&pItemName=' + Replace(Replace(O1,'#',''),'~','') + '&pYCode=' + N1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/L1/M1/N1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用ITEM
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + Replace(D1,'+','*') as URL, "
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACEITEM" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/O1/P1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '*
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(Replace(V1,'#',''),'~','') + '&pYCode=' + U1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/Q1/R1/S1/T1/U1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用ITEM
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + Replace(D1,'+','*') as URL, "
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACEITEM" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/O1/P1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '*
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + E1 + '&pItemName=' + Replace(Replace(U1,'#',''),'~','') + '&pYCode=' + T1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用ITEM
            If AtReplace.Checked = True Then
                sql = "SELECT top 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + B1 + '&pItemName=' + Replace(D1,'+','*') as URL, "
                sql &= "A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType = '" & "REPLACEITEM" & "' "
                sql &= "And Active = 0 "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/O1/P1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '*
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                '
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'UA
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(Replace(O1,'#',''),'~','') + '&pYCode=' + N1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                sql &= "And A1 = 'APP' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1 DESC,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
            '
            '--代用ITEM
            If AtReplace.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + C1 + '&pItemName=' + Replace(Replace(Q1,'#',''),'~','') + '&pYCode=' + P1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                sql &= "And A1 = 'BAG' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + H1 + '&pItemName=' + Replace(Replace(I1,'#',''),'~','') + '&pYCode=' + H1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + H1 + '&pItemName=' + Replace(Replace(I1,'#',''),'~','') + '&pYCode=' + H1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + ''  + '&pItemName=' + Replace(Replace(Replace(G1,'#',''),'~',''),'*','') + '&pYCode=' + B1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        '=========================================================================
        'VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(O1,'+','*') + '&pYCode=' + N1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "
                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        ''=========================================================================
        ''PATAGONIA
        ''=========================================================================
        If pBuyer = "000141" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + C1 + '&pItemName=' + Replace(H1,'+','*') + '&pYCode=' + G1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "

                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If
        '
        ''=========================================================================
        ''ARCTERYX
        ''=========================================================================
        If pBuyer = "000053" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(L1,'+','*') + '&pYCode=' + K1 as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "

                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        ''=========================================================================
        ''SALOMON
        ''=========================================================================
        If pBuyer = "000055" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=' + D1 + '&pItemName=' + Replace(F1,'+','*') + '&pYCode=' as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "

                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        ''=========================================================================
        ''TUMI
        ''=========================================================================
        If pBuyer = "TW0020" Then
            '
            '--BUYERITEM
            If AtBuyer.Checked = True Then
                sql = "SELECT TOP 300 "
                sql &= "'Ｓ' as Mark, "
                sql &= "'http://10.245.0.3/ISOS/FindItemPage.aspx?pUserID=" & UserID & "&pBuyer=" & pBuyer & "&pBuyerItem=&pItemName=' + Replace(G1,'+','*') + '&pYCode=' + F1  as URL, "
                sql &= "A1, B1, C1, D1, E1, "
                sql &= "F1, G1, H1, I1, J1, "
                sql &= "K1, L1, M1, N1, O1, P1, Q1, R1, S1, T1, "
                sql &= "U1, V1, W1, X1, Y1, Z1, "
                '
                sql &= "'' as FormMark, "
                sql &= "'' AS FormURL "
                '
                sql &= "From M_PSCommonData "
                sql &= "Where Buyer = '" & pBuyer & "' "
                sql &= "And DataType IN ('" & "BUYERITEM" & "', '" & "BUYERITEM-B" & "') "
                sql &= "And Active = '0' "

                '
                If DFKEY.Text <> "" Then
                    '使用A1~E1, O1~P1
                    sql &= "And (" & GetSearchField(DFKEY.Text, _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1", _
                                                   "A1/B1/C1/D1/E1/F1/G1/H1/I1/J1/K1/L1/M1/N1/O1/P1/Q1/R1/S1/T1/U1/V1/W1/X1/Y1/Z1") & ") "
                End If
                If DKEY1.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY1.Text & "%' "
                If DKEY2.Text <> "" Then sql &= " And A1+B1+C1+D1+E1+F1+G1+H1+I1+J1+K1+L1+M1+N1+O1+P1+Q1+R1+S1+T1+U1+V1+W1+X1+Y1+Z1 Like '%" & DKEY2.Text & "%' "
                '
                sql &= " Order by A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1,L1,M1,N1,O1,P1,Q1,R1,S1,T1,U1,V1,W1,X1,Y1,Z1 "
                Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
                DBAdapter1.Fill(ds, "item")
                If ds.Tables("item").Rows.Count > 0 Then
                    GridView1.Visible = True
                    GridView1.DataSource = ds
                    GridView1.DataBind()
                Else
                    uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!")
                End If
            End If
        End If

        GoTo LBL_End
        '##
LBL_Error:
        On Error GoTo -1
        If cn.State = ConnectionState.Open Then cn.Close()
        uJavaScript.PopMsg(Me, "指定條件搜尋不到資料,請確認!")
        '##
LBL_End:
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     Get Search Field
    '**
    '*****************************************************************
    Public Function GetSearchField(ByVal pCmd As String, ByVal pFieldStr As String, ByVal pDataStr As String) As String
        Dim RtnString As String = ""
        Dim str As String
        Dim i As Integer
        Dim FieldNames(), DataNames() As String
        '
        Try
            str = UCase(pCmd)
            FieldNames = pFieldStr.Split("/")
            DataNames = pDataStr.Split("/")
            For i = 0 To DataNames.Length - 1
                If FieldNames(i) <> DataNames(i) Then
                    'MsgBox(FieldNames(i) & ";" & DataNames(i))
                    'MsgBox("B=" & str)
                    str = str.Replace(FieldNames(i), DataNames(i))
                    'MsgBox("A=" & str)
                    If str <> UCase(pCmd) Then Exit For
                End If
            Next
            RtnString = str
        Catch ex As Exception
        End Try
        '
        Return RtnString
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     編輯資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        '
        DRemark.Style("left") = -500 & "px"
        '=========================================================================
        '共通
        '=========================================================================
        '**開發
        If AtSPD.Checked = True Then
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Code"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Group Code"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Spec"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "　"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "SPD NO."
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "F1" & "<BR>" & "申請者"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "G1" & "<BR>" & "Status"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "H1" & "<BR>" & "申請日"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 8 To 25
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        If AtIRW.Checked = True Then
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Code"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "ItemName1"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "ItemName2"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "ItemName3"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "IRW NO."
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "F1" & "<BR>" & "申請者"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "G1" & "<BR>" & "Status"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "H1" & "<BR>" & "申請日"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                For i = 8 To 25
                    e.Row.Cells(i).Visible = False
                Next
            End If
        End If
        '
        '**FAS
        If AtFAS.Checked = True Then
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Code"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Tape"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Teeth"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "Slider"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "Other"
                tcl.Add(New TableHeaderCell())
                tcl(5).Text = "F1" & "<BR>" & "Item"
                tcl.Add(New TableHeaderCell())
                tcl(6).Text = "G1" & "<BR>" & "ItemName1"
                tcl.Add(New TableHeaderCell())
                tcl(7).Text = "H1" & "<BR>" & "ItemName2"
                tcl.Add(New TableHeaderCell())
                tcl(8).Text = "I1" & "<BR>" & "ItemName3"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                    e.Row.Cells(5).ForeColor = Color.Red
                    e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                    'e.Row.Cells(26).Text = ""
                Else
                    e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                End If

                For i = 9 To 27
                    If i <> 26 Then
                        e.Row.Cells(i).Visible = False
                    End If
                Next
            End If
        End If
        '
        '--SALES
        If AtSales.Checked = True Then
            DRemark.Text = "BULK & " & DateAdd("m", -3, Now).ToString("yyyyMM") & "01 ~"
            DRemark.Style("left") = 748 & "px"
            '
            'Header
            If (e.Row.RowType = DataControlRowType.Header) Then
                Dim tcl As TableCellCollection = e.Row.Cells
                tcl.Clear() '清除自动生成的表头
                '添加新的表头第一行
                tcl.Add(New TableHeaderCell())
                tcl(0).Text = "A1" & "<BR>" & "Item"
                tcl.Add(New TableHeaderCell())
                tcl(1).Text = "B1" & "<BR>" & "Item Name"
                tcl.Add(New TableHeaderCell())
                tcl(2).Text = "C1" & "<BR>" & "Buyer Item"
                tcl.Add(New TableHeaderCell())
                tcl(3).Text = "D1" & "<BR>" & "Customer"
                tcl.Add(New TableHeaderCell())
                tcl(4).Text = "E1" & "<BR>" & "Qty"
            End If
            '
            'DataRow
            If (e.Row.RowType = DataControlRowType.DataRow) Then
                e.Row.Cells(4).Text = String.Format("{0:N0}", CDbl(e.Row.Cells(4).Text))
                e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
                For i = 5 To 25
                    e.Row.Cells(i).Visible = False
                Next
                e.Row.Cells(27).Visible = False
            End If
        End If
        '
        '=========================================================================
        'TNF
        '=========================================================================
        If pBuyer = "000021" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Base length (in inch, for zipper only)"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Price for basic length (USD/100pcs)"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Price for 1inch up ( USD/100pcs)"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "YKK/Web Code"
                    tcl(9).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "Article" & "<BR>" & "(Puller:Tnf Black僅供參考)"
                    tcl(10).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    If InStr(e.Row.Cells(15).Text, "[1]") > 0 Then
                        e.Row.Cells(14).ForeColor = Color.Red
                        e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(15).Text, "[]") > 0 Then
                            e.Row.Cells(14).ForeColor = Color.Blue
                            e.Row.Cells(15).Text = Replace(e.Row.Cells(15).Text, "[]", "")
                        End If
                    End If

                    For i = 9 To 13
                        e.Row.Cells(i).Visible = False
                    Next
                    '
                    For i = 16 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
            '
            '**代用ITEM
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Substitute"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "A1" & "<BR>" & "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 5 To 25
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'ADIDAS
        '=========================================================================
        If pBuyer = "000001" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Item Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "End Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Kids safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "PLM#"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "YKK ITEM"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Base length"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "Base (US$/PC)"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "add.1 inch (US$/PC)"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "Sample L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "YKK/PLM"
                    tcl(13).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "O1" & "<BR>" & "YKK ITEM" & "<BR>" & "(Puller:095A僅供參考)"
                    tcl(14).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, Chr(10), "<br>")
                    e.Row.Cells(13).Text = Replace(e.Row.Cells(13).Text, Chr(10), "<br>")

                    If InStr(e.Row.Cells(14).Text, "[1]") > 0 Then
                        e.Row.Cells(13).ForeColor = Color.Red
                        e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(14).Text, "[]") > 0 Then
                            e.Row.Cells(13).ForeColor = Color.Blue
                            e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[]", "")
                        End If
                    End If

                    'For i = 8 To 10
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    '
                    For i = 15 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False

                End If
            End If
            '
            '**代用ITEM
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Substitute"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "A1" & "<BR>" & "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 5 To 25
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'NIKE
        '=========================================================================
        If pBuyer = "000013" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "PCX"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "IM Code"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Base" & "<BR>" & "Size(inch)"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Base"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Add 1 inch"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "YPGOLD" & "<BR>" & "(OQG)"
                    tcl(9).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "YKKSHB" & "<BR>" & "(VKL)"
                    tcl(10).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "SHINY ICESIL" & "<BR>" & "(C5V)"
                    tcl(11).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "YKKH6N" & "<BR>" & "(H6N)"
                    tcl(12).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "01M" & "<BR>" & "(H1)"
                    tcl(13).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "O1" & "<BR>" & "OP-O(95K)" & "<BR>" & "OP-H6(95H)"
                    tcl(14).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "P1" & "<BR>" & "Sample L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Q1" & "<BR>" & "Bulk L/T"
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "R1" & "<BR>" & "MCO_zipper"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "S1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "T1" & "<BR>" & "YKK/IM Code"
                    tcl(19).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "U1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:00A僅供參考)"
                    tcl(20).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 23
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(21).Text, "[1]") > 0 Then
                        e.Row.Cells(20).ForeColor = Color.Red
                        e.Row.Cells(21).Text = Replace(e.Row.Cells(21).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(21).Text, "[]") > 0 Then
                            e.Row.Cells(20).ForeColor = Color.Blue
                            e.Row.Cells(21).Text = Replace(e.Row.Cells(21).Text, "[]", "")
                        End If
                    End If

                    e.Row.Cells(6).Visible = False
                    '
                    'For i = 6 To 15
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 22 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
            '
            '**代用ITEM
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Substitute"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "A1" & "<BR>" & "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 5 To 25
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'COLUMBIA
        '=========================================================================
        If pBuyer = "000003" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Brand"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Status"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Kids Safe"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "PDM"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "YKK Dsc"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Base Length"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "L/T" & "<BR>" & "SPL w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "L/T" & "<BR>" & "SPL w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "L/T" & "<BR>" & "PROD w/o"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "L/T" & "<BR>" & "PROD w/"
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "Unit"
                    tcl(12).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "Currency"
                    tcl(13).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "O1" & "<BR>" & "PROD Loc."
                    tcl(14).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "P1" & "<BR>" & "Base"
                    tcl(15).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Q1" & "<BR>" & "Add 1 inch"
                    tcl(16).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(17).Text = "R1" & "<BR>" & "MOQ" & "<BR>" & "SPL/PROD"
                    tcl.Add(New TableHeaderCell())
                    tcl(18).Text = "S1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(19).Text = "T1" & "<BR>" & "YKK/PDM"
                    tcl(19).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(20).Text = "U1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                    tcl(20).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    e.Row.Cells(3).ForeColor = Color.Red

                    For i = 0 To 23
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(20).Text, "[1]") > 0 Then
                        e.Row.Cells(19).ForeColor = Color.Red
                        e.Row.Cells(20).Text = Replace(e.Row.Cells(20).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(20).Text, "[]") > 0 Then
                            e.Row.Cells(19).ForeColor = Color.Blue
                            e.Row.Cells(20).Text = Replace(e.Row.Cells(20).Text, "[]", "")
                        End If
                    End If
                    '
                    'e.Row.Cells(7).Visible = False
                    'For i = 13 To 17
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 21 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
            '
            '**代用ITEM
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Web Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Substitute"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "A1" & "<BR>" & "Article"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 5 To 25
                        e.Row.Cells(2).ForeColor = Color.Red
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'UA
        '=========================================================================
        If pBuyer = "TW0371" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "ZP #"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Description"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "YKK" & "<BR>" & "Description"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "CPSC"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "Length" & "<BR>" & "(Inch)"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "Price"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                    tcl(11).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "Remark"

                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "YKK/ZP"
                    tcl(13).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "O1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                    tcl(14).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    e.Row.Cells(3).ForeColor = Color.Red

                    For i = 0 To 23
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(14).Text, "[1]") > 0 Then
                        e.Row.Cells(13).ForeColor = Color.Red
                        e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(14).Text, "[]") > 0 Then
                            e.Row.Cells(13).ForeColor = Color.Blue
                            e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[]", "")
                        End If
                    End If
                    '
                    'For i = 9 To 11
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 15 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
            '
            '**代用ITEM
            If AtReplace.Checked = True Then
                '
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "ItemType"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "#"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Category"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Description"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Rebranding" & "<BR>" & "ZP #"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Rebraning" & "<BR>" & "YKK Code"
                    tcl(6).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Previous" & "<BR>" & "ZP #"
                    tcl(7).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Previous" & "<BR>" & "YKK Code"
                    tcl(8).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "Length" & "<BR>" & "(Inch)"
                    tcl(11).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "Price"
                    tcl(12).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "1(Inch)" & "<BR>" & "Upcharge"
                    tcl(13).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(14).Text = "O1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(15).Text = "P1" & "<BR>" & "YKK/PDM"
                    tcl(15).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(16).Text = "Q1" & "<BR>" & "Item Name" & "<BR>" & "(Puller:BLACK僅供參考)"
                    tcl(16).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then

                    e.Row.Cells(5).ForeColor = Color.Red

                    For i = 0 To 23
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    If InStr(e.Row.Cells(16).Text, "[1]") > 0 Then
                        e.Row.Cells(15).ForeColor = Color.Red
                        e.Row.Cells(16).Text = Replace(e.Row.Cells(16).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(16).Text, "[]") > 0 Then
                            e.Row.Cells(15).ForeColor = Color.Blue
                            e.Row.Cells(16).Text = Replace(e.Row.Cells(16).Text, "[]", "")
                        End If
                    End If
                    '
                    'For i = 11 To 13
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    For i = 17 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'HELLY HANSEN
        '=========================================================================
        If pBuyer = "000098" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Prod"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())

                    tcl(2).Text = "C1" & "<BR>" & "Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())

                    tcl(4).Text = "E1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(4).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(5).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Remark"

                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "YKK"
                    tcl(7).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Item Name"
                    tcl(8).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 23
                        If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                            e.Row.Cells(7).ForeColor = Color.Red
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                        Else
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                        End If
                        '
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    '
                    For i = 9 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'BURTON
        '=========================================================================
        If pBuyer = "000068" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Prod"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())

                    tcl(2).Text = "C1" & "<BR>" & "Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Description"
                    tcl.Add(New TableHeaderCell())

                    tcl(4).Text = "E1" & "<BR>" & "SMS" & "<BR>" & "Lead Time"
                    tcl(4).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "BULK" & "<BR>" & "Lead Time"
                    tcl(5).BackColor = Color.Blue

                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Remark"

                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "YKK"
                    tcl(7).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Item Name"
                    tcl(8).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 23
                        If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                            e.Row.Cells(7).ForeColor = Color.Red
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                        Else
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                        End If
                        '
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    '
                    For i = 9 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'HERSCHEL
        '=========================================================================
        If pBuyer = "TW5068" Then
            '
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Category"
                    tcl(1).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Part #"
                    tcl(2).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Size"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Matereial Spec"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Color Name"
                    tcl(5).BackColor = Color.Green
                    '
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Supplier Item"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Unit Of Measure"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "FOB USD"
                    tcl(8).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "MOQ"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "Remark"
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    For i = 0 To 23
                        If InStr(e.Row.Cells(8).Text, "[1]") > 0 Then
                            e.Row.Cells(7).ForeColor = Color.Red
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[1]", "")
                        Else
                            e.Row.Cells(8).Text = Replace(e.Row.Cells(8).Text, "[]", "")
                        End If
                        '
                        e.Row.Cells(i).Text = Replace(e.Row.Cells(i).Text, Chr(10), "<br>")
                    Next
                    '
                    '調整特定欄位寬度
                    e.Row.Cells(11).Width = 300
                    '
                    'e.Row.Cells(8).Visible = False
                    For i = 12 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '=========================================================================
        'VERA BRADLEY
        '=========================================================================
        If pBuyer = "TW0655" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Cat."
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Class"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Code"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Material Name"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "YKK NAME"
                    tcl(5).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Length(inch)"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Price(USD)"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Unit"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "1 inch(USD)"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "1inch Unit"
                    tcl(10).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    'tcl(13).Text = "M1" & "<BR>" & "SPPrice"
                    'tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "N1" & "<BR>" & "YKK/Web Code"
                    tcl(12).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "O1" & "<BR>" & "YKK/Web Name"
                    tcl(13).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    '
                    If InStr(e.Row.Cells(14).Text, "[1]") > 0 Then
                        e.Row.Cells(13).ForeColor = Color.Red
                        e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[1]", "")
                    Else
                        If InStr(e.Row.Cells(14).Text, "[]") > 0 Then
                            e.Row.Cells(13).ForeColor = Color.Blue
                            e.Row.Cells(14).Text = Replace(e.Row.Cells(14).Text, "[]", "")
                        End If
                    End If

                    'For i = 6 To 10
                    '    e.Row.Cells(i).Visible = False
                    'Next
                    e.Row.Cells(12).Visible = False
                    '
                    For i = 15 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If

        '=========================================================================
        'PATAGONIA
        '=========================================================================
        If pBuyer = "000141" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "RM#"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Material Name"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "YKK Item Code"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "YKK/Web Code"
                    tcl(6).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "YKK/Web Name"
                    tcl(7).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    '
                    If InStr(e.Row.Cells(7).Text, "[1]") > 0 Then
                        e.Row.Cells(6).ForeColor = Color.Red
                        e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, "[1]", "")
                    End If
                    If InStr(e.Row.Cells(7).Text, "[]") > 0 Then
                        e.Row.Cells(6).ForeColor = Color.Blue
                        e.Row.Cells(7).Text = Replace(e.Row.Cells(7).Text, "[]", "")
                    End If
                    For i = 8 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If

        '=========================================================================
        'ARCTERYX
        '=========================================================================
        If pBuyer = "000053" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "TAS Product"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Centric Code"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Finish"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "YKK Code"
                    tcl(5).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Description"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Import location of YKK"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Lead Time"
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "Remark"
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "YKK/Web Code"
                    tcl(10).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "YKK/Web Name"
                    tcl(11).BackColor = Color.Red
                End If
                '
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    '
                    If InStr(e.Row.Cells(11).Text, "[1]") > 0 Then
                        e.Row.Cells(10).ForeColor = Color.Red
                        e.Row.Cells(11).Text = Replace(e.Row.Cells(11).Text, "[1]", "")
                    End If
                    If InStr(e.Row.Cells(11).Text, "[]") > 0 Then
                        e.Row.Cells(10).ForeColor = Color.Blue
                        e.Row.Cells(11).Text = Replace(e.Row.Cells(11).Text, "[]", "")
                    End If
                    e.Row.Cells(5).Text = Replace(e.Row.Cells(5).Text, Chr(10), "<br>")
                    e.Row.Cells(10).Text = Replace(e.Row.Cells(10).Text, Chr(10), "<br>")
                    For i = 12 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If

        '=========================================================================
        'SALOMON
        '=========================================================================
        If pBuyer = "000055" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Plasma"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Centric Code"
                    tcl(3).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "Buyer Item Description"
                    tcl(4).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "YKK Item Description"
                    tcl(5).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Lead time(W/I)"
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "Lead time(W/O)"
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "Remark"
                End If
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    '
                    For i = 9 To 25
                        e.Row.Cells(i).Visible = False
                    Next
                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If

        '=========================================================================
        'TUMI
        '=========================================================================
        If pBuyer = "TW0020" Then
            '**BUYERITEM
            If AtBuyer.Checked = True Then
                'Header
                If (e.Row.RowType = DataControlRowType.Header) Then
                    Dim tcl As TableCellCollection = e.Row.Cells
                    tcl.Clear() '清除自动生成的表头
                    '添加新的表头第一行
                    tcl.Add(New TableHeaderCell())
                    tcl(0).Text = "A1" & "<BR>" & "Season"
                    tcl.Add(New TableHeaderCell())
                    tcl(1).Text = "B1" & "<BR>" & "Type"
                    tcl.Add(New TableHeaderCell())
                    tcl(2).Text = "C1" & "<BR>" & "Series"
                    tcl.Add(New TableHeaderCell())
                    tcl(3).Text = "D1" & "<BR>" & "Slider/Chain Name"
                    tcl.Add(New TableHeaderCell())
                    tcl(4).Text = "E1" & "<BR>" & "TUMI Description"
                    tcl(4).BackColor = Color.Green
                    tcl.Add(New TableHeaderCell())
                    tcl(5).Text = "F1" & "<BR>" & "Item Code"
                    tcl(5).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(6).Text = "G1" & "<BR>" & "Item Name"
                    tcl(6).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(7).Text = "H1" & "<BR>" & "TUMI Color"
                    tcl(7).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(8).Text = "I1" & "<BR>" & "SMS Lead Time"
                    tcl(8).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(9).Text = "J1" & "<BR>" & "BULK Lead Time"
                    tcl(9).BackColor = Color.Blue
                    tcl.Add(New TableHeaderCell())
                    tcl(10).Text = "K1" & "<BR>" & "Unit"
                    tcl(10).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(11).Text = "L1" & "<BR>" & "NT$"
                    tcl(11).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(12).Text = "M1" & "<BR>" & "US$"
                    tcl(12).BackColor = Color.Red
                    tcl.Add(New TableHeaderCell())
                    tcl(13).Text = "N1" & "<BR>" & "Remark"
                End If
                'DataRow
                If (e.Row.RowType = DataControlRowType.DataRow) Then
                    '
                    e.Row.Cells(13).Visible = False
                    For i = 15 To 25
                        e.Row.Cells(i).Visible = False
                    Next

                    e.Row.Cells(27).Visible = False
                End If
            End If
        End If
        '
        '---------------------------------------
        '
        'Footer
        If e.Row.RowType = DataControlRowType.Footer Then
        End If

    End Sub
    '
    '*****************************************************************
    '**
    '**     CHECKBOX CHANGE
    '**
    '*****************************************************************
    Protected Sub AtBuyer_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtBuyer.CheckedChanged
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtSPD.Checked = False
        AtIRW.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            AtBuyer.Checked = True
        End If
    End Sub

    Protected Sub AtFAS_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtFAS.CheckedChanged
        AtBuyer.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtSPD.Checked = False
        AtIRW.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            AtFAS.Checked = True
        End If
    End Sub

    Protected Sub AtReplace_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtReplace.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtSales.Checked = False
        AtSPD.Checked = False
        AtIRW.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            AtReplace.Checked = True
        End If
    End Sub

    Protected Sub AtSales_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtSales.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSPD.Checked = False
        AtIRW.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            AtSales.Checked = True
        End If
    End Sub

    Protected Sub AtSPD_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtSPD.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtIRW.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtIRW.Checked = False Then
            AtSPD.Checked = True
        End If
    End Sub

    Protected Sub AtIRW_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles AtIRW.CheckedChanged
        AtBuyer.Checked = False
        AtFAS.Checked = False
        AtReplace.Checked = False
        AtSales.Checked = False
        AtSPD.Checked = False
        '
        If AtBuyer.Checked = False And AtFAS.Checked = False And AtReplace.Checked = False And AtSales.Checked = False And AtSPD.Checked = False And AtSPD.Checked = False Then
            AtIRW.Checked = True
        End If
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel
    '**
    '*****************************************************************
    Protected Sub BExcel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        '
        Response.Clear()
        Response.Buffer = True

        Response.AppendHeader("Content-Disposition", "attachment;filename=ItemList.xls")     '程式別不同
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        '
        ShowData()
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()
    End Sub
    '
    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub
    '
    '---------------------------------------------------------------------------------------------------

End Class
