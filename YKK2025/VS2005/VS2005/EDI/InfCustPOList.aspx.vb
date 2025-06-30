Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfCustPOList
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim xBuyer As String            'Buyer
    Dim xUserID As String           'UserID
    Dim xType As String             'Type
    Dim xPO As String               'PO
    Dim xSeq As String              'Seqno
    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oWaves As New Waves.CommonService

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        If Not IsPostBack Then                      'PostBack
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
        Response.Cookies("PGM").Value = "InfCustPOList.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        xBuyer = Request.QueryString("pBuyer")
        xUserID = Request.QueryString("pUserID")
        xType = Request.QueryString("pType")
        xPO = Request.QueryString("pPO")
        xSeq = Request.QueryString("pSeq")
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     資料顯示
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String
        '
        If xType = "P" Then        ' GRBuyer(2)=P (使用自動PO) 
            sql = "SELECT "
            sql &= "A1 + ' / ' + B1 + ' / ' + C1 + ' / ' + D1 + ' / ' + E1 + ' / ' + "
            sql &= "F1 + ' / ' + G1 + ' / ' + H1 + ' / ' + I1 + ' / ' + J1 + ' / ' + "
            sql &= "K1 + ' / ' + L1 + ' / ' + M1 + ' / ' + N1 + ' / ' + O1 + ' / ' + "
            sql &= "P1 + ' / ' + Q1 + ' / ' + R1 + ' / ' + S1 + ' / ' + T1 + ' / ' + "
            sql &= "U1 + ' / ' + V1 + ' / ' + W1 + ' / ' + X1 + ' / ' + Y1 + ' / ' + "
            sql &= "Z1 + ' / ' "
            sql &= "as OrderInf, "
            sql &= "BM1 as PO, BN1 as Seq "
            sql &= "From E_InputSheet "
            sql &= "Where Buyer = '" & xBuyer & "' "
            sql &= "  And BM1 = '" & xPO & "' "
            sql &= "  And BN1 = " & xSeq & " "
            sql &= " Order by BM1, Convert(int,BN1) "
        Else
            sql = "SELECT "
            sql &= "A1 + ' / ' + B1 + ' / ' + C1 + ' / ' + D1 + ' / ' + E1 + ' / ' + "
            sql &= "F1 + ' / ' + G1 + ' / ' + H1 + ' / ' + I1 + ' / ' + J1 + ' / ' + "
            sql &= "K1 + ' / ' + L1 + ' / ' + M1 + ' / ' + N1 + ' / ' + O1 + ' / ' + "
            sql &= "P1 + ' / ' + Q1 + ' / ' + R1 + ' / ' + S1 + ' / ' + T1 + ' / ' + "
            sql &= "U1 + ' / ' + V1 + ' / ' + W1 + ' / ' + X1 + ' / ' + Y1 + ' / ' + "
            sql &= "Z1 + ' / ' "
            sql &= "as OrderInf, "
            sql &= "E1 as PO, D1 as Seq "
            sql &= "From E_InputSheet "
            sql &= "Where Buyer = '" & xBuyer & "' "
            sql &= "  And E1 = '" & xPO & "' "
            sql &= "  And D1 = " & xSeq & " "
            sql &= " Order by E1, Convert(int,D1) "
        End If

        '
        GridView1.DataSource = uDataBase.GetDataTable(sql)
        GridView1.DataBind()

    End Sub

End Class
