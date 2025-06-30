Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class FindItemPage
    Inherits System.Web.UI.Page
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim NowDateTime As String       '現在日期時間
    Dim wUserID As String = ""      '申請者
    Dim wBuyer As String = ""       '申請者
    Dim wBuyerItem As String = ""   '申請者
    Dim wItemname As String = ""    'ItemName

    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim ConnectString As String = uCommon.GetAppSetting("WAVESOLEDB")
    Dim EDLConnect As String = uCommon.GetAppSetting("EDLDB")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Server.ScriptTimeout = 900 ' 設定逾時時間

        SetParameter()                  ' 設定共用參數
        If Not IsPostBack Then
            SetDefaultValue()                   '設定初值
            If DCode.Text <> "" Or DName1.Text <> "" Then
                DataList()
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetParameter)
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")     '現在日時
        wUserID = Request.QueryString("pUserID")
        wBuyer = Request.QueryString("pBuyer")
        wBuyerItem = Request.QueryString("pBuyerItem")
        wItemname = Request.QueryString("pItemName")
        Response.Cookies("PGM").Value = "FindItemPage.aspx"
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        Dim i, j As Integer
        Dim str, key As String
        Dim xItemName As String()
        '
        DCode.Text = ""
        DName1.Text = ""
        DName2.Text = ""
        DName3.Text = ""
        DName4.Text = ""
        '
        'ItemName(1)
        str = wItemname & " "
        xItemName = str.Split(" ")
        i = 0
        str = ""
        For Each str In xItemName
            Select Case i
                Case 3, 4, 5
                    If Mid(str, 1, 1) = "P" Then
                        For j = 0 To i
                            If DName1.Text = "" Then
                                DName1.Text = xItemName(j)
                            Else
                                DName1.Text = DName1.Text & " " & xItemName(j)
                            End If
                        Next
                        Exit For
                    End If
                Case Else
            End Select
            i = i + 1
        Next
        '
        'ItemName(2)(3)(4)
        For j = i + 1 To xItemName.Length - 1
            If j = i + 1 Then DName2.Text = xItemName(j)
            If j = i + 2 Then DName3.Text = xItemName(j)
            If j = i + 3 Then DName4.Text = xItemName(j)
        Next
    End Sub
    '------------------------------------------------------

    Protected Sub BFindItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFindItem.Click
        If DCode.Text <> "" Or DName1.Text <> "" Then
            DataList()
        End If
    End Sub

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView1.PageIndexChanging
        DataList()
        GridView1.PageIndex = e.NewPageIndex
        GridView1.DataBind()
    End Sub

    Sub DataList()
        Dim Sql As String = ""
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        '
        cn.ConnectionString = ConnectString
        Sql = "SELECT "
        Sql = Sql + "TRIM(ITMCA0) As Code, "
        Sql = Sql + "TRIM(IT1IA0) As Name1, "
        Sql = Sql + "TRIM(IT2IA0) As Name2, "
        Sql = Sql + "TRIM(IT3IA0) As Name3 "
        Sql = Sql + "From WAVEDLIB.FA000 "
        '
        If DCode.Text <> "" Then
            Sql = Sql + "Where ITMCA0 LIKE '%" & DCode.Text.ToUpper & "%' "
        Else
            Sql = Sql + "Where TRIM(IT1IA0) || TRIM(IT2IA0) || TRIM(IT3IA0) LIKE '%" & DName1.Text.ToUpper & "%' "
            If DName2.Text <> "" Then
                Sql = Sql + "And TRIM(IT1IA0) || TRIM(IT2IA0) || TRIM(IT3IA0) LIKE '%" & DName2.Text.ToUpper & "%' "
            End If
            If DName3.Text <> "" Then
                Sql = Sql + "And TRIM(IT1IA0) || TRIM(IT2IA0) || TRIM(IT3IA0) LIKE '%" & DName3.Text.ToUpper & "%' "
            End If
            If DName4.Text <> "" Then
                Sql = Sql + "And TRIM(IT1IA0) || TRIM(IT2IA0) || TRIM(IT3IA0) LIKE '%" & DName4.Text.ToUpper & "%' "
            End If
        End If
        Sql = Sql + "And NDPCA0 <> '1' "
        Sql = Sql + "ORDER BY RADUA0 DESC "
        '
        Dim DBAdapter1 As New OleDbDataAdapter(Sql, cn)
        DBAdapter1.Fill(ds, "FA000")
        If ds.Tables("FA000").Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = ds
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If

        cn.Close()
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged
        Dim cn As New OleDbConnection
        Dim ds As New DataSet
        Dim dc As New OleDbCommand
        Dim sql As String
        '
        Dim xCode As String = GridView1.SelectedRow.Cells(1).Text
        cn.ConnectionString = EDLConnect
        '
        If wBuyer <> "" And wUserID <> "" And xCode <> "" Then
            sql = "SELECT Item "
            sql = sql + "From T_Quotation "
            sql = sql + "Where Buyer = '" & wBuyer & "' "
            sql = sql + "And   Item = '" & xCode & "' "
            '
            Dim DBAdapter1 As New OleDbDataAdapter(sql, cn)
            DBAdapter1.Fill(ds, "Quotation")
            If ds.Tables("Quotation").Rows.Count <= 0 Then
                '
                sql = " insert into T_Quotation (Active, Buyer, BuyerItem, Item, CreateUser, CreateTime) "
                sql = sql + "values(0, '" + wBuyer + "','" + wBuyerItem + "','" + xCode + "','" + wUserID + "', getdate()" + ") "
                '
                dc.Connection = cn
                dc.CommandText = sql
                cn.Open()
                dc.ExecuteNonQuery()
                cn.Close()
                '
            End If
        End If

    End Sub
End Class
