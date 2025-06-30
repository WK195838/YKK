Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing

Partial Class InfF_VDPErrorMaint
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    '全域變數
    Dim NowDateTime As String       '現在日期時間
    Dim NowDate As String           '現在日期時間
    Dim Buyer As String             'Buyer
    Dim UserID As String            'UserID

    '外部Object
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     主程式
    '**
    '*****************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetParameter()                              '設定參數
        SetPopupFunction()                          '設定彈出視窗事件

        If Not IsPostBack Then                      'PostBack
            SetDefaultValue()                       '設定初值
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
        Server.ScriptTimeout = 900                                  '設定逾時時間
        Response.Cookies("PGM").Value = "InfF_VDPErrorMaint.aspx"     '程式名
        '-----------------------------------------------------------------
        '-- 共用參數
        '-----------------------------------------------------------------
        NowDateTime = Now.ToString("yyyy/MM/dd HH:mm:ss")   '現在日期時間
        NowDate = Now.ToString("yyyyMMdd")                  '現在日期時間
        Buyer = Request.QueryString("pBuyer")
        If Buyer <> "000013T" And Buyer <> "TW0371T" Then
            Buyer = Mid(Request.QueryString("pBuyer"), 1, 6)
        End If

        UserID = Request.QueryString("pUserID")             'UserID
        '-----------------------------------------------------------------
        '-- 初值
        '-----------------------------------------------------------------
        GridView1.Visible = False
        '-----------------------------------------------------------------
        'it003	徐滿霖
        'mk023	楊惠文
        'mk027	吳馥如
        'mk034	周姵伶, 王舒怡
        'sl035	曾姿栩
        'sl038	王孝軒
        '-----------------------------------------------------------------
        If UCase(UserID) = "IT003" Or UCase(UserID) = "MK023" Or UCase(UserID) = "MK027" Or UCase(UserID) = "MK034" Or UCase(UserID) = "MK042" Or _
           UCase(UserID) = "SL035" Or UCase(UserID) = "SL038" Or UCase(UserID) = "MK005" Or UCase(UserID) = "MK045" Then
            DKey1.Visible = True
            DKey2.Visible = True
            BInq.Visible = True
        Else
            DKey1.Visible = False
            DKey2.Visible = False
            BInq.Visible = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetPopupFunction)
    '**     設定彈出視窗事件
    '**
    '*****************************************************************
    Sub SetPopupFunction()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(SetDefaultValue)
    '**     設定初值
    '**
    '*****************************************************************
    Sub SetDefaultValue()
        DKey1.Text = ""
        DKey2.Text = ""
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ShowData)
    '**     顯示資料
    '**
    '*****************************************************************
    Sub ShowData()
        Dim sql As String
        ' 取得資料
        sql = "SELECT "
        sql &= "Case When CompletedFlag = '1' Then '已出貨' else '未出貨' End As Completed, Finished, ModifyUser, Month, "
        sql &= "Unique_ID, CustWavesCode, OrderNo, Season, Year, PLM, BUYMONTH, Style, TapeColor "
        sql &= "From E_VDPErrorList "
        sql &= "Where Buyer = '" & Buyer & "' "
        '
        '-----------------------------------------------------------------
        'it003	徐滿霖
        'mk023	楊惠文
        'mk027	吳馥如
        'mk034	周姵伶
        'sl035	曾姿栩
        'sl038	王孝軒
        '-----------------------------------------------------------------
        If UCase(UserID) = "IT003" Or UCase(UserID) = "MK023" Or UCase(UserID) = "MK027" Or UCase(UserID) = "MK034" Or UCase(UserID) = "MK042" Or _
           UCase(UserID) = "SL035" Or UCase(UserID) = "SL038" Or UCase(UserID) = "MK005" Or UCase(UserID) = "MK045" Then
            If DKey1.Text <> "" Then
                sql &= " And ( CustWavesCode Like '%" & DKey1.Text & "%' "
                sql &= "    Or OrderNo Like '%" & DKey1.Text & "%' "
                sql &= "    Or Month Like '%" & DKey1.Text & "%' ) "
            End If
            If DKey2.Text <> "" Then
                sql &= " And ( CustWavesCode Like '%" & DKey2.Text & "%' "
                sql &= "    Or OrderNo Like '%" & DKey2.Text & "%' "
                sql &= "    Or Month Like '%" & DKey2.Text & "%' ) "
            End If
        Else
            If UCase(UserID) = "" Then
                sql &= " And Finished = '" & "W" & "' "
                sql &= " And CompletedFlag = '" & "1" & "' "
                sql &= " And ModifyUser = '" & "NOTHING" & "' "
            Else
                sql &= " And Finished = '" & "W" & "' "
                sql &= " And CompletedFlag = '" & "1" & "' "
                sql &= " And ModifyUser = '" & UserID & "' "
            End If
        End If
        sql &= "Order by CustWavesCode, OrderNo, Season, Year, PLM, BUYMONTH, Style, TapeColor "
        '
        Dim dt_ErrorList As DataTable = uDataBase.GetDataTable(sql)
        If dt_ErrorList.Rows.Count > 0 Then
            GridView1.Visible = True
            GridView1.DataSource = dt_ErrorList
            GridView1.DataBind()
        Else
            uJavaScript.PopMsg(Me, "搜尋不到資料,請確認!!")
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RowCancelingEdit)
    '**     GridView 取消編輯
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles GridView1.RowCancelingEdit
        GridView1.EditIndex = -1
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RowCreated)
    '**     GridView 隱藏欄位
    '**
    '*****************************************************************
    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If UCase(UserID) = "IT003" Or UCase(UserID) = "MK023" Or UCase(UserID) = "SL038" Or UCase(UserID) = "MK005" Or UCase(UserID) = "MK045" Then
        Else
            e.Row.Cells(4).Visible = False
            e.Row.Cells(5).Visible = False
            e.Row.Cells(6).Visible = False
            e.Row.Cells(13).Visible = False
        End If
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RowEditing)
    '**     GridView 編輯
    '**
    '*****************************************************************
    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing
        GridView1.EditIndex = e.NewEditIndex
        ShowData()
    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(RowUpdating)
    '**     GridView 更新資料
    '**
    '*****************************************************************
    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles GridView1.RowUpdating
        Dim sql As String
        '
        Dim GVRow As GridViewRow = Me.GridView1.Rows(e.RowIndex)
        Dim xSeason As String = CType(GVRow.Cells(7).Controls(0), TextBox).Text
        Dim xYear As String = CType(GVRow.Cells(8).Controls(0), TextBox).Text
        Dim xBuyMonth As String = CType(GVRow.Cells(9).Controls(0), TextBox).Text
        Dim xPLM As String = CType(GVRow.Cells(10).Controls(0), TextBox).Text
        Dim xStyle As String = CType(GVRow.Cells(11).Controls(0), TextBox).Text
        Dim xTapeColor As String = CType(GVRow.Cells(12).Controls(0), TextBox).Text
        Dim xID As String = CType(GVRow.Cells(13).Controls(0), TextBox).Text
        '
        If UCase(UserID) = "IT003" Or UCase(UserID) = "MK023" Or UCase(UserID) = "SL038" Or UCase(UserID) = "MK005" Or UCase(UserID) = "MK045" Then
            Dim xFinished As String = CType(GVRow.Cells(4).Controls(0), TextBox).Text
            Dim xModifyUser As String = CType(GVRow.Cells(5).Controls(0), TextBox).Text
            '
            If Not IsNumeric(xYear) Or Not IsNumeric(xBuyMonth) Then
                uJavaScript.PopMsg(Me, "有無效數字資料,請確認!!")
            Else
                sql = "Update E_VDPErrorList Set "
                sql &= " Finished = '" & xFinished & "', "
                sql &= " ModifyUser = '" & xModifyUser & "', "
                sql &= " Season = '" & xSeason & "', "
                sql &= " Year = " & xYear & ", "
                sql &= " BuyMonth = " & xBuyMonth & ", "
                sql &= " PLM = '" & xPLM & "', "
                sql &= " Style = '" & xStyle & "', "
                sql &= " TapeColor = '" & xTapeColor & "', "
                sql &= " ModifyTime = '" & NowDateTime & "' "
                sql &= " Where Unique_ID = " & xID & " "
                uDataBase.ExecuteNonQuery(sql)
            End If
        Else
            If xSeason = "" Or xYear = "" Or xBuyMonth = "" Or xPLM = "" Or xStyle = "" Then
                uJavaScript.PopMsg(Me, "有無效空白資料,請確認!!")
            Else
                If Not IsNumeric(xYear) Or Not IsNumeric(xBuyMonth) Then
                    uJavaScript.PopMsg(Me, "有無效數字資料,請確認!!")
                Else
                    sql = "Update E_VDPErrorList Set "
                    sql &= " Season = '" & xSeason & "', "
                    sql &= " Year = " & xYear & ", "
                    sql &= " BuyMonth = " & xBuyMonth & ", "
                    sql &= " PLM = '" & xPLM & "', "
                    sql &= " Style = '" & xStyle & "', "
                    sql &= " TapeColor = '" & xTapeColor & "', "
                    sql &= " Finished = '" & "Y" & "', "
                    sql &= " ModifyTime = '" & NowDateTime & "' "
                    sql &= " Where Unique_ID = " & xID & " "
                    uDataBase.ExecuteNonQuery(sql)
                End If
            End If
        End If
        '
        GridView1.EditIndex = -1
        ShowData()
    End Sub
End Class
