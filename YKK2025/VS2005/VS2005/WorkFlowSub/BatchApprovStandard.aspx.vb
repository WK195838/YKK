Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

Partial Class BatchApprovStandard
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim pFormNo As String
    Dim wLevel As String = ""
    Dim wDivision As String = ""
    Dim wUserID As String = ""
    Dim wName As String = ""
    Dim NowDateTime As String       '現在日期時間
    Dim fpObj As New ForProject
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim oCommon As New Common.CommonService
    Dim ItemCount As Integer = 3 '預先定義欄位數量
    Dim wUserIP As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "BatchApprovStandard.aspx"
        SetParameter()          '設定共用參數
        If Not Me.IsPostBack Then
            CheckAuthority()
            DataList()
        End If

    End Sub
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     設定共用參數
    '**
    '*****************************************************************
    Sub SetParameter()
        NowDateTime = CStr(Now.Date) + " " + _
                          CStr(Now.Hour) + ":" + _
                          CStr(Now.Minute) + ":" + _
                          CStr(Now.Second)     '現在日時
        pFormNo = Request.QueryString("pFormNo")    '表單號碼
        wUserID = Request.QueryString("pUserID")    'UserID
        wUserIP = Request.ServerVariables("REMOTE_ADDR")
    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        SQL = "Select * From V_BatchApprovStandard "
        SQL = SQL & "Where FormNo = '" & pFormNo & "' "
        SQL = SQL + "And Decideid = '" & wUserID & "' "
        SQL = SQL & "Order By FormNo, FormSno, Step, Seqno "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        If DBAdapter1.Rows.Count > 0 Then
            BOK.Enabled = True
        Else
            BOK.Enabled = False
        End If
        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BOK.Click
        Dim i As Integer
        Dim wError As Integer = 0
        Dim Message As String = ""
        Dim wNo As String = ""
        Dim ErrCode As Integer = 0
        '
        Dim wFun As String = ""
        Dim wAction As Integer = 0
        Dim wSts As String = ""
        Dim wStsDes As String = ""
        Dim wFormNo As String = ""
        Dim wFormSno As Integer = 0
        Dim wStep As Integer = 0
        Dim wSeqNo As Integer = 0
        Dim wDecideCalendar As String = ""
        Dim wDecideID As String = ""
        Dim wApplyID As String = ""
        Dim wAgentID As String = ""
        Dim wAllocteID As String = ""
        Dim wDecideDesc As String = ""
        Dim wTableName As String = ""
        '
        For i = 0 To Me.GridView1.Rows.Count - 1 Step i + 1
            'pFun		    pFun=OK, NG1, NG2, SAVE  
            'pAction	    pAction=0:OK, 1:NG1, 2:NG2, 3:Save
            'pSts		    pSts=0:未處理, 1:OK, 2:NG1, 3:NG2, 4:已閱讀, 5:被抽單
            'pStsDes	    pSts說明
            If (CType(GridView1.Rows(i).FindControl("CheckYESDETAIL"), CheckBox)).Checked Then
                wFun = "OK"
                wAction = 0
                wSts = "1"
                wStsDes = "OK"
            ElseIf (CType(GridView1.Rows(i).FindControl("CheckNODETAIL"), CheckBox)).Checked Then
                wFun = "NG1"
                wAction = 1
                wSts = "2"
                wStsDes = "NG"
            End If
            'pFormNo	    表單代號
            'pFormSno	    單號
            'pStep		    工程代號
            'pSeqNo		    序號
            wFormNo = Me.GridView1.Rows(i).Cells(29).Text
            wFormSno = CLng(Me.GridView1.Rows(i).Cells(30).Text)
            wStep = CLng(Me.GridView1.Rows(i).Cells(31).Text)
            wSeqNo = CLng(Me.GridView1.Rows(i).Cells(32).Text)
            'pDecideID	    簽核者	
            'pDecideCalendar行事曆
            'pApplyID	    申請者
            'pAgentID	    被代理者
            'pAllocteID	    指定簽核者
            'pDecideDesc    承認說明
            'pTableName     表單Table
            wDecideID = Me.GridView1.Rows(i).Cells(34).Text
            wDecideCalendar = oCommon.GetCalendarGroup(wDecideID)
            wApplyID = Me.GridView1.Rows(i).Cells(33).Text
            wAgentID = ""
            wAllocteID = ""
            Dim txt As TextBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("TextBox1"), TextBox)
            wDecideDesc = txt.Text
            wTableName = Me.GridView1.Rows(i).Cells(35).Text
            '
            'pNo    表單No
            wNo = Me.GridView1.Rows(i).Cells(16).Text
            'CheckBox
            Dim checkBoxYES As CheckBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("CheckYESDETAIL"), CheckBox)
            Dim checkBoxNO As CheckBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("CheckNODETAIL"), CheckBox)
            '
            If wDecideDesc <> "" And (checkBoxYES.Checked = True Or checkBoxNO.Checked = True) Then
                '採用AgentApprov暫存檔
                ErrCode = oCommon.RunAgentApprov(wFun, _
                                                    wAction, _
                                                    wSts, _
                                                    wStsDes, _
                                                    wFormNo, _
                                                    wFormSno, _
                                                    wStep, _
                                                    wSeqNo, _
                                                    wDecideCalendar, _
                                                    wDecideID, _
                                                    wApplyID, _
                                                    wAgentID, _
                                                    wAllocteID, _
                                                    wDecideDesc, _
                                                    wTableName, _
                                                    wUserIP)
                '--
                '表單資料
                If ErrCode <> 0 Then
                    wError = wError + 1
                    If Message = "" Then
                        Message = wNo & "(BAS)"
                    Else
                        Message = Message & "," & wNo & "(BAS)"
                    End If
                End If

            ElseIf wDecideDesc = "" And (checkBoxYES.Checked = True Or checkBoxNO.Checked = True) Then
                wError = wError + 1
                If Message = "" Then
                    Message = wNo
                Else
                    Message = Message + "," + wNo
                End If
            End If
        Next
        '
        If wError > 0 Then  '若沒有輸入却下原因
            Message = "異常：" + Message + "\n" + "需輸入簽核說明!"
            uJavaScript.PopMsg(Me, Message)
        Else
            BOK.Enabled = False
            '再回到待處理
            'Dim URL As String = uCommon.GetAppSetting("RedirectURL") & "?pUserID=" & Request.Cookies("UserID").Value
            'Response.Redirect(URL)
        End If
        DataList()
        'GridView1.DataSource = Nothing
    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Dim i As Integer
        Dim SQL As String
        '
        '表頭
        If e.Row.RowType = DataControlRowType.Header Then
            '固定隱藏
            For i = 29 To 35
                e.Row.Cells(i).Visible = False
            Next
            '動態隱藏
            For i = 1 To 26
                e.Row.Cells(i).Text = "@"
            Next
            SQL = "Select * From M_BatchApprovTask "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo = '" & pFormNo & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("A1").ToString = "1" Then e.Row.Cells(1).Text = dt.Rows(0)("A1Name").ToString
                If dt.Rows(0)("B1").ToString = "1" Then e.Row.Cells(2).Text = dt.Rows(0)("B1Name").ToString
                If dt.Rows(0)("C1").ToString = "1" Then e.Row.Cells(3).Text = dt.Rows(0)("C1Name").ToString
                If dt.Rows(0)("D1").ToString = "1" Then e.Row.Cells(4).Text = dt.Rows(0)("D1Name").ToString
                If dt.Rows(0)("E1").ToString = "1" Then e.Row.Cells(5).Text = dt.Rows(0)("E1Name").ToString
                If dt.Rows(0)("F1").ToString = "1" Then e.Row.Cells(6).Text = dt.Rows(0)("F1Name").ToString
                If dt.Rows(0)("G1").ToString = "1" Then e.Row.Cells(7).Text = dt.Rows(0)("G1Name").ToString
                If dt.Rows(0)("H1").ToString = "1" Then e.Row.Cells(8).Text = dt.Rows(0)("H1Name").ToString
                If dt.Rows(0)("I1").ToString = "1" Then e.Row.Cells(9).Text = dt.Rows(0)("I1Name").ToString
                If dt.Rows(0)("J1").ToString = "1" Then e.Row.Cells(10).Text = dt.Rows(0)("J1Name").ToString
                '
                If dt.Rows(0)("K1").ToString = "1" Then e.Row.Cells(11).Text = dt.Rows(0)("K1Name").ToString
                If dt.Rows(0)("L1").ToString = "1" Then e.Row.Cells(12).Text = dt.Rows(0)("L1Name").ToString
                If dt.Rows(0)("M1").ToString = "1" Then e.Row.Cells(13).Text = dt.Rows(0)("M1Name").ToString
                If dt.Rows(0)("N1").ToString = "1" Then e.Row.Cells(14).Text = dt.Rows(0)("N1Name").ToString
                If dt.Rows(0)("O1").ToString = "1" Then e.Row.Cells(15).Text = dt.Rows(0)("O1Name").ToString
                If dt.Rows(0)("P1").ToString = "1" Then e.Row.Cells(16).Text = dt.Rows(0)("P1Name").ToString
                If dt.Rows(0)("Q1").ToString = "1" Then e.Row.Cells(17).Text = dt.Rows(0)("Q1Name").ToString
                If dt.Rows(0)("R1").ToString = "1" Then e.Row.Cells(18).Text = dt.Rows(0)("R1Name").ToString
                If dt.Rows(0)("S1").ToString = "1" Then e.Row.Cells(19).Text = dt.Rows(0)("S1Name").ToString
                If dt.Rows(0)("T1").ToString = "1" Then e.Row.Cells(20).Text = dt.Rows(0)("T1Name").ToString
                '
                If dt.Rows(0)("U1").ToString = "1" Then e.Row.Cells(21).Text = dt.Rows(0)("U1Name").ToString
                If dt.Rows(0)("V1").ToString = "1" Then e.Row.Cells(22).Text = dt.Rows(0)("V1Name").ToString
                If dt.Rows(0)("W1").ToString = "1" Then e.Row.Cells(23).Text = dt.Rows(0)("W1Name").ToString
                If dt.Rows(0)("X1").ToString = "1" Then e.Row.Cells(24).Text = dt.Rows(0)("X1Name").ToString
                If dt.Rows(0)("Y1").ToString = "1" Then e.Row.Cells(25).Text = dt.Rows(0)("Y1Name").ToString
                If dt.Rows(0)("Z1").ToString = "1" Then e.Row.Cells(26).Text = dt.Rows(0)("Z1Name").ToString
            End If
            For i = 1 To 26
                If e.Row.Cells(i).Text = "@" Then
                    e.Row.Cells(i).Visible = False
                End If
            Next
        End If
        '
        '資料
        If e.Row.RowType = DataControlRowType.DataRow Then
            '固定隱藏
            For i = 29 To 35
                e.Row.Cells(i).Visible = False
            Next
            '
            '動態隱藏
            For i = 1 To 28
                If e.Row.Cells(i).Text = "@" Then
                    e.Row.Cells(i).Visible = False
                End If
            Next
        End If
        '
        '表尾
        If e.Row.RowType = DataControlRowType.Footer Then
            '固定隱藏
            For i = 29 To 35
                e.Row.Cells(i).Visible = False
            Next
            '動態隱藏
            For i = 1 To 26
                e.Row.Cells(i).Text = "@"
            Next
            SQL = "Select * From M_BatchApprovTask "
            SQL = SQL & " Where Active = 1 "
            SQL = SQL & "   And FormNo = '" & pFormNo & "'"
            Dim dt As DataTable = uDataBase.GetDataTable(SQL)
            If dt.Rows.Count > 0 Then
                If dt.Rows(0)("A1").ToString = "1" Then e.Row.Cells(1).Text = dt.Rows(0)("A1Name").ToString
                If dt.Rows(0)("B1").ToString = "1" Then e.Row.Cells(2).Text = dt.Rows(0)("B1Name").ToString
                If dt.Rows(0)("C1").ToString = "1" Then e.Row.Cells(3).Text = dt.Rows(0)("C1Name").ToString
                If dt.Rows(0)("D1").ToString = "1" Then e.Row.Cells(4).Text = dt.Rows(0)("D1Name").ToString
                If dt.Rows(0)("E1").ToString = "1" Then e.Row.Cells(5).Text = dt.Rows(0)("E1Name").ToString
                If dt.Rows(0)("F1").ToString = "1" Then e.Row.Cells(6).Text = dt.Rows(0)("F1Name").ToString
                If dt.Rows(0)("G1").ToString = "1" Then e.Row.Cells(7).Text = dt.Rows(0)("G1Name").ToString
                If dt.Rows(0)("H1").ToString = "1" Then e.Row.Cells(8).Text = dt.Rows(0)("H1Name").ToString
                If dt.Rows(0)("I1").ToString = "1" Then e.Row.Cells(9).Text = dt.Rows(0)("I1Name").ToString
                If dt.Rows(0)("J1").ToString = "1" Then e.Row.Cells(10).Text = dt.Rows(0)("J1Name").ToString
                '
                If dt.Rows(0)("K1").ToString = "1" Then e.Row.Cells(11).Text = dt.Rows(0)("K1Name").ToString
                If dt.Rows(0)("L1").ToString = "1" Then e.Row.Cells(12).Text = dt.Rows(0)("L1Name").ToString
                If dt.Rows(0)("M1").ToString = "1" Then e.Row.Cells(13).Text = dt.Rows(0)("M1Name").ToString
                If dt.Rows(0)("N1").ToString = "1" Then e.Row.Cells(14).Text = dt.Rows(0)("N1Name").ToString
                If dt.Rows(0)("O1").ToString = "1" Then e.Row.Cells(15).Text = dt.Rows(0)("O1Name").ToString
                If dt.Rows(0)("P1").ToString = "1" Then e.Row.Cells(16).Text = dt.Rows(0)("P1Name").ToString
                If dt.Rows(0)("Q1").ToString = "1" Then e.Row.Cells(17).Text = dt.Rows(0)("Q1Name").ToString
                If dt.Rows(0)("R1").ToString = "1" Then e.Row.Cells(18).Text = dt.Rows(0)("R1Name").ToString
                If dt.Rows(0)("S1").ToString = "1" Then e.Row.Cells(19).Text = dt.Rows(0)("S1Name").ToString
                If dt.Rows(0)("T1").ToString = "1" Then e.Row.Cells(20).Text = dt.Rows(0)("T1Name").ToString
                '
                If dt.Rows(0)("U1").ToString = "1" Then e.Row.Cells(21).Text = dt.Rows(0)("U1Name").ToString
                If dt.Rows(0)("V1").ToString = "1" Then e.Row.Cells(22).Text = dt.Rows(0)("V1Name").ToString
                If dt.Rows(0)("W1").ToString = "1" Then e.Row.Cells(23).Text = dt.Rows(0)("W1Name").ToString
                If dt.Rows(0)("X1").ToString = "1" Then e.Row.Cells(24).Text = dt.Rows(0)("X1Name").ToString
                If dt.Rows(0)("Y1").ToString = "1" Then e.Row.Cells(25).Text = dt.Rows(0)("Y1Name").ToString
                If dt.Rows(0)("Z1").ToString = "1" Then e.Row.Cells(26).Text = dt.Rows(0)("Z1Name").ToString
            End If
            For i = 1 To 26
                If e.Row.Cells(i).Text = "@" Then
                    e.Row.Cells(i).Visible = False
                End If
            Next
            For i = 1 To 26
                e.Row.Cells(i).Text = ""
            Next
        End If
        '
        '其它
        '選擇列時變色 
        If e.Row.RowIndex > -1 Then
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='#e0e0ff'")
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c;")
        End If

        '---------------------------------------------------------
        'If e.Row.RowType = DataControlRowType.Footer Then
        '    e.Row.Cells(5).Text = "合計"
        'End If


        ''寫入超連結網址
        'If e.Row.RowType <> DataControlRowType.Header And e.Row.Cells(5).Text <> "合計" Then
        '    Dim h1 As New HyperLink
        '    h1.Text = e.Row.Cells(1).Text
        '    h1.Target = "_blank"
        '    h1.NavigateUrl = ("FundingSheet_01.aspx?pFormNo=003110" & "&pFormSno=" & CStr(e.Row.Cells(12).Text) & "&pStep=" & CStr(e.Row.Cells(15).Text) & "&pSeqNo=" & "1" & "&pApplyID=" & e.Row.Cells(13).Text & "&pUserID=" & wUserID)
        '    e.Row.Cells(1).Text = ""
        '    e.Row.Cells(1).Controls.Add(h1)
        'End If


        ''寫入超連結網址
        'If e.Row.Cells(13).Text = "..." Then
        '    Dim h1 As New HyperLink
        '    h1.Text = e.Row.Cells(13).Text
        '    h1.Target = "_blank"
        '    h1.NavigateUrl = ("http://10.245.1.10/WorkFlow/BefOPList.aspx?pFormNo=003110&pFormSno=" & CStr(e.Row.Cells(13).Text)) + "&pStep=1&pSeqNo=1&pApplyID=" + e.Row.Cells(14).Text
        '    e.Row.Cells(13).Text = ""
        '    e.Row.Cells(13).Controls.Add(h1)
        'End If

        '若勾則寫入OK
        'If e.Row.RowType = DataControlRowType.DataRow Then
        ' Dim checkBoxYES As CheckBox = CType(e.Row.FindControl("CheckYESDETAIL"), CheckBox)
        ' Dim checkBoxNO As CheckBox = CType(e.Row.FindControl("CheckNODETAIL"), CheckBox)
        'Dim textbox1 As TextBox = CType(e.Row.FindControl("textbox1"), TextBox)
        '  checkBoxYES.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='OK.';} else { document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxNO.ClientID + "').checked=false; }")
        '  checkBoxNO.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='NG.';} else {document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxYES.ClientID + "').checked=false;}")

        'End If
        '---------------------------------------------------------
    End Sub

    Protected Sub BRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BRefresh.Click
        DataList()
    End Sub
End Class
