Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb

 


Partial Class DASW_DISPOSALADMIN
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
    Dim ItemCount As Integer = 5 '預先定義欄位數量


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.Cookies("PGM").Value = "DASW_DISPOSALADMIN.aspx"
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
        wUserID = Request.QueryString("pUserID")      'UserID



    End Sub

    Sub CheckAuthority()
        Response.Cookies("UserID").Value = Request.QueryString("pUserID")
    End Sub

    Sub DataList()
        Dim SQL As String
        'SQL = "  select    no,dutydepo,disposaltype, case when customertoll =1 then 'V' else'' end as customertoll,"
        'SQL = SQL + " case when disposalreason='其它' then disposalreason+'-'+disposalreason1 else disposalreason end as disposalreason,"
        'SQL = SQL + " disposalrule,piece,Meter,yard,kg,price,a.formsno,ApplyID,decideID,step, "
        'SQL = SQL + " '...' as  History,appname,disposalfile1"
        'SQL = SQL + " from ("
        'SQL = SQL & "  select  no,dutydepo,disposaltype,customertoll,disposalreason,disposalreason1,disposalrule,piece,Meter,yard,kg,price,formsno,createuser,formno,appname,disposalfile1 from F_DISPOSALSheet "
        'SQL = SQL & " )a, ("
        'SQL = SQL & " select  formno ,formsno,step,applyid,decideID  from   t_waithandle  where formno = '006001'"
        'SQL = SQL & " and  decideID = '" + Request.QueryString("pUserID") + "'"
        ''SQL = SQL & " and  decideID = 'fas006'"
        'SQL = SQL & " and flowtype <> 0"
        '' SQL = SQL & "  and step in (10,30,40,80,90,100) 
        'SQL = SQL & " and sts in (0,4)"
        'SQL = SQL & " and active =1"
        'SQL = SQL & " )b where  a.formno =b.formno and a.formsno =b.formsno "
        'SQL = SQL & " order by a.formsno"


        SQL = "  select    no,dutydepo,disposaltype, case when customertoll =1 then 'V' else'' end as customertoll, "
        SQL = SQL + " case when disposalreason='其它' then disposalreason+'-'+disposalreason1 else disposalreason end as disposalreason,"
        SQL = SQL + " disposalrule,piece,Meter,yard,kg,price,a.formsno,ApplyID,decideID,step,  '...' as  History,appname,disposalfile1,STYPE from ( "
        SQL = SQL + " select a.no,a.dutydepo,a.disposaltype,a.customertoll,a.disposalreason,a.disposalreason1,a.disposalrule,a.formsno,a.createuser,a.formno,a.appname,a.disposalfile1,STYPE,"
        SQL = SQL + " sum(convert(decimal(15,2),piece))piece,"
        SQL = SQL + " sum(convert(decimal(15,2),meter))meter,"
        SQL = SQL + " sum(convert(decimal(15,2),yard))yard,"
        SQL = SQL + " sum(convert(decimal(15,2),kg))kg,"
        SQL = SQL + " sum(convert(decimal(12,2),price))price"
        SQL = SQL + " from ("
        SQL = SQL + " select  a.no,a.dutydepo,a.disposaltype,a.customertoll,a.disposalreason,a.disposalreason1,a.disposalrule,a.formsno,a.createuser,a.formno,a.appname,a.disposalfile1,STYPE,"
        SQL = SQL + " case when b.piece='' then '0' else  replace(b.piece,',','') end as   piece,"
        SQL = SQL + " case when b.meter='' then '0' else  replace(b.meter,',','') end as   meter,"
        SQL = SQL + " case when b.yard='' then '0' else  replace(b.yard,',','') end as   yard,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.kg,',','') end as   kg,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.price,',','') end as   price"
        SQL = SQL + " from F_DISPOSALSheet a,F_DISPOSALSheetdt b"
        SQL = SQL + " where(a.formsno = b.formsno)"
        SQL = SQL + " )a"
        SQL = SQL + " group by  a.no,a.dutydepo,a.disposaltype,a.customertoll,a.disposalreason,a.disposalreason1,a.disposalrule,a.formsno,a.createuser,a.formno,a.appname,a.disposalfile1,STYPE"
        SQL = SQL & "  )a, ( select  formno ,formsno,step,applyid,decideID  from   t_waithandle  where formno = '006001'"
        SQL = SQL & " and  decideID = '" + Request.QueryString("pUserID") + "'"
        'SQL = SQL & " and  decideID = 'fas006'"
        SQL = SQL & " and flowtype <> 0"
        ' SQL = SQL & "  and step in (10,30,40,80,90,100) 
        SQL = SQL & " and sts in (0,4)"
        SQL = SQL & " and active =1"
        SQL = SQL & " )b where  a.formno =b.formno and a.formsno =b.formsno "
        SQL = SQL & " order by a.formsno"


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
        Dim wFormSNo As String = ""
        Dim wSts As String = ""
        Dim wStep As String = ""
        Dim wDecideDesc As String = ""
        Dim wApplyID As String = ""
        Dim wDecideID As String = ""
        Dim wError As Integer = 0
        Dim Message As String = ""
        Dim wNo As String = ""


        '可以抓到該列的DataKeys的值，我設定的是pk值

        For i = 0 To Me.GridView1.Rows.Count - 1 Step i + 1

            wFormSNo = GridView1.DataKeys(i).Value.ToString()
            If (CType(GridView1.Rows(i).FindControl("CheckYESDETAIL"), CheckBox)).Checked Then
                ' Response.Write(GridView1.DataKeys(i).Value.ToString())
                'GridView1.DataKeys[i].Value.ToString() 可以抓到該列的DataKeys的值，我設定的是pk值
                wSts = "1"
            ElseIf (CType(GridView1.Rows(i).FindControl("CheckNODETAIL"), CheckBox)).Checked Then
                wSts = "2"
            End If

            wApplyID = Me.GridView1.Rows(i).Cells(16).Text
            wDecideID = Me.GridView1.Rows(i).Cells(17).Text
            wStep = Me.GridView1.Rows(i).Cells(18).Text
            wNo = Me.GridView1.Rows(i).Cells(19).Text
            ' Dim checkBoxYES As CheckBox = CType(e.Row.FindControl("CheckYESDETAIL"), CheckBox)
            ' Dim checkBoxNO As CheckBox = CType(e.Row.FindControl("CheckNODETAIL"), CheckBox)
            Dim checkBoxYES As CheckBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("CheckYESDETAIL"), CheckBox)
            Dim checkBoxNO As CheckBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("CheckNODETAIL"), CheckBox)
            '取得原因
            Dim txt As TextBox = CType(Me.GridView1.Rows(i).Cells(0).FindControl("TextBox1"), TextBox)

            wDecideDesc = txt.Text
            If wDecideDesc <> "" And (checkBoxYES.Checked = True Or checkBoxNO.Checked = True) Then
                '新增到暫存檔
                Dim SQL As String
                SQL = " Insert into Q_WaitAutoDisposal (Formno,Formsno,step,ApplyID,DecideID,Sts,DecideDesc,createtime)"
                SQL = SQL & "values('006001','" & wFormSNo & "','" & wStep & "','" & wApplyID & "','" & wDecideID & "','" & wSts & "',N'" & YKK.ReplaceString(wDecideDesc) & "',getdate())"
                uDataBase.ExecuteNonQuery(SQL)
                'Response.Write(SQL)

                '將T_WaitHandle狀態更新成 5
                SQL = "  update T_WaitHandle "
                SQL = SQL + " set sts = 5"
                SQL = SQL + " ,stsdes ='批次處理'"
                SQL = SQL + " ,ModifyUser =DecideID "
                SQL = SQL + " ,ModifyTime =getdate()"
                SQL = SQL + " where    DecideID = '" + wDecideID + "'"
                SQL = SQL + " and formno = '006001' and Active = '1'  and flowtype=1  And (Sts = '0'  Or  Sts = '4')"
                SQL = SQL + " and FormSNo ='" + wFormSNo + "'"
                uDataBase.ExecuteNonQuery(SQL)
            ElseIf wDecideDesc = "" And (checkBoxYES.Checked = True Or checkBoxNO.Checked = True) Then
                wError = wError + 1
                If Message = "" Then
                    Message = wNo
                Else
                    Message = Message + "," + wNo
                End If
            End If

           
        Next
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

    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound

        Dim formatNu As Integer
        formatNu = 2
        For i As Integer = 0 To ItemCount - 1
            '計算合計
            Dim intTotal As Decimal = 0
            For Each gvr As GridViewRow In GridView1.Rows


                intTotal += gvr.Cells(i + 8).Text

                gvr.Cells(i + 8).Text = FormatNumber(gvr.Cells(i + 8).Text, formatNu, TriState.True, TriState.False, TriState.True)

                gvr.Cells(i + 8).HorizontalAlign = HorizontalAlign.Right
                GridView1.FooterRow.Cells(i + 8).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)

                GridView1.FooterRow.Cells(i + 8).HorizontalAlign = HorizontalAlign.Right

            Next
        Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(6).Text = "合計"
        End If


        '寫入超連結網址
        If e.Row.RowType <> DataControlRowType.Header And e.Row.Cells(5).Text <> "合計" Then
            Dim h1 As New HyperLink
            h1.Text = e.Row.Cells(1).Text
            h1.Target = "_blank"
            h1.NavigateUrl = ("DASW_DISPOSALSheet01.aspx?pFormNo=006001" & "&pFormSno=" & CStr(e.Row.Cells(15).Text) & "&pStep=" & CStr(e.Row.Cells(18).Text) & "&pSeqNo=" & "1" & "&pApplyID=" & e.Row.Cells(16).Text & "&pUserID=" & wUserID)
            e.Row.Cells(1).Text = ""
            e.Row.Cells(1).Controls.Add(h1)
        End If


        '寫入超連結網址
        If e.Row.Cells(14).Text = "..." Then
            Dim h1 As New HyperLink
            h1.Text = e.Row.Cells(14).Text
            h1.Target = "_blank"
            h1.NavigateUrl = ("http://10.245.1.10/WorkFlow/BefOPList.aspx?pFormNo=006001&pFormSno=" & CStr(e.Row.Cells(15).Text)) + "&pStep=1&pSeqNo=1&pApplyID=" + e.Row.Cells(16).Text
            e.Row.Cells(14).Text = ""
            e.Row.Cells(14).Controls.Add(h1)
        End If


        If e.Row.RowType <> DataControlRowType.Header And e.Row.RowType <> DataControlRowType.Footer Then
            '寫入超連結網址
            If e.Row.Cells(20).Text <> "" Then
                Dim h1 As New HyperLink
                h1.Text = "報廢明細檔"
                h1.Target = "_blank"
                h1.NavigateUrl = ("http://10.245.1.6/DASW/Document/006001/" & e.Row.Cells(20).Text)
                e.Row.Cells(20).Text = ""
                e.Row.Cells(20).Controls.Add(h1)
            End If

        End If



        '若勾則寫入OK
        'If e.Row.RowType = DataControlRowType.DataRow Then
        ' Dim checkBoxYES As CheckBox = CType(e.Row.FindControl("CheckYESDETAIL"), CheckBox)
        ' Dim checkBoxNO As CheckBox = CType(e.Row.FindControl("CheckNODETAIL"), CheckBox)
        'Dim textbox1 As TextBox = CType(e.Row.FindControl("textbox1"), TextBox)
        '  checkBoxYES.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='OK.';} else { document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxNO.ClientID + "').checked=false; }")
        '  checkBoxNO.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='NG.';} else {document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxYES.ClientID + "').checked=false;}")

        'End If
        ' e.Row.Cells(13).Visible = False
        e.Row.Cells(15).Visible = False
        e.Row.Cells(16).Visible = False
        e.Row.Cells(17).Visible = False
        e.Row.Cells(18).Visible = False
        e.Row.Cells(19).Visible = False

        '選擇列時變色 
        If e.Row.RowIndex > -1 Then
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='#e0e0ff'")
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c;")
        End If



    End Sub
End Class

