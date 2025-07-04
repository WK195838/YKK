Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Data.OleDb
Imports System.IO



 


Partial Class DASW_DISPOSALReport
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
            GetDisposalYM()
            'DataList()
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
        Dim srule As String = ""
        If DDISPOSALRULE.SelectedValue <> "" Then
            srule = " and DISPOSALRULE ='" + DDISPOSALRULE.SelectedValue + "'"
        End If
        Dim stype As String = ""
        If DDISPOSALTYPE.SelectedValue <> "" Then
            stype = " and DISPOSALTYPE like '%" + DDISPOSALTYPE.SelectedValue + "%'"
        End If

        Dim SQL As String
        SQL = " select a.*,isnull(decidename,'完成') as decidename,OPURL  from ("
        SQL = SQL + " select  appname,disposalfile1,no,dutydepo,disposaltype,customertoll,disposalreason,disrule as disposalrule,STYPE,"
        SQL = SQL + " sum(convert(decimal(15,2),piece))piece,"
        SQL = SQL + " sum(convert(decimal(15,2),meter))meter,"
        SQL = SQL + " sum(convert(decimal(15,2),yard))yard,"
        SQL = SQL + " sum(convert(decimal(15,2),kg))kg,"
        SQL = SQL + " sum(convert(decimal(12,2),price))price,formsno"
        SQL = SQL + " from ("
        SQL = SQL + " select  appname,disposalfile1,a.no,dutydepo,disposaltype, case when customertoll =1 then 'V' else'' end as customertoll, "
        SQL = SQL + " case when disposalreason='其它' then disposalreason+'-'+disposalreason1 else disposalreason end as disposalreason, disrule,STYPE, "
        SQL = SQL + " case when b.piece='' then '0' else  replace(b.piece,',','') end as   piece,"
        SQL = SQL + " case when b.meter='' then '0' else  replace(b.meter,',','') end as   meter,"
        SQL = SQL + " case when b.yard='' then '0' else  replace(b.yard,',','') end as   yard,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.kg,',','') end as   kg,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.price,',','') end as   price,b.formsno"
        SQL = SQL + " from F_DISPOSALSheet a,F_disposalsheetdt  b  where a.sts <> 2 and disposalYM='" + DDisposalYM.Text + "'"
        SQL = SQL + srule + stype
        SQL = SQL + " and a.formsno =b.formsno "
        SQL = SQL + " )a group by  no,dutydepo,disposaltype,customertoll,disposalreason,disrule,formsno,appname,disposalfile1,STYPE"
        SQL = SQL + " )a left join (select formsno,decidename,"
        SQL = SQL + "'履歷' As OPURL  "
        SQL = SQL + " from t_waithandle where formno = 6001 "
        SQL = SQL + " and active =1and flowtype <> 0 ) b on a.formsno =b.formsno "
        SQL = SQL + " union all "
        SQL = SQL + " select   '' as appname,'' as disposalfile1,"
        SQL = SQL + " case when disposalrule ='低價法' then   'D999999991' "
        SQL = SQL + " when disposalrule ='非低價法' then   'D999999992' "
        SQL = SQL + " else    'D999999993' end as no,"
        SQL = SQL + " '' dutydepo,'' disposaltype,''  customertoll, ''disposalreason,disposalrule,STYPE,"
        SQL = SQL + " convert(char,sum(convert(decimal(15,2),replace(piece,',',''))))   piece,"
        SQL = SQL + " convert(char, sum(convert(decimal(15,2),replace(meter,',','')))) meter,"
        SQL = SQL + " convert(char, sum(convert(decimal(15,2),replace(yard,',',''))))yard, "
        SQL = SQL + " convert(char,sum(convert(decimal(15,2),replace(kg,',','')))) kg,"
        SQL = SQL + " convert(char,sum(convert(decimal(15,2),replace(price,',','')))) price"
        SQL = SQL + " , 1 as formsno,'' as decidename,'' as OPURL "
        SQL = SQL + " from ("
        SQL = SQL + " select a.*,isnull(decidename,'完成') as decidename,OPURL  from ("
        SQL = SQL + " select  no,dutydepo,disposaltype,customertoll,disposalreason,disrule as disposalrule,STYPE,"
        SQL = SQL + " sum(convert(decimal(15,2),piece))piece,"
        SQL = SQL + " sum(convert(decimal(15,2),meter))meter,"
        SQL = SQL + " sum(convert(decimal(15,2),yard))yard,"
        SQL = SQL + " sum(convert(decimal(15,2),kg))kg,"
        SQL = SQL + " sum(convert(decimal(15,2),price))price,formsno"
        SQL = SQL + " from ("
        SQL = SQL + " select  a.no,dutydepo,disposaltype, case when customertoll =1 then 'V' else'' end as customertoll, "
        SQL = SQL + " case when disposalreason='其它' then disposalreason+'-'+disposalreason1 else disposalreason end as disposalreason, disrule,STYPE, "
        SQL = SQL + " case when b.piece='' then '0' else  replace(b.piece,',','') end as   piece,"
        SQL = SQL + " case when b.meter='' then '0' else  replace(b.meter,',','') end as   meter,"
        SQL = SQL + " case when b.yard='' then '0' else  replace(b.yard,',','') end as   yard,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.kg,',','') end as   kg,"
        SQL = SQL + " case when b.kg='' then '0' else  replace(b.price,',','') end as   price,b.formsno"
        SQL = SQL + " from F_DISPOSALSheet a,F_disposalsheetdt  b  where a.sts <> 2 and disposalYM='" + DDisposalYM.Text + "'"
        SQL = SQL + srule + stype
        SQL = SQL + " and a.formsno =b.formsno "
        SQL = SQL + " )a group by  no,dutydepo,disposaltype,customertoll,disposalreason,disrule,formsno,STYPE"
        SQL = SQL + " )a left join (select formsno,decidename,"
        SQL = SQL + "'履歷' As OPURL  "
        SQL = SQL + " from t_waithandle where formno = 6001 "
        SQL = SQL + " and active =1and flowtype <> 0 ) b on a.formsno =b.formsno "
        SQL = SQL + " )a GROUP BY DISPOSALRULE,stype"
        SQL = SQL + " order by no,disposalrule"


        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)

        GridView1.DataSource = DBAdapter1
        GridView1.DataBind()


    End Sub

 

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound

        Dim formatNu As Integer
        formatNu = 2

    
        For i As Integer = 0 To ItemCount - 1
            '計算合計

            Dim intTotal As Decimal = 0
            For Each gvr As GridViewRow In GridView1.Rows

                Dim sno As String
                sno = gvr.Cells(0).Text

                If sno = "D201811001" Then
                    Dim a As String
                    a = gvr.Cells(8).Text
                End If

                If (sno = "低價法對象合計") Or (sno = "非低價法對象合計") Or (sno = "兩者皆是合計") Then
                    If gvr.RowType = DataControlRowType.DataRow Then
                        intTotal += gvr.Cells(i + 8).Text

                    End If
                End If
                'intTotal += gvr.Cells(i + 6).Text

                If i = 0 Or i = 4 Then
                    gvr.Cells(i + 8).Text = FormatNumber(gvr.Cells(i + 8).Text, 0, TriState.True, TriState.False, TriState.True)
                Else
                    gvr.Cells(i + 8).Text = FormatNumber(gvr.Cells(i + 8).Text, formatNu, TriState.True, TriState.False, TriState.True)

                End If


                gvr.Cells(i + 8).HorizontalAlign = HorizontalAlign.Right


                If i = 0 Or i = 4 Then
                    GridView1.FooterRow.Cells(i + 8).Text = FormatNumber(intTotal, 0, TriState.True, TriState.False, TriState.True)

                Else
                    GridView1.FooterRow.Cells(i + 8).Text = FormatNumber(intTotal, formatNu, TriState.True, TriState.False, TriState.True)

                End If

                GridView1.FooterRow.Cells(i + 8).HorizontalAlign = HorizontalAlign.Right




            Next
        Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowType = DataControlRowType.Footer Then
            e.Row.Cells(0).Text = "總合計"
        End If

        Dim sno As String
        sno = e.Row.Cells(0).Text

    



        If Mid(sno, 1, 9) <> "D99999999" Then
            '寫入超連結網址
            If e.Row.RowType <> DataControlRowType.Header And e.Row.Cells(5).Text <> "合計" Then
                Dim h1 As New HyperLink
                h1.Text = e.Row.Cells(0).Text
                If e.Row.Cells(0).Text <> "總合計" Then
                    h1.Target = "_blank"
                    h1.NavigateUrl = ("DASW_DISPOSALSheet02.aspx?pFormNo=006001" & "&pFormSno=" & e.Row.Cells(13).Text)
                End If

                e.Row.Cells(0).Text = ""
                e.Row.Cells(0).Controls.Add(h1)

               

                Dim h15 As New HyperLink
                Dim h14 As New HyperLink
                If h1.Text <> "總合計" Then


                    h15.Text = "履歷"
                    Dim a As String
                    a = e.Row.Cells(0).Text

                    'If e.Row.Cells(14).Text <> "" Then

                    h15.Target = "_blank"
                    h15.NavigateUrl = ("http://10.245.1.10/WorkFlow/BefOPList.aspx?pFormNo=006001" & "&pFormSno=" & e.Row.Cells(13).Text)
                    e.Row.Cells(16).Text = ""
                    e.Row.Cells(16).Controls.Add(h15)

                    h14.Text = "報廢明細檔"
                    h14.Target = "_blank"
                    h14.NavigateUrl = ("http://10.245.1.6/DASW/Document/006001/" & e.Row.Cells(15).Text)

                    e.Row.Cells(15).Text = ""
                    e.Row.Cells(15).Controls.Add(h14)

                    'End If

                End If







            End If


        Else

            If sno = "D999999991" Then
                e.Row.Cells(0).Text = "低價法對象合計"
            ElseIf sno = "D999999992" Then
                e.Row.Cells(0).Text = "非低價法對象合計"
            Else
                e.Row.Cells(0).Text = "兩者皆是合計"
            End If
            e.Row.Cells(7).Text = ""


        End If

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim i As Integer
            For i = 0 To 12
                If sno = "D999999991" Then
                    e.Row.Cells(i).BackColor = Drawing.Color.Aqua
                ElseIf sno = "D999999992" Then
                    e.Row.Cells(i).BackColor = Drawing.Color.LightPink
                ElseIf sno = "D999999993" Then
                    e.Row.Cells(i).BackColor = Drawing.Color.LightGreen
                End If

            Next
        End If


        '若勾則寫入OK
        'If e.Row.RowType = DataControlRowType.DataRow Then
        ' Dim checkBoxYES As CheckBox = CType(e.Row.FindControl("CheckYESDETAIL"), CheckBox)
        ' Dim checkBoxNO As CheckBox = CType(e.Row.FindControl("CheckNODETAIL"), CheckBox)
        'Dim textbox1 As TextBox = CType(e.Row.FindControl("textbox1"), TextBox)
        '  checkBoxYES.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='OK.';} else { document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxNO.ClientID + "').checked=false; }")
        '  checkBoxNO.Attributes.Add("onclick", "if(this.checked){document.getElementById('" + textbox1.ClientID + "').value='NG.';} else {document.getElementById('" + textbox1.ClientID + "').value=''; document.getElementById('" + checkBoxYES.ClientID + "').checked=false;}")

        'End If
        e.Row.Cells(13).Visible = False



        '選擇列時變色 
        If e.Row.RowIndex > -1 Then
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='#e0e0ff'")
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c;")
        End If



    End Sub


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataList()
        LDISPOSALFILE.NavigateUrl = ""
        GetDisposalFile()
    End Sub

    Sub GetDisposalYM()
        Dim SQL As String
        Dim i As Integer
        SQL = " select distinct disposalym from F_DISPOSALSheet order by disposalym desc "
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        DDisposalYM.Items.Clear()
        DDisposalYM.Items.Add("")
        For i = 0 To DBAdapter1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = DBAdapter1.Rows(i).Item("disposalym")
            ListItem1.Value = DBAdapter1.Rows(i).Item("disposalym")
            DDisposalYM.Items.Add(ListItem1)
        Next
        DDisposalYM.SelectedValue = Now.ToString("yyyy/MM")
        DBAdapter1.Clear()
        '報廢準則
       
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DISPOSALRULE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALRULE.Items.Add("")
        For i = 0 To dtReferp.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp.Rows(i).Item("Data")
            ListItem1.Value = dtReferp.Rows(i).Item("Data")
            DDISPOSALRULE.Items.Add(ListItem1)
        Next
        dtReferp.Clear()

        '報廢品分類
      
        SQL = "  Select  * from M_referp"
        SQL = SQL & " where  cat = '6001'"
        SQL = SQL & " and dkey = 'DISPOSALTYPE'"
        SQL = SQL & " order by unique_id"

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)
        DDISPOSALTYPE.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DDISPOSALTYPE.Items.Add(ListItem1)
        Next
        dtReferp1.Clear()
        

    End Sub


    '*****************************************************************
    '**
    '**     轉Excel共用程式
    '**
    '*****************************************************************
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        'Confirms that an HtmlForm control is rendered for the specified ASP.NET
        ' server control at run time. */
    End Sub

    Private Sub BExcel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BExcel.Click
        Dim wAllowPaging As Boolean = GridView1.AllowPaging   '程式別不同

        ' pFormNo = DFormName.SelectedValue
        GridView1.AllowPaging = False   '程式別不同
        DataList()                      '程式別不同

        Response.AppendHeader("Content-Disposition", "attachment;filename=DisposalReport-" + CStr(DDisposalYM.SelectedValue) + ".xls")     '程式別不同
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5")
        Response.ContentType = "application/vnd.ms-excel"
        Response.Charset = ""
        Me.EnableViewState = False

        Dim tw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Response.Write(tw.ToString())
        Response.End()

        GridView1.AllowPaging = wAllowPaging        '程式別不同
 
    End Sub

  

    Sub GetDisposalFile()
        Dim OpenDir2 As String
        OpenDir2 = ""
        Dim SQL As String
        '主檔資料
        SQL = " select  substring(data,1,15)data,substring(replace(data,'/','\'),1,15)data1  from M_referp"
        SQL = SQL + " where cat = '6001'"
        SQL = SQL + " and dkey ='DisposalFilePath'"
        Dim DBAdapter1 As DataTable = uDataBase.GetDataTable(SQL)
        Dim OpenDir1 As String = ""
        If DBAdapter1.Rows.Count > 0 Then
            'OpenDir1 = "file://" + DBAdapter1.Rows(0).Item("Data")
            OpenDir2 = "\\" + DBAdapter1.Rows(0).Item("Data1")
        End If

        Dim Message As String
        '檢查目錄是否存，不存在就新建一個
        Dim IP As String
        IP = HttpContext.Current.Request.UserHostAddress
        If DDisposalYM.SelectedValue <> "" Then
            Dim tempFolderPath As String
            tempFolderPath = OpenDir2 + "\DASW報廢附檔\" + CStr(Mid(DDisposalYM.SelectedValue, 1, 4)) + CStr(Mid(DDisposalYM.SelectedValue, 6, 2)) + "\"
            Dim dInfo As DirectoryInfo = New DirectoryInfo(tempFolderPath)
            If dInfo.Exists Then
                tempFolderPath = OpenDir2 + "\DASW報廢附檔\" + CStr(Mid(DDisposalYM.SelectedValue, 1, 4)) + CStr(Mid(DDisposalYM.SelectedValue, 6, 2)) + "\"
                '存在就刪除
                '設定刪除目錄和其下檔案和子目錄
                ' My.Computer.FileSystem.DeleteDirectory(tempFolderPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
            Else
                'dInfo.Create()
                My.Computer.FileSystem.CreateDirectory(tempFolderPath)
            End If

            '刪除資料夾的附檔
            Dim FileInfo() As IO.FileSystemInfo = dInfo.GetFileSystemInfos

            For Each info As IO.FileSystemInfo In FileInfo
                info.Delete()
            Next




            '    System.IO.File.Copy("\\10.245.1.10\www$\DASW\Document\006001\" + "\56-DPFILE-20163115434-報廢格式_3月.xlsx", tempFolderPath + "\56-DPFILE-20163115434-報廢格式_3月.xlsx", True)

            Dim i As Integer
            SQL = "  Select no,disposalfile1 as disposalfile1    from f_disposalsheet"
            SQL = SQL + "  where sts <> 2 and disposalYM='" + DDisposalYM.SelectedValue + "'"
            SQL = SQL + " order by no"
            Dim dtReferp1 As DataTable = uDataBase.GetDataTable(SQL)

            'System.IO.File.Delete(tempFolderPath)


            For i = 0 To dtReferp1.Rows.Count - 1
                If dtReferp1.Rows(i).Item("disposalfile1") <> "" Then
                    Dim a As String
                    a = dtReferp1.Rows(i).Item("disposalfile1")
                    ' Dim sourceFile As String = System.IO.Path.Combine("\\10.245.1.10\www$\DASW\Document\006001", dtReferp1.Rows(i).Item("disposalfile1"))
                    'Dim destFile As String = System.IO.Path.Combine(tempFolderPath, dtReferp1.Rows(i).Item("no") + "-" + dtReferp1.Rows(i).Item("disposalfile1"))
                    'System.IO.File.Copy(sourceFile, destFile, True)
                    ' uJavaScript.PopMsg(Me, a)
                    System.IO.File.Copy("\\10.245.1.6\www$\DASW\Document\006001\" + dtReferp1.Rows(i).Item("disposalfile1"), tempFolderPath + dtReferp1.Rows(i).Item("no") + "-" + dtReferp1.Rows(i).Item("disposalfile1"), True)

                End If

            Next


            dtReferp1.Clear()

            '  Message = "報廢檔案已匯出完成，存於\\10.245.1.61\wfs$\DASW報廢附檔" + CStr(Mid(DDisposalYM.SelectedValue, 1, 4)) + CStr(Mid(DDisposalYM.SelectedValue, 6, 2))

            LDISPOSALFILE.NavigateUrl = OpenDir2 + "\DASW報廢附檔\" + CStr(Mid(DDisposalYM.SelectedValue, 1, 4)) + CStr(Mid(DDisposalYM.SelectedValue, 6, 2))
        Else

            Message = "請選擇年月!"
            uJavaScript.PopMsg(Me, Message)
        End If

        'Response.Write(YKK.ShowMessage(Message))

        DataList()
    End Sub
   
    Protected Sub DDisposalYM_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDisposalYM.SelectedIndexChanged
        LDISPOSALFILE.NavigateUrl = ""
        DataList()
        GetDisposalFile()
    End Sub

    Protected Sub DDISPOSALRULE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALRULE.SelectedIndexChanged
        LDISPOSALFILE.NavigateUrl = ""
        DataList()
        GetDisposalFile()
    End Sub

    Protected Sub DDISPOSALTYPE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDISPOSALTYPE.SelectedIndexChanged
        LDISPOSALFILE.NavigateUrl = ""
        DataList()
        GetDisposalFile()
    End Sub
End Class

