Imports System.Data
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Collections.Generic


Partial Class RemarkListCL
    Inherits System.Web.UI.Page
    Dim fpObj As New ForProject     '操作db的物件
    Dim uDataBase As Utility.DataBase = fpObj.GetDataBaseObj()
    Dim uCommon As New Utility.Common
    Dim uJavaScript As New Utility.JScript
    Dim Remarkint As String
    Dim Remark As String
    Dim ConStr As String = ""
    Dim taxtype As String = ""
    Dim ErrCode As Integer = 0


    Protected Sub BAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BAdd.Click

        If CheckRADD(0) = False Then
            Dim Message As String = ""
            '20230228 秀貞需求
            If DCat.SelectedValue = "發票號碼" Then
                If Len(DContentList.Text) <> 10 Then
                    ErrCode = 1
                    Message = "請確認發票號碼應10碼！"

                End If
            ElseIf DCat.SelectedValue = "稅單號碼" Then
                If Len(DContentList.Text) <> 14 Then
                    ErrCode = 1
                    Message = "請確認稅單號碼應14碼！"
                End If
            ElseIf DCat.SelectedValue = "賣方統一編號" Then
                If Len(DContentList.Text) <> 8 Then
                    ErrCode = 1
                    Message = "請確認賣方統一編號應8碼！"
                End If
            End If
            If Message <> "" Then
                uJavaScript.PopMsg(Me, Message)
            End If



            If ErrCode = 0 Then
                '字串ENTER取代成-
                Dim contentstr As String
                contentstr = DContentList.Text.Replace(vbCrLf, "-")

                If DRemarkList.Text = "" Then
                    DRemarkList.Text = "[" & DCat.Text & "=" & Trim(contentstr) & "]"
                Else
                    DRemarkList.Text = DRemarkList.Text & "[" & DCat.Text & "=" & DContentList.Text & "]"
                End If
                DContentList.Text = ""
            End If
        End If
    
    End Sub

    Protected Sub BClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BClose.Click

        If CheckRemark(0) = True Then

            Dim js As String = ""
            If Remarkint = 1 Then
                js &= "var obj = window.opener.document.all.DContent1;"
                js &= "obj.value = '" & DRemarkList.Text & "';"
            Else
                js &= "var obj = window.opener.document.all.DRemark1;"
                js &= "obj.value = '" & DRemarkList.Text & "';"
            End If
            'js &= "var obj = window.opener.document.all.DContent;"
            'js &= "obj.value = '" & DRemarkList.Text & "';"
            'js &= "var obj = window.opener.document.all.DRemark;"
            'js &= "obj.value = '" & DRemarkList.Text & "';"
            js &= "window.opener.document.forms[0].submit();"
            js &= "window.close();"
            Me.ClientScript.RegisterClientScriptBlock(GetType(String), "back", js, True)
        End If



    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then   '不是PostBack
            GetRmarkData()
        Else
            Remarkint = Mid(Request.QueryString("field"), 1, 1)
            Remark = Mid(Request.QueryString("field"), 3, Len(Request.QueryString("field")) - 1)

        End If

    End Sub

    Sub GetRmarkData()

        Remarkint = Mid(Request.QueryString("field"), 1, 1)
        Remark = Mid(Request.QueryString("field"), 3, Len(Request.QueryString("field")) - 1)
        taxtype = Request.QueryString("field1")


        Dim sql As String
        Dim Content As String = ""
        Dim i As Integer

        '找出必填的字串

        sql = "  Select RContent as Data  from M_expitemlistCL"
        sql = sql & " where expcat+'--'+expitem = '" & Remark + "'"
        Dim dtReferp As DataTable = uDataBase.GetDataTable(sql)

        If dtReferp.Rows.Count > 0 Then
            Content = dtReferp.Rows(i).Item("Data")
            Dim str As String = Content
            Dim ary As String() = str.Split(","c)

            For Each e As String In ary
                If ConStr = "" Then
                    ConStr = "'" + e.ToString() + "'"
                Else
                    ConStr = ConStr + ",'" + e.ToString() + "'"
                End If
            Next

            '檢查是不是'電子發票證明聯(V4)','繳納證(V8)','統一發票(V2)','收銀機發票(V4)'
            '2024/10/30 jessica 不需要
            'If taxtype = "電子發票證明聯(V4)" Or taxtype = "收銀機發票(V4)" Or taxtype = "統一發票(V2)" Then
            ' ConStr = ConStr + ",'發票號碼','賣方統一編號'"
            'ElseIf taxtype = "繳納證(V8)" Then
            ' ConStr = ConStr + ",'稅單號碼'"
            'End If
            '2024/10/30 jessica 不需要

            If ConStr <> "" Then
                ConStr = " and  Data in (" + ConStr + ")"
            Else
                ConStr = ""
            End If
        End If

        '再從Master 找出來對應

        sql = "  Select  * from M_referp"
        sql = sql & " where  cat = '3111'"
        If Remarkint = 1 Then
            sql = sql & " and dkey = 'Content'" & ConStr
        Else
            sql = sql & " and dkey = 'Remark'"
        End If

        sql = sql & " order by unique_id"

        DChk.Text = sql

        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(sql)
        DCat.Items.Clear()
        DCat.Items.Add("")
        For i = 0 To dtReferp1.Rows.Count - 1
            Dim ListItem1 As New ListItem
            ListItem1.Text = dtReferp1.Rows(i).Item("Data")
            ListItem1.Value = dtReferp1.Rows(i).Item("Data")
            DCat.Items.Add(ListItem1)
        Next
        dtReferp.Clear()

        dtReferp.Clear()
    End Sub
    '檢查字串是否有必填項目
    Function CheckRemark(ByVal pAction As Integer) As Boolean
        Dim i As Integer
        Dim chk As Boolean
        Dim data As String
        Dim dtReferp1 As DataTable = uDataBase.GetDataTable(DChk.Text)
        For i = 0 To dtReferp1.Rows.Count - 1
            data = DRemarkList.Text
            '判斷字串是否有必填字元
            chk = data.Contains(dtReferp1.Rows(i).Item("Data"))
            If chk = False Then
                uJavaScript.PopMsg(Me, "[" + dtReferp1.Rows(i).Item("Data") + "]" + "為必填項目")
                Exit For
            End If
        Next

        Return chk

    End Function



    '檢查必需必填項目是否重覆
    Function CheckRADD(ByVal pAction As Integer) As Boolean

        Dim chk As Boolean
        Dim data As String


        data = DRemarkList.Text
        '判斷字串是否有必填字元
        chk = data.Contains(DCat.SelectedValue)
        If chk = True Then
            uJavaScript.PopMsg(Me, "必填項目重覆輸入")

        End If


        Return chk

    End Function


    Protected Sub BDEL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BDEL.Click
        DRemarkList.Text = ""
    End Sub
End Class
