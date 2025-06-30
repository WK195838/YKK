Imports System.Data
Imports System.Data.OleDb

Partial Class HR_WorkRatioOtherInfor
    Inherits System.Web.UI.Page
    '*****************************************************************
    '**
    '**     自訂涵式庫
    '**
    '*****************************************************************
    Public YKK As New YKK_SPDClass   'YKK SPD系共通涵式
    '*****************************************************************
    '**
    '**     全域變數
    '**
    '*****************************************************************
    Dim wDepo As String
    Dim wYear As String
    Dim wEmpid As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataList()
    End Sub
    Sub DataList()

        Dim SQL As String
        Dim i As Integer
        Dim DBDataSet1 As New DataSet
        Dim DBTable1 As DataTable
        Dim OleDbConnection1 As New OleDbConnection
        OleDbConnection1.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("SqlConn")  'SQL連結設定

        OleDbConnection1.Open()
        wDepo = Request.QueryString("pDepo")
        wYear = Request.QueryString("pYear")
        wEmpid = Request.QueryString("pEmpId")
        'DDivision.Text = HttpUtility.UrlDecode(Request.QueryString("pEmpId").ToString(), System.Text.Encoding.UTF8)
        '---------------------------------------------------------------------------------------------------------
        '顯示員工個人資料
        '---------------------------------------------------------------------------------------------------------
        'SQL = " select com_code,id,b.hrwdivname,name   from m_emp a,M_users b "
        'SQL &= " where  a.com_code = b.depoid  and a.id = b.empid"
        'SQL &= " and id = '" + wEmpid + "'"

        'Dim DBAdapter4 As New OleDbDataAdapter(SQL, OleDbConnection1)
        'DBAdapter4.Fill(DBDataSet1, "M_Users")
        'DBTable1 = DBDataSet1.Tables("M_Users")
        'If DBTable1.Rows.Count > 0 Then
        '    DDivision.Text = Trim(DBTable1.Rows(0).Item("hrwdivname"))
        '    DName.Text = Trim(DBTable1.Rows(0).Item("name"))
        '    DEmpID.Text = Trim(DBTable1.Rows(0).Item("id"))
        'End If

        '---------------------------------------------------------------------------------------------------------
        'Y2  應出勤時間
        '---------------------------------------------------------------------------------------------------------
        SQL = " select  *  from  m_referp "
        SQL &= " where cat ='016' "
        SQL &= " and dkey like '" + wDepo + "-Y2%" + "'"

        Dim DBAdapter1 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter1.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        DY2Sum.Text = "0"
        If DBTable1.Rows.Count > 0 Then
            For i = 0 To DBTable1.Rows.Count - 1
                If DBTable1.Rows(i).Item("DKEY") = wDepo + "-Y2" Then
                    DY2Min.Text = Trim(DBTable1.Rows(i).Item("Data"))
                    DY2Sum.Text = Trim(DBTable1.Rows(i).Item("Data"))
                Else
                    DY2Desc.Text = Trim(DBTable1.Rows(i).Item("Data"))
                End If
            Next
        End If
        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'Y3  應出勤時間
        '---------------------------------------------------------------------------------------------------------
        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + wYear + "'"
        SQL &= " and adjusttype ='Y3' "
        SQL &= " and empid = '" + wEmpid + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter2 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter2.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")
        Dim wDY3SUM As Integer = 0


        If DBTable1.Rows.Count > 0 Then
            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DY3Min1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DY3Min2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DY3Min3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DY3Min4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DY3Min5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DY3Desc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDY3SUM = wDY3SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DY3Sum.Text = wDY3SUM

        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'X2  實際出勤時間
        '---------------------------------------------------------------------------------------------------------

        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + wYear + "'"
        SQL &= " and adjusttype ='X2'"
        SQL &= " and empid = '" + wEmpid + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter3 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter3.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        Dim wDX2SUM As Integer = 0

        If DBTable1.Rows.Count > 0 Then

            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DX2Min1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DX2Min2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DX2Min3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DX2Min4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DX2Min5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DX2Desc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDX2SUM = wDX2SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DX2Sum.Text = wDX2SUM
        DBDataSet1.Clear()

        '---------------------------------------------------------------------------------------------------------
        'ZZ  曠職
        '---------------------------------------------------------------------------------------------------------

        SQL = " select top 5 * from  M_PersonAdjustTime "
        SQL &= " where active = 1 and yy = '" + wYear + "'"
        SQL &= " and adjusttype ='Z1'"
        SQL &= " and empid = '" + wEmpid + "'"
        SQL &= " order by adjusttype,seqno "

        Dim DBAdapter5 As New OleDbDataAdapter(SQL, OleDbConnection1)
        DBAdapter5.Fill(DBDataSet1, "M_Vacation")
        DBTable1 = DBDataSet1.Tables("M_Vacation")

        Dim wDZ1SUM As Integer = 0

        If DBTable1.Rows.Count > 0 Then

            For i = 0 To DBTable1.Rows.Count - 1
                If i = 0 Then
                    DZZMin1.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc1.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 1 Then
                    DZZMin2.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc2.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 2 Then
                    DZZMin3.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc3.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 3 Then
                    DZZMin4.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc4.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                ElseIf i = 4 Then
                    DZZMin5.Text = Trim(DBTable1.Rows(i).Item("AdjustTime"))
                    DZZDesc5.Text = Trim(DBTable1.Rows(i).Item("Remark"))
                End If
                wDZ1SUM = wDZ1SUM + Int(Trim(DBTable1.Rows(i).Item("AdjustTime")))
            Next
        End If
        DZZSum.Text = wDZ1SUM
        DBDataSet1.Clear()



        DY2Min.Style.Add("TEXT-ALIGN", "right")
        DY2Sum.Style.Add("TEXT-ALIGN", "right")

        DY3Min1.Style.Add("TEXT-ALIGN", "right")
        DY3Min2.Style.Add("TEXT-ALIGN", "right")
        DY3Min3.Style.Add("TEXT-ALIGN", "right")
        DY3Min4.Style.Add("TEXT-ALIGN", "right")
        DY3Min5.Style.Add("TEXT-ALIGN", "right")
        DY3Sum.Style.Add("TEXT-ALIGN", "right")

        DX2Min1.Style.Add("TEXT-ALIGN", "right")
        DX2Min2.Style.Add("TEXT-ALIGN", "right")
        DX2Min3.Style.Add("TEXT-ALIGN", "right")
        DX2Min4.Style.Add("TEXT-ALIGN", "right")
        DX2Min5.Style.Add("TEXT-ALIGN", "right")
        DX2Sum.Style.Add("TEXT-ALIGN", "right")

        DZZMin1.Style.Add("TEXT-ALIGN", "right")
        DZZMin2.Style.Add("TEXT-ALIGN", "right")
        DZZMin3.Style.Add("TEXT-ALIGN", "right")
        DZZMin4.Style.Add("TEXT-ALIGN", "right")
        DZZMin5.Style.Add("TEXT-ALIGN", "right")
        DZZSum.Style.Add("TEXT-ALIGN", "right")
    End Sub
End Class
