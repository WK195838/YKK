Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data.SqlClient

Public Class ForProject

    Dim strConnectionKey = "WF_Con"     ' 連線字串

    '取得資料庫共用函式物件
    Public Function GetDataBaseObj() As Utility.DataBase
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Return uDataBase
    End Function

    '取得某一日期所屬月份的每週第一天清單
    Public Function GetTheFirstDaysOfWeekInMoth(ByVal dt As DateTime) As DateTime()
        Dim dtList As New List(Of DateTime)
        Dim dtTemp As DateTime = GetTheFirstDayOfWeek(GetTheFirstDayOfMonth(dt))        ' 取得某月第一週的第一天
        Dim dtEnd As DateTime = GetTheLastDayOfMonth(dt)

        ' 取得某月的最後一天
        For i As Integer = 0 To 5
            If dtTemp.AddDays(i * 7) <= dtEnd Then                                      ' 此集合所記錄的為所屬月份的每一週
                dtList.Add(dtTemp.AddDays(i * 7))
            End If
        Next
        Return dtList.ToArray()
    End Function

    '取得這個月的第1週的第1天
    Public Function GetTheFirstDayOfWeek(ByVal dt As DateTime) As DateTime
        If CInt(dt.DayOfWeek) = 0 Then
            Return dt.AddDays(7 * -1 + 1)
        Else
            Return dt.AddDays(CInt(dt.DayOfWeek) * -1 + 1)
        End If
    End Function

    '取得這個月的第1天
    Public Function GetTheFirstDayOfMonth(ByVal dt As DateTime) As DateTime
        Return New DateTime(dt.Year, dt.Month, 1)
    End Function

    '取得這個月的最後1天
    Public Function GetTheLastDayOfMonth(ByVal dt As DateTime) As DateTime
        If dt.Month = 12 Then
            Return New DateTime(dt.Year + 1, 1, 1).AddDays(-1)
        Else
            Return New DateTime(dt.Year, dt.Month + 1, 1).AddDays(-1)
        End If

    End Function

    '取得某月有幾週
    Public Function GetMonthWeek(ByVal y As Integer, ByVal m As Integer) As Integer
        Dim dt As New DateTime(Val(y), Val(m), 1)  '增加可選年份所以多加一個年份的變數 2010/12/23 modify by jessica
        Return GetTheFirstDaysOfWeekInMoth(dt).Length
    End Function

    '取得插號內的字串
    Function GetBrackets(ByVal value As String) As String
        Dim str As String = value.Substring(value.LastIndexOf("(") + 1, value.LastIndexOf(")") - value.LastIndexOf("(") - 1)
        Return str
    End Function

    '取得關係人
    Function GetRelated(ByVal userId As String) As ListItemCollection
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Dim sql As String = "select RUserID , RUserName  from M_Related where userid='" & userId & "' and RelatedID='B'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim lic As New ListItemCollection
        For Each dtr As Data.DataRow In dt.Rows
            Dim li As New ListItem(dtr("RUserName"), dtr("RUserID"))
            lic.Add(li)
        Next
        Return lic
    End Function


    '取得代理人
    Function GetAgent(ByVal aFormno As String, ByVal aFormsno As String, ByVal aStep As String) As String
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Dim sql As String = " select agentid  from  t_waithandle "
        sql &= " where formsno ='" + aFormsno + "'"
        sql &= " and formno  ='" + aFormno + "'"
        sql &= " and step = '" + aStep + "'"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim str As String
        If dt.Rows.Count > 0 Then
            str = dt.Rows(0)("agentid").ToString
            Return str
        Else
            str = ""
            Return str
        End If

    End Function


    '取得可申請部門
    Function GetApplyDiv(ByVal userId As String) As ListItemCollection
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Dim sql As String = "select data  from M_Referp Where cat='1998' and dkey = 'APPLYDIV-" & userId & "' order by data"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        Dim lic As New ListItemCollection
        For Each dtr As Data.DataRow In dt.Rows
            Dim li As New ListItem(dtr("data"))
            lic.Add(li)
        Next
        Return lic
    End Function

    '判斷日期是否為假日
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Function IsHoliday(ByVal d As String) As Boolean
    '
    Function IsHoliday(ByVal d As String, ByVal pDepo As String) As Boolean
        'MsgBox(d + "-" + pDepo)
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))

        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'Dim sql As String = "select count (*)  from V_Vacation_01 where depo='cl' and sdate = '" & d & "'"
        '
        Dim sql As String = "select count (*)  from V_Vacation_01 where depo='" & pDepo & "' and sdate = '" & d & "'"
        'Modify-End

        Dim result As String = uDataBase.SelectQuery(sql)
        If result = "0" Then
            Return False
        Else
            Return True
        End If
    End Function

    '傳回日期為平日,假日,國定假日
    'Modify-Start by Joy 2009/11/20(2010行事曆對應)
    'Function GetHolidayType(ByVal d As String) As String
    '
    Function GetHolidayType(ByVal d As String, ByVal pDepo As String) As String
        'Modify-End
        'MsgBox(pDepo)
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))

        'Modify-Start by Joy 2009/11/20(2010行事曆對應)
        'Dim sql As String = "select vacationtype from m_vacation where depo='cl' and active='1' and ymd='" & d & "'"
        '
        Dim sql As String = "select vacationtype from m_vacation where depo='" & pDepo & "' and active='1' and ymd='" & d & "'"
        'Modify-End

        Dim vacationType As String = uDataBase.SelectQuery(sql)
        Select Case vacationType
            Case "0"
                vacationType = "假日"
            Case "1"
                vacationType = "國定假日"
            Case Else
                vacationType = "平日"
        End Select
        Return vacationType
    End Function

    '取得整數時
    Function GetHour(ByVal n As Decimal) As Integer
        Return Int(n)
    End Function

    '取得小數分
    Function GetMin(ByVal n As Decimal) As Decimal
        Dim i As Integer = Int(n)
        Dim d As Decimal = n - i
        Return d
    End Function

    '取得刷卡資料的上班時,上班分,下班時,下班分
    Function GetCardTime(ByVal s As String) As List(Of String)
        Dim CardTime As New List(Of String)
        's = s.Replace("刷卡時間-", "")
        'Dim sl As String() = s.Split("~")
        'If String.IsNullOrEmpty(sl(0)) Then
        '    CardTime.Add("0")
        '    CardTime.Add("0")
        'Else
        '    Dim temp As String() = sl(0).Split(":")
        '    CardTime.Add(temp(0))
        '    CardTime.Add(temp(1))
        'End If
        'If String.IsNullOrEmpty(sl(1)) Then
        '    CardTime.Add("0")
        '    CardTime.Add("0")
        'Else
        '    Dim temp As String() = sl(1).Split(":")
        '    CardTime.Add(temp(0))
        '    CardTime.Add(temp(1))
        'End If
        CardTime.Add("0")
        CardTime.Add("0")
        CardTime.Add("0")
        CardTime.Add("0")
        Return CardTime
    End Function

    '將小於10的補0
    Function TransString(ByVal n As Integer) As String
        Dim s As String = n.ToString

        If n < 10 Then
            s = "0" & n.ToString
        End If
        Return s
    End Function

    '取得BaseTime46
    Function GetBaseTime46(ByVal h As Decimal, ByVal dateType As String) As Decimal

        '如果是平日 , 則傳為原有時數
        '如果是假日 , 則>8傳回8-x , <8則傳為0

        If dateType = "平日" Then
            Return h
        Else
            If h > 8 Then
                Return h - 8
            Else
                Return 0
            End If
        End If
    End Function

    ''' <summary>
    ''' 將簽核者預設選為部門關係
    ''' </summary>
    ''' <param name="nextGate">簽核者下拉選單</param>
    ''' <param name="divisionName">部門名稱</param>
    '''  <param name="applyID">申請人id</param>
    ''' <remarks></remarks>
    Sub GetDeafultNextGate(ByRef nextGate As DropDownList, ByVal divisionName As String, ByVal applyID As String)
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))

        For Each li As ListItem In nextGate.Items
            Dim sql As String = "Select RUserID from M_related where RelatedID = 'B' and UserID = '" & applyID & "' and  RDivision like '%' + (Select distinct rtrim(SupportdivID) from M_Emp where SupportDivName = '" & divisionName & "') + '%'"
            Dim result As String = uCommon.ReplaceDBnull(uDataBase.SelectQuery(sql), "")

            If Not String.IsNullOrEmpty(result) Then
                nextGate.SelectedValue = result
                Exit For
            End If
        Next
    End Sub

    Sub GetMasterTime(ByRef hourList As ListItemCollection, ByRef minList As ListItemCollection, ByRef adjList As ListItemCollection)
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Dim sql As String = "SELECT [Data] , [Dkey] FROM M_Referp WHERE (Cat = '1052') AND (DKey = 'HOUR' or Dkey ='MIN' or Dkey='ADJHOUR') ORDER BY DATA"
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        For Each dr As Data.DataRow In dt.Rows
            Dim li As New ListItem(dr("Data").ToString)
            Select Case dr("Dkey").ToString
                Case "HOUR"

                    hourList.Add(li)

                Case "MIN"
                    minList.Add(li)

                Case "ADJHOUR"
                    adjList.Add(li)
            End Select
        Next
    End Sub

    '置換SQL關鍵字
    Function ReplaceString(ByVal pData As String) As String
        ReplaceString = ""
        Dim Str As String = pData
        Str = Replace(Str, "'", "''")
        Str = Replace(Str, ",", "，")
        ReplaceString = Str
    End Function

    '取得OP CODE
    Function GetOPCode(ByVal pDKey As String) As String
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        '
        Dim Sql As String = ""
        GetOPCode = "@@@@@"
        '
        Sql = "Select Data From M_Referp "
        Sql &= "Where Cat = '" & "2002" & "' "
        Sql &= "  And DKey = '" & pDKey & "' "
        Dim dt_OPCode As Data.DataTable = uDataBase.GetDataTable(Sql)
        If dt_OPCode.Rows.Count > 0 Then
            GetOPCode = dt_OPCode.Rows(0).Item("Data")
        End If
        '
    End Function

End Class
