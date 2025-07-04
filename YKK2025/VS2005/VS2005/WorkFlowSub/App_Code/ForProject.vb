Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data.SqlClient

Public Class ForProject

    Dim strConnectionKey = "WF_Con"     ' 連線字串

    '*********************************************************************
    '取得資料庫共用函式物件
    '*********************************************************************
    Public Function GetDataBaseObj() As Utility.DataBase
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Return uDataBase
    End Function
    '*********************************************************************
    '排序字串陣列
    '*********************************************************************
    Public Function SPRequestSort(ByVal pStr As String) As String
        Dim xStr As String() = pStr.Split("/")
        Dim yStr As String() = pStr.Split("/")
        Dim zStr As String() = pStr.Split("/")
        Dim xEBCDIC As Encoding = Encoding.GetEncoding(37)
        Dim i, j As Integer
        Dim uStr As String = ""
        '
        For i = 0 To xStr.Length - 1
            yStr(i) = BitConverter.ToString(xEBCDIC.GetBytes(xStr(i)))
        Next
        '
        Array.Sort(yStr)
        '
        For i = 0 To zStr.Length - 1
            zStr(i) = ""
        Next
        For i = 0 To yStr.Length - 1
            For j = 0 To xStr.Length - 1
                If yStr(i) = BitConverter.ToString(xEBCDIC.GetBytes(xStr(j))) Then
                    zStr(i) = xStr(j)
                End If
            Next
        Next
        '
        For i = 0 To zStr.Length - 1
            If uStr = "" Then
                uStr = zStr(i)
            Else
                uStr = uStr & "/" & zStr(i)
            End If
        Next
        '
        Return uStr
    End Function
    '*********************************************************************
    '組合Item Name(1)
    '*********************************************************************
    Public Function SetItemName1(ByVal pType As String, ByVal pStr As String) As String
        Dim xStr As String() = pStr.Split("/")
        Dim uStr As String = ""
        '0:Size
        '1:Chain Code
        '2:Class
        '3:Tape
        '4:Slider-1
        '5:Finish-1
        '6:Slider-2
        '7:Finish-2
        '8:Family Code
        '--------------------------------------------------------------------------------
        'SLD
        If pType = "PS" Then
            uStr = xStr(0) & " " & xStr(8) & " " & xStr(4) & " " & xStr(5)
        Else
            '--------------------------------------------------------------------------------
            'CHAIN
            If pType = "CH" Then
                uStr = xStr(0) & " " & xStr(1) & " " & "CHAIN" & " " & xStr(3)
            Else
                '--------------------------------------------------------------------------------
                'TAPE
                If pType = "T" Then
                    uStr = xStr(0) & " " & xStr(1) & " " & xStr(3) & " " & "TAPE"
                Else
                    '--------------------------------------------------------------------------------
                    '引手
                    If pType = "SP" Then
                        uStr = xStr(0) & " " & xStr(1) & " " & xStr(4) & "-" & xStr(3) & " " & xStr(5)
                    Else
                        '--------------------------------------------------------------------------------
                        'ZIP
                        If xStr(1) = "CC" Then      '特殊處理
                            uStr = "CH" & xStr(2) & "-" & GetSize(xStr(0))
                        Else
                            uStr = xStr(1) & xStr(2) & "-" & GetSize(xStr(0))
                        End If
                        If GetSliderFunction(xStr(4) & "/" & xStr(5) & "/" & xStr(6) & "/" & xStr(7)) <> "" Then
                            uStr = uStr & GetSliderFunction(xStr(4) & "/" & xStr(5) & "/" & xStr(6) & "/" & xStr(7))
                        End If
                        If GetSliderFinish(xStr(4) & "/" & xStr(5) & "/" & xStr(6) & "/" & xStr(7)) <> "" Then
                            uStr = uStr & " " & GetSliderFinish(xStr(4) & "/" & xStr(5) & "/" & xStr(6) & "/" & xStr(7))
                        End If
                        If xStr(3) <> "" Then
                            uStr = uStr & " " & xStr(3)
                        End If
                    End If
                End If
            End If
        End If
        '
        Return uStr
    End Function
    '*********************************************************************
    '取得Size
    '*********************************************************************
    Public Function GetSize(ByVal pStr As String) As String
        Dim uStr As String = pStr
        If CInt(pStr) < 10 Then
            uStr = CStr(CInt(pStr)).Trim
        End If
        '
        Return uStr
    End Function
    '*********************************************************************
    '取得Slider Function
    '*********************************************************************
    Public Function GetSliderFunction(ByVal pStr As String) As String
        Dim xStr As String() = pStr.Split("/")
        Dim xKey As String = "SLDFUN-"
        Dim uStr As String = ""
        '
        If xStr(2) <> "" And xStr(3) <> "" Then     '雙拉頭
            xKey = xKey & "*"
        Else                                        '單拉頭
            xKey = xKey & Mid(xStr(0), 2, 1)
        End If
        '
        Dim uDataBase As New Utility.DataBase
        Dim uCommon As New Utility.Common
        uDataBase.DefaultConnection = New SqlConnection(uCommon.GetConnectionString(strConnectionKey))
        Dim sql As String = "SELECT [Data] FROM M_Referp " & _
                            "WHERE Cat  = '1151' " & _
                            "  AND DKey = '" & xKey & "' "
        Dim dt As Data.DataTable = uDataBase.GetDataTable(sql)
        For Each dr As Data.DataRow In dt.Rows
            uStr = dr("Data").ToString
        Next
        '
        Return uStr
    End Function
    '*********************************************************************
    '取得Slider and Finish
    '*********************************************************************
    Public Function GetSliderFinish(ByVal pStr As String) As String
        Dim xStr As String() = pStr.Split("/")
        Dim uStr As String = ""
        '
        If xStr(0) <> "" And xStr(1) <> "" Then
            If xStr(2) <> "" And xStr(3) <> "" Then     '雙拉頭
                uStr = xStr(0) & " " & xStr(1) & _
                       "/" & _
                       xStr(2) & " " & xStr(3)
            Else                                        '單拉頭
                uStr = xStr(0) & " " & xStr(1)
            End If
        End If
        '
        Return uStr
    End Function

End Class
