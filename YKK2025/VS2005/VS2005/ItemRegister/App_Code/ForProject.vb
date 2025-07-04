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
        Dim xStr As String() = pStr.Split("!")
        Dim yStr As String() = pStr.Split("!")
        Dim zStr As String() = pStr.Split("!")
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
                uStr = uStr & "!" & zStr(i)
            End If
        Next
        '
        Return uStr
    End Function
    '*********************************************************************
    '組合Item Name(1)
    '*********************************************************************
    Public Function SetItemName1(ByVal pType As String, ByVal pStr As String) As String
        Dim xStr As String() = pStr.Split("!")
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
                        If Mid(xStr(1), 1, 2) = "CC" Then      '特殊處理
                            uStr = Replace(xStr(1), "CC", "CH") & xStr(2) & "-" & GetSize(xStr(0))
                        Else
                            uStr = xStr(1) & xStr(2) & "-" & GetSize(xStr(0))
                        End If
                        If GetSliderFunction(xStr(4) & "!" & xStr(5) & "!" & xStr(6) & "!" & xStr(7)) <> "" Then
                            uStr = uStr & GetSliderFunction(xStr(4) & "!" & xStr(5) & "!" & xStr(6) & "!" & xStr(7))
                        End If
                        If GetSliderFinish(xStr(4) & "!" & xStr(5) & "!" & xStr(6) & "!" & xStr(7)) <> "" Then
                            uStr = uStr & " " & GetSliderFinish(xStr(4) & "!" & xStr(5) & "!" & xStr(6) & "!" & xStr(7))
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
        Dim xStr As String() = pStr.Split("!")
        Dim xKey As String = "SLDFUN-"
        Dim uStr As String = ""
        '
        If xStr(2) <> "" And xStr(3) <> "" Then     '雙拉頭
            xKey = xKey & "*"
        Else                                        '單拉頭
            If Mid(xStr(0), 1, 1) = "*" Then        'SLIDER1第一字元為"*"特別處理 2017/6/6
                xKey = xKey & Mid(xStr(0), 1, 1)
            Else
                xKey = xKey & Mid(xStr(0), 2, 1)
            End If
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
        Dim xStr As String() = pStr.Split("!")
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
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyTranData)
    '**     更新交易資料
    '**
    '*****************************************************************
    Public Function ModifyTranData(ByVal pFun As String, _
                                   ByVal pSts As String, _
                                   ByVal pFormNo As String, _
                                   ByVal pFormSno As Integer, _
                                   ByVal pStep As Integer, _
                                   ByVal pSeqNo As Integer, _
                                   ByVal pNowDateTime As String, _
                                   ByVal pDecideID As String) As Integer
        Dim RtnCode As Integer = 0
        Dim SQl As String
        SQl = "Update T_WaitHandle Set "
        SQl = SQl + " Active = '" & "0" & "',"
        SQl = SQl + " Sts = '" & pSts & "',"
        SQl = SQl + " StsDes = '" & "完成" & "',"
        SQl = SQl + " AEndTime = '" & pNowDateTime & "',"
        SQl = SQl + " CompletedTime = '" & pNowDateTime & "',"
        SQl = SQl + " DecideDesc = N'" & "OK(自動核定)" & "',"
        SQl = SQl + " ModifyUser = '" & pDecideID & "',"
        SQl = SQl + " ModifyTime = '" & pNowDateTime & "' "
        SQl = SQl + " Where FormNo  =  '" & pFormNo & "'"
        SQl = SQl + "   And FormSno =  '" & CStr(pFormSno) & "'"
        SQl = SQl + "   And Step    =  '" & CStr(pStep) & "'"
        SQl = SQl + "   And SeqNo   =  '" & CStr(pSeqNo) & "'"
        SQl = SQl + "   And Active =  '1' "
        Try
            Dim uDataBase As Utility.DataBase = GetDataBaseObj()
            uDataBase.ExecuteNonQuery(SQl)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ModifyData)
    '**     更新表單資料
    '**
    '*****************************************************************
    Public Function ModifyData(ByVal pFun As String, _
                               ByVal pSts As String, _
                               ByVal pFormNo As String, _
                               ByVal pFormSno As Integer, _
                               ByVal pNowDateTime As String, _
                               ByVal pDecideID As String, _
                               ByVal pTableName As String) As Integer
        Dim RtnCode As Integer = 0
        Dim sql As String
        sql = "Update " & pTableName & " Set "
        sql &= " Sts = '" & pSts & "',"
        sql &= " CompletedTime = '" & pNowDateTime & "',"
        sql &= " ModifyUser = '" & pDecideID & "',"
        sql &= " ModifyTime = '" & pNowDateTime & "' "
        sql &= " Where FormNo  =  '" & pFormNo & "'"
        sql &= "   And FormSno =  '" & CStr(pFormSno) & "'"
        '
        Try
            Dim uDataBase As Utility.DataBase = GetDataBaseObj()
            uDataBase.ExecuteNonQuery(sql)
        Catch ex As Exception
            RtnCode = 1
        End Try
        '
        Return RtnCode
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceSliderString)
    '**     置換Slider關鍵字
    '**     SK(表示換PIN顏色)，-B(表示引手反面組立)，-E(磁鐵)會出現在Slider的字尾，若同時出現為SK-B，需去除才能取得真正的Puller+Color
    '**
    '*****************************************************************
    Public Function ReplaceSliderString(ByVal pData As String) As String
        Dim str As String = pData
        If Right(str, 2) = "-B" Then
            str = Left(str, Len(str) - 2)
        End If
        If Right(str, 2) = "SK" Then
            str = Left(str, Len(str) - 2)
        End If
        If Right(str, 2) = "-M" Then
            str = Left(str, Len(str) - 2)
        End If
        'ADD-START BY 20250409 BY JOY (ITM)
        If Right(str, 2) = "-Z" Then
            str = Left(str, Len(str) - 2)
        End If
        'ADD-END
        '
        Return str
    End Function
    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(ReplaceColorString)
    '**     置換Color關鍵字
    '**
    '*****************************************************************
    Public Function ReplaceColorString(ByVal pData As String) As String
        Dim i As Integer
        Dim str As String = pData
        Dim key As String()
        '從最長的key開始做搜尋取代
        key = "SLSTC-,SLSBC-,SLSTC,SLSBC,SLS-T,SLS#,SLS-,SLS".Split(",")
        For i = 0 To key.Length - 1
            If InStr(str, key(i)) > 0 Then
                str = Replace(str, key(i), "")
                If InStr(str, "SLS") = 0 Then
                    Exit For
                End If
            End If
        Next
        '
        Return str
    End Function

    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**(GetCompareVal)
    '**     '取得比對欄位的值
    '**
    '*****************************************************************
    Function GetCompareVal(ByVal pStr As String, ByVal pFD As String) As String
        Dim strRes As String
        Dim aryStr() = pStr.Split("!")
        '
        Select Case UCase(pFD)
            Case "SIZE"
                strRes = aryStr(0)
            Case "CHAIN"
                strRes = aryStr(1)
            Case "CLASS"
                strRes = aryStr(2)
            Case "TAPE"
                strRes = aryStr(3)
            Case "SLIDER1"
                strRes = aryStr(4)
            Case "FINISH1"
                strRes = aryStr(5)
            Case "SLIDER2"
                strRes = aryStr(6)
            Case "FINISH2"
                strRes = aryStr(7)
            Case "SREQUEST1"
                strRes = aryStr(8)
            Case "FAMILY"
                strRes = aryStr(9)
            Case Else
                strRes = pFD
        End Select
        Return strRes
    End Function

End Class
