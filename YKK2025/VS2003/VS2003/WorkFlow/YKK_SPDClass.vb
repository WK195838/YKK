Public Class YKK_SPDClass

    '判斷是否帳號遺失
    Function Data_Loss(ByVal ID As String) As Integer
        Data_Loss = 0
        If ID = "" Then Data_Loss = 9000
    End Function

    '取得登錄地資訊
    Function Get_LoginPlace(ByVal IP As String) As String
        Get_LoginPlace = "Other"
        If (IP >= "10.245.0.0" And IP <= "10.245.0.99") Or _
              (IP >= "10.245.0.100" And IP <= "10.245.0.255") Then
            Get_LoginPlace = "台北"
        End If

        If (IP >= "10.245.1.0" And IP <= "10.245.1.99") Or _
           (IP >= "10.245.1.100" And IP <= "10.245.1.255") Then
            Get_LoginPlace = "中壢一廠"
        End If

        If (IP >= "10.245.2.0" And IP <= "10.245.2.99") Or _
           (IP >= "10.245.2.100" And IP <= "10.245.2.255") Then
            Get_LoginPlace = "中壢二三廠"
        End If

        If (IP >= "10.245.5.0" And IP <= "10.245.5.99") Or _
           (IP >= "10.245.5.100" And IP <= "10.245.5.255") Then
            Get_LoginPlace = "高雄"
        End If

        If (IP >= "10.245.62.1" And IP <= "10.245.62.30") Then
            Get_LoginPlace = "台南"
        End If

        If (IP >= "10.245.62.31" And IP <= "10.245.62.70") Then
            Get_LoginPlace = "台中"
        End If

        If (IP >= "10.245.62.71" And IP <= "10.245.62.99") Or _
           (IP >= "10.245.62.100" And IP <= "10.245.62.255") Then
            Get_LoginPlace = "楊梅"
        End If
    End Function

    '顯示訊息
    Function ShowMessage(ByVal Msg As String) As String
        ShowMessage = "<" + "Script language=""VBScript"">" + Chr(13) + _
                      "Msgbox """ + Msg + " "",, ""訊息"" " + Chr(13) + _
                      "</" + "Script>"
    End Function

    '整數檢查
    Function IntegerData(ByVal pData As String) As Boolean
        IntegerData = IsNumeric(pData)
    End Function

    '置換SQL關鍵字
    Function ReplaceString(ByVal pData As String) As String
        ReplaceString = ""
        Dim Str As String = pData
        Str = Replace(Str, "'", "''")
        Str = Replace(Str, ",", "，")
        ReplaceString = Str
    End Function


    '上一頁網頁
    Function HistoryBack() As String
        HistoryBack = "<script>History.Back();</script>"
    End Function

    '關閉網頁
    Function Close() As String
        Close = "<script>window.close();</script>"
    End Function

End Class
