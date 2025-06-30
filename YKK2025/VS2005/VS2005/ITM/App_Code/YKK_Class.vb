Imports Microsoft.VisualBasic

Public Class YKK_Class

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


    '置換SQL關鍵字
    Function ReplaceString(ByVal pData As String) As String
        ReplaceString = ""
        Dim Str As String = pData
        Str = Replace(Str, "'", "''")
        Str = Replace(Str, ",", "，")
        Str = Replace(Str, "[未歸還]", "")
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


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     檢查上傳檔案
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".xls", ".doc", ".ppt"}   '定義允許的檔案格式
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '取得檔案格式
        For i = 0 To allowedExtensions.Length - 1           '逐一檢查允許的格式中是否有上傳的格式
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check上傳檔案Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If
    End Function

End Class
