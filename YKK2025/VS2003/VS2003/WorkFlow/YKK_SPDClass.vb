Public Class YKK_SPDClass

    '�P�_�O�_�b����
    Function Data_Loss(ByVal ID As String) As Integer
        Data_Loss = 0
        If ID = "" Then Data_Loss = 9000
    End Function

    '���o�n���a��T
    Function Get_LoginPlace(ByVal IP As String) As String
        Get_LoginPlace = "Other"
        If (IP >= "10.245.0.0" And IP <= "10.245.0.99") Or _
              (IP >= "10.245.0.100" And IP <= "10.245.0.255") Then
            Get_LoginPlace = "�x�_"
        End If

        If (IP >= "10.245.1.0" And IP <= "10.245.1.99") Or _
           (IP >= "10.245.1.100" And IP <= "10.245.1.255") Then
            Get_LoginPlace = "���c�@�t"
        End If

        If (IP >= "10.245.2.0" And IP <= "10.245.2.99") Or _
           (IP >= "10.245.2.100" And IP <= "10.245.2.255") Then
            Get_LoginPlace = "���c�G�T�t"
        End If

        If (IP >= "10.245.5.0" And IP <= "10.245.5.99") Or _
           (IP >= "10.245.5.100" And IP <= "10.245.5.255") Then
            Get_LoginPlace = "����"
        End If

        If (IP >= "10.245.62.1" And IP <= "10.245.62.30") Then
            Get_LoginPlace = "�x�n"
        End If

        If (IP >= "10.245.62.31" And IP <= "10.245.62.70") Then
            Get_LoginPlace = "�x��"
        End If

        If (IP >= "10.245.62.71" And IP <= "10.245.62.99") Or _
           (IP >= "10.245.62.100" And IP <= "10.245.62.255") Then
            Get_LoginPlace = "����"
        End If
    End Function

    '��ܰT��
    Function ShowMessage(ByVal Msg As String) As String
        ShowMessage = "<" + "Script language=""VBScript"">" + Chr(13) + _
                      "Msgbox """ + Msg + " "",, ""�T��"" " + Chr(13) + _
                      "</" + "Script>"
    End Function

    '����ˬd
    Function IntegerData(ByVal pData As String) As Boolean
        IntegerData = IsNumeric(pData)
    End Function

    '�m��SQL����r
    Function ReplaceString(ByVal pData As String) As String
        ReplaceString = ""
        Dim Str As String = pData
        Str = Replace(Str, "'", "''")
        Str = Replace(Str, ",", "�A")
        ReplaceString = Str
    End Function


    '�W�@������
    Function HistoryBack() As String
        HistoryBack = "<script>History.Back();</script>"
    End Function

    '��������
    Function Close() As String
        Close = "<script>window.close();</script>"
    End Function

End Class
