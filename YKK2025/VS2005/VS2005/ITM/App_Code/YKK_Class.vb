Imports Microsoft.VisualBasic

Public Class YKK_Class

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


    '�m��SQL����r
    Function ReplaceString(ByVal pData As String) As String
        ReplaceString = ""
        Dim Str As String = pData
        Str = Replace(Str, "'", "''")
        Str = Replace(Str, ",", "�A")
        Str = Replace(Str, "[���k��]", "")
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


    '---------------------------------------------------------------------------------------------------
    '*****************************************************************
    '**
    '**     �ˬd�W���ɮ�
    '**
    '*****************************************************************
    Function UPFileIsNormal(ByVal UPFile As HtmlInputFile) As Integer
        Dim fileExtension As String     '�ŧi�@���ܼƦs���ɮ׮榡(���ɦW)
        Dim allowedExtensions As String() = {".bmp", ".jpg", ".jpeg", ".gif", ".pdf", ".xls", ".doc", ".ppt"}   '�w�q���\���ɮ׮榡
        Dim i As Integer

        UPFileIsNormal = 0

        fileExtension = IO.Path.GetExtension(UPFile.PostedFile.FileName).ToLower   '���o�ɮ׮榡
        For i = 0 To allowedExtensions.Length - 1           '�v�@�ˬd���\���榡���O�_���W�Ǫ��榡
            If fileExtension = allowedExtensions(i) Then
                UPFileIsNormal = 0
                Exit For
            Else
                UPFileIsNormal = 9010
            End If
        Next
        'Check�W���ɮ�Size
        If UPFileIsNormal = 0 Then
            If UPFile.PostedFile.ContentLength <= 1000 * 1024 Then
                UPFileIsNormal = 0
            Else
                UPFileIsNormal = 9020
            End If
        End If
    End Function

End Class
