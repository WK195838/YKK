Public Class SPDClass

    'Mail Body
    Function SetBody(ByVal NowDateTime As String, ByVal MsgType As String, ByVal FormNo As String, ByVal FormSno As String, ByVal OP As String) As String
        Dim str As String
        str = "<table border=0 width=61% cellspacing=0 cellpadding=0>"
        str = str & "<tr><td width=100% colspan=3><p align=center><img border=0 src=http://10.245.0.178/SSD/Images/MailTop.jpg width=584 height=112></td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���G&nbsp;&nbsp;" & "</p></td><td width=80%>" & NowDateTime & "</td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "�T�������G&nbsp;&nbsp;" & "</td><td width=80%>" & MsgType & "</p></td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��G&nbsp;&nbsp;" & "</p></td><td width=80%>" & FormNo & "</td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "�y&nbsp;&nbsp;��&nbsp;&nbsp;���G&nbsp;&nbsp;" & "</p></td><td width=80%>" & FormSno & "</td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "�u&nbsp;&nbsp;�{&nbsp;&nbsp;�W�G&nbsp;&nbsp;" & "</p></td><td width=80%>" & OP & "</td></tr>"
        str = str & "<tr><td width=10%><p align=right>" & "�s&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���G&nbsp;&nbsp;" & "</p></td><td width=80%><a href=" & "http://10.245.0.178/SSD/Default_T.aspx" & " target=_blank>" & "�t�γs��" & "</a></td></tr>"
        str = str & "<tr><td width=100% colspan=3><p align=center><img border=0 src=http://10.245.0.178/SSD/Images/MailBottom.jpg width=582 height=36></td></tr>"
        str = str & "</table>"
        SetBody = str
    End Function

End Class
