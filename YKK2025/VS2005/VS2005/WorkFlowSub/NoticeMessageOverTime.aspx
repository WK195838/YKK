<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NoticeMessageOverTime.aspx.vb" Inherits="NoticeMessageOverTime" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>NoticeMessageOverTime</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/HRWMSG_160101.jpg" style="z-index: 1; left: 10px; position: absolute;
            top: 15px" />
        <asp:CheckBox ID="DReadFlag" style="Z-INDEX: 102; POSITION: absolute; TOP: 633px; LEFT: 9px" runat="server" BackColor="DarkGray" Font-Size="10pt" ForeColor="Red" Text="已閱讀並瞭解今後不顯示" Width="515px" />
        <asp:Button ID="BClose" runat="server" BackColor="DarkGray" ForeColor="Black" Height="28px"
            Style="z-index: 110; left: 8px; position: absolute; top: 656px" Text="關閉視窗" Width="515px" />
    </div>
    </form>
</body>
</html>
