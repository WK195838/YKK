<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MessagePage.aspx.vb" Inherits="MessagePage" aspCompat="True"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>訊息</title>
 
    
</head>
	<body MS_POSITIONING="GridLayout" onLoad="closeWin()" scroll="no">





		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
                <asp:TextBox ID="FocusBox" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                    ForeColor="White" Style="left: 0px; position: absolute; top: 0px"></asp:TextBox>
                &nbsp;
				<asp:label id="Label1" style="Z-INDEX: 103; LEFT: 24px; POSITION: absolute; TOP: 120px" runat="server"
					Height="20px" Width="48px">訊息：</asp:label><asp:textbox id="Message" style="Z-INDEX: 104; LEFT: 72px; POSITION: absolute; TOP: 120px" runat="server"
					Height="104px" Width="504px" BorderStyle="Groove" TextMode="MultiLine" Font-Size="10pt" ReadOnly="True">fffff</asp:textbox>
                <asp:Image ID="MessageTop" runat="server" Height="112px" ImageUrl="~/Images/MailTop2.jpg" style="LEFT: 3px; POSITION: absolute; TOP: 0px"
                    Width="584px" /><asp:Image ID="MessageBottom" runat="server" Height="36px" ImageUrl="Images\MailBottom.jpg" style="LEFT: 8px; POSITION: absolute; TOP: 232px"
                    Width="582px" />
                </FONT></form>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
            <div id="second">
</div>

<script language="javascript" type="text/javascript"> 
var second = 2;
 
function closeWin(){ 
if (second != 0){ 
second -= 1; 
document.getElementById('second').innerHTML = '<strong><font face="Arial" font color=#cc33ff size="7" top="580">視窗於'+second+'秒後自動轉換網頁'; 
 
} 
else{ 
 
return 
} 
setTimeout("closeWin()",1000) 
}
</script> 
  

	
	</body>
</html>
  