<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MPM_NewCommission.aspx.vb" Inherits="MPM_NewCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<HTML>
	<HEAD>
		<title>新委託</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
				<asp:Label id="DSystemTitle" style="Z-INDEX: 100; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="200px">機械加工各項申請</asp:Label>
                <asp:DropDownList ID="DFormNo" runat="server" Style="z-index: 101; left: 15px;
                    position: absolute; top: 67px" Width="199px">
                </asp:DropDownList>
                <asp:TextBox ID="TextBox1" runat="server" Height="85px" ReadOnly="True" Style="z-index: 102;
                    left: 271px; position: absolute; top: 28px" TextMode="MultiLine" Width="250px">先選擇委託單之後點選需要的按鈕  新委託：發行新的委託單</asp:TextBox>
                <asp:Button ID="BNew" runat="server" Style="z-index: 104; left: 101px; position: absolute;
                    top: 112px" Text="新委託" Width="109px" />
            </FONT></form>
	</body>
</HTML>