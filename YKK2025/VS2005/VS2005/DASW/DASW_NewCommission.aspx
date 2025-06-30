<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DASW_NewCommission.aspx.vb" Inherits="DASW_NewCommission" %>

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
			<FONT face="新細明體">
				<asp:image id="DSheet1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					ImageUrl="Images\DASW_NewCommission.jpg" Width="780px" Height="500px"></asp:image>
                <asp:HyperLink ID="LFun11" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="OverTimeSheet.aspx" Style="z-index: 106; left: 18px; position: absolute;
                    top: 104px" Target="_self" Width="160px">報廢處理申請表</asp:HyperLink>
                <asp:HyperLink ID="LFun12" runat="server" Font-Size="12pt" Height="16px"
                    NavigateUrl="DASW_NOCopy.aspx" Style="z-index: 106; left: 18px; position: absolute;
                    top: 140px" Target="_self" Width="160px" Enabled="False">報廢扣帳明細上傳</asp:HyperLink>
                &nbsp;
				<asp:Label id="Label1" style="Z-INDEX: 114; LEFT: 708px; POSITION: absolute; TOP: 48px" runat="server"
					Width="76px" Font-Size="14pt" ForeColor="Navy" Font-Bold="True" Font-Names="Times New Roman">DASW</asp:Label>
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
				<asp:Label id="DSystemTitle" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="200px">報廢申請</asp:Label></FONT></form>
	</body>
</HTML>