<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InqCommissionSales.aspx.vb" Inherits="InqCommissionSales" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<HEAD>
		<title>調閱資料</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:image id="DSheet1" style="Z-INDEX: 100; LEFT: 16px; POSITION: absolute; TOP: 16px" runat="server"
					ImageUrl="Images\NewCommission.jpg" Width="780px" Height="500px"></asp:image>
                <asp:HyperLink ID="LFun04" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="CustomerInfoinqCommission.aspx" Style="z-index: 102; left: 56px;
                    position: absolute; top: 200px" Target="_self" Width="224px">04.顧客變更</asp:HyperLink>
                <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="DiscountinqCommission.aspx" Style="z-index: 102; left: 56px; position: absolute;
                    top: 152px" Target="_self" Width="224px">02. 顧客折扣單價</asp:HyperLink>
                <asp:HyperLink ID="LFun03" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="CustomerInfoinqCommission.aspx" Style="z-index: 102; left: 56px; position: absolute;
                    top: 176px" Target="_self" Width="224px">03.顧客建檔</asp:HyperLink>
                &nbsp; &nbsp; &nbsp;
                &nbsp;&nbsp;
				<asp:Label id="Label1" style="Z-INDEX: 114; LEFT: 742px; POSITION: absolute; TOP: 48px" runat="server"
					Width="42px" Font-Size="14pt" ForeColor="Navy" Font-Bold="True" Font-Names="Times New Roman">N2W</asp:Label>
                &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
                &nbsp;
				<asp:Label id="DSystemTitle" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="260px">Notes To Web 各項調閱</asp:Label></FONT></form>
	</body>
</HTML>