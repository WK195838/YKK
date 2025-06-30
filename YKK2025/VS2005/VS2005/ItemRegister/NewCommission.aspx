<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NewCommission.aspx.vb" Inherits="NewCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
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
					ImageUrl="Images\NewCommission.jpg" Width="780px" Height="500px"></asp:image>
                <asp:HyperLink ID="LFun11" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="ItemRegisterZIPSheet.aspx" Style="z-index: 106; left: 232px; position: absolute;
                    top: 128px" Target="_self" Width="160px">11. ZIP登錄申請單</asp:HyperLink>
				<asp:hyperlink id="LFun15" style="Z-INDEX: 115; LEFT: 232px; POSITION: absolute; TOP: 256px" runat="server"
					Height="16px" Width="160px" Font-Size="12pt" Enabled="False" Target="_self" NavigateUrl="BoardEdit.aspx">15. </asp:hyperlink>
				<asp:Label id="Label1" style="Z-INDEX: 114; LEFT: 742px; POSITION: absolute; TOP: 48px" runat="server"
					Width="42px" Font-Size="14pt" ForeColor="Navy" Font-Bold="True" Font-Names="Times New Roman">IRW</asp:Label>
				<asp:hyperlink id="LFun33" style="Z-INDEX: 113; LEFT: 600px; POSITION: absolute; TOP: 192px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" BackColor="White"
					Enabled="False">33. </asp:hyperlink>
				<asp:hyperlink id="LFun32" style="Z-INDEX: 112; LEFT: 600px; POSITION: absolute; TOP: 160px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">32. </asp:hyperlink>
				<asp:hyperlink id="LFun31" style="Z-INDEX: 111; LEFT: 600px; POSITION: absolute; TOP: 128px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">31. ITEM未受注報告書</asp:hyperlink>
				<asp:hyperlink id="LFun21" style="Z-INDEX: 110; LEFT: 416px; POSITION: absolute; TOP: 128px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">21. 單價承認申請書</asp:hyperlink>
				<asp:hyperlink id="LFun14" style="Z-INDEX: 108; LEFT: 232px; POSITION: absolute; TOP: 224px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">14. </asp:hyperlink>
				<asp:hyperlink id="LFun13" style="Z-INDEX: 107; LEFT: 232px; POSITION: absolute; TOP: 192px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="ItemRegisterCHSheet.aspx" Target="_self" Enabled="False">13. CH登錄申請單</asp:hyperlink>
				<asp:hyperlink id="LFun12" style="Z-INDEX: 106; LEFT: 232px; POSITION: absolute; TOP: 160px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="ItemRegisterSLDSheet.aspx" Target="_self"
					Enabled="False">12.  SLD登錄申請單</asp:hyperlink>
				<asp:hyperlink id="LFun04" style="Z-INDEX: 105; LEFT: 56px; POSITION: absolute; TOP: 224px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False"
					Visible="False">04. </asp:hyperlink>
				<asp:hyperlink id="LFun03" style="Z-INDEX: 104; LEFT: 56px; POSITION: absolute; TOP: 192px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">03. </asp:hyperlink>
				<asp:hyperlink id="LFun02" style="Z-INDEX: 103; LEFT: 56px; POSITION: absolute; TOP: 160px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_self" Enabled="False">02. SLD登錄申請單</asp:hyperlink>
				<asp:hyperlink id="LFun01" style="Z-INDEX: 102; LEFT: 56px; POSITION: absolute; TOP: 128px" runat="server"
					Width="160px" Height="16px" Font-Size="12pt" NavigateUrl="ItemRegisterSheet.aspx" Target="_self" Enabled="False">01. 登錄申請單</asp:hyperlink>
				<asp:Label id="DSystemTitle" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="260px">Wave's Item Code各項申請</asp:Label></FONT></form>
	</body>
</HTML>