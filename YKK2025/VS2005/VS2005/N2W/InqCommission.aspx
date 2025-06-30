<%@ Page Language="VB" AutoEventWireup="false" CodeFile="InqCommission.aspx.vb" Inherits="InqCommission" %>

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
                <asp:HyperLink ID="LFun07" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="FinalcheckModinqCommission.aspx" Style="z-index: 102; left: 56px;
                    position: absolute; top: 200px" Target="_self" Width="224px">07.最終再檢驗處理報告書-修改</asp:HyperLink>
                <asp:HyperLink ID="LFun06" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="FinalcheckinqCommission.aspx" Style="z-index: 102; left: 56px; position: absolute;
                    top: 176px" Target="_self" Width="224px">06.最終再檢驗處理報告書</asp:HyperLink>
                <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px"
                    NavigateUrl="ExpenseinqCommission.aspx" Style="z-index: 102; left: 56px; position: absolute;
                    top: 152px" Target="_self" Width="224px">05. 交際費申請調閱</asp:HyperLink>
                &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                &nbsp; &nbsp; &nbsp;
                &nbsp;&nbsp;
				<asp:Label id="Label1" style="Z-INDEX: 114; LEFT: 742px; POSITION: absolute; TOP: 48px" runat="server"
					Width="42px" Font-Size="14pt" ForeColor="Navy" Font-Bold="True" Font-Names="Times New Roman">N2W</asp:Label>
                &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
				<asp:hyperlink id="LFun01" style="Z-INDEX: 102; LEFT: 56px; POSITION: absolute; TOP: 128px" runat="server"
					Width="224px" Height="16px" Font-Size="12pt" NavigateUrl="ISMSinqCommission.aspx" Target="_self" Enabled="False">01. ISMS資訊工作日誌調閱</asp:hyperlink>
				<asp:Label id="DSystemTitle" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="260px">Notes To Web 各項調閱</asp:Label></FONT></form>
	</body>
</HTML>