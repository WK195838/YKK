<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Progress.aspx.vb" Inherits="SPD.Progress"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Progress</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 114px" runat="server"
					BackColor="White" Height="50px" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px"
					BorderColor="#CC9966" Font-Size="9pt" BorderStyle="None" Width="776px">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="�e�U��">
							<HeaderStyle Width="200px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="DelayURL" DataNavigateUrlFormatString="{0}"
							DataTextField="Delay" HeaderText="����">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="ReadDelayURL" DataNavigateUrlFormatString="{0}"
							DataTextField="ReadDelay" HeaderText="���\Ū">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="NormalURL" DataNavigateUrlFormatString="{0}"
							DataTextField="Normal" HeaderText="���`">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="OKURL" DataNavigateUrlFormatString="{0}" DataTextField="OK"
							HeaderText="�ݢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="NGURL" DataNavigateUrlFormatString="{0}" DataTextField="NG"
							HeaderText="�ܢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="CancelURL" DataNavigateUrlFormatString="{0}"
							DataTextField="Cancel" HeaderText="����">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="OKYURL" DataNavigateUrlFormatString="{0}" DataTextField="OKY"
							HeaderText="�ݢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="NGYURL" DataNavigateUrlFormatString="{0}" DataTextField="NGY"
							HeaderText="�ܢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="CancelYURL" DataNavigateUrlFormatString="{0}"
							DataTextField="CancelY" HeaderText="����">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="OKBYURL" DataNavigateUrlFormatString="{0}"
							DataTextField="OKBY" HeaderText="�ݢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="NGBYURL" DataNavigateUrlFormatString="{0}"
							DataTextField="NGBY" HeaderText="�ܢ�">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="CancelBYURL" DataNavigateUrlFormatString="{0}"
							DataTextField="CancelBY" HeaderText="����">
							<HeaderStyle Width="48px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:imagebutton id="BExcel" style="Z-INDEX: 150; LEFT: 760px; POSITION: absolute; TOP: 40px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:textbox id="DRefresh" style="Z-INDEX: 112; LEFT: 8px; POSITION: absolute; TOP: 88px" runat="server"
					BackColor="White" Height="24px" BorderStyle="None" Width="160px" ForeColor="Blue" ReadOnly="True">DDateTime</asp:textbox><asp:hyperlink id="Hyperlink1" style="Z-INDEX: 111; LEFT: 528px; POSITION: absolute; TOP: 8px"
					runat="server" Height="8px" Font-Size="12pt" Width="99px" Target="_self" NavigateUrl="Developing_List.aspx">�}�o�i�פ@��</asp:hyperlink><asp:hyperlink id="LDevelop" style="Z-INDEX: 110; LEFT: 636px; POSITION: absolute; TOP: 8px" runat="server"
					Height="8px" Font-Size="12pt" Width="152px" Target="_self" NavigateUrl="Developing_Suppiler.aspx">�~�`�O�}�o�i�פ@��</asp:hyperlink>
				<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 108; LEFT: 8px; WIDTH: 296px; COLOR: #000000; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�}�o���G�w�w��������\Ū������</DIV>
				<asp:textbox id="DDateTime" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 64px" runat="server"
					BackColor="White" Height="24px" BorderStyle="None" Width="160px" ForeColor="Blue" ReadOnly="True">DDateTime</asp:textbox>
				<DIV style="DISPLAY: inline; Z-INDEX: 104; LEFT: 180px; WIDTH: 151px; COLOR: white; POSITION: absolute; TOP: 64px; HEIGHT: 49px; BACKGROUND-COLOR: #cc6600; TEXT-ALIGN: center"
					align="center" ms_positioning="FlowLayout">�}�o��</DIV>
				<DIV style="DISPLAY: inline; Z-INDEX: 107; LEFT: 332px; WIDTH: 450px; COLOR: white; POSITION: absolute; TOP: 64px; HEIGHT: 24px; BACKGROUND-COLOR: #cc6600; TEXT-ALIGN: center"
					align="center" ms_positioning="FlowLayout">����</DIV>
				<DIV style="DISPLAY: inline; Z-INDEX: 103; LEFT: 332px; WIDTH: 149px; COLOR: white; POSITION: absolute; TOP: 89px; HEIGHT: 24px; BACKGROUND-COLOR: #cc6600; TEXT-ALIGN: center"
					align="center" ms_positioning="FlowLayout">���</DIV>
				<DIV style="DISPLAY: inline; Z-INDEX: 102; LEFT: 482px; WIDTH: 149px; COLOR: white; POSITION: absolute; TOP: 89px; HEIGHT: 24px; BACKGROUND-COLOR: #cc6600; TEXT-ALIGN: center"
					align="center" ms_positioning="FlowLayout">��~(�t���)</DIV>
				<DIV style="DISPLAY: inline; Z-INDEX: 101; LEFT: 632px; WIDTH: 150px; COLOR: white; POSITION: absolute; TOP: 89px; HEIGHT: 24px; BACKGROUND-COLOR: #cc6600; TEXT-ALIGN: center"
					align="center" ms_positioning="FlowLayout">��~�H�e</DIV>
				<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 109; LEFT: 320px; WIDTH: 192px; COLOR: #000000; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�����G�}�o������</DIV>
			</FONT>
		</form>
	</body>
</HTML>
