<%@ Page Language="vb" AutoEventWireup="false" Codebehind="HR_ErrorCard.aspx.vb" Inherits="SPD.HR_ErrorCard"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����d</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 100; POSITION: absolute; TOP: 8px; LEFT: 8px" runat="server"
					AutoPostBack="True" Height="40px" Width="168px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DName" style="Z-INDEX: 106; POSITION: absolute; TOP: 8px; LEFT: 176px" runat="server"
					Height="40px" Width="168px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:dropdownlist id="DCardType" style="Z-INDEX: 105; POSITION: absolute; TOP: 8px; LEFT: 344px" runat="server"
					Height="40px" Width="72px" ForeColor="Blue" BackColor="Yellow">
					<asp:ListItem Value="0" Selected="True">ALL</asp:ListItem>
					<asp:ListItem Value="1">���`</asp:ListItem>
					<asp:ListItem Value="2">���`</asp:ListItem>
				</asp:dropdownlist><asp:datagrid id="DataGrid1" style="Z-INDEX: 104; POSITION: absolute; TOP: 40px; LEFT: 8px" runat="server"
					Height="100px" Width="790px" BackColor="White" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px"
					BorderColor="#CC9966" Font-Size="9pt" PageSize="15" BorderStyle="None">
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="DateDesc" ReadOnly="True" HeaderText="��d��">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="WeekDesc" ReadOnly="True" HeaderText="�P��">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="NameDesc" ReadOnly="True" HeaderText="�m�W">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="HRWDivName" ReadOnly="True" HeaderText="����">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="JobName" ReadOnly="True" HeaderText="¾��">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeA" ReadOnly="True" HeaderText="�W�Z��d">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="TimeB" ReadOnly="True" HeaderText="�U�Z��d">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OTURL" DataNavigateUrlFormatString="{0}" DataTextField="OTDesc"
							HeaderText="�[�Z">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="VAURL" DataNavigateUrlFormatString="{0}" DataTextField="VADesc"
							HeaderText="�а�">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="AwayURL" DataNavigateUrlFormatString="{0}"
							DataTextField="AwayDesc" HeaderText="�~�X">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="AworkURL" DataNavigateUrlFormatString="{0}"
							DataTextField="AworkDesc" HeaderText="�ɤu">
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:HyperLinkColumn>
						<asp:BoundColumn DataField="Remark" ReadOnly="True" HeaderText="�Ƶ�">
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid><asp:imagebutton id="BExcel" style="Z-INDEX: 103; POSITION: absolute; TOP: 8px; LEFT: 472px" runat="server"
					Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:button id="Go" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 424px" runat="server"
					Height="24px" Width="40px" ForeColor="Blue" BackColor="White" Text="Go"></asp:button></FONT></form>
	</body>
</HTML>
