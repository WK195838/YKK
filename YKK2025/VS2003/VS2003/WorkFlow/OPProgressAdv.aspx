<%@ Page Language="vb" AutoEventWireup="false" Codebehind="OPProgressAdv.aspx.vb" Inherits="SPD.OPProgressAdv"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�i���i�׬d��</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="javascript" type="text/javascript">
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���"></FONT>
			<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 117; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">�z�ﶵ�ءG</DIV>
			<asp:dropdownlist id="DActive" style="Z-INDEX: 118; LEFT: 400px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="1" Selected="True">�ݳB�z</asp:ListItem>
				<asp:ListItem Value="0">����</asp:ListItem>
			</asp:dropdownlist>
			<asp:datagrid id="DataGrid1" style="Z-INDEX: 119; LEFT: 8px; POSITION: absolute; TOP: 104px" runat="server"
				Height="300px" Width="640px" BackColor="White" BorderStyle="None" Font-Size="9pt" AllowPaging="True"
				BorderColor="#CC9966" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False">
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="�u�{��T">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="FormName" HeaderText="�e�U��">
						<HeaderStyle Width="96px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="FormSno" ReadOnly="True" HeaderText="�渹">
						<HeaderStyle Width="48px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="�u�{">
						<HeaderStyle Width="160px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="���O">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="�e�U�H">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="�֩w�H">
						<HeaderStyle Width="84px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
				</Columns>
				<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid>
			<asp:dropdownlist id="DStepName" style="Z-INDEX: 120; LEFT: 264px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="132px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="0">�q��</asp:ListItem>
				<asp:ListItem Value="1">�֩w</asp:ListItem>
			</asp:dropdownlist>
			<asp:dropdownlist id="DFormName" style="Z-INDEX: 121; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue" AutoPostBack="True"></asp:dropdownlist>
			<asp:dropdownlist id="DSts" style="Z-INDEX: 122; LEFT: 504px; POSITION: absolute; TOP: 8px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="1" Selected="True">����</asp:ListItem>
				<asp:ListItem Value="0">���`</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 123; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">�e�U����G</DIV>
			<asp:textbox id="DSDate" style="Z-INDEX: 124; LEFT: 96px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DSDate</asp:textbox>
			<asp:button id="BSDate" style="Z-INDEX: 125; LEFT: 184px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="�z�ﶵ�ءG" style="DISPLAY: inline; Z-INDEX: 126; LEFT: 208px; WIDTH: 16px; COLOR: #0000ff; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout">��</DIV>
			<asp:textbox id="DEDate" style="Z-INDEX: 127; LEFT: 232px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" BorderStyle="Groove" ReadOnly="True">DEDate</asp:textbox>
			<asp:button id="BEDate" style="Z-INDEX: 128; LEFT: 320px; POSITION: absolute; TOP: 40px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="�ƧǡG" style="DISPLAY: inline; Z-INDEX: 129; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 72px; HEIGHT: 24px"
				ms_positioning="FlowLayout">�ƧǡG</DIV>
			<asp:dropdownlist id="DSortKey" style="Z-INDEX: 130; LEFT: 96px; POSITION: absolute; TOP: 72px" runat="server"
				Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="FormSno" Selected="True">���渹</asp:ListItem>
				<asp:ListItem Value="ApplyTime">�e�U�ɶ�</asp:ListItem>
				<asp:ListItem Value="ReceiptTime">����ɶ�</asp:ListItem>
				<asp:ListItem Value="ReadTimeLimit">�\Ū����</asp:ListItem>
				<asp:ListItem Value="FirstReadTime">�Ĥ@���\Ū�ɶ�</asp:ListItem>
				<asp:ListItem Value="LastReadTime">�̫�@���\Ū�ɶ�</asp:ListItem>
				<asp:ListItem Value="BStartTime">�w�w�}�l�ɶ�</asp:ListItem>
				<asp:ListItem Value="BEndTime">�w�w�����ɶ�</asp:ListItem>
				<asp:ListItem Value="AStartTime">��ڶ}�l�ɶ�</asp:ListItem>
				<asp:ListItem Value="AEndTime">��ڧ����ɶ�</asp:ListItem>
			</asp:dropdownlist>
			<asp:dropdownlist id="DSort" style="Z-INDEX: 131; LEFT: 264px; POSITION: absolute; TOP: 72px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ASC" Selected="True">�@��</asp:ListItem>
				<asp:ListItem Value="DESC">����</asp:ListItem>
			</asp:dropdownlist>
			<asp:button id="Go" style="Z-INDEX: 132; LEFT: 368px; POSITION: absolute; TOP: 72px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button>
		</form>
	</body>
</HTML>
