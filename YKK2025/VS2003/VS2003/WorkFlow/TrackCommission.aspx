<%@ Page Language="vb" AutoEventWireup="false" Codebehind="TrackCommission.aspx.vb" Inherits="SPD.TrackCommission"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�l�ܸ�R</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function calendarPicker(strField)
			{
				window.open('DatePicker.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			function MapPicker(strField)
			{
				window.open('MapPicker.aspx?field=' + strField,'MapPopup','width=168,height=328,resizable=yes');
			}
		    function OpenSimulationSheet(URL) {
				window.open(URL,'Simulation','width=630,height=580,resizable=yes,scrollbars=yes, menubar=yes');
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���"></FONT>
			<DIV title="�z�ﶵ�ءG" style="Z-INDEX: 100; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 8px; LEFT: 8px"
				ms_positioning="FlowLayout">�z�ﶵ�ءG</DIV>
			<asp:datagrid id="DataGrid1" style="Z-INDEX: 119; POSITION: absolute; TOP: 104px; LEFT: 8px" runat="server"
				PageSize="15" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px"
				BorderColor="#CC9966" Font-Size="9pt" Height="300px" Width="790px" BackColor="White" DataKeyField="ID"
				AllowPaging="True">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="No" HeaderText="�e�UNo">
						<HeaderStyle Width="60px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Person" ReadOnly="True" HeaderText="�e�U�H">
						<HeaderStyle Width="55px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="ApplyTime" ReadOnly="True" HeaderText="�e�U��">
						<HeaderStyle Width="110px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="���O">
						<HeaderStyle Width="50px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepName" ReadOnly="True" HeaderText="�ݳB�z�u�{">
						<HeaderStyle Width="140px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="StepPerson" ReadOnly="True" HeaderText="�u�{���">
						<HeaderStyle Width="55px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="OPTime" ReadOnly="True" HeaderText="�u�{�w�w">
						<HeaderStyle Width="200px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="�i��">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:ButtonColumn Text="���" HeaderText="����" CommandName="Select">
						<HeaderStyle Width="30px"></HeaderStyle>
					</asp:ButtonColumn>
				</Columns>
				<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid><asp:dropdownlist id="DSortKey" style="Z-INDEX: 116; POSITION: absolute; TOP: 72px; LEFT: 96px" runat="server"
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
			</asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 114; POSITION: absolute; TOP: 8px; LEFT: 96px" runat="server"
				Height="40px" Width="165px" BackColor="Yellow" ForeColor="Blue"></asp:dropdownlist><FONT face="�s�ө���"></FONT><asp:dropdownlist id="DFlowType" style="Z-INDEX: 102; POSITION: absolute; TOP: 8px; LEFT: 264px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ALL">ALL</asp:ListItem>
				<asp:ListItem Value="0">�q��</asp:ListItem>
				<asp:ListItem Value="1" Selected="True">�֩w</asp:ListItem>
			</asp:dropdownlist><asp:button id="BEDate" style="Z-INDEX: 113; POSITION: absolute; TOP: 40px; LEFT: 320px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button><asp:textbox id="DEDate" style="Z-INDEX: 112; POSITION: absolute; TOP: 40px; LEFT: 232px" runat="server"
				BorderStyle="Groove" Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DEDate</asp:textbox>
			<DIV title="�z�ﶵ�ءG" style="Z-INDEX: 111; POSITION: absolute; WIDTH: 16px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 40px; LEFT: 208px"
				ms_positioning="FlowLayout">��</DIV>
			<asp:textbox id="DSDate" style="Z-INDEX: 110; POSITION: absolute; TOP: 40px; LEFT: 96px" runat="server"
				BorderStyle="Groove" Height="20px" Width="88px" BackColor="Yellow" ForeColor="Blue" ReadOnly="True">DSDate</asp:textbox><asp:button id="BSDate" style="Z-INDEX: 109; POSITION: absolute; TOP: 40px; LEFT: 184px" runat="server"
				Height="20px" Width="20px" Text="....."></asp:button>
			<DIV title="�z�ﶵ�ءG" style="Z-INDEX: 108; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 40px; LEFT: 8px"
				ms_positioning="FlowLayout">�ӽФ���G</DIV>
			<asp:button id="Go" style="Z-INDEX: 105; POSITION: absolute; TOP: 72px; LEFT: 368px" runat="server"
				Height="24px" Width="40px" BackColor="White" ForeColor="Blue" Text="Go"></asp:button><asp:dropdownlist id="DSort" style="Z-INDEX: 106; POSITION: absolute; TOP: 72px; LEFT: 264px" runat="server"
				Height="40px" Width="100px" BackColor="Yellow" ForeColor="Blue">
				<asp:ListItem Value="ASC">�@��</asp:ListItem>
				<asp:ListItem Value="DESC" Selected="True">����</asp:ListItem>
			</asp:dropdownlist>
			<DIV title="�ƧǡG" style="Z-INDEX: 107; POSITION: absolute; WIDTH: 88px; DISPLAY: inline; HEIGHT: 24px; COLOR: #0000ff; TOP: 72px; LEFT: 8px"
				ms_positioning="FlowLayout">�ƧǡG</DIV>
		</form>
	</body>
</HTML>
