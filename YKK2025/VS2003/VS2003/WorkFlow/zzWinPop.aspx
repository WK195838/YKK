<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="zzWinPop.aspx.vb" Inherits="SPD.WinPop"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>WinPop</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			var wPop;
			function aaa()
			{
				wPop=window.open('WinPop1.aspx','aaa','width=400,height=400,resizable=yes');
				setTimeout("SendToChild('" + document.Form1.TextBox1.value + "')",500);
				
			}
		    function SendToChild(data){
			alert(data);
				wPop.document.Form2.TextBox1.value=data;
			}
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:textbox id="TextBox1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Width="168px" Height="96px" TextMode="MultiLine">111aaa</asp:textbox>
				<asp:hyperlink id="Hyperlink1" style="Z-INDEX: 116; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">���Y�ӥ�</asp:hyperlink>
				<asp:hyperlink id="LContact" style="Z-INDEX: 115; LEFT: 8px; POSITION: absolute; TOP: 48px" runat="server"
					Height="8px" Width="80px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">�~�ȳs����</asp:hyperlink>
				<asp:hyperlink id="LSliderDetail" style="Z-INDEX: 112; LEFT: 536px; POSITION: absolute; TOP: 1120px"
					runat="server" Height="8px" Width="64px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">���Y�ӥ�</asp:hyperlink>
				<asp:hyperlink id="LOPContact" style="Z-INDEX: 114; LEFT: 104px; POSITION: absolute; TOP: 8px"
					runat="server" Height="8px" Width="80px" Font-Size="12pt" NavigateUrl="BoardEdit.aspx" Target="_blank">�u�{�s����</asp:hyperlink>
				<asp:textbox id="DStatus" style="Z-INDEX: 113; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="64px" Width="203px" Font-Size="10pt" BorderStyle="Groove" BackColor="Red" ForeColor="White"
					ReadOnly="True">�ק�u�{�i�椤(xxxxxxxxxxxxxxxxxx)</asp:textbox>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 110; LEFT: 24px; POSITION: absolute; TOP: 296px"
					runat="server" Height="381px" Width="712px" BackColor="White" AutoGenerateColumns="False"
					CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" AllowPaging="True" BorderStyle="None"
					Font-Size="9pt">
					<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
					<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
					<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
					<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
					<Columns>
						<asp:BoundColumn DataField="DelaySts" ReadOnly="True" HeaderText="�i��">
							<HeaderStyle Width="24px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StsDesc" ReadOnly="True" HeaderText="���A">
							<HeaderStyle Width="54px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormName" ReadOnly="True" HeaderText="�e�U��">
							<HeaderStyle Width="72px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FormSno" ReadOnly="True" HeaderText="�渹">
							<HeaderStyle Width="24px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Step" ReadOnly="True" HeaderText="No">
							<HeaderStyle Width="12px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="StepNameDesc" ReadOnly="True" HeaderText="�u�{">
							<HeaderStyle Width="104px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="FlowTypeDesc" ReadOnly="True" HeaderText="���O">
							<HeaderStyle Width="24px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="DecideName" ReadOnly="True" HeaderText="���">
							<HeaderStyle Width="36px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
					</Columns>
					<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
						BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>
				<asp:button id="Button6" style="Z-INDEX: 108; LEFT: 360px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Text="QA2"></asp:button>
				<asp:button id="Button5" style="Z-INDEX: 107; LEFT: 272px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Text="QA1"></asp:button><asp:button id="Button4" style="Z-INDEX: 106; LEFT: 184px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Text="Price"></asp:button><asp:button id="Button3" style="Z-INDEX: 105; LEFT: 96px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Text="Sample"></asp:button><asp:button id="Button1" style="Z-INDEX: 101; LEFT: 104px; POSITION: absolute; TOP: 120px" runat="server"
					Width="64px" Text="Button"></asp:button><asp:textbox id="TextBox2" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 168px" runat="server"
					Width="304px" Height="64px" TextMode="MultiLine">[3,CF,DA,h6,100p]</asp:textbox><asp:button id="Button2" style="Z-INDEX: 103; LEFT: 8px; POSITION: absolute; TOP: 256px" runat="server"
					Width="80px" Text="Slider"></asp:button><asp:textbox id="TextBox3" style="Z-INDEX: 104; LEFT: 328px; POSITION: absolute; TOP: 216px"
					runat="server" Width="128px"></asp:textbox>
				<asp:Button id="Button7" style="Z-INDEX: 109; LEFT: 272px; POSITION: absolute; TOP: 336px" runat="server"
					Height="32px" Width="104px" Text="call java"></asp:Button></FONT></form>
	</body>
</HTML>
