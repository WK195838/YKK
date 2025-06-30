<%@ Page Language="vb" AspCompat="true" AutoEventWireup="false" Codebehind="zzGetSeqNo_T.aspx.vb" Inherits="SPD.GetSeqNo_T" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>GetSeqNo_T</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript">
			function QAPicker(strField)
			{
				window.open('QA1Picker.aspx?field=' + strField,'QA1Popup','width=330,height=20,resizable=yes');
			}
			function NewOption(val)
			{ 
				//alert('The pager number & (' + val + ')') 
				sel=document.getElementById('DQAContent'); 
				sel.options[sel.options.length]=new Option(val,val,true,true); 
			} 
			function QA2Picker(strField)
			{
				window.open('QA2Picker.aspx?field=' + strField,'QA2Popup','width=330,height=20,resizable=yes');
			}
			function NewOption(val)
			{ 
				//alert('The pager number & (' + val + ')') 
				sel=document.getElementById('DQAContent1'); 
				sel.options[sel.options.length]=new Option(val,val,true,true); 
			} 
			function Show(val)
			{ 
				alert('message=(' + val + ')');
			} 
		
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"><FONT face="新細明體">
					<asp:textbox id="DSeqno" style="Z-INDEX: 105; LEFT: 128px; POSITION: absolute; TOP: 80px" runat="server"
						ForeColor="Red" BackColor="Yellow" Width="150px" Height="28px"></asp:textbox>
					<asp:listbox id="DQAContent1" style="Z-INDEX: 116; LEFT: 168px; POSITION: absolute; TOP: 504px"
						runat="server" Height="96px" Width="512px"></asp:listbox><asp:button id="Button4" style="Z-INDEX: 115; LEFT: 32px; POSITION: absolute; TOP: 504px" runat="server"
						Width="128px" Height="32px" Text="QA-2"></asp:button><asp:listbox id="DQAContent" style="Z-INDEX: 114; LEFT: 168px; POSITION: absolute; TOP: 408px"
						runat="server" Width="512px" Height="96px"></asp:listbox><asp:dropdownlist id="Dsize" style="Z-INDEX: 112; LEFT: 16px; POSITION: absolute; TOP: 360px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Width="48px" Height="20px">
						<asp:ListItem Value="3">3</asp:ListItem>
					</asp:dropdownlist><asp:dropdownlist id="DChainType" style="Z-INDEX: 108; LEFT: 64px; POSITION: absolute; TOP: 360px"
						runat="server" ForeColor="Blue" BackColor="Yellow" Width="56px" Height="20px">
						<asp:ListItem Value="cf">CF</asp:ListItem>
					</asp:dropdownlist><asp:dropdownlist id="DBody" style="Z-INDEX: 109; LEFT: 120px; POSITION: absolute; TOP: 360px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Width="80px" Height="20px">
						<asp:ListItem Value="cf">CF</asp:ListItem>
					</asp:dropdownlist><asp:imagebutton id="BSpec" style="Z-INDEX: 111; LEFT: 200px; POSITION: absolute; TOP: 360px" runat="server"
						Width="20px" Height="20px" ImageUrl="Images\arrow1r.jpg"></asp:imagebutton><asp:textbox id="DSpec" style="Z-INDEX: 110; LEFT: 224px; POSITION: absolute; TOP: 360px" runat="server"
						ForeColor="Blue" BackColor="Yellow" Width="320px" Height="20px" TextMode="MultiLine" BorderStyle="Groove"></asp:textbox><asp:datagrid id="DataGrid1" style="Z-INDEX: 106; LEFT: 312px; POSITION: absolute; TOP: 48px"
						runat="server" BackColor="White" Width="462px" Height="300px" BorderStyle="None" AllowPaging="True" BorderWidth="1px" CellPadding="4" AutoGenerateColumns="False" Font-Size="9pt" BorderColor="#CC9966">
						<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
						<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
						<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
						<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
						<Columns>
							<asp:HyperLinkColumn Target="_self" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
								DataTextField="FormName" HeaderText="委託單">
								<HeaderStyle Width="96px"></HeaderStyle>
								<ItemStyle HorizontalAlign="Left"></ItemStyle>
							</asp:HyperLinkColumn>
							<asp:BoundColumn DataField="T_FormSno" ReadOnly="True" HeaderText="單號">
								<HeaderStyle Width="60px"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="ApplyName" ReadOnly="True" HeaderText="委託人">
								<HeaderStyle Width="72px"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="ApplyTime" ReadOnly="True" HeaderText="委託時間">
								<HeaderStyle Width="148px"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
						</Columns>
						<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
							BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
					</asp:datagrid><asp:label id="Label2" style="Z-INDEX: 104; LEFT: 16px; POSITION: absolute; TOP: 88px" runat="server"
						Width="120px" Height="24px">流水號：</asp:label><FONT face="新細明體"><asp:textbox id="DFormNo" style="Z-INDEX: 102; LEFT: 128px; POSITION: absolute; TOP: 40px" runat="server"
							BackColor="Yellow" Width="150px" Height="28px">000001</asp:textbox></FONT><asp:button id="Button1" style="Z-INDEX: 103; LEFT: 64px; POSITION: absolute; TOP: 128px" runat="server"
						Width="128px" Height="40px" Text="Go"></asp:button><asp:label id="Label1" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 40px" runat="server"
						Width="120px" Height="24px">表單號碼：</asp:label><asp:button id="Button2" style="Z-INDEX: 107; LEFT: 168px; POSITION: absolute; TOP: 224px" runat="server"
						Width="120px" Height="40px" Text="Button"></asp:button><asp:button id="Button3" style="Z-INDEX: 113; LEFT: 32px; POSITION: absolute; TOP: 408px" runat="server"
						Width="128px" Height="32px" Text="QA-1"></asp:button>
					<asp:Button id="Button5" style="Z-INDEX: 117; LEFT: 576px; POSITION: absolute; TOP: 358px" runat="server"
						Height="34px" Width="152px" Text="Call Javascript Function"></asp:Button></FONT></FONT></form>
	</body>
</HTML>
