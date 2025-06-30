<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OPInqCommission_Manuf.aspx.vb" Inherits="OPInqCommission_Manuf" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>調閱製造委託書</title>
</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體"></FONT>
			<DIV title="篩選項目：" style="DISPLAY: inline; Z-INDEX: 100; LEFT: 8px; WIDTH: 88px; COLOR: #0000ff; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout">篩選項目：</DIV>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 122; LEFT: 920px; POSITION: absolute; TOP: 36px" runat="server"
				Height="21px" Width="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:dropdownlist id="DProgress" style="Z-INDEX: 121; LEFT: 352px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">開發中</asp:ListItem>
				<asp:ListItem Value="2">開發完成</asp:ListItem>
			</asp:dropdownlist><asp:dropdownlist id="DSts" style="Z-INDEX: 120; LEFT: 480px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px">
				<asp:ListItem Value="ALL" Selected="True">ALL</asp:ListItem>
				<asp:ListItem Value="1">OK</asp:ListItem>
				<asp:ListItem Value="2">NG</asp:ListItem>
				<asp:ListItem Value="3">取消</asp:ListItem>
			</asp:dropdownlist><asp:textbox id="DCpsc" style="Z-INDEX: 118; LEFT: 736px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DCpsc</asp:textbox><asp:textbox id="DNo" style="Z-INDEX: 116; LEFT: 96px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DNo</asp:textbox><asp:textbox id="DSpec" style="Z-INDEX: 115; LEFT: 608px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DSpec</asp:textbox>
			<asp:textbox id="DSliderCode" style="Z-INDEX: 113; LEFT: 480px; POSITION: absolute; TOP: 35px"
				runat="server" ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DSliderCode</asp:textbox><asp:textbox id="DMapNo" style="Z-INDEX: 112; LEFT: 352px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DMapNo</asp:textbox><asp:textbox id="DPerson" style="Z-INDEX: 111; LEFT: 737px; POSITION: absolute; TOP: 8px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DPerson</asp:textbox><asp:textbox id="DBuyer" style="Z-INDEX: 101; LEFT: 224px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="Yellow" Height="20px" Width="120px" BorderStyle="Groove" ReadOnly="True">DBuyer</asp:textbox><asp:dropdownlist id="DDivision" style="Z-INDEX: 102; LEFT: 609px; POSITION: absolute; TOP: 8px"
				runat="server" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="120px"></asp:dropdownlist><asp:dropdownlist id="DFormName" style="Z-INDEX: 103; LEFT: 96px; POSITION: absolute; TOP: 8px" runat="server"
				AutoPostBack="True" ForeColor="Blue" BackColor="Yellow" Height="40px" Width="248px"></asp:dropdownlist>
			<asp:button id="Go" style="Z-INDEX: 110; LEFT: 867px; POSITION: absolute; TOP: 35px" runat="server"
				ForeColor="Blue" BackColor="White" Height="24px" Width="40px" Text="Go"></asp:button>

        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 70px" 
 BorderStyle="None" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" >
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle> 

            <Columns>
                <asp:TemplateField>
                    <ItemTemplate>  
                        <asp:Image ID="Image1" runat="server" Height="100" Width="115" ImageUrl='<%# Container.DataItem("ImageUrl") %>'  />  
                    </ItemTemplate>  
                </asp:TemplateField>

                    <asp:HyperLinkField HeaderText="圖樣" Target="_blank" DataTextField="Field1" DataNavigateUrlFields="ViewURL" DataNavigateUrlFormatString="{0}" >
                        <HeaderStyle Width="90px" />
                    </asp:HyperLinkField>
					<asp:BoundField DataField="Field2" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field3" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field4" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field5" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field6" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field7" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field8" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					
                    <asp:HyperLinkField HeaderText="履歷" Target="_blank" DataTextField="WorkFlow" DataNavigateUrlFields="OPURL" DataNavigateUrlFormatString="{0}" >
                        <HeaderStyle Width="30px" />
                    </asp:HyperLinkField>

            </Columns>
        </asp:GridView>
        
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" style="Z-INDEX: 114; LEFT: 8px; POSITION: absolute; TOP: 70px" 
 BorderStyle="None" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" >
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle> 

            <Columns>
                    <asp:HyperLinkField HeaderText="圖樣" Target="_blank" DataTextField="Field1" DataNavigateUrlFields="ViewURL" DataNavigateUrlFormatString="{0}" >
                        <HeaderStyle Width="90px" />
                    </asp:HyperLinkField>
					<asp:BoundField DataField="Field2" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field3" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field4" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field5" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field6" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field7" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					<asp:BoundField DataField="Field8" ReadOnly="True">
                        <HeaderStyle Width="110px" />
					</asp:BoundField>
					
                    <asp:HyperLinkField HeaderText="履歷" Target="_blank" DataTextField="WorkFlow" DataNavigateUrlFields="OPURL" DataNavigateUrlFormatString="{0}" >
                        <HeaderStyle Width="30px" />
                    </asp:HyperLinkField>

            </Columns>
        </asp:GridView>
        
        </form>
	</body>

</html>
