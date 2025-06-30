<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_WorkTime.aspx.vb" Inherits="HR_WorkTime" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<HTML>
	<HEAD>
		<title>出缺勤月報</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:dropdownlist id="DDivision" style="Z-INDEX: 100; LEFT: 216px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="Yellow" ForeColor="Blue" Width="168px" Height="40px" AutoPostBack="True"></asp:dropdownlist>
				<asp:dropdownlist id="DMonth" style="Z-INDEX: 107; LEFT: 112px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist>
				<asp:dropdownlist id="DYear" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 8px" runat="server"
					Height="40px" Width="100px" ForeColor="Blue" BackColor="Yellow"></asp:dropdownlist><asp:imagebutton id="BExcel" style="Z-INDEX: 103; LEFT: 614px; POSITION: absolute; TOP: 8px" runat="server"
					Width="21px" Height="21px" ImageUrl="Images\msexcel.gif"></asp:imagebutton><asp:button id="Go" style="Z-INDEX: 102; LEFT: 566px; POSITION: absolute; TOP: 8px" runat="server"
					BackColor="White" ForeColor="Blue" Width="40px" Height="24px" Text="Go"></asp:button>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
                    BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
                    Style="z-index: 114; left: 10px; position: absolute; top: 40px" Width="900px" ShowFooter="True">
                    <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
                    <Columns>
                        <asp:BoundField DataField="NameDesc" HeaderText="姓名" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="HRWDivName" HeaderText="部門" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="DateDesc" HeaderText="出勤日" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="WeekDesc" HeaderText="星期" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TimeA" HeaderText="上班刷卡" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TimeB" HeaderText="下班刷卡" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:HyperLinkField DataNavigateUrlFields="OTURL" DataNavigateUrlFormatString="{0}"
                            DataTextField="OTDesc" HeaderText="加班" Target="_blank" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:HyperLinkField>

                        <asp:HyperLinkField DataNavigateUrlFields="VAURL" DataNavigateUrlFormatString="{0}"
                            DataTextField="VADesc" HeaderText="請假" Target="_blank" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:HyperLinkField>

                        <asp:HyperLinkField DataNavigateUrlFields="AwayURL" DataNavigateUrlFormatString="{0}"
                            DataTextField="AwayDesc" HeaderText="外出" Target="_blank" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:HyperLinkField>

                        <asp:HyperLinkField DataNavigateUrlFields="AworkURL" DataNavigateUrlFormatString="{0}"
                            DataTextField="AworkDesc" HeaderText="補工" Target="_blank" >
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:HyperLinkField>
                    
                        <asp:BoundField DataField="Remark" HeaderText="備註" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>

                        <asp:BoundField DataField="SOverTime" HeaderText="加班時數" ReadOnly="True" >
                        </asp:BoundField>
                        <asp:BoundField DataField="STimeOff" HeaderText="請假日數" ReadOnly="True" >
                        </asp:BoundField>
                        <asp:BoundField DataField="SAway" HeaderText="外出日數" ReadOnly="True" >
                        </asp:BoundField>
                        <asp:BoundField DataField="SAddWork" HeaderText="補工時數" ReadOnly="True" >
                        </asp:BoundField>
                        <asp:BoundField DataField="ABS_FormSno1" HeaderText="ABS_FormSno1" ReadOnly="True" >
                        </asp:BoundField>
                        <asp:BoundField DataField="ABS_FormSno2" HeaderText="ABS_FormSno2" ReadOnly="True" >
                        </asp:BoundField>
 
                    </Columns>
                </asp:GridView>
                <asp:DropDownList ID="DName" runat="server" BackColor="Yellow" ForeColor="Blue" Height="40px"
                    Style="z-index: 106; left: 389px; position: absolute; top: 8px" Width="168px">
                </asp:DropDownList>
            </FONT></form>
	</body>
</HTML>
