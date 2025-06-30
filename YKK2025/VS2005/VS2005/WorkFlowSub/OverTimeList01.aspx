<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OverTimeList01.aspx.vb" Inherits="OverTimeList01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>選取加班申請書</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 8px; position: absolute; top: 29px" Width="230px" ShowFooter="True">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" /><Columns>
                <asp:TemplateField>
                    <ItemTemplate>
                        <asp:CheckBox ID="CheckBox1" runat="server" />
                    </ItemTemplate>
                    <HeaderStyle Width="20px" />
                </asp:TemplateField>
                
                <asp:HyperLinkField DataNavigateUrlFields="OverTimeURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="加班單No." Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:HyperLinkField>
                
                <asp:BoundField DataField="OverTimeDateDesc" HeaderText="加班日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="60px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>

                <asp:BoundField DataField="Total" HeaderText="時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:Button ID="BOK" runat="server" Height="20px" Style="z-index: 103; left: 156px;
            position: absolute; top: 6px" Text="選取完成" Width="80px" /><asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 247px; position: absolute; top: 29px" Width="230px" ShowFooter="True">
                <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                <Columns>
                    <asp:HyperLinkField DataNavigateUrlFields="VacationURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="請假單No." Target="_blank">
                        <FooterStyle HorizontalAlign="Center" />
                        <HeaderStyle Width="100px" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:HyperLinkField>

                    <asp:BoundField DataField="AVacationDate" HeaderText="請假日時" ReadOnly="True">
                        <FooterStyle HorizontalAlign="Center" />
                        <HeaderStyle Width="60px" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>

                    <asp:BoundField DataField="Total" HeaderText="時數" ReadOnly="True">
                        <FooterStyle HorizontalAlign="Center" />
                        <HeaderStyle Width="50px" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:BoundField>

                </Columns>
                <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
    </div>
    </form>
</body>
</html>
