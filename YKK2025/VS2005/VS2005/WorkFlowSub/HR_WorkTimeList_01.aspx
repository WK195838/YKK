<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_WorkTimeList_01.aspx.vb" Inherits="HR_WorkTimeList_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>出缺勤月報</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp;
        <asp:TextBox ID="DWorkDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" ReadOnly="True" Style="z-index: 118; left: 16px;
            position: absolute; top: 11px" Width="139px"></asp:TextBox>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            Style="z-index: 114; left: 11px; position: absolute; top: 45px" Width="300px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="加班單No." Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="Field1" HeaderText="目的地" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Time" HeaderText="時數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
    </form>
</body>
</html>
