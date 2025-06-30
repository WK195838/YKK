<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_OverTimeList.aspx.vb" Inherits="HR_OverTimeList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>加班一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 8px; position: absolute; top: 7px" Width="345px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="StsDesc" HeaderText="*" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="45px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:HyperLinkField DataNavigateUrlFields="OverTimeURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="加班單No." Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="OverTimeDateDesc" HeaderText="加班日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="65px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Common" HeaderText="平日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="45px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="Vacation" HeaderText="假日" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="45px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="CVacation" HeaderText="國假" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="45px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
