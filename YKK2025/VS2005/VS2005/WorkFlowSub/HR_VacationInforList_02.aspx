<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_VacationInforList_02.aspx.vb" Inherits="HR_VacationInforList_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>休假一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 7px; position: absolute; top: 8px" Width="669px">
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="URL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="請假單No." Target="_blank">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:HyperLinkField>
                <asp:BoundField DataField="Vacation" HeaderText="假別" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="100px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="VacationTime" HeaderText="期間" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="200px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
                <asp:BoundField DataField="ADays" HeaderText="日數" ReadOnly="True">
                    <FooterStyle HorizontalAlign="Center" />
                    <HeaderStyle Width="50px" />
                    <ItemStyle HorizontalAlign="Center" />
                </asp:BoundField>
            </Columns>
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
