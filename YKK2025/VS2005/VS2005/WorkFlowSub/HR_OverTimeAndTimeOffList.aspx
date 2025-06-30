<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HR_OverTimeAndTimeOffList.aspx.vb" Inherits="HR_OverTimeAndTimeOffList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>調休一覽</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 5px; position: absolute; top: 5px" Width="300px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Right" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundField DataField="NameDesc" HeaderText="姓名" ReadOnly="True" />
                <asp:BoundField DataField="CVacationDesc" HeaderText="調休" ReadOnly="True" />
                <asp:BoundField DataField="OverTimeDateDesc" HeaderText="加班日期" ReadOnly="True" />
                <asp:BoundField DataField="Total" HeaderText="時數" ReadOnly="True"  >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>

        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
            BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True"
            Style="z-index: 114; left: 308px; position: absolute; top: 5px" Width="400px">
            <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Right" />
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
            <Columns>
                <asp:BoundField DataField="NameDesc" HeaderText="姓名" ReadOnly="True" />
                <asp:BoundField DataField="VacationDesc" HeaderText="假別" ReadOnly="True" />
                <asp:BoundField DataField="VacationDateDesc" HeaderText="請假日期" ReadOnly="True" />
                <asp:BoundField DataField="Total" HeaderText="日時" ReadOnly="True" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
            </Columns>
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>
