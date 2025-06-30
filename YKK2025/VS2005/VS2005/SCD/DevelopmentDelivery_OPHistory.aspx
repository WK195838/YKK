<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DevelopmentDelivery_OPHistory.aspx.vb" Inherits="DevelopmentDelivery_OPHistory" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>試作明細</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="20" Style="z-index: 114; left: 10px; position: absolute; top: 10px"
            Width="800px">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="OPPER" HeaderText="擔當" />
                <asp:BoundField DataField="OPBTIME" HeaderText="預定時間" />
                <asp:BoundField DataField="OPBHOURS" HeaderText="預定分數" >
                               <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="OPATIME" HeaderText="實際時間" />
                <asp:BoundField DataField="OPAHOURS" HeaderText="實際分數" >
                                               <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                <asp:BoundField DataField="OPDELAYC1" HeaderText="遲納類別" />
                <asp:BoundField DataField="OPDELAYC2" HeaderText="原因類別" />
                <asp:BoundField DataField="OPREM" HeaderText="作業說明" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" BorderStyle="Groove" Font-Bold="True" ForeColor="White" />
            <FooterStyle BorderStyle="Groove" />
        </asp:GridView>
    
    </div>
    </form>
</body>
</html>