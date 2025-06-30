<%@ Page Language="VB" AutoEventWireup="false" CodeFile="POPList.aspx.vb" Inherits="POPList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>POP LIST</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 10px; position: absolute; top: 7px"
            Width="1100px" ShowFooter="True" AllowSorting="True">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:BoundField DataField="OP" HeaderText="工程" />
                <asp:BoundField DataField="Code" HeaderText="Code" />
                <asp:BoundField DataField="ItemName" HeaderText="Item Name" />

                <asp:BoundField DataField="WaitTime" HeaderText="生產等待" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>

                <asp:BoundField DataField="StartTime" HeaderText="生產開始" />
                <asp:BoundField DataField="EndTime" HeaderText="生產完成" />
                
                <asp:BoundField DataField="ProdTime" HeaderText="生產時間" >
                    <ItemStyle HorizontalAlign="Right" />
                </asp:BoundField>
                
                <asp:BoundField DataField="MachineNo" HeaderText="機台No." />
                <asp:BoundField DataField="WaveTime" HeaderText="Wave's完成" />
                <asp:BoundField DataField="PRNO" HeaderText="PR / OR / No" />
                <asp:BoundField DataField="ORDNX2" HeaderText="OR" />
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" BorderStyle="Groove" />
            <FooterStyle BorderStyle="Groove" BackColor="#FFFFCC" HorizontalAlign="Right"/>
        </asp:GridView>
        <asp:TextBox ID="DMessage" runat="server" BackColor="White" BorderStyle="None" ForeColor="Red"
            Height="20px" ReadOnly="True" Style="z-index: 107; left: 13px; position: absolute;
            top: 10px" Width="332px">無POP資料</asp:TextBox>
    </div>
    </form>
</body>
</html>
