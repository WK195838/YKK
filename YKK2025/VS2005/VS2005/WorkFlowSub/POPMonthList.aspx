<%@ Page Language="VB" AutoEventWireup="false" CodeFile="POPMonthList.aspx.vb" Inherits="POPMonthList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>POP月報</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div ms_positioning="FlowLayout" style="display: inline; z-index: 105; left: 8px;
            width: 96px; color: #0000ff; position: absolute; top: 8px; height: 24px" title="委託年月：">
            委託年月：</div>
        <asp:DropDownList ID="DStartY" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="28px" Style="z-index: 104; left: 104px; position: absolute;
            top: 8px" Width="72px">
        </asp:DropDownList>
        <asp:DropDownList ID="DStartM" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="28px" Style="z-index: 102; left: 181px; position: absolute;
            top: 8px" Width="48px">
        </asp:DropDownList>
        <asp:Button ID="Go" runat="server" BackColor="White" ForeColor="Blue" Height="24px"
            Style="z-index: 108; left: 358px; position: absolute; top: 7px" Text="Go" Width="40px" />


        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
            PageSize="25" Style="z-index: 114; left: 6px; position: absolute; top: 41px"
            Width="1300px" AllowSorting="True">
            <RowStyle BackColor="White" ForeColor="Blue" />
            <Columns>
                <asp:HyperLinkField DataNavigateUrlFields="ViewURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="No" HeaderText="SPD-No." Target="_blank">
                </asp:HyperLinkField>

                <asp:HyperLinkField DataNavigateUrlFields="OPURL" DataNavigateUrlFormatString="{0}"
                    DataTextField="History" Target="_blank">
                </asp:HyperLinkField>
               
                <asp:BoundField DataField="ORNo" HeaderText="OR-No." />
                <asp:BoundField DataField="PRNo" HeaderText="PR-No." />
            
                <asp:BoundField DataField="OP" HeaderText="工程" />
                <asp:BoundField DataField="Code" HeaderText="Code" />
                <asp:BoundField DataField="ItemName" HeaderText="Item Name" />

                <asp:BoundField DataField="ReceiptTime" HeaderText="SPD-收到時間" />
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
            </Columns>
            <PagerStyle ForeColor="#8C4510" HorizontalAlign="Center" />
            <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#FF8000" Font-Bold="True" ForeColor="White" BorderStyle="Groove" />
            <FooterStyle BorderStyle="Groove" BackColor="#FFFFCC" HorizontalAlign="Right"/>
        </asp:GridView>
        <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 103; left: 234px; position: absolute; top: 7px"
            Width="112px">No</asp:TextBox>
    
    </div>
    </form>
</body>
</html>
