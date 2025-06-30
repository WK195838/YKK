<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SP_ControlRecordLog.aspx.vb" Inherits="SP_ControlRecordLog" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Control History</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
            <asp:GridView ID="GridView1" runat="server" WIDTH="1500px" AutoGenerateColumns="false"  Style="z-index: 103; left: 8px; position: absolute; top: 40px" CellPadding="6" Font-Size="10pt" GridLines="Vertical" PageSize="20" ForeColor="Black" BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px" >
                <Columns>
                    <asp:BoundField DataField="User"  />
                    <asp:BoundField DataField="AccessTime"  />
                    <asp:BoundField DataField="Cat"  />
                    <asp:BoundField DataField="Active" />
                    <asp:BoundField DataField="Code" />
                    <asp:BoundField DataField="Name"  />
                    
                    <asp:BoundField DataField="Customer" />
                    <asp:BoundField DataField="Import" />
                    <asp:BoundField DataField="Demand"  />
                    <asp:BoundField DataField="ActPlan"  />
                    <asp:BoundField DataField="ImpActPlan"  />
                    <asp:BoundField DataField="RspActPlan"  />
                    <asp:BoundField DataField="KPInterface"  />

                    <asp:BoundField DataField="RspWINGS"  />
                    <asp:BoundField DataField="PILSheet"  />
                    <asp:BoundField DataField="Final"  />
                    <asp:BoundField DataField="ChgFinal"  />
                    <asp:BoundField DataField="Progress" />                    

                </Columns>
                <HeaderStyle BackColor="Black" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" Font-Bold="True" />
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="Chartreuse" Font-Bold="False" ForeColor="Black" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>    
    
    </div>
    </form>
</body>
</html>
