<%@ Page Language="VB" AutoEventWireup="false" CodeFile="HomeWorkSign.aspx.vb" Inherits="HomeWorkSign" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>在家勤務</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
           <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#CC9966"
               BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
               Style="z-index: 104; left: 8px; position: absolute; top: 10px" Width="1080px">
               <Columns>
                   <asp:BoundField DataField="UserName" HeaderText="姓名" >
                   </asp:BoundField>
                   <asp:BoundField DataField="Division" HeaderText="部門" >
                   </asp:BoundField>
                   <asp:BoundField DataField="ClickTime" HeaderText="刷卡時間" >
                   </asp:BoundField>
                   <asp:BoundField DataField="IPAddress" HeaderText="IP" >
                   </asp:BoundField>

               </Columns>
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
                
        </div>
        

    </form>
</body>
</html>
