<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DiscountExtend.aspx.vb" Inherits="DiscountExtend" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客折扣單價</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="579px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
    <asp:HyperLinkField DataNavigateUrlFields="viewURL" HeaderText="連結" Target="_blank" Text="連結" />
<asp:BoundField DataField="STS" HeaderText="狀態"></asp:BoundField>
    <asp:BoundField DataField="DATE" HeaderText="申請日期" />
<asp:BoundField DataField="NO" HeaderText="No"></asp:BoundField>
    <asp:BoundField DataField="Name" HeaderText="擔當者" />
    <asp:BoundField DataField="ASDate" HeaderText="有效日期起" />
    <asp:BoundField DataField="AEDate" HeaderText="有效日期迄" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> <br />
        &nbsp;<br />
        &nbsp;&nbsp; &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
