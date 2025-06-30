<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CustomerInfoCopy.aspx.vb" Inherits="CustomerInfoCopy" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客資料</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 160px; position: absolute; top: 24px" Height="17px" Width="368px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 552px; position: absolute; top: 24px" />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="579px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="STS" HeaderText="狀態"></asp:BoundField>
<asp:BoundField DataField="NO" HeaderText="No"></asp:BoundField>
    <asp:BoundField DataField="Name" HeaderText="擔當者" />
    <asp:BoundField DataField="CustomerCode" HeaderText="客戶代號" />
    <asp:BoundField DataField="NameCH" HeaderText="客戶名稱" />
    <asp:HyperLinkField DataNavigateUrlFields="viewURL" HeaderText="連結" Target="_blank" Text="連結" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> 
        <asp:Label ID="Label1" runat="server" Style="z-index: 104; left: 16px; position: absolute;
            top: 24px" Text="客戶代號/客戶名稱"></asp:Label>
        &nbsp; &nbsp;&nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
