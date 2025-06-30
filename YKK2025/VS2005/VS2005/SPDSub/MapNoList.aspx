<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MapNoList.aspx.vb" Inherits="MapNoList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>客戶資料</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp; 圖號&nbsp;&nbsp;<asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 63px; position: absolute; top: 13px;position: absolute;text-transform : uppercase;" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 234px; position: absolute; top: 12px" />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="300px" Width="688px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px" style="z-index: 103; left: 20px; position: absolute; top: 45px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
    <asp:BoundField DataField="Title" HeaderText="內製/外注" ReadOnly="True" />
<asp:BoundField DataField="No" HeaderText="No"></asp:BoundField>
<asp:BoundField DataField="MapNo" HeaderText="圖號"></asp:BoundField>
    <asp:BoundField DataField="Buyer" HeaderText="Buyer" />
    <asp:BoundField DataField="Sno" HeaderText="單號" />
    <asp:BoundField DataField="SliderCode" HeaderText="拉頭/拉片品名" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> &nbsp; &nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
    </form>
</body>
</html>
