<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ExpItemListCL.aspx.vb" Inherits="ExpItemListCL" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head id="Head1" runat="server">
    <title>經費資料</title>
</head>
<body>
    <form id="form1" runat="server">
        尋找
        <asp:DropDownList style="Z-INDEX: 134; LEFT: 48px; POSITION: absolute; TOP: 8px" id="DExpCat" runat="server" Width="152px"  Height="24px" BackColor="yellow">
               </asp:DropDownList>    
        <asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 200px; position: absolute;;text-transform : uppercase; top: 8px" Height="24px" Width="304px"  BackColor="yellow"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 512px; position: absolute; top: 8px" Height="32px" Width="64px" />
        <br />
        <br />
        <asp:Panel ID="Panel1" runat="server"  Height="300px" Width="598px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">    
         
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Left" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="ExpCat" HeaderText="種類"></asp:BoundField>
<asp:BoundField DataField="acremark" HeaderText="會計科目"></asp:BoundField>
    <asp:BoundField DataField="ACID" HeaderText="會計編號" />
    <asp:BoundField DataField="ExpItem" HeaderText="用途" />
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
