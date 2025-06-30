<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PostCodeList.aspx.vb" Inherits="PostCodeList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>客戶資料</title>
 
  
</head>
<body>
    <form id="form1" runat="server">
        &nbsp; &nbsp;&nbsp;<asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 100; left: 616px; position: absolute; top: 216px" Height="40px" Width="120px" />
        &nbsp;&nbsp;<br />
        <asp:Image ID="Image1" runat="server" ImageUrl="~/images/PostCode_01.jpg" />
        <asp:DropDownList ID="DCounty2" runat="server" AutoPostBack="True" BackColor="Yellow"
            Font-Size="Medium" Style="z-index: 101; left: 184px; position: absolute; top: 136px"
            Width="152px">
        </asp:DropDownList>
        &nbsp;&nbsp;
        <asp:TextBox ID="DAddress1" runat="server" BackColor="Yellow" Height="16px" Style="z-index: 102;
            left: 352px; position: absolute; top: 136px" Width="56px" AutoPostBack="True"></asp:TextBox>
        <asp:DropDownList ID="DCounty1" runat="server" AutoPostBack="True" BackColor="Yellow"
            Font-Size="Medium" Style="z-index: 103; left: 48px; position: absolute; top: 136px"
            Width="112px">
        </asp:DropDownList>
        <asp:DropDownList ID="DAddress2" runat="server" BackColor="Yellow" Font-Size="Medium"
            Style="z-index: 104; left: 424px; position: absolute; top: 136px" Width="168px" AutoPostBack="True">
        </asp:DropDownList>
        &nbsp;&nbsp;
        <asp:TextBox ID="DAddress3" runat="server" BackColor="Yellow" Font-Size="Larger"
            Height="40px" Style="z-index: 107; left: 48px; position: absolute; top: 192px"
            Width="528px"></asp:TextBox>
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="300px" Width="712px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px" style="z-index: 106; left: 32px; position: absolute; top: 280px">     
        <asp:GridView id="GridView1" runat="server" Height="34px"  CellPadding="4"  EmptyDataText="查無資料" Width="554px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="Code" HeaderText="郵遞區號"></asp:BoundField>
    <asp:BoundField DataField="County1" HeaderText="縣市" />
<asp:BoundField DataField="County2" HeaderText="鄉鎮市區"></asp:BoundField>
    <asp:BoundField DataField="Address" HeaderText="道路或街名" />
    <asp:BoundField DataField="AddressNo" HeaderText="門牌號碼" />
    <asp:BoundField DataField="TJCode" HeaderText="大榮簡碼" />
</Columns>

<HeaderStyle BackColor="#FFCC99"></HeaderStyle>

<AlternatingRowStyle BackColor="#CCFFCC"></AlternatingRowStyle>
</asp:GridView>
 </asp:Panel> &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;
            <br />
   
    </form>
</body>
</html>
