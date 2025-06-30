<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DivisionEmpList.aspx.vb" Inherits="DivisionEmpList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>員工資料</title>
</head>
<body>
    <form id="form1" runat="server">
        &nbsp; 員工卡號或姓名 &nbsp;&nbsp;<asp:TextBox ID="DNo" runat="server" style="z-index: 100; left: 136px; position: absolute; top: 16px" Height="17px"></asp:TextBox><asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 312px; position: absolute; top: 16px" />
        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="DNo"
            ErrorMessage="請輸入加班申請書單號" Style="left: 224px; position: relative; top: 0px; z-index: 102;" Width="240px"></asp:RequiredFieldValidator>&nbsp;<br />
        <br />
        &nbsp;&nbsp;<asp:Panel ID="Panel1" runat="server"  Height="408px" Width="448px" ScrollBars="Auto" BorderStyle="Solid" BorderWidth="1px">     
        <asp:GridView id="GridView1" runat="server" Height="1px"  CellPadding="4"  EmptyDataText="只能查詢本部門員工卡號" AllowPaging="True" Width="392px" Font-Size="Small" AutoGenerateColumns="False" AllowSorting="True" PageSize="15">
<RowStyle HorizontalAlign="Center" BackColor="#FFFF99"></RowStyle>
<Columns>
<asp:CommandField ShowSelectButton="True"></asp:CommandField>
<asp:BoundField DataField="hrwdivname" HeaderText="部門"></asp:BoundField>
<asp:BoundField DataField="id" HeaderText="員工編號"></asp:BoundField>
<asp:BoundField DataField="name" HeaderText="姓名"></asp:BoundField>
<asp:BoundField DataField="com_code" HeaderText="com_code" Visible="False"></asp:BoundField>
<asp:BoundField DataField="com_name" HeaderText="com_name" Visible="False"></asp:BoundField>
<asp:BoundField DataField="Hrwdivid" HeaderText="Hrwdivid" Visible="False"></asp:BoundField>
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
