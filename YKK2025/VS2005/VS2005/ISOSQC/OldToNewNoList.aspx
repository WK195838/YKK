<%@ Page Language="VB" AutoEventWireup="false" CodeFile="OldToNewNoList.aspx.vb" Inherits="OldToNewNoList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>舊換新</title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:TextBox ID="DData" runat="server" style="z-index: 100; left: 96px; position: absolute;;text-transform : uppercase; top: 16px" Height="17px" Width="264px"></asp:TextBox>
        &nbsp;<asp:Button ID="Button1" runat="server" Text="查詢" style="z-index: 101; left: 376px; position: absolute; top: 16px" />
        &nbsp; &nbsp;型別組 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;<br />
        &nbsp;&nbsp;<br />
        <br />
        &nbsp;&nbsp;
            <br />
        <asp:Panel ID="Panel2" runat="server" BorderWidth="1px" Height="248px" ScrollBars="Auto"
            Style="z-index: 113; left: 24px; position: absolute; top: 48px" Width="904px">
            &nbsp;
            <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AutoGenerateSelectButton="True"
                BackColor="White" BorderColor="#999999" BorderStyle="Solid" BorderWidth="1px"
                CellPadding="3" DataKeyNames="Unique_ID" Font-Size="9pt" ForeColor="Black" GridLines="Vertical"
                PageSize="100" Style="z-index: 103; left: 8px; position: absolute; top: 8px" Width="824px">
                <Columns>
                    <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="30px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="No" HeaderText="依賴書No" />
                    <asp:BoundField DataField="Supplier" HeaderText="廠商">
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Code" HeaderText="CODE">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Size" HeaderText="SIZE">
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Family" HeaderText="FAMILY">
                        <HeaderStyle Height="20px" />
                        <ItemStyle HorizontalAlign="Right" Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Body" HeaderText="BODY">
                        <HeaderStyle Height="20px" />
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="Puller" HeaderText="PULLER">
                        <HeaderStyle Height="20px" />
                        <ItemStyle HorizontalAlign="Right" Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="COLOR" HeaderText="COLOR" />
                    <asp:BoundField DataField="Finish" HeaderText="FINISH">
                        <ItemStyle Width="50px" />
                    </asp:BoundField>
                    <asp:BoundField DataField="OldCancel" HeaderText="舊廢除" />
                    <asp:BoundField DataField="result1" HeaderText="TestResult" />
                </Columns>
                <FooterStyle BackColor="#CCCCCC" />
                <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center"
                    VerticalAlign="Middle" />
                <AlternatingRowStyle BackColor="#CCCCCC" />
            </asp:GridView>
        </asp:Panel>
    </form>
</body>
</html>
