<%@ Page Language="VB" AutoEventWireup="false" CodeFile="StatusList.aspx.vb" Inherits="StatusList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>工程進度</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="#400040"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" HorizontalAlign="Left" Width="800px"
            Style="z-index: 114; left: 9px; position: absolute; top: 2px" >
            <Columns>
                <asp:BoundField DataField="Field1" HeaderText="營業" ><ItemStyle Width="30px" /></asp:BoundField>
                <asp:BoundField DataField="Field2" HeaderText="工程" ><ItemStyle Width="100px" /></asp:BoundField>
                <asp:BoundField DataField="Field3" HeaderText="擔當" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field4" HeaderText="代理/兼職" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field5" HeaderText="類別" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field6" HeaderText="處理進度" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field7" HeaderText="處理結果" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field8" HeaderText="說明" ><ItemStyle Width="160px" /></asp:BoundField>
                <asp:BoundField DataField="Field9" HeaderText="參考資料" ><ItemStyle Width="200px" /></asp:BoundField>

            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BorderColor="#400040"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" HorizontalAlign="Left" Width="800px"
            Style="z-index: 114; left: 9px; position: absolute; top: 149px">
            <Columns>
                <asp:BoundField DataField="Field1" HeaderText="ZIP" ><ItemStyle Width="30px" /></asp:BoundField>
                <asp:BoundField DataField="Field2" HeaderText="工程" ><ItemStyle Width="100px" /></asp:BoundField>
                <asp:BoundField DataField="Field3" HeaderText="擔當" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field4" HeaderText="代理/兼職" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field5" HeaderText="類別" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field6" HeaderText="處理進度" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field7" HeaderText="處理結果" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field8" HeaderText="說明" ><ItemStyle Width="160px" /></asp:BoundField>
                <asp:BoundField DataField="Field9" HeaderText="參考資料" ><ItemStyle Width="200px" /></asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderColor="#400040"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" HorizontalAlign="Left" Width="800px"
            Style="z-index: 114; left: 9px; position: absolute; top: 297px">
            <Columns>
                <asp:BoundField DataField="Field1" HeaderText="SLD" ><ItemStyle Width="30px" /></asp:BoundField>
                <asp:BoundField DataField="Field2" HeaderText="工程" ><ItemStyle Width="100px" /></asp:BoundField>
                <asp:BoundField DataField="Field3" HeaderText="擔當" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field4" HeaderText="代理/兼職" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field5" HeaderText="類別" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field6" HeaderText="處理進度" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field7" HeaderText="處理結果" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field8" HeaderText="說明" ><ItemStyle Width="160px" /></asp:BoundField>
                <asp:BoundField DataField="Field9" HeaderText="參考資料" ><ItemStyle Width="200px" /></asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <asp:GridView ID="GridView4" runat="server" AutoGenerateColumns="False" BorderColor="#400040"
            BorderStyle="Groove" BorderWidth="1px" CellPadding="4" Font-Size="9pt" HorizontalAlign="Left" Width="800px"
            Style="z-index: 114; left: 9px; position: absolute; top: 446px">
            <Columns>
                <asp:BoundField DataField="Field1" HeaderText="CH" ><ItemStyle Width="30px" /></asp:BoundField>
                <asp:BoundField DataField="Field2" HeaderText="工程" ><ItemStyle Width="100px" /></asp:BoundField>
                <asp:BoundField DataField="Field3" HeaderText="擔當" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field4" HeaderText="代理/兼職" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field5" HeaderText="類別" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field6" HeaderText="處理進度" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field7" HeaderText="處理結果" ><ItemStyle Width="60px" /></asp:BoundField>
                <asp:BoundField DataField="Field8" HeaderText="說明" ><ItemStyle Width="160px" /></asp:BoundField>
                <asp:BoundField DataField="Field9" HeaderText="參考資料" ><ItemStyle Width="200px" /></asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
        </asp:GridView>
        <br />
    </div>
    </form>
</body>
</html>
