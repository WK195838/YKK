<%@ Page Language="VB" AutoEventWireup="false" CodeFile="EOESMenu.aspx.vb" Inherits="EOESMenu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>EOES</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:TextBox ID="TextBox1" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="20px" Style="z-index: 103; left: 108px; position: absolute; top: 47px"
                Width="200px">[★]範本檔設定</asp:TextBox>
            <asp:ImageButton ID="TempLate" runat="server" Height="114px" ImageUrl="~/iMages/EOES_VBAExcel.jpg"
                Style="z-index: 103; left: 108px; position: absolute; top: 71px" Width="109px" />
            <asp:ImageButton ID="BTemplate" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
                left: 256px; width: 50px; position: absolute; top: 82px; height: 50px" />


            <asp:TextBox ID="TextBox2" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="20px" Style="z-index: 103; left: 108px; position: absolute; top: 197px"
                Width="200px">[★]ITEM轉換設定</asp:TextBox>
            <asp:ImageButton ID="Item" runat="server" Height="114px" ImageUrl="~/iMages/EOES_VBAExcel.jpg"
                Style="z-index: 103; left: 108px; position: absolute; top: 221px" Width="109px" />
            <asp:ImageButton ID="BItem" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
                left: 256px; width: 50px; position: absolute; top: 232px; height: 50px" />

            <asp:TextBox ID="TextBox3" runat="server" BackColor="Black" BorderStyle="None" ForeColor="White"
                Height="20px" Style="z-index: 103; left: 108px; position: absolute; top: 347px"
                Width="200px">[★]COLOR轉換設定</asp:TextBox>
            <asp:ImageButton ID="Color" runat="server" Height="114px" ImageUrl="~/iMages/EOES_VBAExcel.jpg"
                Style="z-index: 103; left: 108px; position: absolute; top: 371px" Width="109px" />
            <asp:ImageButton ID="BColor" runat="server" ImageUrl="iMages/Go.jpg" Style="z-index: 1;
                left: 256px; width: 50px; position: absolute; top: 382px; height: 50px" />


        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath1" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath2" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="DFilePath3" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>


        </div>
    </form>
</body>
</html>
