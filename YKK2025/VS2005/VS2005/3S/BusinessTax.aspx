<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BusinessTax.aspx.vb" Inherits="BusinessTax" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>SSES System Ver 1.0</title>
	<script language="javascript" src="ForProject.js"></script>
</head>

<body>
    <form id="form1" runat="server">
    <div>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="Images/BusinessTax.jpg" style="z-index: 1; left: -4px; position: absolute;top: 4px; width: 1070px; height: 470px;"/>
        &nbsp;&nbsp;
        <asp:DropDownList ID="DTAX" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="24px" Style="z-index: 112; left: 42px; position: absolute; top: 405px"
            Width="160px">
            <asp:ListItem>12345678901234567890</asp:ListItem>
        </asp:DropDownList>
        &nbsp; &nbsp;&nbsp;<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  各按鈕                                                                              ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:ImageButton ID="BINTAX" runat="server" style="z-index: 1; left: 416px; position: absolute;top: 144px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" CausesValidation="False" />
        &nbsp;
        <asp:ImageButton ID="BASSETSTAX" runat="server" style="z-index: 1; left: 728px; position: absolute;top: 144px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" CausesValidation="False" />
        <asp:ImageButton ID="BOUTTAX" runat="server" style="z-index: 1; left: 416px; position: absolute;top: 336px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" CausesValidation="False" />
        <asp:ImageButton ID="BDATA" runat="server" style="z-index: 1; left: 984px; position: absolute;top: 272px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" CausesValidation="False" />
        &nbsp; &nbsp; &nbsp;&nbsp;<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  執行跑馬燈                                                                       ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ --><!-- ++  機能選擇                                                                            ++ --><!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        &nbsp;&nbsp;

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:TextBox ID="xFun" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 13px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox><asp:TextBox ID="xFunID" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 23px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="xFileName" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 33px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:TextBox ID="xFilePath" runat="server" Height="16px" Style="z-index: 318; left: 20px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
        <asp:ImageButton ID="BTAX" runat="server" style="z-index: 1; left: 200px; position: absolute;top: 392px; width: 50px; height: 50px;" ImageUrl="iMages/Go.png" CausesValidation="False" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <asp:Label ID="LINTAX" runat="server" Height="16px" Style="z-index: 318; left: 410px;
            position: absolute; top: 200px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:Label>
        <asp:Label ID="LZEROTAX" runat="server" BorderStyle="None" BorderWidth="0px" Height="16px"
            Style="font-size: 10px; z-index: 318; background: none transparent scroll repeat 0% 0%;
            left: 715px; position: absolute; top: 390px" Width="116px"></asp:Label>
        <asp:Label ID="LASSETSTAX" runat="server" BorderStyle="None" BorderWidth="0px" Height="16px"
            Style="font-size: 10px; z-index: 318; background: none transparent scroll repeat 0% 0%;
            left: 715px; position: absolute; top: 200px" Width="116px"></asp:Label>
        <asp:Label ID="LOUTTAX" runat="server" BorderStyle="None" BorderWidth="0px" Height="16px"
            Style="font-size: 10px; z-index: 318; background: none transparent scroll repeat 0% 0%;
            left: 410px; position: absolute; top: 392px" Width="116px"></asp:Label>
    </div>
    </form>
</body>
</html>
