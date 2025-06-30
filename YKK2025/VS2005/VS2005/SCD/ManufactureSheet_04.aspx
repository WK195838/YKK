<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManufactureSheet_04.aspx.vb" Inherits="ManufactureSheet_04" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>製造委託書</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>

                     <asp:Image ID="Image2" runat="server" ImageUrl="~/Images/ManufactureCommission_04.jpg"
                     Style="z-index: 99; left: 2px; position: absolute; top: -1px" />
        <asp:HyperLink ID="LOldMSheet" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 18px; position: absolute; top: 51px" Target="_blank" Width="104px">舊-製造委託書</asp:HyperLink>

        <asp:TextBox ID="DOP8REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1292px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP8DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1318px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP8DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1292px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP8AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1266px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1266px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1240px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1240px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1240px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 1240px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP8PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 1240px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1188px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP7DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1214px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP7DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1188px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP7AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1162px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1162px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1136px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1136px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1136px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 1136px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP7PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 1136px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1084px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP6DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1110px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP6DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1084px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP6AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1058px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1058px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 1032px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 1032px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 1032px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 1032px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP6PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 1032px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 980px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP5DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 1006px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP5DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 980px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP5AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 954px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 954px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 928px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 928px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 928px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 928px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP5PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 928px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 876px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP4DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 902px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP4DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 876px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP4AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 850px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 850px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 824px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 824px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 824px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 824px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP4PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 824px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 772px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP3DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 798px"
            Width="144px" AutoPostBack="True">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        
        <asp:DropDownList ID="DOP3DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 772px"
            Width="144px" AutoPostBack="True">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        
        
        <asp:TextBox ID="DOP3AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 746px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 746px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 720px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 720px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 720px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 720px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP3PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 720px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 668px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP2DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 694px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP2DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 668px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DOP2AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 642px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 642px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 616px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 616px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 616px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 616px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP2PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 616px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1REM" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 564px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:DropDownList ID="DOP1DELAYC2" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 590px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DOP1DELAYC1" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 269px; position: absolute; top: 564px"
            Width="144px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTASPEC" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="99px"   Style="z-index: 126; left: 333px;
            position: absolute; top: 243px; text-align: left" TextMode="MultiLine" Width="433px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1AHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 538px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1ATIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 538px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1BHOURS" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 354px;
            position: absolute; top: 512px; text-align: RIGHT" Width="56px"  ></asp:TextBox>
        <asp:TextBox ID="DPLEN1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 564px;
            position: absolute; top: 374px; text-align: left" Width="98px"  ></asp:TextBox>
        <asp:TextBox ID="DCCOL1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 333px;
            position: absolute; top: 348px; text-align: left" Width="119px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1CON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="46px"   Style="z-index: 126; left: 416px;
            position: absolute; top: 512px; text-align: left" TextMode="MultiLine" Width="351px"  ></asp:TextBox>
        <asp:TextBox ID="DTHSPEC" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="47px"   Style="z-index: 126; left: 123px; position: absolute;
            top: 373px; text-align: left" TextMode="MultiLine" Width="329px"  ></asp:TextBox>
            
            
        <asp:Image ID="LHINTFILE" runat="server" BorderStyle="Groove" Height="145px" ImageUrl="http://10.245.0.178//WorkFlow/Document/000003/1-Map2-2006124111750-2.jpg"
            Style="z-index: 241; left: 20px; position: absolute; top: 194px" Width="200px" />
        <asp:TextBox ID="DDEVTITLE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 80px; text-align: left" Width="642px"  ></asp:TextBox>
        <asp:TextBox ID="DMANUFTYPE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 460px; text-align: left" Width="642px"  ></asp:TextBox>
        <asp:TextBox ID="DDEVNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 106px; text-align: left" Width="264px"  ></asp:TextBox>
        <asp:TextBox ID="DCODENO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 502px; position: absolute;
            top: 106px; text-align: left" Width="264px"  ></asp:TextBox>
        <asp:TextBox ID="DDEVPER1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 502px; position: absolute;
            top: 132px; text-align: left" Width="264px"  ></asp:TextBox>
        <asp:TextBox ID="DISSDATE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 132px; text-align: left" Width="264px"  ></asp:TextBox>
        <asp:TextBox ID="DUPSTK1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 333px; position: absolute;
            top: 192px; text-align: left" Width="160px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1BTIME" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 207px;
            position: absolute; top: 512px; text-align: left" Width="143px"  ></asp:TextBox>
        <asp:TextBox ID="DEALEN1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 564px;
            position: absolute; top: 400px; text-align: left" Width="98px"  ></asp:TextBox>
        <asp:TextBox ID="DEAQTY1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 669px;
            position: absolute; top: 400px; text-align: left" Width="98px"  ></asp:TextBox>
        <asp:TextBox ID="DPQTY1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 669px;
            position: absolute; top: 374px; text-align: left" Width="98px"  ></asp:TextBox>
        <asp:TextBox ID="DECOL1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 123px; position: absolute;
            top: 348px; text-align: left" Width="98px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 17px;
            position: absolute; top: 512px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOP1PER" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="98px"   Style="z-index: 126; left: 101px;
            position: absolute; top: 512px; text-align: left" Width="79px"  ></asp:TextBox>
        <asp:TextBox ID="DOPPART1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 333px; position: absolute;
            top: 218px; text-align: left" Width="433px"  ></asp:TextBox>
        <asp:TextBox ID="DLOSTK1" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 606px; position: absolute;
            top: 192px; text-align: left" Width="160px"  ></asp:TextBox>
    
           <asp:HyperLink ID="HOP1" runat="server" Font-Size="10pt" Height="100px" 
              Style="z-index: 126; left: 17px; position: absolute;
            top: 512px" Target="_blank" Width="82px"></asp:HyperLink>    
     
       <asp:HyperLink ID="HOP2" runat="server" Font-Size="10pt" Height="100px"
              Style="z-index: 126; left: 17px; position: absolute;
            top: 617px" Target="_blank" Width="82px"></asp:HyperLink>   
         
                      <asp:HyperLink ID="HOP3" runat="server" Font-Size="10pt" Height="100px"
            Style="z-index: 126; left: 17px; position: absolute;
            top: 720px" Target="_blank" Width="82px"></asp:HyperLink>    
                    
       <asp:HyperLink ID="HOP4" runat="server" Font-Size="10pt" Height="100px"
              Style="z-index: 126; left: 17px; position: absolute;
            top: 824px" Target="_blank" Width="82px"></asp:HyperLink>    
            
       <asp:HyperLink ID="HOP5" runat="server" Font-Size="10pt" Height="100px"
            Style="z-index: 126; left: 17px; position: absolute;
            top: 928px" Target="_blank" Width="82px"></asp:HyperLink>        
      
       <asp:HyperLink ID="HOP6" runat="server" Font-Size="10pt" Height="100px"
             Style="z-index: 126; left: 17px; position: absolute;
            top: 1032px" Target="_blank" Width="82px"></asp:HyperLink>   
            
         <asp:HyperLink ID="HOP7" runat="server" Font-Size="10pt" Height="100px"
            Style="z-index: 126; left: 17px; position: absolute;
            top: 1136px" Target="_blank" Width="82px"></asp:HyperLink>         
      
            
          <asp:HyperLink ID="HOP8" runat="server" Font-Size="10pt" Height="100px"
              Style="z-index: 126; left: 17px; position: absolute;
            top: 1240px" Target="_blank" Width="82px"></asp:HyperLink> 
            
                                       
    </div>
    </form>
</body>
</html>
