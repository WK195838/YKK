<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PS_MENU.aspx.vb" Inherits="PS_MENU" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >

<head runat="server">
    <title>ISOS Menu</title>
	    <script language="javascript" src="ForProject.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  底圖                                                                                ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <img src="iMages/PS_MenuFY22.png" style="z-index: 1; left: 40px; position: absolute;top: 0px; width: 860px; height: 480px;" onclick="alert('請點選[BUYER] OR [直接使用] !');" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  客戶選擇                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- 客戶  -->
        <asp:DropDownList ID="DBuyer" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 196px; position: absolute; top: 60px" Width="135px" Height="24px">
            <asp:ListItem Value="000021">TNF</asp:ListItem>
            <asp:ListItem Selected="True" Value="999999">選擇</asp:ListItem>
        </asp:DropDownList>

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  各Action按鈕                                                                        ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->

        <asp:ImageButton ID="BBuyer" runat="server" style="z-index: 1; left: 328px; position: absolute;top: 50px; width: 50px; height: 50px;" ImageUrl="iMages/Go.jpg"  />

        <asp:ImageButton ID="BUPLISTPRICE" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 96px;" ImageUrl="iMages/UP_LISTPRICE.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BUPBUYERITEM" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 150px;" ImageUrl="iMages/UP_BUYERITEM.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BUPBUYERCOLOR" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 204px;" ImageUrl="iMages/UP_BUYERCOLOR.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BUPREPLACEITEM" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 258px;" ImageUrl="iMages/UP_COMITEM.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BUPREPLACECOLOR" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 312px;" ImageUrl="iMages/UP_COMCOLOR.png" Height="42px" Width="128px" />

        <asp:ImageButton ID="BDIFFINF" runat="server" style="z-index: 1; left: 616px; position: absolute;top:78px;" ImageUrl="iMages/DIFFINF.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BITEM" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 132px;" ImageUrl="iMages/ITEM.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BCOLOR" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 186px;" ImageUrl="iMages/COLOR.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BDLLISTPRICE" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 240px;" ImageUrl="iMages/LISTPRICE.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BSALESPRICE" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 294px;" ImageUrl="iMages/SALESPRICE.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BINTEGRATEPRICE" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 332px;" ImageUrl="iMages/INTEGRATEPRICE.png" Height="42px" Width="128px" />
        <asp:ImageButton ID="BTOOLBOX" runat="server" style="z-index: 1; left: 192px; position: absolute;top: 368px;" ImageUrl="iMages/TOOLBOX.png" Height="42px" Width="128px" />

        <asp:ImageButton ID="BPRICEVAL" runat="server" style="z-index: 1; left: 754px; position: absolute;top: 78px;" ImageUrl="~/iMages/PRICEVAL.png" Height="40px" Width="144px" />
        <asp:ImageButton ID="BIRW" runat="server" style="z-index: 1; left: 754px; position: absolute;top: 222px;" ImageUrl="~/iMages/IRW.png" Height="40px" Width="144px" />
        <asp:ImageButton ID="BITEMSUITABLE" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 78px;" ImageUrl="~/iMages/ITEMSuitable.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BITEMTDF00" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 115px;" ImageUrl="~/iMages/IRWTDF00.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BITEMINQSA" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 152px;" ImageUrl="~/iMages/ITEMINQSA.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BITEMDashboard" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 189px;" ImageUrl="~/iMages/ITEMDashboard.PNG" Height="40px" Width="128px" />
        <asp:ImageButton ID="BITEMINQMOLD" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 226px;" ImageUrl="~/iMages/ITEMINQMOLD.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BEDXRegister" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 263px;" ImageUrl="~/iMages/EDXRegister.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BSliderImage3" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 300px;" ImageUrl="~/iMages/SliderImage3.png" Height="40px" Width="128px" />
        <asp:ImageButton ID="BITMaster" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 337px;" ImageUrl="~/iMages/ITMaster1.png" Height="40px" Width="128px" />

        <asp:ImageButton ID="BShopping" runat="server" style="z-index: 1; left: 896px; position: absolute;top: 374px;" ImageUrl="~/iMages/ShoppingList.png" Height="40px" Width="128px" />

        <asp:ImageButton ID="BISIP" runat="server" style="z-index: 1; left: 616px; position: absolute;top: 376px;" ImageUrl="~/iMages/ISIP.png" Height="42px" Width="128px" />

        <asp:ImageButton ID="BACTFCT" runat="server" style="z-index: 1; left: -500px; position: absolute;top: 376px;" ImageUrl="~/iMages/ACT-FCT.png" Height="42px" Width="128px" />

<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
<!-- ++  隱藏欄位                                                                       ++ -->
<!-- ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ -->
        <!-- FileData  -->
        <asp:TextBox ID="DFileName" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DLISTPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DBUYERITEMFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DBUYERCOLORFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DREPLACEITEMFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DREPLACECOLORFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DSALESPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINTEGRATEPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINQLISTPRICEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DISOS2FASFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINQBUYERITEMFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DINQBUYERCOLORFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DItemValuationFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DITEMSUITABLEFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DITEMSUITABLEIRWFile" runat="server" Height="16px" Style="z-index: 318; left: 202px;
            position: absolute; top: 43px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:TextBox ID="DEDXComment1" runat="server" Height="32px" Style="z-index: 318; left: 1024px;
            position: absolute; top: 264px;font-size:14px;background:transparent" Width="176px"   BorderStyle="None" BorderWidth="0px" TextMode="MultiLine">無法開啟時請複製以下連結至其它BROWSER使用</asp:TextBox>
        <asp:TextBox ID="DEDXComment2" runat="server" Height="24px" Style="z-index: 318; left: 1024px;
            position: absolute; top: 304px;font-size:14px;background:transparent" Width="300px"   BorderStyle="None" BorderWidth="0px">https://forms.office.com/r/MDyLjRE2V2</asp:TextBox>

        </div>
    </form>
</body>


</html>
