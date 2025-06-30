<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CommissionSheet.aspx.vb" Inherits="CommissionSheet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>

              <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DevelopmentCommission_02.png"
             Style="z-index: 99; left: 2px; position: absolute; top: -1px" />
        <asp:TextBox ID="DTHRRLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="12" Style="z-index: 126; left: 650px;
            position: absolute; top: 772px; text-align: left" Width="116px"></asp:TextBox>
        <asp:TextBox ID="DTHRRLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="20" Style="z-index: 126; left: 712px;
            position: absolute; top: 746px; text-align: left" Width="54px"></asp:TextBox>
        <asp:DropDownList ID="DTHRRLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 650px; position: absolute; top: 746px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHRRUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="12" Style="z-index: 126; left: 650px;
            position: absolute; top: 720px; text-align: left" Width="116px"></asp:TextBox>
        <asp:TextBox ID="DTHRRUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="20" Style="z-index: 126; left: 712px;
            position: absolute; top: 696px; text-align: left" Width="54px"></asp:TextBox>
        <asp:DropDownList ID="DTHRRUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 650px; position: absolute; top: 696px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
           
        <asp:Button ID="BGetYColor" runat="server" Height="22px" Style="z-index: 111; left: 376px;
            position: absolute; top: 1084px" Text="取Y色號" Width="76px"/>
        <asp:FileUpload ID="DMAPFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 120px; position: absolute; top: 1222px; background-color: #c0ffff"
            Width="644px" />
        <asp:FileUpload ID="DFORTYPEFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 128px; position: absolute; top: 1282px; background-color: #c0ffff"
            Width="267px" />
        <asp:HyperLink ID="LFORTYPEFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 128px; position: absolute; top: 1282px" Target="_blank"
            Width="104px">適用型別</asp:HyperLink>
        <asp:FileUpload ID="DSAMPLECONFIRMFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 504px; position: absolute; top: 1282px; background-color: #c0ffff"
            Width="267px" />
        <asp:HyperLink ID="LSAMPLECONFIRMFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 504px; position: absolute; top: 1282px" Target="_blank"
            Width="104px">樣品確認</asp:HyperLink>
        <asp:FileUpload ID="DQCCHECKFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 128px; position: absolute; top: 1306px; background-color: #c0ffff"
            Width="267px" />
        <asp:HyperLink ID="LQCCHECKFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 128px; position: absolute; top: 1306px" Target="_blank"
            Width="104px">品質檢測項目</asp:HyperLink>

        <asp:FileUpload ID="DMANUFAUTHORITYFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 504px; position: absolute; top: 1306px; background-color: #c0ffff"
            Width="267px" />
        <asp:HyperLink ID="LMANUFAUTHORITYFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 504px; position: absolute; top: 1306px" Target="_blank"
            Width="104px">製造授權</asp:HyperLink>

        <asp:FileUpload ID="DCONTACTFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 128px; position: absolute; top: 1332px; background-color: #c0ffff"
            Width="267px" />
        <asp:HyperLink ID="LCONTACTFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 128px; position: absolute; top: 1336px" Target="_blank"
            Width="104px">客戶切結書</asp:HyperLink>
        <asp:FileUpload ID="DMAPREFFILE" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 124px; position: absolute; top: 236px; background-color: #c0ffff"
            Width="644px" />
        <asp:HyperLink ID="LMAPREFFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl=""
            Style="z-index: 261; left: 125px; position: absolute; top: 239px" Target="_blank"
            Width="64px">草圖</asp:HyperLink>
        <asp:DropDownList ID="DLEVEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 1196px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DMAKEMAP" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 1170px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DMAPNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 124px;
            position: absolute; top: 1170px; text-align: left" Width="264px" MaxLength="20"  ></asp:TextBox>
        <asp:TextBox ID="DOTCON" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 124px;
            position: absolute; top: 1110px; text-align: left" Width="643px" MaxLength="255"  ></asp:TextBox>
        <asp:TextBox ID="DLYLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 376px;
            position: absolute; top: 1056px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DLYYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 648px;
            position: absolute; top: 1056px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DLYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px;
            position: absolute; top: 1056px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DLYCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 1056px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DHMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 376px;
            position: absolute; top: 1032px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DHMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 648px;
            position: absolute; top: 1032px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DHMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 560px;
            position: absolute; top: 1032px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DHMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 1032px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DGMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"  Style="z-index: 126; left: 376px;
            position: absolute; top: 1006px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DGMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 1006px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DGMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 1006px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DGMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 1006px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DFMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 980px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DFMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 980px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DFMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 980px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DFMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 980px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DEMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 954px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DEMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 954px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DEMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 954px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DEMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 954px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DDMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 928px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DDMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 928px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DDMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 928px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DDMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 928px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DCMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 902px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DCMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 902px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DCMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 902px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DCMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 902px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DBMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 876px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DBMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 876px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DBMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 876px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DBMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 876px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DAMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 850px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DAMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 850px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DAMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 850px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DAMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 850px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DXMLEN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 376px;
            position: absolute; top: 824px; text-align: center" Width="58px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DXMYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 648px;
            position: absolute; top: 824px; text-align: left" Width="120px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DXMCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 560px;
            position: absolute; top: 824px; text-align: left" Width="86px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DXMCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 440px; position: absolute; top: 824px"
            Width="120px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHRLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 524px;
            position: absolute; top: 772px; text-align: left" Width="116px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHLLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 272px;
            position: absolute; top: 772px; text-align: left" Width="116px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHLRLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="12" Style="z-index: 126; left: 398px;
            position: absolute; top: 772px; text-align: left" Width="116px"></asp:TextBox>
        <asp:TextBox ID="DTHLRLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="20" Style="z-index: 126; left: 460px;
            position: absolute; top: 746px; text-align: left" Width="54px"></asp:TextBox>
        <asp:DropDownList ID="DTHLRLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 398px; position: absolute; top: 746px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLRUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="12" Style="z-index: 126; left: 398px;
            position: absolute; top: 720px; text-align: left" Width="116px"></asp:TextBox>
        <asp:TextBox ID="DTHLRUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="20" Style="z-index: 126; left: 460px;
            position: absolute; top: 696px; text-align: left" Width="54px"></asp:TextBox>
        <asp:DropDownList ID="DTHLRUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 398px; position: absolute; top: 696px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLOYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 772px; text-align: left" Width="140px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHRLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 586px;
            position: absolute; top: 746px; text-align: left" Width="54px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHRLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 524px; position: absolute; top: 746px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 336px;
            position: absolute; top: 746px; text-align: left" Width="54px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHLLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 272px; position: absolute; top: 746px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLOCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 192px;
            position: absolute; top: 746px; text-align: left" Width="72px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHLOCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 746px"
            Width="64px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHRUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 524px;
            position: absolute; top: 720px; text-align: left" Width="116px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHLUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 272px;
            position: absolute; top: 720px; text-align: left" Width="116px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHUPYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 720px; text-align: left" Width="140px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTHRUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 586px;
            position: absolute; top: 696px; text-align: left" Width="54px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHRUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 524px; position: absolute; top: 696px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHLUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 336px;
            position: absolute; top: 696px; text-align: left" Width="54px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHLUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 272px; position: absolute; top: 696px"
            Width="56px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTHUPCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 192px;
            position: absolute; top: 696px; text-align: left" Width="72px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTHUPCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 696px"
            Width="64px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTARYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 524px;
            position: absolute; top: 640px; text-align: left" Width="240px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTALYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 272px;
            position: absolute; top: 640px; text-align: left" Width="240px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTAYCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 642px; text-align: left" Width="140px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DTARCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 656px;
            position: absolute; top: 616px; text-align: left" Width="108px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTARCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 524px; position: absolute; top: 616px"
            Width="128px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTALCOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 408px;
            position: absolute; top: 616px; text-align: left" Width="104px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTALCOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 272px; position: absolute; top: 616px"
            Width="128px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTACOLNO" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 192px;
            position: absolute; top: 616px; text-align: left" Width="72px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTACOL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 616px"
            Width="64px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DCCOL" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 634px;
            position: absolute; top: 564px; text-align: left" Width="131px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DCCOLSEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 564px"
            Width="130px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DECOL" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 256px;
            position: absolute; top: 564px; text-align: left" Width="131px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DTATYPE" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 538px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DSIZENO" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 512px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DITEM" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 512px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DNEEDMAP" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 210px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DTAWIDTH" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 502px;
            position: absolute; top: 538px; text-align: left" Width="263px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DSPSPEC" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 452px; text-align: left" Width="641px" MaxLength="255"  ></asp:TextBox>
        <asp:TextBox ID="DLOSTK" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 565px;
            position: absolute; top: 426px; text-align: left" Width="200px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DUPSTK" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 565px;
            position: absolute; top: 400px; text-align: left" Width="200px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DLOFIN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 334px;
            position: absolute; top: 426px; text-align: left" Width="222px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DUPFIN" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 334px;
            position: absolute; top: 400px; text-align: left" Width="222px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DLOSLI" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 426px; text-align: left" Width="200px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DUPSLI" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 400px; text-align: left" Width="200px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DEAQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 607px; position: absolute;
            top: 348px; text-align: left" Width="74px" MaxLength="12"  ></asp:TextBox>
        <asp:DropDownList ID="DEAQTYUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 684px; position: absolute; top: 348px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DPQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 607px; position: absolute;
            top: 322px; text-align: left" Width="74px" MaxLength="12"  ></asp:TextBox>
        <asp:DropDownList ID="DPQTYUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 684px; position: absolute; top: 322px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DEALEN" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 439px; position: absolute;
            top: 348px; text-align: left" Width="74px" MaxLength="12"  ></asp:TextBox>
        <asp:DropDownList ID="DEALENUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 516px; position: absolute; top: 348px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DPRO" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 123px; position: absolute; top: 322px"
            Width="204px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DPLEN" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 439px; position: absolute;
            top: 322px; text-align: left" Width="74px" MaxLength="12"  ></asp:TextBox>
        <asp:TextBox ID="DOPPART" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 348px; text-align: left" Width="200px" MaxLength="20"  ></asp:TextBox>
        <asp:DropDownList ID="DECOLSEL" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 564px"
            Width="130px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>


        <asp:HyperLink ID="LMAPFILE" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 261; left: 128px; position: absolute; top: 1224px" Target="_blank"
            Width="64px">圖檔</asp:HyperLink>
        <asp:DropDownList ID="DUSAGE" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 502px; position: absolute; top: 184px"
            Width="266px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="DAPPBUYER" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 124px; position: absolute; top: 132px"
            Width="266px">
            <asp:ListItem Value="徐">徐</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DORNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 210px; text-align: left" Width="263px" MaxLength="20"  ></asp:TextBox>
        <asp:TextBox ID="DCUSTITEM" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 184px; text-align: left" Width="263px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DEXPDEL" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 502px; position: absolute;
            top: 158px; text-align: left" Width="240px" MaxLength="10" ></asp:TextBox>
        <asp:Button ID="BEXPDEL" runat="server" Height="20px" Style="z-index: 111; left: 747px;
            position: absolute; top: 160px" Text="....." Width="20px" />
        <asp:TextBox ID="DESYQTY" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 158px; text-align: left" Width="263px" MaxLength="30"  ></asp:TextBox>
        <asp:TextBox ID="DSellVendor" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 502px;
            position: absolute; top: 132px; text-align: left" Width="263px" MaxLength="50"  ></asp:TextBox>
        <asp:DropDownList ID="DPLENUN" runat="server" BackColor="LightGray" ForeColor="Blue"
            Height="22px" Style="z-index: 266; left: 516px; position: absolute; top: 322px"
            Width="84px">
            <asp:ListItem Value="3">3</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DAPPPER" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 502px; position: absolute;
            top: 106px; text-align: left" Width="263px" MaxLength="20"  ></asp:TextBox>
        <asp:TextBox ID="DAPPDEPT" runat="server" BackColor="LightGray" BorderStyle="None"
            ForeColor="Black" Height="20px"   Style="z-index: 126; left: 124px;
            position: absolute; top: 106px; text-align: left" Width="263px" MaxLength="20"  ></asp:TextBox>
        <asp:TextBox ID="DAPPDATE" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 502px; position: absolute;
            top: 80px; text-align: left" Width="263px" MaxLength="10"  ></asp:TextBox>
        <asp:TextBox ID="DNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 126; left: 124px; position: absolute;
            top: 80px; text-align: left" Width="120px" MaxLength="20"  ></asp:TextBox>
<asp:TextBox ID="DQCLT" runat="server" Height="20px" Style="z-index: 318; left: 128px;
            position: absolute; top: 1254px; text-align: right" Width="69px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>

        <asp:FileUpload ID="DQCFILE1" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 16px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE1" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 16px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:FileUpload ID="DQCFILE2" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 144px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE2" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 144px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:FileUpload ID="DQCFILE3" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 272px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE3" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 272px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:FileUpload ID="DQCFILE4" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 392px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE4" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 400px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:FileUpload ID="DQCFILE5" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 520px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE5" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 520px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:FileUpload ID="DQCFILE6" runat="server" BackColor="LightGray" Height="20px"
            Style="z-index: 121; left: 648px; position: absolute; top: 1422px; background-color: #c0ffff"
            Width="123px" />
        <asp:HyperLink ID="LQCFILE6" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 261;
            left: 648px; position: absolute; top: 1424px" Target="_blank" Width="62px">QC文件</asp:HyperLink>

        <asp:CheckBox ID="DNeedSample" runat="server" style="z-index: 174; left: 128px; position: absolute; top: 296px" Width="130px" />
        <asp:CheckBox ID="DNeedItemRegister" runat="server" style="z-index: 174; left: 24px; position: absolute; top: 296px" Width="130px" />

        <asp:Button ID="BREFNO" runat="server" Height="20px" Style="z-index: 111; left: 370px;
            position: absolute; top: 82px" Text="....." Width="20px" />
        <asp:TextBox ID="DREFNO" runat="server" BackColor="LightGray" BorderStyle="None" ForeColor="Black"
            Height="20px"   Style="z-index: 261;left: 248px; position: absolute; top: 80px" Width="120px"></asp:TextBox>
        <asp:HyperLink ID="LREFNO" runat="server" Height="8px" Style="z-index: 261;
            left: 248px; position: absolute; top: 82px" Target="_blank" Width="120px"></asp:HyperLink>

        <asp:DropDownList ID="DQCPEOPLE" runat="server" BackColor="White" ForeColor="Blue"
           Height="22px" Style="z-index: 266; left: 120px; position: absolute; top: 1368px"
           Width="266px">
          <asp:ListItem Value="3">3</asp:ListItem>
          </asp:DropDownList>
            

    </div>
    </form>
</body>
</html>
