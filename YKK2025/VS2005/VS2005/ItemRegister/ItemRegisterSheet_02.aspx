<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemRegisterSheet_02.aspx.vb" Inherits="ItemRegisterSheet_02" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>Wave's Item登錄申請單</title>
    <script language="javascript" src="RegisterItem.js"></script>
    <style type="text/css"> 
    td { 
      text-align:center;  
    } 
    .textbox
    {
        background-color: transparent;
        border : none;
    }
    </style>     
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/RegisterSheet_001.jpg" style="z-index: 1; left: 6px; position: absolute;top: 7px" />
        <asp:HyperLink ID="LSPDNOURL" runat="server" BorderColor="Black" Font-Size="10pt"
            Height="8px" NavigateUrl="BoardEdit.aspx" Style="z-index: 124; left: 592px; position: absolute;
            top: 104px" Target="_blank" Width="70px">SPDNO</asp:HyperLink>
        &nbsp;
        <asp:TextBox ID="DDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 116px;
            position: absolute; top: 74px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 116px; position: absolute; top: 100px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 452px; position: absolute; top: 100px; text-align: left" Width="130px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 452px;
            position: absolute; top: 74px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 18px;
            position: absolute; top: 14px">DNo</asp:TextBox>
        <img src="images/RegisterSheet_004.jpg" style="z-index: 1; left: 6px; position: absolute;top: 130px" />

        <asp:HyperLink ID="LIRWSPDInf" runat="server" Font-Size="10pt" Height="16px" NavigateUrl="http://10.245.0.3/ISOS/iMages/MCU_TW1741.png"
            Style="z-index: 102; left: 536px; position: absolute; top: 136px" Target="_blank"
            Width="80px">[上SLD-SPD]</asp:HyperLink>

        <asp:HyperLink ID="LIRWSPDInf1" runat="server" Font-Size="10pt" Height="16px" NavigateUrl="http://10.245.0.3/ISOS/iMages/MCU_TW1741.png"
            Style="z-index: 102; left: 632px; position: absolute; top: 136px" Target="_blank"
            Width="80px">[下SLD-SPD]</asp:HyperLink>

        <asp:HyperLink ID="LISIPInf" runat="server" Font-Size="10pt" Height="16px" NavigateUrl="http://10.245.0.3/ISOS/iMages/MCU_TW1741.png"
            Style="z-index: 102; left: 148px; position: absolute; top: 134px" Target="_blank"
            Width="80px">[上ISIP]</asp:HyperLink>

        <asp:HyperLink ID="LISIPInf1" runat="server" Font-Size="10pt" Height="16px" NavigateUrl="http://10.245.0.3/ISOS/iMages/MCU_TW1741.png"
            Style="z-index: 102; left: 248px; position: absolute; top: 134px" Target="_blank"
            Width="80px">[下ISIP]</asp:HyperLink>
            
        <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 105; left: 181px; position: absolute;
            top: 771px" Width="193px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 108; left: 181px; position: absolute; top: 803px"
            Width="193px"></asp:TextBox>
        <asp:TextBox ID="DForUse" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 108; left: 180px; position: absolute;
            top: 837px" Width="515px"></asp:TextBox>
        <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 122; left: 538px; position: absolute;
            top: 771px" Width="158px"></asp:TextBox>
        <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
            BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 123; left: 537px;
            position: absolute; top: 802px" Width="158px"></asp:TextBox>
        <asp:CheckBox ID="DA211" runat="server" style="z-index: 174; left: 213px; position: absolute; top: 734px" Checked="True" Font-Size="9pt" Text="A211" Width="48px" />
        <asp:CheckBox ID="DK206" runat="server" style="z-index: 174; left: 311px; position: absolute; top: 734px" Font-Size="9pt" Text="K206" Width="48px" />
        <asp:CheckBox ID="DK211" runat="server" style="z-index: 174; left: 360px; position: absolute; top: 734px" Checked="True" Font-Size="9pt" Text="K211" Width="48px" />
        <asp:CheckBox ID="DA206" runat="server" style="z-index: 174; left: 164px; position: absolute; top: 734px" Font-Size="9pt" Text="A206" Width="48px" />
        <asp:CheckBox ID="DA999" runat="server" style="z-index: 174; left: 262px; position: absolute; top: 734px" Font-Size="9pt" Text="A999" Width="48px" />
        <asp:CheckBox ID="DA001" runat="server" style="z-index: 174; left: 115px; position: absolute; top: 734px" Font-Size="9pt" Text="A001" Width="48px" />
        <asp:TextBox ID="DPriceDescr" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" MaxLength="1" Style="z-index: 126; left: 410px;
            position: absolute; top: 734px; text-align: left" Width="184px" Font-Size="9pt"></asp:TextBox>
        <asp:TextBox ID="DPriceNo" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" MaxLength="1" Style="z-index: 126; left: 597px;
            position: absolute; top: 734px; text-align: left" Width="101px" Font-Size="9pt"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <img src="images/RegisterSheet_006.jpg" style="z-index: 1; left: 7px; position: absolute;top: 870px" />
        <asp:HyperLink ID="LAttachfile2" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 124; left: 120px; position: absolute; top: 952px" Target="_blank"
            Width="100px">參考文件</asp:HyperLink>
        &nbsp;&nbsp;
        <asp:HyperLink ID="LAttachfile1" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 124; left: 120px; position: absolute; top: 928px" Target="_blank"
            Width="100px">參考文件</asp:HyperLink>
        &nbsp;&nbsp;
        <asp:HyperLink ID="LCHSheet" runat="server" BorderColor="Black" Font-Size="10pt"
            Height="8px" NavigateUrl="BoardEdit.aspx" Style="z-index: 124; left: 162px; position: absolute;
            top: 50px" Target="_blank" Width="70px">有CH申請</asp:HyperLink>
        <asp:HyperLink ID="LWaitHandle" runat="server" BorderColor="Black" Font-Size="10pt"
            Height="8px" NavigateUrl="BoardEdit.aspx" Style="z-index: 124; left: 237px; position: absolute;
            top: 50px" Target="_self" Width="70px">待簽核中</asp:HyperLink>
        <asp:HyperLink ID="LSLDSheet" runat="server" BorderColor="Black" Font-Size="10pt"
            Height="8px" NavigateUrl="BoardEdit.aspx" Style="z-index: 124; left: 86px; position: absolute;
            top: 50px" Target="_blank" Width="70px">有SLD申請</asp:HyperLink>
        <asp:HyperLink ID="LZIPSheet" runat="server" BorderColor="Black" Font-Size="10pt"
            Height="8px" NavigateUrl="BoardEdit.aspx" Style="z-index: 124; left: 16px; position: absolute;
            top: 48px" Target="_blank" Width="70px">有ZIP申請</asp:HyperLink>
        &nbsp;&nbsp;

        <asp:TextBox ID="DRCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 187px; text-align: left" Width="289px" MaxLength="7" AutoPostBack="True"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DRItemName1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 213px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 239px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 265px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRSize" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 291px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRChain" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 317px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DRClass" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 343px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRTape" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 369px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DRSlider1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 395px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 421px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSlider2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 447px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 473px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSRequest1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 499px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 525px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 551px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 577px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 603px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 629px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRFamily" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 655px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DRST1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 157px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 198px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 239px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 280px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 321px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 362px;
            position: absolute; top: 681px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRNoDisplay" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 707px; text-align: left" Width="289px" MaxLength="1"></asp:TextBox>
  
        <asp:TextBox ID="DItemName1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 213px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" MaxLength="35" Style="z-index: 126; left: 410px; position: absolute;
            top: 187px; text-align: left" Width="289px"></asp:TextBox>
        <asp:TextBox ID="DItemName2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 239px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DItemName3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 265px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DSize" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 291px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DChain" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 317px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DClass" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 343px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DTape" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 369px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DSlider1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 395px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 421px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSlider2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 447px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 473px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSRequest1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 499px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 525px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 551px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 577px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 603px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 629px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DFamily" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 655px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DST1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 451px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 492px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 533px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 574px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 615px;
            position: absolute; top: 681px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 656px;
            position: absolute; top: 681px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DNoDisplay" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 707px; text-align: left" Width="289px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DSLDPrice" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="43px" MaxLength="16" Style="z-index: 126; left: 599px;
            position: absolute; top: 877px; text-align: left" Width="101px"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="43px" Style="z-index: 126; left: 117px;
            position: absolute; top: 877px; text-align: left" Width="479px" MaxLength="100" TextMode="MultiLine"></asp:TextBox>


        <asp:TextBox ID="DSPDNOUrl" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Bold="False" Font-Size="Medium" ForeColor="Black" Height="480px" Style="z-index: 103;
            left: 712px; position: absolute; top: 136px" Width="700px" TextMode="MultiLine"></asp:TextBox>


        <asp:HyperLink ID="LEDXLink1" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="http://10.245.1.6/IRW/EDXList01.aspx?&pSize=&pPuller=&pColor="
            Style="z-index: 124; left: 712px; position: absolute; top: 112px" Target="_blank"
            Width="100px">EDX-1 (Click)</asp:HyperLink>

        <asp:HyperLink ID="LEDXLink2" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="http://10.245.1.6/IRW/EDXList01.aspx?&pSize=&pPuller=&pColor="
            Style="z-index: 124; left: 812px; position: absolute; top: 112px" Target="_blank"
            Width="100px">EDX-2 (Click)</asp:HyperLink>
        <asp:TextBox ID="DOrderDate" runat="server" BackColor="LightGray" BorderColor="Black"
            BorderStyle="Solid" Font-Bold="False" Font-Size="Medium" ForeColor="Red" Height="17px"
            Style="z-index: 103; left: 416px; position: absolute; top: 136px" Width="104px" ReadOnly="True"></asp:TextBox>
        <asp:TextBox ID="LOrderDate" runat="server" BorderStyle="None" Font-Bold="False"
            Font-Size="Medium" ForeColor="Black" Height="16px" ReadOnly="True" Style="z-index: 103;
            left: 344px; position: absolute; top: 136px" Width="70px">客戶了解</asp:TextBox>

        </div>
    </form>
</body>
</html>
