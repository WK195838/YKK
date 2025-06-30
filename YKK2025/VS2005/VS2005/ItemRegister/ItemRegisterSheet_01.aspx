<%@ Page Language="VB" Debug="true" CodeFile="ItemRegisterSheet_01.aspx.vb" Inherits="RegisterSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Wave's Item登錄申請單</title>
	    <script language="javascript" src="RegisterItem.js">
	    </script>

		<script language="javascript" type="text/javascript">
            function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    btn.disabled="disabled";
				    var answer = confirm("是否執行作業嗎？ 請確認....");
				    if (answer) {
                        document.forms[0].__EVENTTARGET.value = btn.name;
                        document.forms[0].__EVENTARGUMENT.value = '';
                        document.forms[0].submit();
				    }                    
                    else    {
                        btn.disabled="";
                        return false;   
                    }				    
                }
                else    {
                    return false;
                }
            }
		</script>
	    
	   <script language="javascript" type="text/javascript">
	      function GetCustomer()
{
        
    window.open('CustomerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 

   function GetBuyer()
{
        
    window.open('BuyerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 
	   </script> 	
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <img src="images/RegisterSheet_001.jpg" style="z-index: 1; left: 6px; position: absolute;top: 7px" />
        <asp:TextBox ID="DDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 116px;
            position: absolute; top: 74px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 116px; position: absolute; top: 100px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 452px; position: absolute; top: 100px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 452px;
            position: absolute; top: 74px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 18px;
            position: absolute; top: 14px">DNo</asp:TextBox>
        <img src="images/RegisterSheet_004.jpg" style="z-index: 1; left: 6px; position: absolute;top: 130px" />
        <asp:CheckBox ID="DA211" runat="server" style="z-index: 174; left: 213px; position: absolute; top: 734px" Checked="True" Font-Size="9pt" Text="A211" Width="48px" />
        <asp:CheckBox ID="DK206" runat="server" style="z-index: 174; left: 311px; position: absolute; top: 734px" Font-Size="9pt" Text="K206" Width="48px" />
        <asp:CheckBox ID="DK211" runat="server" style="z-index: 174; left: 360px; position: absolute; top: 734px" Checked="True" Font-Size="9pt" Text="K211" Width="48px" />
        <asp:CheckBox ID="DA206" runat="server" style="z-index: 174; left: 164px; position: absolute; top: 734px" Font-Size="9pt" Text="A206" Width="48px" />
        <asp:CheckBox ID="DA999" runat="server" style="z-index: 174; left: 262px; position: absolute; top: 734px" Font-Size="9pt" Text="A999" Width="48px" />
        <asp:CheckBox ID="DA001" runat="server" style="z-index: 174; left: 115px; position: absolute; top: 734px" Font-Size="9pt" Text="A001" Width="48px" />
        <asp:TextBox ID="DPriceDescr" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" MaxLength="1" Style="z-index: 126; left: 410px;
            position: absolute; top: 734px; text-align: left" Width="185px" Font-Size="9pt"></asp:TextBox>
        &nbsp;
        <asp:Button ID="BCheckItem" runat="server" CausesValidation="False" Height="22px" Style="z-index: 131;
            left: 241px; position: absolute; top: 135px" Text="檢測新Item" Width="100px" />
        <asp:DropDownList ID="DSPDPerson" runat="server" BackColor="Yellow" ForeColor="Blue"
            Style="z-index: 112; left: 628px; position: absolute; top: 928px" Width="76px">
            <asp:ListItem Selected="True" Value="徐徐徐">徐徐徐</asp:ListItem>
        </asp:DropDownList>
        <img src="images/RegisterSheet_003.jpg" style="z-index: 1; left: 8px; position: absolute;top: 869px" />
        <asp:TextBox ID="DPriceNo" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" MaxLength="1" Style="z-index: 126; left: 598px;
            position: absolute; top: 734px; text-align: left" Width="101px" Font-Size="9pt"></asp:TextBox>
        <asp:FileUpload ID="DAttachfile1" runat="server" style="z-index: 121; left: 119px; position: absolute; top: 927px; background-color: #C0FFFF" Height="24px" Width="505px"/>
        <asp:HyperLink ID="LAttachfile1" runat="server" Font-Size="12pt" Height="8px" NavigateUrl="BoardEdit.aspx"
            Style="z-index: 124; left: 126px; position: absolute; top: 932px" Target="_blank"
            Width="100px">參考文件</asp:HyperLink>
        <asp:TextBox ID="DRCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 187px; text-align: left" Width="289px" MaxLength="7" AutoPostBack="True"></asp:TextBox>
        <asp:TextBox ID="DMark1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 221px; position: absolute;
            top: 137px; text-align: left" Width="17px">→</asp:TextBox>
        &nbsp; &nbsp;
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
            position: absolute; top: 875px; text-align: left" Width="101px"></asp:TextBox>
        <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="43px" Style="z-index: 126; left: 118px;
            position: absolute; top: 876px; text-align: left" Width="479px" MaxLength="100" TextMode="MultiLine"></asp:TextBox>
        <asp:Button ID="BSAVE" runat="server" Height="23px" Style="z-index: 128;
            left: 354px; position: absolute; top: 1144px" Text="儲存" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG2" runat="server" Height="23px" Style="z-index: 129;
            left: 445px; position: absolute; top: 1144px" Text="NG2" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BNG1" runat="server" Height="23px" Style="z-index: 130;
            left: 537px; position: absolute; top: 1144px" Text="NG1" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <asp:Button ID="BOK" runat="server" Height="23px" Style="z-index: 131;
            left: 629px; position: absolute; top: 1144px" Text="OK" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 12px;
            position: absolute; top: 1034px" visible="false" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: 59px; position: absolute; top: 965px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 133; left: 243px; position: absolute; top: 1039px" Visible="False"
            Width="352px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 134;
            left: 167px; position: absolute; top: 1042px" Visible="False" Width="64px" BackColor="Yellow">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 135; left: 171px; position: absolute; top: 1072px"
            Visible="False" Width="424px"></asp:TextBox>
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
            Style="z-index: 137; left: 14px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 136; left: 13px; position: absolute; top: 1179px" Width="780px">
            <RowStyle BackColor="White" Font-Size="9pt" ForeColor="#330099" />
            <Columns>
                <asp:BoundField DataField="StepNameDesc" HeaderText="工程">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="DecideName" HeaderText="擔當" />
                <asp:BoundField DataField="AgentName" HeaderText="代理/兼職" />
                <asp:BoundField DataField="FlowTypeDesc" HeaderText="類別" />
                <asp:BoundField DataField="StsDesc" HeaderText="處理結果" />
                <asp:BoundField DataField="DecideDescA" HeaderText="說明">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="Description" HeaderText="核定時間">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#CC6600" Font-Size="9pt" ForeColor="#FFFFCC" HorizontalAlign="Center"
                VerticalAlign="Middle" />
        </asp:GridView>
        <img id="DDescSheet" runat="server" src="images/Sheet_Remark.jpg" style="z-index: 1;
            left: 13px; position: absolute; top: 959px" />
        <input id="DStatus" runat="server" name="Hidden1" style="z-index: 174; left: 14px;
            width: 31px; position: absolute; top: 46px" type="hidden" />
        <input id="DStatusItem" runat="server" name="Hidden2" style="z-index: 174; left: 57px;
            width: 31px; position: absolute; top: 46px" type="hidden" />
        <asp:CheckBox ID="DReRegister" runat="server" style="z-index: 174; left: 624px; position: absolute; top: 136px" Checked="True" Font-Size="10pt" Text="繼續申請" Width="76px" /><asp:Button ID="BFindItem" runat="server" Height="22px" Style="z-index: 131;
            left: 118px; position: absolute; top: 135px" Text="搜尋參考Item" Width="100px" />
        <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 105; left: 181px; position: absolute;
            top: 771px" Width="193px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="24px" Style="z-index: 108; left: 181px; position: absolute; top: 808px"
            Width="193px"></asp:TextBox>
        &nbsp;
        <input id="BBuyer" runat="server" style="z-index: 152; left: 381px; width: 24px;
            position: absolute; top: 808px" type="button" value="..." />
        <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="24px" Style="z-index: 122; left: 538px; position: absolute;
            top: 771px" Width="158px"></asp:TextBox>
        <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
            BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 123; left: 537px;
            position: absolute; top: 802px" Width="158px"></asp:TextBox><asp:DropDownList ID="DForUse" runat="server" Height="20px" Style="z-index: 134;
            left: 184px; position: absolute; top: 840px" Visible="False" Width="516px" BackColor="Yellow">
            </asp:DropDownList>
        <input id="BCustomer" runat="server" style="z-index: 151; left: 382px; width: 24px;
            position: absolute; top: 774px" type="button" value="..." />
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 132; left: 768px; position: absolute;
               top: 300px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 768px; position: absolute; top: 336px"
               Width="190px"></asp:TextBox>

        <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
            <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px;position: absolute; top: 100px; text-align: left">AAA</asp:TextBox>
        &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DMemo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Bold="True" Font-Size="Medium" ForeColor="Red" Height="32px" Style="z-index: 103;
            left: 712px; position: absolute; top: 728px" Width="896px"></asp:TextBox>
       
    </form>
</body>
</html>
