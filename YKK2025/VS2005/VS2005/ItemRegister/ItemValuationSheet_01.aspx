<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ItemValuationSheet_01.aspx.vb" Inherits="ItemValuationSheet_01" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>估價申請單</title>
	    <script language="javascript" src="RegisterItem.js">
	    </script>

		<script language="javascript" type="text/javascript">
            function ConfirmMe(btn) {
                if(Page_ClientValidate())   {
                    //--btn.disabled="disabled";
				    var answer = confirm("是否執行估價嗎？");
				    if (answer) {
                        //--啟動Excel程式        
                        var xFilePath = document.getElementById('DItemValuationFile');
                        var executableFullPath = "excel.exe " + xFilePath.value;

                        try
                        {
                            var shellActiveXObject = new ActiveXObject("WScript.Shell");

                            if ( !shellActiveXObject )
                            {
                                alert('Could not get reference to WScript.Shell');
                            }
                            shellActiveXObject.Run(executableFullPath, 1, false);
                            shellActiveXObject = null;
                        }
                        catch (errorObject)
                        {
                            alert('Error:\n' + errorObject.message);
                        }            
                        //--啟動Excel程式        
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
       
        <asp:TextBox ID="DDate" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 896px;
            position: absolute; top: 8px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DJobTitle" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 896px; position: absolute; top: 32px; text-align: left" Width="226px"></asp:TextBox>
        <asp:TextBox ID="DDivision" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
            BorderStyle="None" ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126;
            left: 992px; position: absolute; top: 32px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DName" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 992px;
            position: absolute; top: 8px; text-align: left" Width="246px"></asp:TextBox>
        <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Names="Times New Roman" Font-Size="11pt" Style="z-index: 100; left: 648px;
            position: absolute; top: 8px">DNo</asp:TextBox>
        <img src="images/RegisterSheet_004.jpg" style="z-index: 1; left: 6px; position: absolute;top: 88px" />
        <asp:CheckBox ID="DK206" runat="server" style="z-index: 174; left: -311px; position: absolute; top: 734px" Font-Size="9pt" Text="K206" Width="48px" />
        <asp:CheckBox ID="DA206" runat="server" style="z-index: 174; left: -164px; position: absolute; top: 734px" Font-Size="9pt" Text="A206" Width="48px" />

        <asp:CheckBox ID="DA001" runat="server" style="z-index: 174; left: 112px; position: absolute; top: 64px" Font-Size="9pt" Text="A001" Width="48px" />
        <asp:CheckBox ID="DA999" runat="server" style="z-index: 174; left: 184px; position: absolute; top: 64px" Font-Size="9pt" Text="A999" Width="48px" />
        <asp:CheckBox ID="DA211" runat="server" style="z-index: 174; left: 256px; position: absolute; top: 64px" Checked="True" Font-Size="9pt" Text="A211" Width="48px" />
        <asp:CheckBox ID="DK211" runat="server" style="z-index: 174; left: 328px; position: absolute; top: 64px" Font-Size="9pt" Text="K211" Width="48px" />
        <asp:TextBox ID="DPriceDescr" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" MaxLength="1" Style="z-index: 126; left: 410px;
            position: absolute; top: 734px; text-align: left" Width="185px" Font-Size="9pt"></asp:TextBox>

        <asp:Button ID="BCheckItem" runat="server" CausesValidation="False" Height="22px" Style="z-index: 131;
            left: 248px; position: absolute; top: 92px" Text="檢測新Item" Width="100px" />

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
            top: 145px; text-align: left" Width="289px" MaxLength="7" AutoPostBack="True"></asp:TextBox>
        <asp:TextBox ID="DMark1" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 224px; position: absolute;
            top: 92px; text-align: left" Width="17px">→</asp:TextBox>
        <asp:TextBox ID="DMark2" runat="server" BackColor="White" BorderStyle="None" ForeColor="Black"
            Height="20px" ReadOnly="True" Style="z-index: 126; left: 354px; position: absolute;
            top: 92px; text-align: left" Width="17px">→</asp:TextBox>

        <asp:TextBox ID="DRItemName1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 172px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
            Height="20px" Style="z-index: 126; left: 116px; position: absolute;
            top: 198px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DRItemName3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 224px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
            
        <asp:TextBox ID="DRSize" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 250px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRChain" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 276px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DRClass" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 302px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DRTape" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 328px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DRSlider1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 354px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 380px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSlider2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 406px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DRFinish2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 432px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DRSRequest1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 458px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 484px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 510px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 536px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 562px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRSRequest6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 588px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DRFamily" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 614px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DRST1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 157px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 198px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 239px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 280px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 321px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRST7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 362px;
            position: absolute; top: 640px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DRNoDisplay" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Black" Height="20px" Style="z-index: 126; left: 116px;
            position: absolute; top: 666px; text-align: left" Width="289px" MaxLength="1"></asp:TextBox>
  
        <asp:TextBox ID="DCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" MaxLength="35" Style="z-index: 126; left: 410px; position: absolute;
            top: 145px; text-align: left" Width="289px"></asp:TextBox>
        <asp:TextBox ID="DItemName1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 172px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DItemName2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Red"
            Height="20px" Style="z-index: 126; left: 410px; position: absolute;
            top: 198px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>
        <asp:TextBox ID="DItemName3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 224px; text-align: left" Width="289px" MaxLength="35"></asp:TextBox>

        <asp:TextBox ID="DSize" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 250px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DChain" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 276px; text-align: left" Width="289px" MaxLength="6"></asp:TextBox>
        <asp:TextBox ID="DClass" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 302px; text-align: left" Width="289px" MaxLength="2"></asp:TextBox>
        <asp:TextBox ID="DTape" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 328px; text-align: left" Width="289px" MaxLength="5"></asp:TextBox>
        <asp:TextBox ID="DSlider1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 354px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 380px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSlider2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 406px; text-align: left" Width="289px" MaxLength="15"></asp:TextBox>
        <asp:TextBox ID="DFinish2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 432px; text-align: left" Width="289px" MaxLength="3"></asp:TextBox>
        <asp:TextBox ID="DSRequest1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 458px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 484px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 510px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 536px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 562px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DSRequest6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 588px; text-align: left" Width="289px" MaxLength="10"></asp:TextBox>
        <asp:TextBox ID="DFamily" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 614px; text-align: left" Width="289px" MaxLength="4"></asp:TextBox>
        <asp:TextBox ID="DST1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 451px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 492px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 533px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 574px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 615px;
            position: absolute; top: 640px; text-align: left" Width="38px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DST7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 656px;
            position: absolute; top: 640px; text-align: left" Width="43px" MaxLength="1"></asp:TextBox>
        <asp:TextBox ID="DNoDisplay" runat="server" BackColor="#C0FFFF" BorderStyle="None"
            ForeColor="Red" Height="20px" Style="z-index: 126; left: 410px;
            position: absolute; top: 666px; text-align: left" Width="289px" MaxLength="1"></asp:TextBox>
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
            left: 378px; position: absolute; top: 1144px" Text="OK" Width="80px" OnClientClick="return ConfirmMe(this)" UseSubmitBehavior="false"/>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 12px;
            position: absolute; top: 1034px" visible="false" />
        <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 132; left: -700px; position: absolute; top: 965px"
            TextMode="MultiLine" Width="536px"></asp:TextBox>
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 133; left: -500px; position: absolute; top: 1039px" Visible="False"
            Width="352px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 134;
            left: -500px; position: absolute; top: 1042px" Visible="False" Width="64px" BackColor="Yellow">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 135; left: -500px; position: absolute; top: 1072px"
            Visible="False" Width="424px"></asp:TextBox>
        <asp:Label ID="DHistoryLabel" runat="server" Font-Size="11pt" ForeColor="Blue" Height="20px"
            Style="z-index: 137; left: -500px; position: absolute; top: 1157px" Text="核定履歷"></asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 136; left: -1000px; position: absolute; top: 1179px" Width="780px">
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
        <input id="DStatus" runat="server" name="Hidden1" style="z-index: 174; left: 640px;
            width: 31px; position: absolute; top: 40px" type="hidden" />
        <input id="DStatusItem" runat="server" name="Hidden2" style="z-index: 174; left: 688px;
            width: 31px; position: absolute; top: 40px" type="hidden" />
        <asp:CheckBox ID="DReRegister" runat="server" style="z-index: 174; left: -500Px; position: absolute; top: 136px" Font-Size="10pt" Text="繼續申請" Width="76px" />
        <asp:Button ID="BFindItem" runat="server" Height="22px" Style="z-index: 131;
            left: 120px; position: absolute; top: 92px" Text="搜尋參考Item" Width="100px" />
        <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 105; left: 112px; position: absolute;
            top: 8px" Width="193px"></asp:TextBox>
        <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
            Height="20px" Style="z-index: 108; left: 112px; position: absolute; top: 32px"
            Width="193px"></asp:TextBox>
        <input id="BBuyer" runat="server" style="z-index: 152; left: 312px; width: 24px;
            position: absolute; top: 32px" type="button" value="..." />
        <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 122; left: 344px; position: absolute;
            top: 8px" Width="158px"></asp:TextBox>
        <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
            BorderStyle="Groove" ForeColor="Blue" Height="20px" Style="z-index: 123; left: 344px;
            position: absolute; top: 32px" Width="158px"></asp:TextBox><asp:DropDownList ID="DForUse" runat="server" Height="20px" Style="z-index: 134;
            left: -1000px; position: absolute; top: 840px" Visible="False" Width="516px" BackColor="Yellow">
            </asp:DropDownList>
        <input id="BCustomer" runat="server" style="z-index: 151; left: 312px; width: 24px;
            position: absolute; top: 8px" type="button" value="..." />
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
        <asp:TextBox ID="DMemo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
            Font-Bold="True" Font-Size="Medium" ForeColor="Red" Height="32px" Style="z-index: 103;
            left: -1000px; position: absolute; top: 728px" Width="896px"></asp:TextBox>
        &nbsp;

        <asp:TextBox ID="DBlank2" runat="server" AutoPostBack="True" BackColor="White"
            BorderStyle="None" Font-Bold="False" Font-Size="20pt" ForeColor="White" Height="440px"
            MaxLength="7" Style="z-index: 126; left: 8px; position: absolute; top: 692px; text-align: left"
            Width="696px"></asp:TextBox>
'
        <asp:TextBox ID="DItemValuationFile" runat="server" Height="16px" Style="z-index: 318; left: 832px;
            position: absolute; top: 40px;font-size:10px;background:transparent" Width="116px"   BorderStyle="None" BorderWidth="0px"></asp:TextBox>
       
        <asp:Label ID="LCustomer" runat="server" Font-Size="Large" Style="z-index: 103; left: 8px;
            position: absolute; top: 8px">CUSTOMER</asp:Label>

        <asp:Label ID="LBuyer" runat="server" Font-Size="Large" Style="z-index: 103; left: 8px;
            position: absolute; top: 32px">BUYER</asp:Label>

        <asp:Label ID="LPriceV" runat="server" Font-Size="Large" Style="z-index: 103; left: 8px;
            position: absolute; top: 64px">單價設定</asp:Label>

       
    </form>
</body>
</html>
