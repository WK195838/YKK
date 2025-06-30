<%@ Page Language="vb" AutoEventWireup="false" Inherits="CustomerInfoSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="CustomerInfoSheet_02.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客建檔</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			   function GetAddres(Field)
{
        
    window.open('PostCodeList.aspx?field=' +Field,'','status=0,toolbar=0,width=820,height=650,top=10,resizable=yes');
   
}	



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
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/CustomerInfo_03.jpg" style="z-index: 99; left: 8px; position: absolute;
                   top: 16px" />
               &nbsp;&nbsp;
               <asp:DropDownList ID="DPaytype" runat="server" BackColor="#C0FFFF" Height="20px"
                   Style="z-index: 129; left: 136px; position: absolute; top: 616px"
                   Width="352px">
               </asp:DropDownList>
               <asp:TextBox ID="DPayTypeOld" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="32px" MaxLength="255" Style="z-index: 101; left: 136px;
                   position: absolute; top: 640px; text-align: left" TextMode="MultiLine" Width="352px"></asp:TextBox>
               <asp:TextBox ID="DTJCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 137; left: 248px;
                   position: absolute; top: 424px; text-align: left" Width="48px"></asp:TextBox>
               &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DRemark" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="56px" MaxLength="255" Style="z-index: 101; left: 136px; position: absolute;
                   top: 552px; text-align: left" TextMode="MultiLine" Width="560px"></asp:TextBox>
               <asp:TextBox ID="DNameCH" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="255" Style="z-index: 102; left: 160px;
                   position: absolute; top: 320px; text-align: left" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DCustoms" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 103; left: 472px;
                   position: absolute; top: 288px; text-align: left" Width="120px"></asp:TextBox>
               <asp:DropDownList ID="DDelivery" runat="server" BackColor="#C0FFFF" Height="20px"
                   Style="z-index: 104; left: 136px; position: absolute; top: 288px"
                   Width="224px">
               </asp:DropDownList>
               &nbsp;
               <asp:TextBox ID="DTEL1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="15" Style="z-index: 106; left: 136px;
                   position: absolute; top: 224px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp;
               <input id="Button1" runat="server" style="z-index: 142; left: 336px; width: 24px;
                   position: absolute; top: 256px" type="button" value="..." />
               &nbsp;
               <asp:TextBox ID="DNameEN" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="255" Style="z-index: 107; left: 160px;
                   position: absolute; top: 352px; text-align: left" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DInvoiceCH" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="255" Style="z-index: 108; left: 160px;
                   position: absolute; top: 488px; text-align: left" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DInvoiceEN" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="255" Style="z-index: 109; left: 160px;
                   position: absolute; top: 520px; text-align: left" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DAddCH" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="56px" MaxLength="255" Style="z-index: 110; left: 304px;
                   position: absolute; top: 384px; text-align: left" Width="392px"></asp:TextBox>
               <asp:TextBox ID="DAddEN" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="255" Style="z-index: 111; left: 160px;
                   position: absolute; top: 456px; text-align: left" Width="536px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DCustomerCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 112; left: 136px;
                   position: absolute; top: 158px; text-align: left" Width="220px"></asp:TextBox>
               &nbsp;
               <asp:DropDownList ID="DLocation" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 114;
                   left: 136px; position: absolute; top: 192px" Width="224px">
               </asp:DropDownList>
               <asp:TextBox ID="DFAX1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="15" Style="z-index: 115; left: 472px;
                   position: absolute; top: 224px; text-align: left" Width="216px"></asp:TextBox>
               <asp:TextBox ID="DRelationCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 116; left: 472px;
                   position: absolute; top: 158px; text-align: left" Width="220px"></asp:TextBox>
               <asp:TextBox ID="DIDNumber" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="8" Style="z-index: 117; left: 472px;
                   position: absolute; top: 192px; text-align: left" Width="220px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 118; left: 136px; position: absolute;
                   top: 90px; text-align: left" Width="220px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 119;
                   left: 472px; position: absolute; top: 88px">DNo</asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 120; left: 136px;
                   position: absolute; top: 125px; text-align: left" Width="220px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DAppName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 121; left: 472px;
                   position: absolute; top: 125px; text-align: left" Width="220px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 128; left: 768px; position: absolute; top: 300px"
                   Width="190px"></asp:TextBox><asp:DropDownList ID="DSales" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 129;
                   left: 136px; position: absolute; top: 256px" Width="224px">
                   </asp:DropDownList>
               <asp:DropDownList ID="DGoods" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 130;
                   left: 472px; position: absolute; top: 256px" Width="224px">
               </asp:DropDownList>
               <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 131; left: 769px; position: absolute; top: 337px"
                   Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 132; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp;
               <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                   ErrorMessage="不可為空白" Width="96px"></asp:RequiredFieldValidator>
               &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;
           </div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DPostCode" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 137; left: 248px;
               position: absolute; top: 392px; text-align: left" Width="48px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:HyperLink ID="LAttachfile" runat="server" Style="z-index: 138; left: 144px;
               position: absolute; top: 688px">附檔</asp:HyperLink>
           &nbsp;&nbsp;
  	 
      </div>
    </form>
</body>
</html>
