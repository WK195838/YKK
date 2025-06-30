<%@ Page Language="vb" AutoEventWireup="false" Inherits="DiscountSheet_01" aspCompat="True" EnableEventValidation = "false"  CodeFile="DiscountSheet_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>顧客折扣單價申請</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
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
            
            			
   function GetCustomer()
{
        
    window.open('CustomerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 
            


   function GetBuyer()
{
        
    window.open('BuyerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	


  function Copy(wformno)
{
        
    window.open('DiscountCopy.aspx?formno='+wformno,'','status=0,toolbar=0,width=700,height=650,top=10,resizable=yes');
   
}	 
	

		</script>
		
		<style type="text/css">
        .TextUpper
        {
            text-transform:uppercase;
        }
</style>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/discount_02.jpg" style="z-index: 1; left: 8px; position: absolute;
                   top: 8px" />
               <asp:Label ID="Label1" runat="server" Style="z-index: 142; left: 496px; position: absolute;
                   top: 296px" Text="原始表單"></asp:Label>
               <asp:Label ID="Label3" runat="server" Style="z-index: 142; left: 248px; position: absolute;
                   top: 288px" Text="~"></asp:Label>
               <input id="BCopySheet" runat="server" style="z-index: 156; left: 656px; width: 40px;
                   position: absolute; top: 88px" type="button" value="複製" visible="true" />
               <asp:CheckBox ID="DReRegister" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 174;
                   left: 624px; position: absolute; top: 56px" Text="繼續申請" Width="76px" />
               &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDiscount0" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
                   position: absolute; top: 456px; text-align: left" Width="96px"></asp:TextBox>
               <asp:TextBox ID="DSize1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 456px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DChain0" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 456px; text-align: left" Width="64px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DCOM0" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 456px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="TextBox12" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 376px;
                   position: absolute; top: 488px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 488px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="TextBox18" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 456px;
                   position: absolute; top: 488px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="TextBox19" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 536px;
                   position: absolute; top: 488px; text-align: left" Width="56px"></asp:TextBox>
               <asp:TextBox ID="DItemCode0" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
                   position: absolute; top: 456px; text-align: left" Width="120px"></asp:TextBox>
               <asp:DropDownList ID="DST12" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 520px" Visible="False" Width="40px">
                   </asp:DropDownList><asp:DropDownList ID="DST21" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 208px; position: absolute; top: 488px" Visible="False" Width="40px">
                   </asp:DropDownList>
               <asp:DropDownList ID="DST31" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 488px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST41" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 488px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST51" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 488px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST11" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 488px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST20" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 456px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST30" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 456px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST40" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 456px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST50" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 456px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST13" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 552px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST14" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 584px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST15" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 616px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST16" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 648px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST17" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 688px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST18" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 720px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST19" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 752px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST10" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 456px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST29" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 752px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST39" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 752px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST49" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 752px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST59" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 752px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST28" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 720px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST38" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 720px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST48" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 720px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST58" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 720px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST27" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 688px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST37" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 688px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST47" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 688px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST57" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 688px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST26" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 648px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST36" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 648px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST46" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 648px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST56" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 648px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST25" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 616px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST35" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 616px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST45" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 616px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST55" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 616px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST24" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 584px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST34" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 584px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST44" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 584px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST54" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 584px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST23" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 552px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST33" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 552px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST43" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 552px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST53" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 552px" Visible="False" Width="40px">
               </asp:DropDownList>
               <asp:TextBox ID="DItemCode1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
                   position: absolute; top: 488px; text-align: left" Width="120px"></asp:TextBox><asp:TextBox ID="DItemCode2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
                   position: absolute; top: 520px; text-align: left" Width="120px"></asp:TextBox>
               <asp:TextBox ID="DChain1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 456px;
                   position: absolute; top: 488px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 488px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DSize3" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 520px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize4" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 552px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize5" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 584px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize6" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 620px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize7" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 652px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize8" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 684px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize9" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 720px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DSize0" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 376px; position: absolute;
                   top: 752px; text-align: left" Width="64px"></asp:TextBox>
               <asp:TextBox ID="DChain9" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 752px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM9" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 752px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain8" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 720px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM8" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 720px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain7" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 684px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM7" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 684px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain6" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 652px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM6" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 652px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain5" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 620px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM5" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 620px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain4" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 584px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM4" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 584px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain3" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 552px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM3" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 552px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <asp:TextBox ID="DChain2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 456px; position: absolute;
                   top: 520px; text-align: left" Width="64px" CssClass="TextUpper"></asp:TextBox>
               <asp:TextBox ID="DCOM2" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 536px; position: absolute;
                   top: 520px; text-align: left" Width="56px"  CssClass="TextUpper" ></asp:TextBox>
               <input id="BBuyer" runat="server" style="z-index: 152; left: 672px; width: 24px;
                   position: absolute; top: 192px" type="button" value="..." />
               <asp:TextBox ID="DAEDate" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="20px" Style="z-index: 110; left: 272px; position: absolute;
                   top: 288px" Width="72px"></asp:TextBox>
               <asp:Button ID="BASDate" runat="server" Height="26px" Style="z-index: 111; left: 216px;
                   position: absolute; top: 288px" Text="....." Width="28px" />
               <asp:Button ID="BAEDate" runat="server" Height="26px" Style="z-index: 112; left: 352px;
                   position: absolute; top: 288px" Text="....." Width="28px" />
               <asp:TextBox ID="DASDate" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 144px; position: absolute;
                   top: 288px" Width="72px"></asp:TextBox>
               <asp:TextBox ID="DVersion" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 136px;
                   position: absolute; top: 256px; text-align: left" Width="560px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DCurrency" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 136px; position: absolute; top: 224px" Visible="False" Width="224px">
               </asp:DropDownList>
               <asp:DropDownList ID="DJob" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 136px; position: absolute; top: 192px" Visible="False" Width="224px">
               </asp:DropDownList>
               <asp:TextBox ID="DAReason" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="56px" MaxLength="7" Style="z-index: 126; left: 136px; position: absolute;
                   top: 320px; text-align: left" TextMode="MultiLine" Width="560px"></asp:TextBox>
               &nbsp;
               <input id="BCustomer" runat="server" style="z-index: 151; left: 336px; width: 24px;
                   position: absolute; top: 160px" type="button" value="..." />
               &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DDate" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 136px; position: absolute;
                   top: 88px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DNo" runat="server" BackColor="White" BorderColor="White" BorderStyle="None"
                   Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
                   left: 472px; position: absolute; top: 88px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDepName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 136px;
                   position: absolute; top: 120px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DAppName" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 472px;
                   position: absolute; top: 120px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 1; left: 24px;
                   position: absolute; top: 1032px" visible="false" />
               <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 132; left: 72px; position: absolute; top: 960px"
                   TextMode="MultiLine" Width="536px"></asp:TextBox>
               <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
                   Style="z-index: 133; left: 248px; position: absolute; top: 1040px" Visible="False"
                   Width="352px"></asp:TextBox>
               <asp:DropDownList ID="DReasonCode" runat="server" BackColor="Yellow" Height="20px"
                   Style="z-index: 134; left: 176px; position: absolute; top: 1040px" Visible="False"
                   Width="64px">
                   <asp:ListItem>01</asp:ListItem>
                   <asp:ListItem>02</asp:ListItem>
               </asp:DropDownList>
               <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
                   Height="56px" Style="z-index: 135; left: 176px; position: absolute; top: 1072px"
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
                   left: 24px; position: absolute; top: 952px" />
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 132; left: 776px; position: absolute; top: 344px"
                   Width="190px"></asp:TextBox>
               <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 133; left: 776px; position: absolute; top: 312px"
                   Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp;
         
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 114; left: 354px; position: absolute; top: 1144px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 115; left: 445px; position: absolute; top: 1144px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 116; left: 537px; position: absolute; top: 1144px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 117; left: 629px; position: absolute; top: 1144px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" />
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DItemCode3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 552px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 584px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 620px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 652px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 684px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode8" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 720px; text-align: left" Width="120px"></asp:TextBox>
           <asp:TextBox ID="DItemCode9" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 32px;
               position: absolute; top: 752px; text-align: left" Width="120px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<asp:DropDownList ID="DST22" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 208px; position: absolute; top: 520px" Visible="False" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST32" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 520px" Visible="False" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST42" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 520px" Visible="False" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST52" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 520px" Visible="False" Width="40px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DDiscount1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 488px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 520px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 552px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount4" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 584px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 620px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount6" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 652px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount7" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 684px; text-align: left" Width="96px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DDiscount8" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 720px; text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DDiscount9" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 600px;
               position: absolute; top: 752px; text-align: left" Width="96px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
               ErrorMessage="不可為空白" Width="96px"></asp:RequiredFieldValidator>
           <asp:TextBox ID="DNO1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 104; left: 32px; position: absolute; top: 40px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DCustomer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
               ForeColor="Black" Height="24px" Style="z-index: 109; left: 480px; position: absolute;
               top: 157px" Width="216px"></asp:TextBox>
           <asp:TextBox ID="DCustomerCode" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
               ForeColor="Black" Height="24px" Style="z-index: 126; left: 136px; position: absolute;
               top: 157px" Width="200px"></asp:TextBox>
           <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" BorderStyle="Groove" ForeColor="Black"
               Height="24px" Style="z-index: 112; left: 480px; position: absolute; top: 220px"
               Width="216px"></asp:TextBox>
           <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
               BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 127; left: 480px;
               position: absolute; top: 188px" Width="192px"></asp:TextBox>
           <asp:Button ID="DAttachfile" runat="server" CausesValidation="False" Style="z-index: 274;
               left: 136px; position: absolute; top: 779px" Text="開啟附檔資料夾" Width="261px" />
           &nbsp;
           <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 133; left: 768px; position: absolute; top: 416px"
               Width="190px"></asp:TextBox><asp:CheckBox ID="DSaveFile" runat="server" Font-Size="10pt" Style="z-index: 174;
                   left: 416px; position: absolute; top: 784px" Text="附檔保留" Width="76px" />
           <asp:TextBox ID="D4" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 133; left: 776px; position: absolute;
               top: 376px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DOFormSno" runat="server" BackColor="#FFFF80" BorderColor="White" BorderStyle="Groove"
               Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black" Style="z-index: 100;
               left: 568px; position: absolute; top: 288px" Width="88px" ReadOnly="True"></asp:TextBox>
           <asp:HyperLink ID="LNo" runat="server" Style="z-index: 115; left: 656px; position: absolute;
               top: 296px" Target="_blank" Visible="False">LINK</asp:HyperLink>
           <asp:CheckBox ID="ChkYear" runat="server" Font-Size="10pt" Style="z-index: 174;
                   left: 384px; position: absolute; top: 288px" Text="特價延長" Width="104px" AutoPostBack="True" ForeColor="Blue" Visible="False" />
           <asp:TextBox ID="DADDSDATE" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 133; left: 768px; position: absolute;
               top: 496px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DADDEDATE" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 133; left: 768px; position: absolute;
               top: 456px" Width="190px"></asp:TextBox>
       </div>
    </form>
</body>
</html>
