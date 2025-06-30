<%@ Page Language="vb" AutoEventWireup="false" Inherits="DiscountSheet_02" aspCompat="True" EnableEventValidation = "false"  CodeFile="DiscountSheet_02.aspx.vb" %>
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
               <asp:Label ID="Label3" runat="server" Style="z-index: 142; left: 240px; position: absolute;
                   top: 288px" Text="~"></asp:Label>
               <asp:Label ID="Label1" runat="server" Style="z-index: 142; left: 496px; position: absolute;
                   top: 296px" Text="原始表單"></asp:Label>
               <asp:TextBox ID="DOFormSno" runat="server" BackColor="#FFFF80" BorderColor="White"
                   BorderStyle="Groove" Font-Names="Times New Roman" Font-Size="11pt" ForeColor="Black"
                   ReadOnly="True" Style="z-index: 100; left: 568px; position: absolute; top: 288px"
                   Width="88px"></asp:TextBox>
               <asp:HyperLink ID="LNo" runat="server" Style="z-index: 115; left: 656px; position: absolute;
                   top: 296px" Target="_blank" Visible="False">LINK</asp:HyperLink>
               <asp:CheckBox ID="ChkYear" runat="server" AutoPostBack="True" Font-Size="10pt" ForeColor="Blue"
                   Style="z-index: 174; left: 368px; position: absolute; top: 288px" Text="特價延長一年"
                   Visible="False" Width="104px" />
               <asp:HyperLink ID="LYearNo" runat="server" Style="z-index: 115; left: 624px; position: absolute;
                   top: 56px" Target="_blank" Visible="False">延長申請</asp:HyperLink>
               <asp:Button ID="DAttachfile" runat="server" CausesValidation="False" Style="z-index: 274;
                   left: 136px; position: absolute; top: 779px" Text="開啟附檔資料夾" Width="261px" />
               &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
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
                   left: 160px; position: absolute; top: 520px" Width="40px">
                   </asp:DropDownList><asp:DropDownList ID="DST21" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 208px; position: absolute; top: 488px" Width="40px">
                   </asp:DropDownList>
               <asp:DropDownList ID="DST31" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 488px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST41" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 488px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST51" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 488px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST11" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 488px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST20" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 456px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST30" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 456px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST40" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 456px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST50" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 456px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST13" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 552px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST14" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 584px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST15" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 616px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST16" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 648px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST17" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 688px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST18" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 720px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST19" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 752px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST10" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 160px; position: absolute; top: 456px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST29" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 752px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST39" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 752px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST49" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 752px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST59" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 752px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST28" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 720px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST38" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 720px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST48" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 720px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST58" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 720px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST27" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 688px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST37" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 688px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST47" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 688px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST57" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 688px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST26" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 648px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST36" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 648px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST46" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 648px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST56" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 648px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST25" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 616px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST35" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 616px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST45" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 616px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST55" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 616px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST24" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 584px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST34" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 584px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST44" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 584px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST54" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 584px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST23" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 205px; position: absolute; top: 552px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST33" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 552px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST43" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 552px" Width="40px">
               </asp:DropDownList>
               <asp:DropDownList ID="DST53" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 552px" Width="40px">
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
               &nbsp;
               <asp:TextBox ID="DAEDate" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="20px" Style="z-index: 110; left: 280px; position: absolute;
                   top: 288px" Width="72px"></asp:TextBox>
               &nbsp;&nbsp;
               <asp:TextBox ID="DASDate" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="20px" Style="z-index: 113; left: 144px; position: absolute;
                   top: 288px" Width="72px"></asp:TextBox>
               <asp:TextBox ID="DVersion" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 136px;
                   position: absolute; top: 256px; text-align: left" Width="560px"></asp:TextBox>
               <asp:TextBox ID="DBuyer" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 472px;
                   position: absolute; top: 224px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DCurrency" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 136px; position: absolute; top: 224px" Width="224px">
               </asp:DropDownList>
               <asp:DropDownList ID="DJob" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 136px; position: absolute; top: 192px" Width="224px">
               </asp:DropDownList>
               <asp:TextBox ID="DAReason" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="56px" MaxLength="7" Style="z-index: 126; left: 136px; position: absolute;
                   top: 320px; text-align: left" TextMode="MultiLine" Width="560px"></asp:TextBox>
               <asp:TextBox ID="DCustomerCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 136px; position: absolute;
                   top: 160px; text-align: left" Width="192px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DCustomer" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" MaxLength="7" Style="z-index: 126; left: 472px;
                   position: absolute; top: 160px; text-align: left" Width="224px"></asp:TextBox>
               &nbsp;
               <asp:TextBox ID="DBuyerCode" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" MaxLength="7" Style="z-index: 126; left: 472px; position: absolute;
                   top: 192px; text-align: left" Width="192px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp;
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
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 132; left: 768px; position: absolute; top: 300px"
                   Width="190px"></asp:TextBox>
               <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 133; left: 769px; position: absolute; top: 337px"
                   Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 126; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp; &nbsp; &nbsp;&nbsp;
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
                   left: 208px; position: absolute; top: 520px" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST32" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 248px; position: absolute; top: 520px" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST42" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 288px; position: absolute; top: 520px" Width="40px">
           </asp:DropDownList>
           <asp:DropDownList ID="DST52" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 134;
                   left: 328px; position: absolute; top: 520px" Width="40px">
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
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DNO1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 104; left: 32px; position: absolute; top: 40px"
               Width="142px"></asp:TextBox>
       </div>
    </form>
</body>
</html>
