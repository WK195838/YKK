<%@ Page Language="vb" AutoEventWireup="false" Inherits="ComplaintOutSheet_03" aspCompat="True" EnableEventValidation = "false"  CodeFile="ComplaintOutSheet_03.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>客訴處理報告書</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			               function GetDep()
{
        
    window.open('DepList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
		
			               function GetError(Str)
{
        
    window.open('ErrorList.aspx?Str=' + Str,'','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
			
		
		            			
   function GetCustomer()
{
        
    window.open('CustomerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 
            


   function GetBuyer()
{
        
    window.open('BuyerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	
	
	
   function GetSPEC()
{
        
    window.open('SPECList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
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
            


function IMG1_onclick() {

}

		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/ComplaintOutSheet_Sales.jpg" style="z-index: 99; left: 16px; position: absolute;
                   top: 0px" id="IMG1" onclick="return IMG1_onclick()" /><asp:CheckBox ID="DMAT" runat="server" Style="z-index: 183;
               left: 600px; position: absolute; top: 1096px" Width="104px" ForeColor="Red" Text="物料抱怨單" /><asp:CheckBox ID="DRELY" runat="server" Style="z-index: 183;
               left: 600px; position: absolute; top: 1072px" Width="104px" ForeColor="Red" Text="轉分析依賴" />
               <asp:TextBox ID="DEPCONTENT" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="152px" MaxLength="250" Style="z-index: 117; left: 80px;
                   position: absolute; top: 600px" TextMode="MultiLine" Width="296px"></asp:TextBox>
               &nbsp;&nbsp;
               <asp:Image ID="LMapFile" runat="server" BorderStyle="Groove" Height="320px" Style="z-index: 100;
                   left: 400px; position: absolute; top: 400px" Width="304px" />
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DACCDEP2NO" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
                   BorderStyle="None" ForeColor="Black" Height="20px" Style="z-index: 101; left: 336px;
                   position: absolute; top: 1136px" Width="48px"></asp:TextBox>
               <asp:DropDownList ID="DSCORE1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 102;
               left: 256px; position: absolute; top: 1768px" Visible="False" Width="112px">
               </asp:DropDownList>
               <asp:DropDownList ID="DMANUREASON1" runat="server" BackColor="Yellow" Height="20px" Style="z-index: 103; left: 504px; position: absolute;
                   top: 1200px" Width="192px" Visible="False">
                   </asp:DropDownList>
               &nbsp;&nbsp;
               <asp:TextBox ID="DHOUR3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 104; left: 176px; position: absolute;
                   top: 1368px" Width="48px"></asp:TextBox><asp:TextBox ID="DHOUR4" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
                   Height="20px" Style="z-index: 105; left: 544px; position: absolute; top: 1368px"
                   Width="48px"></asp:TextBox>
               <asp:TextBox ID="DHOUR5" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 106; left: 152px; position: absolute;
                   top: 1400px" Width="56px"></asp:TextBox>
               <input id="BYFGCDATE" runat="server" style="z-index: 199; left: 616px; width: 24px;
               position: absolute; top: 1800px" type="button" value="..." />
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="TextBox2" runat="server" BackColor="Transparent" BorderStyle="None"
                   ForeColor="White" Height="24px" Style="z-index: 107; left: 1384px; position: absolute;
                   top: 40px" Width="190px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              
            
               <asp:DropDownList ID="DMANUREASON" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 108; left: 144px; position: absolute;
                   top: 1200px" Width="240px" Visible="False" AutoPostBack="True">
                   </asp:DropDownList>
               &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DACCDEP1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 109; left: 168px; position: absolute;
                   top: 1000px" Width="136px" Visible="False" AutoPostBack="True">
               </asp:DropDownList>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DBUYER" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 110; left: 248px; position: absolute;
                   top: 304px" Width="144px"></asp:TextBox>
               <input id="BBuyer" runat="server" style="z-index: 202; left: 216px; width: 24px;
                   position: absolute; top: 304px" type="button" value="..." />
               <asp:TextBox ID="DBUYERCODE" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
                   BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 111; left: 152px;
                   position: absolute; top: 304px" Width="56px"></asp:TextBox>
               <asp:TextBox ID="DCUSTOMER" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 112; left: 248px; position: absolute;
                   top: 136px" Width="136px"></asp:TextBox>
               <input id="BCustomer" runat="server" style="z-index: 201; left: 216px; width: 24px;
                   position: absolute; top: 136px" type="button" value="..." />
               <asp:TextBox ID="DCUSTOMERCODE" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
                   ForeColor="Black" Height="24px" Style="z-index: 113; left: 152px; position: absolute;
                   top: 136px" Width="64px"></asp:TextBox>
               &nbsp;
               <input id="BORDERDATE" runat="server" style="z-index: 198; left: 272px; width: 24px;
                   position: absolute; top: 240px" type="button" value="..." />
               &nbsp;&nbsp;&nbsp;&nbsp;
               <asp:DropDownList ID="DACCDEP2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 114; left: 168px; position: absolute;
                   top: 1136px" Width="112px" Visible="False" AutoPostBack="True">
               </asp:DropDownList>
               &nbsp; &nbsp;
               <asp:DropDownList ID="DJUDGE" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 115; left: 520px; position: absolute;
                   top: 1136px" Width="184px" Visible="False" AutoPostBack="True">
               </asp:DropDownList>
               <asp:DropDownList ID="DITEM" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 116; left: 400px; position: absolute;
                   top: 768px" Width="296px" Visible="False">
               </asp:DropDownList>
               &nbsp; &nbsp;
               <asp:TextBox ID="DCPCONTENT" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="152px" Style="z-index: 117; left: 80px; position: absolute;
                   top: 432px" Width="296px" TextMode="MultiLine" MaxLength="250"></asp:TextBox>
               DACCDEP12
               <asp:TextBox ID="DSHIPDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 118; left: 512px; position: absolute;
                   top: 240px" Width="72px"></asp:TextBox>
               <input id="BSHIPDATE" runat="server" style="z-index: 194; left: 608px; width: 24px;
                   position: absolute; top: 240px" type="button" value="..." />
               &nbsp;
               <asp:DropDownList ID="DGLOBAL" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 119;
                   left: 152px; position: absolute; top: 104px" Width="128px" Visible="False">
                   </asp:DropDownList>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DBIGGOODS" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 120; left: 152px; position: absolute;
                   top: 272px" Width="240px" Visible="False">
               </asp:DropDownList>
               &nbsp;&nbsp;
               <asp:TextBox ID="DOORNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 121; left: 168px; position: absolute;
                   top: 176px" Width="216px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp; &nbsp;
               <asp:TextBox ID="DORDERDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 122; left: 152px; position: absolute;
                   top: 240px" Width="112px"></asp:TextBox>
               <input id="BDATE" runat="server" style="z-index: 196; left: 616px; width: 24px;
                   position: absolute; top: 104px" type="button" value="..." />
               &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:TextBox ID="DNORNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 123; left: 528px; position: absolute;
                   top: 176px" Width="160px" Visible="False"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;
               &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:DropDownList ID="DAPPNAME" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 124;
                   left: 512px; position: absolute; top: 144px" Visible="False" Width="176px">
               </asp:DropDownList>
               &nbsp;
               <asp:TextBox ID="DACCEMPNAME" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 125;
                   left: 144px; position: absolute; top: 1032px" Width="78px" BorderStyle="None" ForeColor="Black"></asp:TextBox>
               &nbsp; &nbsp;
               &nbsp; &nbsp;
               <asp:TextBox ID="DQCNO" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 144px; position: absolute;
                   top: 832px" Width="78px"></asp:TextBox>
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DORQTY"  onkeyup="if(isNaN(value))execCommand('undo')"   runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 127; left: 512px; position: absolute;
                   top: 208px" Width="64px"  onafterpaste="if(isNaN(value))execCommand('undo')" ></asp:TextBox>
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 128; left: 512px; position: absolute;
                   top: 112px" Width="96px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;&nbsp;
               <asp:TextBox ID="DNO" runat="server" BackColor="White" BorderStyle="None"
                   ForeColor="Black" Height="24px" Style="z-index: 129; left: 560px; position: absolute;
                   top: 72px" Width="144px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;
               &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp;
               &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<input id="BDELIVERYDATE" runat="server" style="z-index: 196; left: 288px; width: 24px;
                   position: absolute; top: 336px" type="button" value="..." />
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 137; left: 1168px; position: absolute; top: 144px"
                   Width="190px"></asp:TextBox>
               <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 138; left: 1152px; position: absolute; top: 200px"
                   Width="190px"></asp:TextBox>
                    <asp:TextBox ID="DCODE" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                   Height="24px" Style="z-index: 138; left: 1104px; position: absolute; top: 256px"
                   Width="190px"></asp:TextBox>
               <!--************************************************************ 
            ** 防止重覆按 [BUTTON 2 次]                               **     
            ************************************************************ -->
               <asp:TextBox ID="TextBox1" runat="server" Style="z-index: 139; left: -500px; position: absolute;
                   top: 100px; text-align: left">AAA</asp:TextBox>
               &nbsp;
               <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                   ErrorMessage="不可為空白" Width="96px"></asp:RequiredFieldValidator>
               &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               &nbsp; &nbsp;&nbsp;&nbsp;
               &nbsp; &nbsp;&nbsp;
               <asp:TextBox ID="DDELIVERYDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
                   ForeColor="Black" Height="20px" Style="z-index: 126; left: 152px; position: absolute;
                   top: 336px" Width="128px"></asp:TextBox>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="24px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 140; left: 624px; position: absolute; top: 1968px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               &nbsp; &nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="D3" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 144; left: 784px; position: absolute; top: 80px"
               Width="190px"></asp:TextBox>
           <asp:TextBox ID="D4" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 145; left: 1160px; position: absolute; top: 64px"
               Width="190px"></asp:TextBox>
           &nbsp;
           <asp:DropDownList ID="DRESPONS" runat="server" BackColor="#C0FFFF" Height="20px"
               Style="z-index: 146; left: 168px; position: absolute; top: 1232px" Visible="False"
               Width="216px">
           </asp:DropDownList><asp:DropDownList ID="DHappen" runat="server" BackColor="#C0FFFF" Height="20px"
               Style="z-index: 146; left: 48px; position: absolute; top: 1304px" Visible="False"
               Width="216px">
           </asp:DropDownList>
           <input id="BQCDATE" runat="server" style="z-index: 195; left: 608px; width: 24px;
               position: absolute; top: 1440px" type="button" value="..." />
           <asp:TextBox ID="DQCDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 147; left: 504px; position: absolute; top: 1440px"
               Width="96px"></asp:TextBox>
           <asp:DropDownList ID="DREMARK" runat="server" BackColor="#C0FFFF" Height="20px"
               Style="z-index: 148; left: 504px; position: absolute; top: 1504px" Visible="False"
               Width="168px">
           </asp:DropDownList>
           <asp:TextBox ID="DBCOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 149; left: 504px; position: absolute; top: 1536px"
               Width="96px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DACOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 150; left: 152px; position: absolute; top: 1536px"
               Width="96px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DCCOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 151; left: 152px; position: absolute; top: 1568px"
               Width="96px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DDCOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 152; left: 504px; position: absolute; top: 1568px"
               Width="96px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DECOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 153; left: 176px; position: absolute; top: 1600px"
               Width="96px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DALLCOST" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 154; left: 592px; position: absolute;
               top: 1632px" Width="96px"></asp:TextBox>
           <asp:DropDownList ID="DPOINT1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 155;
               left: 144px; position: absolute; top: 1664px" Visible="False" Width="88px" AutoPostBack="True">
           </asp:DropDownList>
           &nbsp;
           <asp:TextBox ID="DPOINT3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 156; left: 208px; position: absolute;
               top: 1704px" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DCOST" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 157; left: 144px; position: absolute;
               top: 1736px; text-align: left" Width="552px"></asp:TextBox>
           <asp:TextBox ID="DREMARK1" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="56px" Style="z-index: 158; left: 144px;
               position: absolute; top: 1064px; text-align: left" Width="448px"></asp:TextBox>
           <asp:TextBox ID="DTYPE1" runat="server" AutoPostBack="True" BackColor="#C0FFFF"
               BorderStyle="None" ForeColor="Black" Height="20px" Style="z-index: 101; left: 336px;
               position: absolute; top: 864px" Width="184px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DREMARK2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 159; left: 144px;
               position: absolute; top: 1864px; text-align: left" Width="552px"></asp:TextBox>
           <asp:TextBox ID="DREMARK3" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 160; left: 232px;
               position: absolute; top: 1896px; text-align: left" Width="472px"></asp:TextBox>
           <asp:TextBox ID="DCUSTDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 161; left: 152px; position: absolute;
               top: 1800px" Width="96px"></asp:TextBox>
           <input id="BCUSTDATE" runat="server" style="z-index: 197; left: 256px; width: 24px;
               position: absolute; top: 1800px" type="button" value="..." />
           <asp:Button ID="DAttachfile1" runat="server" CausesValidation="False" Style="z-index: 162;
               left: 520px; position: absolute; top: 792px" Text="開啟附檔" Width="80px" BackColor="#FFC080" />
           <asp:Button ID="DAttachfile2" runat="server" CausesValidation="False" Style="z-index: 163;
               left: 528px; position: absolute; top: 1024px" Text="開啟附檔" Width="80px" BackColor="#FFC080" />
           <asp:Button ID="DAttachfile3" runat="server" CausesValidation="False" Style="z-index: 164;
               left: 528px; position: absolute; top: 1400px" Text="開啟附檔" Width="80px" BackColor="#FFC080" />
           <asp:Button ID="DAttachfile4" runat="server" CausesValidation="False" Style="z-index: 165;
               left: 568px; position: absolute; top: 1696px" Text="開啟附檔" Width="80px" BackColor="#FFC080" />
           <asp:Button ID="DAttachfile5" runat="server" CausesValidation="False" Style="z-index: 166;
               left: 168px; position: absolute; top: 1928px" Text="開啟附檔" Width="80px" BackColor="#FFC080" />
           <asp:DropDownList ID="DREPLAYLAN" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 167; left: 168px; position: absolute;
                   top: 800px" Width="168px" Visible="False">
           </asp:DropDownList>
           <asp:DropDownList ID="DCUSTOMERTYPE" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 168; left: 312px; position: absolute;
                   top: 1032px" Width="80px" Visible="False">
           </asp:DropDownList><input id="BSAMPLEDATE" runat="server" style="z-index: 200; left: 256px; width: 24px;
               position: absolute; top: 1432px" type="button" value="..." />
           <asp:TextBox ID="DCPQTY" onkeyup="if(isNaN(value))execCommand('undo')"  runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 169; left: 512px; position: absolute; top: 304px"
               Width="56px" onafterpaste="if(isNaN(value))execCommand('undo')"  ></asp:TextBox>
           <asp:TextBox ID="DHOUR1" runat="server" BackColor="#C0FFFF" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 170; left: 208px; position: absolute; top: 1336px"
               Width="48px"></asp:TextBox>
           <asp:TextBox ID="DANSWERDAYS" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 171; left: 152px; position: absolute;
               top: 1504px" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DSAMPLEDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 172; left: 152px; position: absolute;
               top: 1440px" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DYFGCDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 173; left: 512px; position: absolute;
               top: 1800px" Width="96px"></asp:TextBox><asp:DropDownList ID="DTYPE" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 174; left: 168px; position: absolute;
                   top: 864px" Width="136px" Visible="False" AutoPostBack="True">
               </asp:DropDownList>
           <asp:TextBox ID="DHOUR2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 175; left: 504px; position: absolute;
               top: 1336px" Width="48px"></asp:TextBox>
        
           <asp:CheckBox ID="DCPQTYCHK" runat="server" Style="z-index: 176;
               left: 576px; position: absolute; top: 304px" Width="104px" Text="數量確認中" />
           &nbsp;
               <asp:TextBox ID="DWORKPAY" runat="server" BackColor="Transparent" BorderStyle="None"
                   ForeColor="White" Height="24px" Style="z-index: 177; left: 848px; position: absolute;
                   top: 216px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="DPOINT2" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 178; left: 504px; position: absolute;
               top: 1672px" Width="96px"></asp:TextBox><asp:DropDownList ID="DSCORE2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 179;
               left: 416px; position: absolute; top: 1768px" Visible="False" Width="120px">
               </asp:DropDownList>
           <asp:DropDownList ID="DSCORE3" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 180;
               left: 592px; position: absolute; top: 1768px" Visible="False" Width="112px">
           </asp:DropDownList>
           <asp:DropDownList ID="DPLACE" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 181; left: 504px; position: absolute;
                   top: 272px" Width="200px" Visible="False">
           </asp:DropDownList><asp:CheckBox ID="DChkData1" runat="server" Style="z-index: 182;
               left: 608px; position: absolute; top: 800px" Width="104px" ForeColor="Red" />
           <asp:CheckBox ID="DChkData2" runat="server" Style="z-index: 183;
               left: 608px; position: absolute; top: 1032px" Width="104px" ForeColor="Red" />
           <asp:CheckBox ID="DChkData3" runat="server" Style="z-index: 184;
               left: 608px; position: absolute; top: 1408px" Width="104px" ForeColor="Red" />
           <asp:CheckBox ID="DChkData4" runat="server" Style="z-index: 185;
               left: 648px; position: absolute; top: 1704px" Width="104px" ForeColor="Red" />
           <asp:CheckBox ID="DChkData5" runat="server" Style="z-index: 186;
               left: 256px; position: absolute; top: 1928px" Width="104px" ForeColor="Red" />
           <asp:TextBox ID="chktemp" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 187; left: 848px; position: absolute;
               top: 296px" Width="190px"></asp:TextBox>
           <asp:FileUpload ID="DMapFile" runat="server" Height="24px" Style="z-index: 188; left: 400px;
               position: absolute; top: 728px; background-color: #c0ffff" Width="304px" />
           <asp:TextBox ID="DMappath" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 189; left: 848px; position: absolute;
               top: 256px" Width="190px"></asp:TextBox>
           <asp:HyperLink ID="LMapFile1" runat="server" ForeColor="Blue" Height="24px" Style="z-index: 190;
               left: 400px; position: absolute; top: 704px" Target="_blank" ToolTip="簡圖放大" Width="304px">簡圖放大</asp:HyperLink>
           &nbsp;&nbsp;
           <asp:TextBox ID="DGCOST" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 153; left: 208px; position: absolute;
               top: 1632px" Width="56px"></asp:TextBox>
           <asp:TextBox ID="DHCOST" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 153; left: 336px; position: absolute;
               top: 1632px" Width="56px"></asp:TextBox>
           <asp:TextBox ID="DFCOST" runat="server" AutoPostBack="True" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 153; left: 592px; position: absolute;
               top: 1600px" Width="96px"></asp:TextBox>
           <asp:DropDownList ID="DPL" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 109; left: 312px; position: absolute;
                   top: 832px" Width="72px" Visible="False">
           </asp:DropDownList>
           <asp:TextBox ID="DSPEC" runat="server" AutoCompleteType="Disabled" BackColor="#C0FFFF"
               BorderStyle="Groove" ForeColor="Black" Height="24px" Style="z-index: 127; left: 152px;
               position: absolute; top: 205px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="DSPECNAME" runat="server" BackColor="#C0FFFF" BorderStyle="Groove"
               ForeColor="Black" Height="24px" Style="z-index: 112; left: 240px; position: absolute;
               top: 205px" Width="152px"></asp:TextBox>
           <input id="BSPEC" runat="server" style="z-index: 152; left: 216px; width: 24px;
               position: absolute; top: 208px" type="button" value="..." /><asp:DropDownList ID="DACCDEP12" runat="server" BackColor="Yellow" Height="20px" Style="z-index: 109; left: 336px; position: absolute;
                   top: 1000px" Width="176px" Visible="False" AutoPostBack="True">
               </asp:DropDownList>
           <asp:DropDownList ID="DACCDEP13" runat="server" BackColor="Yellow" Height="20px" Style="z-index: 109; left: 544px; position: absolute;
                   top: 1000px" Width="160px" Visible="False">
           </asp:DropDownList>
           <asp:TextBox ID="DQCCDESC" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="56px" MaxLength="250" Style="z-index: 117; left: 168px;
               position: absolute; top: 928px" TextMode="MultiLine" Width="224px"></asp:TextBox>
           <asp:TextBox ID="DQCEDESC" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="56px" MaxLength="250" Style="z-index: 117; left: 416px;
               position: absolute; top: 928px" TextMode="MultiLine" Width="288px"></asp:TextBox><img src="images/ComplaintOutSheet_QC.jpg" style="z-index: 99; left: 16px; position: absolute;
                   top: 824px" id="Img2" onclick="return IMG1_onclick()" /><input id="BQCACCDATE" runat="server" style="z-index: 196; left: 672px; width: 24px;
                   position: absolute; top: 832px" type="button" value="..." /><asp:CheckBox ID="DMach1" runat="server" Style="z-index: 183;
               left: 288px; position: absolute; top: 1304px" Width="56px" ForeColor="Red" Text="機台" Font-Size="Smaller" />
           <asp:CheckBox ID="DMach2" runat="server" Style="z-index: 183;
               left: 544px; position: absolute; top: 1304px" Width="80px" ForeColor="Red" Text="非機台" Font-Size="Smaller" />
           <asp:TextBox ID="DMACHNO" runat="server" BackColor="Yellow" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 157; left: 336px; position: absolute; top: 1304px;
               text-align: left" Width="107px"></asp:TextBox>
           <asp:Label ID="Label1" runat="server" Font-Size="9pt" ForeColor="Black" Height="20px"
               Style="z-index: 135; left: 448px; position: absolute; top: 1312px" Text="該工程POP機台號"
               Width="104px"></asp:Label><asp:DropDownList ID="DJUDGE1" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 115; left: 168px; position: absolute;
                   top: 1168px" Width="200px" AutoPostBack="True">
               </asp:DropDownList>
           <asp:DropDownList ID="DJUDGE2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 115; left: 520px; position: absolute;
                   top: 1168px" Width="184px">
           </asp:DropDownList>
           <asp:DropDownList ID="DTYPE2" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 115; left: 544px; position: absolute;
                   top: 864px" Width="160px" AutoPostBack="True" Visible="False">
           </asp:DropDownList>
           <asp:TextBox ID="DCARDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 152px; position: absolute;
               top: 1472px" Width="128px"></asp:TextBox>
           <asp:TextBox ID="DFIRSTDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 160px; position: absolute;
               top: 1832px" Width="88px"></asp:TextBox>
           <asp:TextBox ID="DMIDDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 528px; position: absolute;
               top: 1832px" Width="128px"></asp:TextBox>
           <asp:TextBox ID="DQCACCDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 536px; position: absolute;
               top: 832px" Width="128px"></asp:TextBox><asp:DropDownList ID="DSAMPLE" runat="server" BackColor="#C0FFFF" Height="20px" Style="z-index: 124;
                   left: 512px; position: absolute; top: 336px" Visible="False" Width="192px">
               </asp:DropDownList><asp:DropDownList ID="DQCDESCTYPE" runat="server" BackColor="Yellow" Height="20px" Style="z-index: 115; left: 528px; position: absolute;
                   top: 896px" Width="176px">
               </asp:DropDownList><input id="BREPORTDATE" runat="server" style="z-index: 196; left: 256px; width: 24px;
                   position: absolute; top: 896px" type="button" value="..." />
           <asp:TextBox ID="DREPORTDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 128; left: 152px; position: absolute;
               top: 896px" Width="96px"></asp:TextBox><asp:Button ID="DAttachfile6" runat="server" CausesValidation="False" Style="z-index: 164;
               left: 400px; position: absolute; top: 1232px" Text="開啟附檔-CAR簽核版" Width="168px" BackColor="#FFC080" />
           <img src="images/ComplaintOutSheet_Manu.jpg" style="z-index: 99; left: 16px; position: absolute;
                   top: 1128px" id="Img3" onclick="return IMG1_onclick()" />
           <input id="BMFGDATE" runat="server" style="z-index: 200; left: 368px; width: 24px;
               position: absolute; top: 1400px" type="button" value="..." />
           <asp:TextBox ID="DMFGDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 172; left: 296px; position: absolute;
               top: 1400px" Width="64px"></asp:TextBox>
           <img src="images/ComplaintOutSheet_QC2.jpg" style="z-index: 99; left: 16px; position: absolute;
                   top: 1432px" id="Img4" onclick="return IMG1_onclick()" />
           <asp:TextBox ID="DFINALDATE" runat="server" BackColor="#C0FFFF" BorderStyle="None"
               ForeColor="Black" Height="20px" Style="z-index: 126; left: 288px; position: absolute;
               top: 1832px" Width="64px"></asp:TextBox>
           <input id="BFINALDATE" runat="server" style="z-index: 199; left: 352px; width: 24px;
               position: absolute; top: 1832px" type="button" value="..." />
           <input id="BCARDATE" runat="server" style="z-index: 200; left: 288px; width: 24px;
               position: absolute; top: 1472px" type="button" value="..." /><input id="BMIDDATE" runat="server" style="z-index: 199; left: 656px; width: 24px;
               position: absolute; top: 1832px" type="button" value="..." /><input id="BFIRSTDATE" runat="server" style="z-index: 199; left: 256px; width: 24px;
               position: absolute; top: 1832px" type="button" value="..." /><asp:CheckBox ID="DChkData6" runat="server" Style="z-index: 186;
               left: 576px; position: absolute; top: 1240px" Width="104px" ForeColor="Red" />
           </div>
    </form>
</body>
</html>
