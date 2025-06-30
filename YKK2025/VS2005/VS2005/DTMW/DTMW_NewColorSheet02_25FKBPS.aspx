<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorSheet02_25FKBPS" aspCompat="True" CodeFile="DTMW_NewColorSheet02_25FKBPS.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴書拉鏈ZIPPER CHAIN (05 CNL SBS16)</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
   function GetCustomer()
{
        
    window.open('CustomerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 

   function GetBuyer()
{
        
    window.open('BuyerList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 


   function GetYKKColor()
{
        
    window.open('YKKColorList.aspx?','','status=0,toolbar=0,width=620,height=650,top=10,resizable=yes');
   
}	 


   function GetYKKColorCode()
{
        
    window.open('YKKColorCodeList.aspx?','','status=0,toolbar=0,width=700,height=650,top=10,resizable=yes');
   
}	 


		    function Button(F, MSG) {
				//alert(F);
				document.cookie="RunBOK=False";
				document.cookie="RunBNG1=False";
				document.cookie="RunBNG2=False";
				document.cookie="RunBSAVE=False";

				answer = confirm("是否執行[" + MSG + "]作業嗎？ 請確認....");
				if (answer) {
					//OK Button
					//FOK=document.getElementById("BOK");
					//if(FOK!=null) document.Form1.BOK.disabled=true;  	
					//NG-1 Button
					//FNG1=document.getElementById("BNG1");
					//if(FNG1!=null) document.Form1.BNG1.disabled=true;  	
					//NG-2 Button
					//FNG2=document.getElementById("BNG2");
					//if(FNG2!=null) document.Form1.BNG2.disabled=true;  	
					//Save Button
					//FSAVE=document.getElementById("BSAVE");
					//if(FSAVE!=null) document.Form1.BSAVE.disabled=true;  	
						
					if (F=="OK")   document.cookie="RunBOK=True";
					if (F=="NG1")  document.cookie="RunBNG1=True";
					if (F=="NG2")  document.cookie="RunBNG2=True";
					if (F=="SAVE") document.cookie="RunBSAVE=True";
				}
			}
		   
		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
          <div>
              &nbsp; &nbsp; &nbsp;
              <asp:Image ID="Image2" runat="server" ImageUrl="~/Images/DTMW_NewColorSheetNO.jpg"
                  Style="z-index: 143; left: 619px; position: absolute; top: 1043px" />
              <asp:Label ID="LFormSno" runat="server" Font-Size="Small" ForeColor="Blue" Style="z-index: 145;
                  left: 51px; position: absolute; top: 1040px"></asp:Label>
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
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
              &nbsp;
              <br />
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
              &nbsp; &nbsp;
              <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                  Height="24px" Style="z-index: 135; left: 768px; position: absolute; top: 300px"
                  Width="190px"></asp:TextBox>
              <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
                  Height="24px" Style="z-index: 136; left: 769px; position: absolute; top: 337px"
                  Width="190px"></asp:TextBox>
              &nbsp; &nbsp;&nbsp;<br />
              <br />
              <br />
              &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              <br />
              <br />
              <br />
              <br />
              &nbsp; &nbsp;
              <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DTMW_NewColorSheet09.jpg"
                  Style="z-index: 100; left: 26px; position: absolute; top: 16px" />
              <asp:DropDownList ID="DAgain" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 119; left: 528px; position: absolute; top: 980px"
                  Width="64px">
              </asp:DropDownList>
              <asp:TextBox ID="DPFBWire" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 109; left: 459px;
                  text-transform: uppercase; position: absolute; top: 749px" Width="110px"></asp:TextBox>
              <asp:TextBox ID="DPFOpenParts" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 109; left: 459px;
                  text-transform: uppercase; position: absolute; top: 782px" Width="110px"></asp:TextBox>
              &nbsp;
              <asp:HyperLink ID="LComplete" runat="server" NavigateUrl="BoardEdit.aspx" Style="z-index: 140;
                  left: 56px; position: absolute; top: 126px" Target="_blank">有新色依賴完成表</asp:HyperLink>
              <asp:HyperLink ID="LAgain" runat="server" NavigateUrl="BoardEdit.aspx" Style="z-index: 140;
                  left: 228px; position: absolute; top: 126px" Target="_blank">有新色再現完成表</asp:HyperLink>
              <asp:DropDownList ID="DNewOldColor" runat="server" AutoPostBack="True" BackColor="Yellow"
                  ForeColor="Blue" Height="20px" Style="z-index: 102; left: 459px; position: absolute;
                  top: 877px" Width="117px">
              </asp:DropDownList>
              &nbsp;
              <asp:TextBox ID="DNO1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
                  Height="24px" Style="z-index: 104; left: 52px; position: absolute; top: 47px"
                  Width="142px"></asp:TextBox>
              <input id="BCopySheet" runat="server" style="z-index: 156; left: 414px; width: 24px;
                  position: absolute; top: 186px" type="button" value="..." visible="true" />
              &nbsp;&nbsp;
              <asp:CheckBox ID="DVCACheck" runat="server" AutoPostBack="True" BorderStyle="None"
                  Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 105; left: 615px;
                  position: absolute; top: 124px" Text="VCA" Width="78px" />
              <asp:CheckBox ID="DFactoryCheck" runat="server" AutoPostBack="True" BorderStyle="None"
                  Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 106; left: 492px;
                  position: absolute; top: 123px" Text="工廠自核" Width="99px" />
              <asp:CheckBox ID="DCustomerCheck" runat="server" AutoPostBack="True" BorderStyle="None"
                  Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 107; left: 367px;
                  position: absolute; top: 123px" Text="客戶自核" Width="108px" />
              <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" Style="z-index: 108; left: 217px; position: absolute;
                  top: 216px" Width="219px"></asp:TextBox>
              <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" Style="z-index: 109; left: 217px; position: absolute;
                  top: 250px" Width="190px"></asp:TextBox>
              &nbsp;&nbsp;
              <asp:TextBox ID="DCustomerColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 110; left: 217px;
                  text-transform: uppercase; position: absolute; top: 317px" Width="219px"></asp:TextBox>
              <asp:CheckBox ID="DNOCCS" runat="server" AutoPostBack="True" Font-Size="Small" Height="1px"
                  Style="z-index: 111; left: 235px; position: absolute; top: 485px" Width="1px" />
              <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                  Height="24px" Style="z-index: 112; left: 217px; position: absolute; top: 284px"
                  Width="188px"></asp:TextBox>
              &nbsp; &nbsp;&nbsp;
              <asp:TextBox ID="DVersion" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" Style="z-index: 113; left: 503px; position: absolute;
                  top: 909px; text-align: center" Width="37px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              <asp:TextBox ID="DYKKColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  Font-Bold="True" ForeColor="Blue" Height="24px" Style="z-index: 114; left: 459px;
                  text-transform: uppercase; position: absolute; top: 844px" Width="94px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              <input id="BCustomer" runat="server" style="z-index: 157; left: 415px; width: 24px;
                  position: absolute; top: 249px" type="button" value="..." />
              <input id="BBuyer" runat="server" style="z-index: 158; left: 415px; width: 24px;
                  position: absolute; top: 282px" type="button" value="..." />
              &nbsp;&nbsp;
              <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                  Height="24px" Style="z-index: 114; left: 597px; position: absolute; top: 184px"
                  Width="115px"></asp:TextBox>
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
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              <asp:DropDownList ID="DDevSeason" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 115; left: 327px; position: absolute; top: 419px"
                  Width="94px">
              </asp:DropDownList>
              <asp:DropDownList ID="DColorLight2" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 116; left: 450px; position: absolute; top: 451px"
                  Width="136px">
              </asp:DropDownList>
              <asp:DropDownList ID="DColorLight1" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 117; left: 217px; position: absolute; top: 450px"
                  Width="115px">
              </asp:DropDownList>
              <asp:DropDownList ID="DCustomerSample" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="25px" Style="z-index: 118; left: 66px; position: absolute; top: 746px"
                  Width="221px">
              </asp:DropDownList>
              <asp:DropDownList ID="DYKKColorType" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 119; left: 459px; position: absolute; top: 811px"
                  Width="117px">
              </asp:DropDownList>
              <asp:DropDownList ID="DDevYear" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="20px" Style="z-index: 120; left: 217px; position: absolute; top: 421px"
                  Width="76px">
              </asp:DropDownList>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp;
              <asp:TextBox ID="DReColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 121; left: 216px;
                  text-transform: uppercase; position: absolute; top: 384px" Width="219px"></asp:TextBox>
              <asp:TextBox ID="DOverseaYKKCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 121; left: 216px;
                  text-transform: uppercase; position: absolute; top: 352px" Width="219px"></asp:TextBox>
              <asp:TextBox ID="DPANTONECode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 122; left: 597px;
                  text-transform: uppercase; position: absolute; top: 367px" Width="117px"></asp:TextBox>
              &nbsp;
              <asp:TextBox ID="DCustomerColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 123; left: 597px;
                  text-transform: uppercase; position: absolute; top: 313px" Width="115px"></asp:TextBox>
              &nbsp;&nbsp;
              <asp:TextBox ID="DReceiveDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 126; left: 704px;
                  position: absolute; top: 488px" Width="74px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp;&nbsp;
              <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                  Height="24px" Style="z-index: 125; left: 597px; position: absolute; top: 216px"
                  Width="115px"></asp:TextBox>
              <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" Style="z-index: 126; left: 597px; position: absolute;
                  top: 247px" Width="115px"></asp:TextBox>
              <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
                  BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 127; left: 597px;
                  position: absolute; top: 282px" Width="115px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
              <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
                  Height="24px" Style="z-index: 128; left: 217px; position: absolute; top: 187px"
                  Width="190px"></asp:TextBox>
              <asp:TextBox ID="DDuplicateNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" Style="z-index: 129; left: 680px; position: absolute;
                  top: 154px" Width="115px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
              &nbsp; &nbsp;&nbsp;<br />
              &nbsp;
              <br />
              &nbsp;
              <asp:TextBox ID="DNOCCSReason" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="33px" MaxLength="255" Style="z-index: 130; left: 348px; position: absolute;
                  top: 483px" TextMode="MultiLine" Width="240px"></asp:TextBox>
              <asp:TextBox ID="DCustomerNGColor" runat="server" BackColor="Yellow" ForeColor="Blue"
                  Height="30px" MaxLength="255" Style="z-index: 131; left: 217px; position: absolute;
                  top: 566px" TextMode="MultiLine" Width="578px"></asp:TextBox>
              <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" ForeColor="Blue" Height="34px"
                  MaxLength="255" Style="z-index: 132; left: 217px; position: absolute; top: 527px"
                  TextMode="MultiLine" Width="578px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
              &nbsp; &nbsp;
              <asp:TextBox ID="TextBox1" runat="server" BackColor="Transparent" BorderStyle="None"
                  ForeColor="White" Height="24px" Style="z-index: 136; left: 768px; position: absolute;
                  top: 300px" Width="190px"></asp:TextBox>
              <asp:TextBox ID="TextBox2" runat="server" BackColor="Transparent" BorderStyle="None"
                  ForeColor="White" Height="24px" Style="z-index: 137; left: 769px; position: absolute;
                  top: 337px" Width="190px"></asp:TextBox>
              &nbsp; &nbsp; &nbsp;
              <asp:TextBox ID="DDeliveryDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
                  ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 122; left: 597px;
                  text-transform: uppercase; position: absolute; top: 415px" Width="117px"></asp:TextBox>
              <br />
              <br />
              <br />
              &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;<br />
              <asp:TextBox ID="DColorSystem" runat="server" BackColor="Yellow" ForeColor="Blue"
                  MaxLength="3" onkeyup="return ValidateNumber(this,value)" Style="z-index: 139;
                  left: 459px; ime-mode: disabled; text-transform: uppercase; position: absolute;
                  top: 716px" Width="78px"></asp:TextBox>
              &nbsp;
              <asp:TextBox ID="DCheckNo" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="2"
                  onkeyup="return ValidateNumber(this,value)" Style="z-index: 140; left: 463px;
                  ime-mode: disabled; text-transform: uppercase; position: absolute; top: 946px"
                  Width="89px"></asp:TextBox>
              <asp:TextBox ID="DSLDColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
                  onkeyup="return ValidateNumber1(this,value)" Style="z-index: 141; left: 459px;
                  ime-mode: disabled; text-transform: uppercase; position: absolute; top: 978px"
                  Width="64px"></asp:TextBox>
              <asp:TextBox ID="DVFColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
                  onkeyup="return ValidateNumber1(this,value)" Style="z-index: 142; left: 459px;
                  ime-mode: disabled; text-transform: uppercase; position: absolute; top: 1011px"
                  Width="111px"></asp:TextBox>
              &nbsp;<br />
              <br />
              <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="z-index: 144; left: 567px;
                  position: absolute; top: 155px" Text="重覆依賴編號"></asp:Label>
              &nbsp;<br />
              <br />
              &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
              <asp:Label ID="LSheetName" runat="server" Font-Bold="True" Font-Size="Medium" Style="z-index: 148;
                  left: 48px; position: absolute; top: 90px; text-align: center" Text="Label" Width="671px"></asp:Label>
              &nbsp; &nbsp; &nbsp;<br />
              &nbsp; &nbsp; &nbsp; &nbsp;<br />
          </div>
    </form>
</body>
</html>
