<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorSheet02_VB" aspCompat="True" CodeFile="DTMW_NewColorSheet02_VB.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴書拉鏈ZIPPER CHAIN (03 CF P12)</title>
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

   function CopyNewColor(wformno)
{
        
    window.open('DTMW_NewColorCopy.aspx?formno='+wformno,'','status=0,toolbar=0,width=700,height=650,top=10,resizable=yes');
   
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
			
    function ValidateNumber(e, pnumber) 
{
    if (!/^[-A-Za-z-]?$/.test(pnumber)) 
    {
        var newValue = /^[-A-Za-z-]+/.exec(e.value);
        
        if (newValue != null) 
        {  
            e.value = newValue;  
        }
        else 
        {  
            e.value = ""; 
        }
    }
    return false;
}

    function ValidateNumber1(e, pnumber) 
{
    if (!/^[a-zA-Z0-9]?$/.test(pnumber)) 
    {
        var newValue = /^[a-zA-Z0-9]+/.exec(e.value);
        
        if (newValue != null) 
        {  
            e.value =  newValue;          }
        else 
        {  
            e.value = ""; 
        }
    }
    return false;
}
		   
		</script>
		
 
  
		
</head>
<body>

  	<form id="Form1" runat="server">
          &nbsp; &nbsp; &nbsp;&nbsp;
          <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DTMW_NewColorSheetUA_05.jpg"
              Style="z-index: 100; left: 28px; position: absolute; top: 18px" />
          <asp:DropDownList ID="DAgain" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 119; left: 536px; position: absolute; top: 928px"
              Width="60px">
          </asp:DropDownList>
          &nbsp; &nbsp;
          &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
          <asp:TextBox ID="DVFColor1" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
              onkeyup="return ValidateNumber1(this,value)" Style="z-index: 138; left: 473px;
              ime-mode: disabled; text-transform: uppercase; position: absolute; top: 962px"
              Width="111px" Height="21px"></asp:TextBox>
          <input id="DYKKColorCodeVF" runat="server" name="DDeliveryDate" style="z-index: 146;
              left: 475px; width: 85px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 814px; background-color: yellow;
              border-bottom-style: groove; height: 21px;" type="text" />
          <input id="DYKKColorCode" runat="server" name="DDeliveryDate" style="z-index: 146;
              left: 475px; width: 85px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 786px; background-color: yellow;
              border-bottom-style: groove; height: 21px;" type="text" />
          <input id="DYKKColorCodeSLD" runat="server" name="DDeliveryDate" style="z-index: 146;
              left: 475px; width: 85px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 844px; background-color: yellow;
              border-bottom-style: groove; height: 21px;" type="text" />
          <asp:CheckBox ID="DReRegister" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 101;
              left: 51px; position: absolute; top: 144px" Text="繼續申請" Width="76px" />
          &nbsp;
          <asp:HyperLink ID="LComplete" runat="server" NavigateUrl="BoardEdit.aspx" Style="z-index: 103;
              left: 56px; position: absolute; top: 119px" Target="_blank" Visible="False">有新色依賴完成表</asp:HyperLink>
          <asp:TextBox ID="DNO1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
              Height="24px" Style="z-index: 104; left: 52px; position: absolute; top: 47px"
              Width="142px"></asp:TextBox>
          &nbsp; &nbsp;&nbsp;
          <asp:CheckBox ID="DVCACheck" runat="server" AutoPostBack="True" BorderStyle="None"
              Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 105; left: 623px;
              position: absolute; top: 120px" Text="VCA" Width="78px" />
          <asp:CheckBox ID="DFactoryCheck" runat="server" AutoPostBack="True" BorderStyle="None"
              Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 106; left: 500px;
              position: absolute; top: 119px" Text="工廠自核" Width="99px" />
          <asp:CheckBox ID="DCustomerCheck" runat="server" AutoPostBack="True" BorderStyle="None"
              Font-Bold="True" Font-Size="Medium" Height="16px" Style="z-index: 107; left: 375px;
              position: absolute; top: 119px" Text="客戶自核" Width="108px" />
          <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" Style="z-index: 108; left: 219px; position: absolute;
              top: 208px" Width="219px"></asp:TextBox>
          <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" Style="z-index: 109; left: 219px; position: absolute;
              top: 238px" Width="190px"></asp:TextBox>
          &nbsp;&nbsp;
          <asp:TextBox ID="DCustomerColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 110; left: 219px;
              text-transform: uppercase; position: absolute; top: 295px" Width="219px"></asp:TextBox>
          <asp:CheckBox ID="DNOCCS" runat="server" AutoPostBack="True" Font-Size="Small" Height="1px"
              Style="z-index: 111; left: 236px; position: absolute; top: 444px" Width="1px" />
          <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
              Height="24px" Style="z-index: 112; left: 219px; position: absolute; top: 263px"
              Width="188px"></asp:TextBox>
          &nbsp; &nbsp;&nbsp;
          <asp:TextBox ID="DVersion" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="23px" Style="z-index: 113; left: 512px; position: absolute;
              top: 874px; text-align: center" Width="37px"></asp:TextBox>
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          <input id="DDeliveryDate" runat="server" name="DDeliveryDate" style="z-index: 152;
              left: 598px; width: 91px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 380px; background-color: yellow;
              border-bottom-style: groove" type="text" />
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp;&nbsp;
          <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
              Height="24px" Style="z-index: 114; left: 598px; position: absolute; top: 176px"
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
          &nbsp;&nbsp;
          <asp:DropDownList ID="DDevSeason" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 115; left: 331px; position: absolute; top: 380px"
              Width="94px">
          </asp:DropDownList>
          <asp:DropDownList ID="DColorLight2" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 116; left: 479px; position: absolute; top: 409px"
              Width="58px">
          </asp:DropDownList>
          <asp:DropDownList ID="DColorLight1" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 117; left: 219px; position: absolute; top: 409px"
              Width="58px">
          </asp:DropDownList>
          <asp:DropDownList ID="DCustomerSample" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="25px" Style="z-index: 118; left: 72px; position: absolute; top: 670px"
              Width="221px">
          </asp:DropDownList>
          &nbsp;
          <asp:DropDownList ID="DDevYear" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 120; left: 219px; position: absolute; top: 381px"
              Width="76px">
          </asp:DropDownList>
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp;
          <asp:TextBox ID="DReColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 121; left: 219px;
              text-transform: uppercase; position: absolute; top: 352px" Width="219px"></asp:TextBox>
          <asp:TextBox ID="DOverseaYKKCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 121; left: 219px;
              text-transform: uppercase; position: absolute; top: 322px" Width="219px"></asp:TextBox>
          <asp:TextBox ID="DPANTONECode" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 122; left: 598px;
              text-transform: uppercase; position: absolute; top: 337px" Width="117px"></asp:TextBox>
          &nbsp;
          <asp:TextBox ID="DCustomerColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" MaxLength="25" Style="z-index: 123; left: 598px;
              text-transform: uppercase; position: absolute; top: 295px" Width="115px"></asp:TextBox>
          <asp:TextBox ID="DReceiveDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" Style="z-index: 124; left: 645px; position: absolute;
              top: 437px" Width="70px"></asp:TextBox>
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
          <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
              Height="24px" Style="z-index: 125; left: 598px; position: absolute; top: 206px"
              Width="115px"></asp:TextBox>
          <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" Style="z-index: 126; left: 598px; position: absolute;
              top: 238px" Width="115px"></asp:TextBox>
          <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
              BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 127; left: 598px;
              position: absolute; top: 265px" Width="115px"></asp:TextBox>
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
          <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
              Height="24px" Style="z-index: 128; left: 219px; position: absolute; top: 175px"
              Width="190px"></asp:TextBox>
          <asp:TextBox ID="DDuplicateNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
              ForeColor="Blue" Height="24px" Style="z-index: 129; left: 598px; position: absolute;
              top: 144px" Width="115px"></asp:TextBox>
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
              Height="27px" MaxLength="255" Style="z-index: 130; left: 346px; position: absolute;
              top: 435px" TextMode="MultiLine" Width="200px"></asp:TextBox>
          <asp:TextBox ID="DCustomerNGColor" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" MaxLength="255" Style="z-index: 131; left: 219px; position: absolute;
              top: 498px" TextMode="MultiLine" Width="495px"></asp:TextBox>
          <asp:TextBox ID="DRemark" runat="server" BackColor="Yellow" ForeColor="Blue" Height="22px"
              MaxLength="255" Style="z-index: 132; left: 221px; position: absolute; top: 466px"
              TextMode="MultiLine" Width="496px"></asp:TextBox>
          &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
          &nbsp;
          <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
              Height="16px" Style="z-index: 134; left: 52px; position: absolute; top: 1059px"
              Width="64px">核定履歷</asp:Label>
          &nbsp; &nbsp;
          <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
              Height="24px" Style="z-index: 136; left: 768px; position: absolute; top: 300px"
              Width="190px"></asp:TextBox>
          <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
              Height="24px" Style="z-index: 137; left: 769px; position: absolute; top: 337px"
              Width="190px"></asp:TextBox>
          &nbsp; &nbsp;<br />
          <br />
          <br />
          &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;<br />
          &nbsp;
          <asp:TextBox ID="DCheckNo" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="2"
              onkeyup="return ValidateNumber(this,value)" Style="z-index: 140; left: 475px;
              ime-mode: disabled; text-transform: uppercase; position: absolute; top: 902px"
              Width="89px" Height="21px"></asp:TextBox>
          <asp:TextBox ID="DSLDColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
              onkeyup="return ValidateNumber1(this,value)" Style="z-index: 141; left: 474px;
              ime-mode: disabled; text-transform: uppercase; position: absolute; top: 932px"
              Width="60px" Height="21px"></asp:TextBox>
          <asp:TextBox ID="DVFColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
              onkeyup="return ValidateNumber1(this,value)" Style="z-index: 142; left: 473px;
              ime-mode: disabled; text-transform: uppercase; position: absolute; top: 989px"
              Width="111px" Height="21px"></asp:TextBox>
          &nbsp;<br />
          <br />
          <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="z-index: 144; left: 482px;
              position: absolute; top: 145px" Text="重覆依賴編號"></asp:Label>
          &nbsp;<br />
          <br />
          &nbsp; &nbsp; &nbsp; &nbsp;<br />
          <asp:Label ID="LSheetName" runat="server" Font-Bold="True" Font-Size="Medium" Style="z-index: 148;
              left: 48px; position: absolute; top: 80px; text-align: center" Text="Label" Width="671px"></asp:Label>
          <asp:HyperLink ID="LDTSheet" runat="server" NavigateUrl="BoardEdit.aspx" Style="z-index: 140;
              left: 210px; position: absolute; top: 122px" Target="_blank" Visible="False">有追加核可單</asp:HyperLink>
          &nbsp; &nbsp; &nbsp;&nbsp;
          <asp:DropDownList ID="DColorLight3" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="20px" Style="z-index: 116; left: 708px; position: absolute; top: 410px"
              Width="58px">
          </asp:DropDownList>
          <asp:DropDownList ID="DNewOldColor" runat="server" AutoPostBack="True" BackColor="Yellow"
              ForeColor="Blue" Height="16px" Style="z-index: 115; left: 475px; position: absolute;
              top: 755px" Width="117px">
          </asp:DropDownList>
          <input id="DPFBWire" runat="server" name="DDeliveryDate" style="z-index: 148; left: 475px;
              width: 85px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 671px; height: 18px; background-color: yellow;
              border-bottom-style: groove" type="text" />
          <input id="DPFOpenParts" runat="server" name="DDeliveryDate" style="z-index: 149;
              left: 475px; width: 85px; color: blue; border-top-style: groove; border-right-style: groove;
              border-left-style: groove; position: absolute; top: 698px; height: 18px; background-color: yellow;
              border-bottom-style: groove" type="text" />
          <asp:DropDownList ID="DYKKColorType" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="16px" Style="z-index: 115; left: 475px; position: absolute; top: 723px"
              Width="117px">
          </asp:DropDownList>
          <asp:TextBox ID="DColorSystem" runat="server" BackColor="Yellow" ForeColor="Blue"
              Height="18px" MaxLength="3" onkeyup="return ValidateNumber(this,value)" Style="z-index: 135;
              left: 475px; ime-mode: disabled; text-transform: uppercase; position: absolute;
              top: 643px" Width="78px"></asp:TextBox>
          <br />
          .<br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
              BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
              Style="z-index: 131; left: 50px; position: absolute; top: 1090px" Width="780px">
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
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
    </form>
</body>
</html>
