<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorCompleteUA_01" aspCompat="True" CodeFile="DTMW_NewColorCompleteUA_01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴書完成表</title>
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
		   
		   
		     function ValidateNumber(e, pnumber) 
{
    if (!/^\d+[.]?[1-9]?$/.test(pnumber)) 
    {
        var newValue = /^\d+/.exec(e.value);
        
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

		</script>
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
           &nbsp; &nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DTMW_NewColorCompleteUA.jpg" style="z-index: 100; left: 10px; position: absolute; top: 6px" />
           <asp:TextBox ID="DFormName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 121; left: 193px;
               position: absolute; top: 691px" Width="190px"></asp:TextBox>
           <asp:DropDownList ID="DCheckType" runat="server" BackColor="GreenYellow" ForeColor="Blue"
               Height="20px" Style="z-index: 101; left: 196px; position: absolute; top: 661px"
               Width="188px" AutoPostBack="True">
           </asp:DropDownList><asp:DropDownList ID="DColor" runat="server" BackColor="GreenYellow" ForeColor="Blue"
               Height="20px" Style="z-index: 102; left: 548px; position: absolute; top: 528px"
               Width="127px">
           </asp:DropDownList>
           <asp:CheckBox ID="DVCACheck" runat="server" Style="z-index: 103; left: 424px; position: absolute;
               top: 133px" Width="78px" Font-Size="Medium" Height="16px" Text="VCA" BorderStyle="None" Font-Bold="True" AutoPostBack="True" Enabled="False" />
           <asp:CheckBox ID="DFactoryCheck" runat="server" Style="z-index: 104; left: 308px; position: absolute;
               top: 131px" Width="99px" Font-Size="Medium" Height="16px" Text="工廠自核" BorderStyle="None" Font-Bold="True" AutoPostBack="True" Enabled="False" />
           <asp:CheckBox ID="DCustomerCheck" runat="server" Style="z-index: 105; left: 183px; position: absolute;
               top: 130px" Width="108px" Font-Size="Medium" Height="16px" Text="客戶自核" BorderStyle="None" Font-Bold="True" AutoPostBack="True" Enabled="False" />
           <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 106; left: 191px; position: absolute;
               top: 211px" Width="220px" ReadOnly="True"></asp:TextBox><asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 107; left: 191px; position: absolute;
               top: 248px" Width="220px" ReadOnly="True"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DCustomerColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 108; left: 192px; position: absolute;text-transform : uppercase;
               top: 315px" Width="220px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 109; left: 191px; position: absolute;
               top: 282px" Width="220px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;
           <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 110; left: 567px; position: absolute;
               top: 182px" Width="115px" ReadOnly="True"></asp:TextBox>
         
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DOverseaYKKCode" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 111; left: 193px; position: absolute; top: 346px;text-transform : uppercase"
               Width="220px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DPANTONECode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 112; left: 296px; position: absolute;text-transform : uppercase;
               top: 390px" Width="385px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DCustomerColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 113; left: 567px; position: absolute;text-transform : uppercase;
               top: 315px" Width="115px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DDReceiveDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 114; left: 567px; position: absolute;
               top: 429px" Width="116px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DDevYear" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 115; left: 190px; position: absolute;
               top: 429px" Width="71px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DColorLight1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 116; left: 191px; position: absolute;
               top: 462px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DVersion" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 117; left: 265px; position: absolute;TEXT-ALIGN: center;
               top: 528px" Width="55px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DVersion1"   runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 118; left: 424px;
               position: absolute; top: 26px; TEXT-ALIGN: center" Width="32px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DTMSDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 119; left: 194px; position: absolute;
               top: 562px" Width="190px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DTMEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 193px; position: absolute;
               top: 594px" Width="190px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DTMDays" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 121; left: 193px; position: absolute;
               top: 626px" Width="190px" ReadOnly="True"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DCheckReason" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="57px" Style="z-index: 122; left: 190px; position: absolute;
               top: 724px" Width="487px" TextMode="MultiLine"></asp:TextBox>
           <asp:TextBox ID="DColorTimes" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 123; left: 549px; position: absolute;
               top: 561px" Width="132px" MaxLength="2" onkeyup="return ValidateNumber(this,value)"></asp:TextBox>
               
                

           <asp:TextBox ID="DSampleTimes" runat="server" BackColor="GreenYellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 124; left: 549px; position: absolute;
               top: 592px" Width="132px" MaxLength="2" onkeyup="return ValidateNumber(this,value)"></asp:TextBox>
           <asp:TextBox ID="DYKKColorType" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 125; left: 549px; position: absolute;
               top: 626px" Width="132px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DYKKColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 126; left: 549px; position: absolute;
               top: 657px" Width="132px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DColorLight2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 127; left: 568px; position: absolute;
               top: 461px" Width="112px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DDevSeason" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 128; left: 305px; position: absolute;
               top: 430px" Width="71px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 129; left: 567px; position: absolute; top: 214px"
               Width="115px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 130; left: 567px; position: absolute;
               top: 249px" Width="115px" ReadOnly="True"></asp:TextBox>
           <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 131; left: 567px;
               position: absolute; top: 280px" Width="115px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
          
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 132; left: 190px; position: absolute; top: 182px"
               Width="220px" ReadOnly="True"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
           &nbsp; 
        <br />
           &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <asp:Label ID="LSheetName" runat="server" Font-Bold="True" Font-Size="Medium" Style="z-index: 160;
               left: 22px; position: absolute;TEXT-ALIGN: center; top: 73px" Text="Label" Width="669px"></asp:Label>
           &nbsp;&nbsp;<br />
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;<br />
        <br />
        <br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <br />
        <br />
        <br />
        <br />
        <br />
           &nbsp; &nbsp;
        &nbsp;&nbsp;&nbsp;<br />
           <asp:Button ID="Button1" runat="server" Style="z-index: 135; left: 626px; position: absolute;
               top: 792px" Text="儲存" Width="53px" />
           <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true"
               ShowSummary="false" Style="z-index: 136; left: 16px; position: absolute; top: 769px" />
           <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"
               Style="z-index: 137; left: 15px; position: absolute; top: 831px"></asp:CustomValidator>
        </div>
    </form>
</body>
</html>
