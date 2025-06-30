<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorSheet01_UAKIPLINGDT" aspCompat="True" CodeFile="DTMW_NewColorSheet01_UAKIPLINGDT.aspx.vb" %>
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
  	 <div>
  	 
  
           &nbsp; &nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/NewColorUAKIPLING03.jpg" style="z-index: 100; left: 26px; position: absolute; top: 16px" />
           <asp:TextBox ID="DPFBWire" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="3"
               onkeyup="return ValidateNumber(this,value)" Style="z-index: 101; left: 199px;
               ime-mode: disabled; text-transform: uppercase; position: absolute; top: 272px"
               Width="160px"></asp:TextBox>
           <asp:DropDownList ID="DDTSheet" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 102; left: 516px; position: absolute;
            top: 237px" Width="158px">
           </asp:DropDownList>
           <asp:CheckBox ID="DReRegister" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 103;
               left: 50px; position: absolute; top: 105px" Text="繼續申請" Width="76px" />
           <asp:DropDownList ID="DNewOldColor" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 104; left: 199px; position: absolute;
               top: 370px" Width="160px">
           </asp:DropDownList>
           &nbsp; &nbsp;
           <input id="BCopySheet" runat="server" style="z-index: 131; left: 690px; width: 24px;
               position: absolute; top: 171px" type="button" value="..." visible="true" />
           &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <input id="DYKKColorCode" runat="server" style="z-index: 129; left: 516px; width: 159px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 337px; background-color: yellow; border-bottom-style: groove"
            type="text" name="DDeliveryDate" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 135;
            left: 393px; width: 80px; position: absolute; top: 1121px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 134;
            left: 559px; width: 80px; position: absolute; top: 1121px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 133;
            left: 477px; width: 80px; position: absolute; top: 1121px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 132;
            left: 642px; width: 80px; position: absolute; top: 1121px; height: 28px" type="button"
            value="OK" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;
           <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 105; left: 200px; position: absolute;
               top: 173px" Width="160px"></asp:TextBox>
           &nbsp; &nbsp;
           <input id="BYKKColorCode" runat="server" style="z-index: 130; left: 684px; width: 24px;
            position: absolute; top: 335px" type="button" value="..." />
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
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:DropDownList ID="DYKKColorType" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 106; left: 199px; position: absolute;
            top: 335px" Width="160px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DNO1" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 107; left: 516px; position: absolute; top: 174px"
               Width="114px"></asp:TextBox>
           <asp:HyperLink ID="HyperLink2" runat="server" Style="z-index: 123; left: 635px; position: absolute;
               top: 174px" Target="_blank" Visible="False">連結</asp:HyperLink>
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 108; left: 200px; position: absolute;
               top: 139px" Width="160px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DOFormSno" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 107; left: 741px; position: absolute;
               top: 170px" Width="75px" Visible="False"></asp:TextBox>
           <br />
           &nbsp; 
        <br />
           &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 109; left: 99px; position: absolute; top: 1285px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 128; left: 51px;
            width: 592px; position: absolute; top: 1347px; height: 112px" visible="false" />
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 99;
               left: 50px; position: absolute; top: 1277px" />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 110; left: 55px; position: absolute; top: 1468px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 111; left: 53px; position: absolute; top: 1491px" Width="780px">
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
           &nbsp;
           <asp:TextBox ID="D1" runat="server" BackColor="Transparent" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 112; left: 768px; position: absolute;
               top: 300px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 113; left: 769px; position: absolute; top: 337px"
               Width="190px"></asp:TextBox>
           &nbsp;&nbsp;<br />
        <br />
        <br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" style="z-index: 114; left: 42px; position: absolute; top: 1060px" />
           &nbsp;<br />
      
      <asp:TextBox ID="DColorSystem" runat="server" style="ime-mode:disabled; z-index: 115; left: 199px; text-transform : uppercase;position: absolute;top: 240px;" onkeyup="return ValidateNumber(this,value)" BackColor="Yellow" ForeColor="Blue" Width="160px" MaxLength="3" ></asp:TextBox>
           &nbsp; &nbsp;
           <asp:TextBox ID="DSLDColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
               onkeyup="return ValidateNumber1(this,value)" Style="z-index: 116; left: 200px;text-transform : uppercase;
               ime-mode: disabled; position: absolute; top: 304px" Width="160px"></asp:TextBox>
           <asp:TextBox ID="DVFColor" runat="server" BackColor="Yellow" ForeColor="Blue" MaxLength="5"
               onkeyup="return ValidateNumber1(this,value)" Style="z-index: 117; left: 516px;text-transform : uppercase;
               ime-mode: disabled; position: absolute; top: 303px" Width="160px"></asp:TextBox>
                  
               
               
        <br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator" style="z-index: 118; left: 41px; position: absolute; top: 1132px"></asp:CustomValidator>
        <br />
        <br />
        <br />
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 119; left: 241px; position: absolute; top: 1357px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 120;
            left: 177px; position: absolute; top: 1357px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 121; left: 175px; position: absolute; top: 1392px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           <asp:Label ID="MNOSts" runat="server" Font-Size="Small" ForeColor="Red" Height="22px"
               Style="z-index: 122; left: 559px; position: absolute; top: 205px" Text="Label"
               Visible="False" Width="429px"></asp:Label>
           <br />
           &nbsp; &nbsp;
           <asp:HyperLink ID="HyperLink1" runat="server" Style="z-index: 123; left: 518px; position: absolute;
               top: 207px" Visible="False" Target="_blank">履歷</asp:HyperLink>
           <asp:TextBox ID="DPFOpenParts" runat="server" BackColor="Yellow" ForeColor="Blue"
               MaxLength="3" onkeyup="return ValidateNumber(this,value)" Style="z-index: 124;
               left: 516px; ime-mode: disabled; text-transform: uppercase; position: absolute;
               top: 271px" Width="160px"></asp:TextBox>
           <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 125; left: 199px; position: absolute;
               top: 204px" Width="160px"></asp:TextBox>
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 126; left: 516px; position: absolute; top: 139px"
               Width="160px"></asp:TextBox>
        </div>
    </form>
</body>
</html>
