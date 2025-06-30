<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_COMBISheet01" aspCompat="True" CodeFile="DTMW_COMBISheet01.aspx.vb" %>
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

   function CopyCombiSheet()
{
        
    window.open('DTMW_CombiSheetCopy.aspx?','','status=0,toolbar=0,width=700,height=650,top=10,resizable=yes');
   
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
    if (!/^[A-Za-z]?$/.test(pnumber)) 
    {
        var newValue = /^[A-Za-z]+/.exec(e.value);
        
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
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/DTMW_COMBISheet02.jpg" style="z-index: 100; left: 31px; position: absolute; top: 19px" />
           <asp:CheckBox ID="DReRegister" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 101;
               left: 57px; position: absolute; top: 108px" Text="繼續申請" Width="76px" />
           <asp:TextBox ID="DNO1" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 102; left: 229px; position: absolute; top: 104px"
               Width="142px"></asp:TextBox>
           &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 103; left: 227px; position: absolute;
               top: 174px" Width="168px"></asp:TextBox>
           <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 104; left: 227px; position: absolute;
               top: 206px" Width="142px"></asp:TextBox>
           &nbsp;&nbsp;&nbsp; &nbsp;
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 105; left: 227px; position: absolute;
               top: 238px" Width="143px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           <input id="DYKKColorCode" runat="server" style="z-index: 137; left: 227px; width: 116px;
            color: blue; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 307px; background-color: yellow; border-bottom-style: groove"
            type="text" name="DDeliveryDate" />
            
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
               &nbsp; &nbsp;&nbsp;
               <asp:Button ID="BSAVE" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 133; left: 368px; position: absolute; top: 728px" Text="儲存"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG2" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 134; left: 464px; position: absolute; top: 728px" Text="NG2"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BNG1" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 135; left: 552px; position: absolute; top: 728px" Text="NG1"
                   UseSubmitBehavior="false" Width="80px" />
               <asp:Button ID="BOK" runat="server" Height="23px" OnClientClick="return ConfirmMe(this)"
                   Style="z-index: 136; left: 648px; position: absolute; top: 728px" Text="OK"
                   UseSubmitBehavior="false" Width="80px" />
                   
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <input id="BCustomer" runat="server" style="z-index: 140; left: 375px; width: 24px;
            position: absolute; top: 208px" type="button" value="..." /><input id="BCopySheet" runat="server" style="z-index: 139; left: 375px; width: 24px;
            position: absolute; top: 143px" type="button" value="..." /><input id="BBuyer" runat="server" style="z-index: 141; left: 375px; width: 24px;
            position: absolute; top: 237px" type="button" value="..." />
           &nbsp;&nbsp;
           <asp:TextBox ID="DDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 106; left: 564px; position: absolute;
               top: 142px" Width="150px"></asp:TextBox>
           <asp:TextBox ID="DVFLTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 107; left: 223px; position: absolute;text-transform : uppercase;
               top: 406px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFLChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 108; left: 352px; position: absolute;text-transform : uppercase;
               top: 406px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFRChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 109; left: 477px; position: absolute;text-transform : uppercase;
               top: 407px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFRTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 110; left: 605px; position: absolute;text-transform : uppercase;
               top: 408px" Width="115px" MaxLength="5" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DVFMLTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 111; left: 222px; position: absolute;text-transform : uppercase;
               top: 470px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFMLChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 112; left: 351px; position: absolute;text-transform : uppercase;
               top: 470px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DPFMFLTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 113; left: 268px; position: absolute;text-transform : uppercase;
               top: 571px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DPFMFRTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 114; left: 524px; position: absolute;text-transform : uppercase;
               top: 571px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFMRChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 115; left: 476px; position: absolute;text-transform : uppercase;
               top: 471px" Width="115px" MaxLength="5"></asp:TextBox>
           <asp:TextBox ID="DVFMRTape" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 116; left: 604px; position: absolute;text-transform : uppercase;
               top: 472px" Width="115px" MaxLength="5"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;
           <input id="BYKKColorCode" runat="server" style="z-index: 138; left: 350px; width: 24px;
            position: absolute; top: 306px" type="button" value="..." />
     

           <asp:DropDownList ID="DYKKColorType" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 117; left: 227px; position: absolute;
            top: 274px" Width="142px">
           </asp:DropDownList>
           <asp:DropDownList ID="DCOMBIItem" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 118; left: 567px; position: absolute;
            top: 275px" Width="150px" AutoPostBack="True">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 119; left: 566px; position: absolute; top: 176px"
               Width="150px"></asp:TextBox>
           <asp:TextBox ID="DCustomerCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 120; left: 566px; position: absolute;
               top: 210px" Width="150px"></asp:TextBox>
           <asp:TextBox ID="DBuyerCode" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="24px" Style="z-index: 121; left: 566px;
               position: absolute; top: 241px" Width="150px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 122; left: 228px; position: absolute; top: 142px"
               Width="142px"></asp:TextBox>
           <asp:TextBox ID="DDuplicateNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 123; left: 565px; position: absolute;
               top: 307px" Width="150px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
           &nbsp; 
        <br />
           &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
           <asp:TextBox ID="DDecideDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="58px" Style="z-index: 124; left: 86px; position: absolute; top: 820px"
               TextMode="MultiLine" Width="536px"></asp:TextBox>
        <img id="DDelay" runat="server" src="images/Sheet_Delay.jpg" style="z-index: 136; left: 37px;
            width: 592px; position: absolute; top: 890px; height: 112px" visible="false" />
           <img id="DDescSheet" runat="server" src="images/MapSheet_004.jpg" style="z-index: 99;
               left: 37px; position: absolute; top: 812px" />
        <asp:Label ID="DHistoryLabel" runat="server" Font-Names="新細明體" Font-Size="11pt" ForeColor="Blue"
            Height="16px" Style="z-index: 125; left: 45px; position: absolute; top: 1017px"
            Width="64px">核定履歷</asp:Label>
        <asp:GridView ID="GridView2" runat="server" AutoGenerateColumns="False" BackColor="White"
            BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="100px"
            Style="z-index: 126; left: 43px; position: absolute; top: 1040px" Width="780px">
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
               ForeColor="White" Height="24px" Style="z-index: 127; left: 768px; position: absolute;
               top: 300px" Width="190px"></asp:TextBox>
           <asp:TextBox ID="D2" runat="server" BackColor="Transparent" BorderStyle="None" ForeColor="White"
               Height="24px" Style="z-index: 128; left: 769px; position: absolute; top: 337px"
               Width="190px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp;<br />
        <br />
        <br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false" style="z-index: 129; left: 53px; position: absolute; top: 640px" />
           &nbsp;&nbsp;<br />
                  
               
               
        <br />
         <asp:CustomValidator  ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator" style="z-index: 130; left: 211px; position: absolute; top: 644px"></asp:CustomValidator>
        <br />
        <br />
        <br />
        <asp:TextBox ID="DReason" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
            Style="z-index: 131; left: 253px; position: absolute; top: 903px" Visible="False"
            Width="360px"></asp:TextBox>
        <asp:DropDownList ID="DReasonCode" runat="server" Height="20px" Style="z-index: 132;
            left: 189px; position: absolute; top: 903px" Visible="False" Width="64px">
            <asp:ListItem>01</asp:ListItem>
            <asp:ListItem>02</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="DReasonDesc" runat="server" BackColor="Yellow" ForeColor="Blue"
            Height="56px" Style="z-index: 133; left: 187px; position: absolute; top: 938px"
            Visible="False" Width="432px"></asp:TextBox>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
           <asp:Label ID="DITEMNAME1" runat="server" Font-Bold="True" Font-Italic="False" Style="z-index: 134;
               left: 66px; position: absolute; top: 395px" Text="塑鋼 VS (一般)" Width="148px"></asp:Label>
           &nbsp;
           &nbsp; &nbsp;
      
        </div>
    </form>
</body>
</html>
