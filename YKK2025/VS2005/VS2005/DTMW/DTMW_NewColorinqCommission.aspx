<%@ Page Language="vb" AutoEventWireup="false" Inherits="DTMW_NewColorinqCommission" aspCompat="True"  EnableEventValidation = "false"  CodeFile="DTMW_NewColorinqCommission.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>新色依賴書調閱資料</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
   function GetStsPicker()
{
        
    window.open('StsPicker.aspx?','','status=0,toolbar=0,width=250,height=190,top=10,resizable=yes');
   
}	 

			
   function GetFormPicker()
{
        
    window.open('FormPicker.aspx?','','status=0,toolbar=0,width=280,height=220,top=10,resizable=yes');
   
}

   function GetDivisionPicker()
{
        
    window.open('DivisionPicker.aspx?','','status=0,toolbar=0,width=250,height=190,top=10,resizable=yes');
   
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
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 792px; POSITION: absolute; TOP: 200px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px" Visible="False"></asp:imagebutton>
           <asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
               BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px"
               CellPadding="4" Font-Size="9pt" Height="100px" PageSize="15" Style="z-index: 101;
               left: 8px; position: absolute; top: 232px" Width="994px">
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <SelectedItemStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
               <ItemStyle BackColor="White" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
               <Columns>
                   <asp:HyperLinkColumn DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="Field1" Target="_blank">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:HyperLinkColumn>
                   <asp:BoundColumn DataField="Field2" HeaderText="狀態" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field3" HeaderText="依賴日" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="CompleteDate" HeaderText="完成日"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field4" HeaderText="依賴表單" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field5" HeaderText="依賴部門" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field6" HeaderText="依賴者" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field7" HeaderText="客戶" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field8" HeaderText="Buyer" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="CustomerColor" HeaderText="客戶色名"></asp:BoundColumn>
                   <asp:BoundColumn DataField="CustomerColorCode" HeaderText="客戶色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="OverSeaYkkcode" HeaderText="海外YKK色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="PantoneCode" HeaderText="PANTONECode"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field9" HeaderText="YKK色別"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field10" HeaderText="YKK色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field11" HeaderText="兼用色拉頭"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field12" HeaderText="兼用色VF上/下止"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field13" HeaderText="兼用色VF上/下止"></asp:BoundColumn>
                   <asp:BoundColumn DataField="NewOldColor" HeaderText="新舊色"></asp:BoundColumn>
                   <asp:BoundColumn HeaderText="對應YKK色號" DataField="ReColorCode"></asp:BoundColumn>
                   <asp:HyperLinkColumn DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="WorkFlow" HeaderText="履歷" Target="_blank">
                       <HeaderStyle Width="30px" />
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:HyperLinkColumn>
               </Columns>
               <PagerStyle BackColor="#FFFFCC" ForeColor="#993300" HorizontalAlign="Center" Mode="NumericPages"
                   NextPageText="下一頁" PrevPageText="上一頁" />
           </asp:DataGrid>
           <asp:textbox id="DSts1" style="Z-INDEX: 102; LEFT: 99px; POSITION: absolute; TOP: 10px" runat="server"
				Height="22px" Width="254px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" AutoCompleteType="Disabled"></asp:textbox>
           <asp:TextBox ID="DSts2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 103; left: 944px;
               position: absolute; top: 12px" Width="190px" BorderColor="White"></asp:TextBox>
           <asp:TextBox ID="DFormNo2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 104; left: 946px;
               position: absolute; top: 36px" Width="424px"></asp:TextBox>
          
           <asp:TextBox ID="DFEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 105; left: 669px;
               position: absolute; top: 116px" Width="88px"></asp:TextBox>
           <asp:Button ID="BFSDate" runat="server" Height="26px" Style="z-index: 106; left: 615px;
               position: absolute; top: 112px" Text="....." Width="28px" /><asp:Button ID="BFEDate" runat="server" Height="26px" Style="z-index: 107; left: 760px;
               position: absolute; top: 110px" Text="....." Width="28px" />
           <asp:TextBox ID="DFSDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 108; left: 521px;
               position: absolute; top: 116px" Width="93px"></asp:TextBox>
           &nbsp;
           <asp:Label ID="Label14" runat="server" Height="1px" Style="z-index: 109; left: 651px;
               position: absolute; top: 115px" Text="~"></asp:Label>
           <asp:textbox id="DAEDate" style="Z-INDEX: 110; LEFT: 258px; POSITION: absolute; TOP: 114px" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
			<asp:button id="BASDate" style="Z-INDEX: 111; LEFT: 198px; POSITION: absolute; TOP: 111px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button>
				<asp:button id="BAEDate" style="Z-INDEX: 112; LEFT: 352px; POSITION: absolute; TOP: 111px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button>	
				<asp:textbox id="DASDate" style="Z-INDEX: 113; LEFT: 99px; POSITION: absolute; TOP: 113px" runat="server"
				Height="20px" Width="93px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 114; left: 460px; position: absolute;
               top: 9px" Width="138px"></asp:TextBox>
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 115; left: 666px; position: absolute;
               top: 11px" Width="145px"></asp:TextBox>
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 116; left: 288px;
               position: absolute; top: 63px" Width="99px"></asp:TextBox>
           <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 117; left: 99px; position: absolute;
               top: 62px" Width="96px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DYKKColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 118; left: 287px;
               position: absolute; top: 89px" Width="99px"></asp:TextBox>
           <asp:TextBox ID="DSLDColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 119; left: 521px;
               position: absolute; top: 90px" Width="74px"></asp:TextBox>
           <asp:TextBox ID="DVFColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 120; left: 738px;
               position: absolute; top: 90px" Width="77px"></asp:TextBox>
           &nbsp;
			<asp:button id="Go" style="Z-INDEX: 121; LEFT: 744px; POSITION: absolute; TOP: 200px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp;
            <asp:Button ID="BSts" runat="server" Height="25px" Style="z-index: 122;
                left: 356px; position: absolute; top: 6px" Text="....." Width="28px" />
           <asp:TextBox ID="DDivision1" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="22px" Style="z-index: 123; left: 460px;
               position: absolute; top: 61px" Width="104px"></asp:TextBox>
           <asp:TextBox ID="DDivision2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 124; left: 947px;
               position: absolute; top: 63px" Width="186px"></asp:TextBox>
           <asp:Button ID="BDivision" runat="server" Height="25px" Style="z-index: 125;
                left: 570px; position: absolute; top: 60px" Text="....." Width="28px" /><asp:Button ID="BFormNo" runat="server" Height="25px" Style="z-index: 126;
                left: 570px; position: absolute; top: 34px" Text="....." Width="28px" />
           <asp:TextBox ID="DFormNo1" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="22px" Style="z-index: 127; left: 99px;
               position: absolute; top: 36px" Width="467px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Label ID="Label12" runat="server" Height="1px" Style="z-index: 128; left: 235px;
               position: absolute; top: 117px" Text="~"></asp:Label>
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;
           <asp:DropDownList ID="DYKKColortype" runat="server" BackColor="Yellow" Style="z-index: 129;
               left: 99px; position: absolute; top: 86px" Width="100px">
           </asp:DropDownList>
           &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/NewColorInq_05.jpg" Style="z-index: 99;
               left: 0px; position: absolute; top: 0px" />
           <asp:TextBox ID="DAFEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 110; left: 256px; position: absolute;
               top: 168px" Width="88px"></asp:TextBox>
           <asp:Button ID="BAFSDate" runat="server" Height="26px" Style="z-index: 111; left: 200px;
               position: absolute; top: 160px" Text="....." Width="28px" />
           <asp:Button ID="BAFEDate" runat="server" Height="26px" Style="z-index: 112; left: 352px;
               position: absolute; top: 160px" Text="....." Width="28px" />
           <asp:TextBox ID="DAFSDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 96px; position: absolute;
               top: 168px" Width="93px"></asp:TextBox>
           <asp:Label ID="Label2" runat="server" Height="1px" Style="z-index: 128; left: 232px;
               position: absolute; top: 168px" Text="~"></asp:Label>
           <asp:TextBox ID="DReColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 131; left: 736px; position: absolute;
               top: 64px" Width="77px"></asp:TextBox>
           <asp:TextBox ID="DCustomerColor" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 132; left: 98px; position: absolute;
               top: 141px" Width="93px"></asp:TextBox>
           <asp:TextBox ID="DCustomerColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 133; left: 289px; position: absolute;
               top: 142px" Width="93px"></asp:TextBox>
           <asp:TextBox ID="DOverSeaYkkcode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 134; left: 522px; position: absolute;
               top: 141px" Width="73px"></asp:TextBox>
           <asp:TextBox ID="DPANTONECode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 135; left: 739px; position: absolute;
               top: 141px" Width="75px"></asp:TextBox>
           <asp:DropDownList ID="DNewOldColor" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 136; left: 667px; position: absolute;
               top: 35px" Width="141px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem>新色</asp:ListItem>
               <asp:ListItem>舊色</asp:ListItem>
           </asp:DropDownList>
           <asp:Button ID="BClear" runat="server" BackColor="Lime" ForeColor="Blue" Style="z-index: 137;
               left: 8px; position: absolute; top: 192px" Text="清除條件" />
           <asp:DataGrid ID="Datagrid2" runat="server" AllowPaging="True" AutoGenerateColumns="False"
               BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px"
               CellPadding="4" Font-Size="9pt" Height="100px" PageSize="15" Style="z-index: 138;
               left: 8px; position: absolute; top: 384px" Width="994px">
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <SelectedItemStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
               <PagerStyle BackColor="#FFFFCC" ForeColor="#993300" HorizontalAlign="Center" Mode="NumericPages"
                   NextPageText="下一頁" PrevPageText="上一頁" />
               <ItemStyle BackColor="White" ForeColor="#330099" />
               <Columns>
                   <asp:HyperLinkColumn DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="Field1" Target="_blank">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:HyperLinkColumn>
                   <asp:BoundColumn DataField="Field2" HeaderText="狀態" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field3" HeaderText="依賴日" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="CompleteDate" HeaderText="完成日"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field4" HeaderText="依賴表單" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field5" HeaderText="依賴部門" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field6" HeaderText="依賴者" ReadOnly="True">
                       <HeaderStyle HorizontalAlign="Center" />
                       <ItemStyle HorizontalAlign="Left" />
                   </asp:BoundColumn>
                   <asp:BoundColumn DataField="Field7" HeaderText="YKK色別"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field8" HeaderText="YKK色號"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field9" HeaderText="兼用色拉頭"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field10" HeaderText="兼用色VF上/下止"></asp:BoundColumn>
                   <asp:BoundColumn DataField="Field11" HeaderText="新舊色"></asp:BoundColumn>
                   <asp:BoundColumn DataField="DTSHeet" HeaderText="核可單種類"></asp:BoundColumn>
                   <asp:HyperLinkColumn DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}"
                       DataTextField="WorkFlow" HeaderText="履歷" Target="_blank">
                       <HeaderStyle Width="30px" />
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:HyperLinkColumn>
               </Columns>
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:DataGrid><asp:DropDownList ID="DTABLE" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 139; left: 104px; position: absolute;
               top: 200px" Width="141px">
               <asp:ListItem>未封存</asp:ListItem>
               <asp:ListItem>已封存</asp:ListItem>
           </asp:DropDownList>
           <asp:Label ID="Label1" runat="server" Height="16px" Style="z-index: 140; left: 256px;
               position: absolute; top: 200px" Text="未封存-三年內      已封存-超過三年" ForeColor="Red"></asp:Label>
  	 
      </div>
    </form>
</body>
</html>
