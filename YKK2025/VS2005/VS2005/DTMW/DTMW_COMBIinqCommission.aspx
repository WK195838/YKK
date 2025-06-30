<%@ Page Language="vb" AutoEventWireup="false" Inherits=" DTMW_COMBIinqCommission" aspCompat="True" EnableEventValidation = "false"  CodeFile="DTMW_COMBIinqCommission.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>COMBI依賴書調閱資料</title>
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

			
   function GetCombiItemPicker()
{
        
    window.open('CombiItemPicker.aspx?','','status=0,toolbar=0,width=250,height=190,top=10,resizable=yes');
   
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
        	<FONT face="新細明體"></FONT>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 792px; POSITION: absolute; TOP: 152px" runat="server"
				ImageUrl="Images\msexcel.gif" Height="21px" Width="21px" Visible="False"></asp:imagebutton>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 24px; POSITION: absolute; TOP: 200px" runat="server"
				Height="100px" Width="941px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" PageSize="15">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<HeaderStyle HorizontalAlign="Center" ForeColor="#FFFFCC" VerticalAlign="Middle" BackColor="#CC6600"></HeaderStyle>
				<Columns>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="ViewURL" DataNavigateUrlFormatString="{0}"
						DataTextField="Field1">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="Field2" ReadOnly="True" HeaderText="狀態">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field3" ReadOnly="True" HeaderText="依賴日">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
                    <asp:BoundColumn DataField="CompletedDate" HeaderText="完成日"></asp:BoundColumn>
					<asp:BoundColumn DataField="Field4" ReadOnly="True" HeaderText="依賴表單">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field5" ReadOnly="True" HeaderText="依賴部門">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field6" ReadOnly="True" HeaderText="依賴者">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field7" ReadOnly="True" HeaderText="客戶">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="Field8" ReadOnly="True" HeaderText="Buyer">
						<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
						<ItemStyle HorizontalAlign="Left"></ItemStyle>
					</asp:BoundColumn>
                    <asp:BoundColumn DataField="Field9" HeaderText="YKK色別"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field10" HeaderText="YKK色號"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field11" HeaderText="項目"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field12" HeaderText="左布帶"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field13" HeaderText="左齒"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field14" HeaderText="右齒"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field15" HeaderText="右布帶"></asp:BoundColumn>
					<asp:HyperLinkColumn Target="_blank" DataNavigateUrlField="OPURL" DataNavigateUrlFormatString="{0}" DataTextField="WorkFlow"
						HeaderText="履歷">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid>
           <asp:TextBox ID="DSts2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 102; left: 944px;
               position: absolute; top: 12px" Width="190px" BorderColor="White"></asp:TextBox>
           <asp:TextBox ID="DCombiItem2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 103; left: 946px;
               position: absolute; top: 36px" Width="424px"></asp:TextBox>
          
           <asp:TextBox ID="DFEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 104; left: 698px;
               position: absolute; top: 128px" Width="88px"></asp:TextBox>
           <asp:Button ID="BFSDate" runat="server" Height="26px" Style="z-index: 105; left: 644px;
               position: absolute; top: 124px" Text="....." Width="28px" /><asp:Button ID="BFEDate" runat="server" Height="26px" Style="z-index: 106; left: 791px;
               position: absolute; top: 124px" Text="....." Width="28px" />
           <asp:TextBox ID="DFSDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 107; left: 548px;
               position: absolute; top: 128px" Width="93px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:textbox id="DAEDate" style="Z-INDEX: 108; LEFT: 260px; POSITION: absolute; TOP: 129px" runat="server"
				Height="20px" Width="88px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
			<asp:button id="BASDate" style="Z-INDEX: 109; LEFT: 204px; POSITION: absolute; TOP: 124px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button>
				<asp:button id="BAEDate" style="Z-INDEX: 110; LEFT: 351px; POSITION: absolute; TOP: 125px" runat="server"
				Height="26px" Width="28px" Text="....."></asp:button>	
				<asp:textbox id="DASDate" style="Z-INDEX: 111; LEFT: 105px; POSITION: absolute; TOP: 129px" runat="server"
				Height="20px" Width="93px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove"></asp:textbox>
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 112; left: 315px; position: absolute;
               top: 24px" Width="94px"></asp:TextBox>
           <asp:TextBox ID="DName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 113; left: 711px; position: absolute;
               top: 52px" Width="102px"></asp:TextBox>
           <asp:TextBox ID="DBuyer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 114; left: 315px;
               position: absolute; top: 75px" Width="95px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DYKKColorCode" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 115; left: 711px;
               position: absolute; top: 77px" Width="102px"></asp:TextBox>
           &nbsp; &nbsp;
			<asp:button id="Go" style="Z-INDEX: 116; LEFT: 744px; POSITION: absolute; TOP: 152px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp;
            <asp:Button ID="BSts" runat="server" Height="25px" Style="z-index: 117;
                left: 212px; position: absolute; top: 22px" Text="....." Width="28px" />
           <asp:TextBox ID="DDivision1" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="22px" Style="z-index: 118; left: 106px;
               position: absolute; top: 50px" Width="482px"></asp:TextBox>
           <asp:TextBox ID="DDivision2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 119; left: 947px;
               position: absolute; top: 63px" Width="186px"></asp:TextBox>
           <asp:Button ID="BDivision" runat="server" Height="25px" Style="z-index: 120;
                left: 590px; position: absolute; top: 49px" Text="....." Width="28px" /><asp:Button ID="BFormNo" runat="server" Height="25px" Style="z-index: 121;
                left: 790px; position: absolute; top: 21px" Text="....." Width="28px" />
           <asp:TextBox ID="DCombiItem1" runat="server" AutoCompleteType="Disabled" BackColor="Yellow"
               BorderStyle="Groove" ForeColor="Blue" Height="22px" Style="z-index: 122; left: 503px;
               position: absolute; top: 23px" Width="283px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DCustomer" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 123; left: 104px; position: absolute;
               top: 75px" Width="136px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;
           <asp:DropDownList ID="DYKKColortype" runat="server" BackColor="Yellow" Style="z-index: 124;
               left: 503px; position: absolute; top: 76px" Width="117px">
           </asp:DropDownList>
           <asp:Button ID="BClear" runat="server" Style="z-index: 125; left: 16px; position: absolute;
               top: 152px" Text="清除條件" ForeColor="Blue" BackColor="Lime" />
           <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/CombiInq_01.jpg" Style="z-index: 99;
               left: 6px; position: absolute; top: 15px" />
           <asp:Label ID="Label3" runat="server" ForeColor="Red" Height="16px" Style="z-index: 140;
               left: 248px; position: absolute; top: 152px" Text="未封存-三年內      已封存-超過三年"></asp:Label>
           <asp:DropDownList ID="DTABLE" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 139; left: 104px; position: absolute;
               top: 152px" Width="141px">
               <asp:ListItem>未封存</asp:ListItem>
               <asp:ListItem>已封存</asp:ListItem>
           </asp:DropDownList>
           <asp:TextBox ID="DLTape" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 127; left: 105px; position: absolute; top: 104px"
               Width="136px"></asp:TextBox>
           <asp:TextBox ID="DLChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 128; left: 315px; position: absolute;
               top: 102px" Width="95px"></asp:TextBox>
           <asp:TextBox ID="DRChain" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 129; left: 503px; position: absolute;
               top: 103px" Width="117px"></asp:TextBox>
           <asp:TextBox ID="DRTape" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="20px" Style="z-index: 130; left: 711px; position: absolute; top: 104px"
               Width="102px"></asp:TextBox>
       
         
           <asp:textbox id="DSts1" style="Z-INDEX: 131; LEFT: 104px; POSITION: absolute; TOP: 22px" runat="server"
				Height="22px" Width="105px" ForeColor="Blue" BackColor="Yellow" BorderStyle="Groove" AutoCompleteType="Disabled"></asp:textbox>
           <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 132;
               left: 239px; position: absolute; top: 130px; text-align: center" Text="~" Width="6px"></asp:Label>
           <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 133;
               left: 677px; position: absolute; top: 131px; text-align: center" Text="~" Width="6px"></asp:Label>
  	 
      </div>
    </form>
</body>
</html>
