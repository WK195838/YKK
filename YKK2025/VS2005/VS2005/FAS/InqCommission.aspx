<%@ Page Language="vb" AutoEventWireup="false" Inherits="InqCommission" aspCompat="True" EnableEventValidation = "false"  CodeFile="InqCommission.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>報廢處理申請書調閱資料</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
			function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
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
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 493px; POSITION: absolute; TOP: 92px" runat="server"
				ImageUrl="~/Images/msexcel.gif" Height="21px" Width="21px" Visible="False"></asp:imagebutton>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 17px; POSITION: absolute; TOP: 165px" runat="server"
				Height="100px" Width="941px" BackColor="White" BorderStyle="None" AutoGenerateColumns="False" CellPadding="4" BorderWidth="1px" BorderColor="#CC9966" Font-Size="9pt" AllowPaging="True" PageSize="15">
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
                    <asp:BoundColumn DataField="Field2"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field3"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field4"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field5"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field6"></asp:BoundColumn>
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
			<asp:button id="Go" style="Z-INDEX: 121; LEFT: 446px; POSITION: absolute; TOP: 89px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DDivision2" runat="server" AutoCompleteType="Disabled" BackColor="White"
               BorderStyle="None" ForeColor="White" Height="22px" Style="z-index: 124; left: 947px;
               position: absolute; top: 63px" Width="186px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/InqCommission.jpg" Style="z-index: 99;
               left: 7px; position: absolute; top: 14px" />
           <asp:TextBox ID="DDivision" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 260; left: 158px; position: absolute;
               top: 66px" Width="329px"></asp:TextBox>
           <asp:TextBox ID="DAEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 110; left: 314px; position: absolute;
               top: 92px" Width="88px"></asp:TextBox>
           <asp:Button ID="BASDate" runat="server" Height="26px" Style="z-index: 111; left: 258px;
               position: absolute; top: 88px" Text="....." Width="28px" />
           <asp:Button ID="BAEDate" runat="server" Height="26px" Style="z-index: 112; left: 407px;
               position: absolute; top: 88px" Text="....." Width="28px" />
           <asp:TextBox ID="DASDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 158px; position: absolute;
               top: 92px" Width="93px"></asp:TextBox>
           <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 160;
               left: 291px; position: absolute; top: 91px; text-align: center" Text="~" Width="6px"></asp:Label>
           &nbsp; &nbsp;&nbsp;&nbsp;
           <asp:DropDownList ID="DSTS" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 255; left: 159px; position: absolute; top: 40px"
               Width="99px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="0">核定中</asp:ListItem>
               <asp:ListItem Value="1">完成</asp:ListItem>
               <asp:ListItem Value="2">取消</asp:ListItem>
           </asp:DropDownList>
           &nbsp; &nbsp;
           <asp:DropDownList ID="DTYPE" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 259; left: 579px; position: absolute;
               top: 66px" Width="200px">
           </asp:DropDownList>
           &nbsp;
           <asp:TextBox ID="DAppName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 260; left: 578px; position: absolute;
               top: 39px" Width="202px"></asp:TextBox>
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 261; left: 327px; position: absolute; top: 40px"
               Width="162px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp;
  	 
      </div>
    </form>
</body>
</html>
