<%@ Page Language="vb" AutoEventWireup="false" Inherits="FundinginqCommissionTC" aspCompat="True" EnableEventValidation = "false"  CodeFile="FundinginqCommissionTC.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>經費申請單調閱</title>
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
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 648px; POSITION: absolute; TOP: 120px" runat="server"
				ImageUrl="~/Images/msexcel.gif" Height="21px" Width="21px"></asp:imagebutton>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 160px" runat="server"
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
			<asp:button id="Go" style="Z-INDEX: 121; LEFT: 608px; POSITION: absolute; TOP: 120px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/Fundinginq_04.jpg" Style="z-index: 99;
               left: 8px; position: absolute; top: 16px" />
           <asp:TextBox ID="DNO" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 128px; position: absolute;
               top: 96px" Width="216px"></asp:TextBox>
           <asp:TextBox ID="DDepName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 128px; position: absolute;
               top: 72px" Width="216px"></asp:TextBox>
           <asp:TextBox ID="DAppName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 464px; position: absolute;
               top: 72px" Width="200px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DAppDateS" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 464px; position: absolute;
               top: 96px" Width="64px"></asp:TextBox>
           <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 160;
               left: 560px; position: absolute; top: 96px; text-align: center" Text="~" Width="6px"></asp:Label>
           <asp:TextBox ID="DAppDateE" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 110; left: 576px; position: absolute;
               top: 96px" Width="64px"></asp:TextBox>
           <asp:Button ID="BASDate1" runat="server" Height="24px" Style="z-index: 111; left: 528px;
               position: absolute; top: 96px" Text="..." Width="24px" />
           <asp:Button ID="BAEDate1" runat="server" Height="24px" Style="z-index: 112; left: 640px;
               position: absolute; top: 96px" Text="..." Width="24px" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:DropDownList ID="DProcess" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 256; left: 128px; position: absolute; top: 45px"
               Width="216px">
               </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;&nbsp;
           <asp:DropDownList ID="DSTS" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 255; left: 464px; position: absolute; top: 45px"
               Width="200px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="0">核定中</asp:ListItem>
               <asp:ListItem Value="1">完成</asp:ListItem>
               <asp:ListItem Value="2">取消</asp:ListItem>
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:DropDownList ID="DDepo" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 256; left: 128px; position: absolute; top: 120px"
               Width="216px">
               <asp:ListItem>台北</asp:ListItem>
               <asp:ListItem>中壢</asp:ListItem>
           </asp:DropDownList>
           <asp:TextBox ID="DName2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 464px; position: absolute;
               top: 120px" Width="136px"></asp:TextBox>
  	 
      </div>
    </form>
</body>
</html>
