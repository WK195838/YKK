<%@ Page Language="vb" AutoEventWireup="false" Inherits="ExpenseinqCommission" aspCompat="True" EnableEventValidation = "false"  CodeFile="ExpenseinqCommission.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>�J�ء����S�ӽЪ�վ\���</title>
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

				answer = confirm("�O�_����[" + MSG + "]�@�~�ܡH �нT�{....");
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
        	<FONT face="�s�ө���"></FONT>
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 736px; POSITION: absolute; TOP: 144px" runat="server"
				ImageUrl="~/Images/msexcel.gif" Height="21px" Width="21px"></asp:imagebutton>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 200px" runat="server"
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
						HeaderText="�i��">
						<HeaderStyle Width="30px"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
					</asp:HyperLinkColumn>
				</Columns>
				<PagerStyle NextPageText="�U�@��" PrevPageText="�W�@��" HorizontalAlign="Center" ForeColor="#993300"
					BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
			</asp:datagrid>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
			<asp:button id="Go" style="Z-INDEX: 121; LEFT: 688px; POSITION: absolute; TOP: 144px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/Expenseinq.jpg" Style="z-index: 99;
               left: -14px; position: absolute; top: 16px" />
           <asp:TextBox ID="DASDateS" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 120px; position: absolute;
               top: 120px" Width="64px"></asp:TextBox>
           <asp:DropDownList ID="DDepName" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 257; left: 120px; position: absolute; top: 62px"
               Width="200px">
           </asp:DropDownList>
           <asp:DropDownList ID="DPayman" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 257; left: 120px; position: absolute; top: 88px"
               Width="200px">
           </asp:DropDownList>          
           &nbsp;
           <asp:TextBox ID="DASDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 120px; position: absolute;
               top: 144px" Width="64px"></asp:TextBox>
           <asp:TextBox ID="DAEDateE" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 240px; position: absolute;
               top: 120px" Width="64px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp;
           <asp:Label ID="Label2" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 160;
               left: 216px; position: absolute; top: 144px; text-align: center" Text="~" Width="6px"></asp:Label>
           <asp:Label ID="Label1" runat="server" Font-Bold="False" Font-Size="Small" Style="z-index: 160;
               left: 216px; position: absolute; top: 120px; text-align: center" Text="~" Width="6px"></asp:Label>
           <asp:DropDownList ID="DProcess" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 256; left: 120px; position: absolute; top: 35px"
               Width="216px">
               </asp:DropDownList>
           <asp:TextBox ID="DAEDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 110; left: 240px; position: absolute;
               top: 144px" Width="64px"></asp:TextBox>
           <asp:Button ID="BASDate1" runat="server" Height="26px" Style="z-index: 111; left: 184px;
               position: absolute; top: 112px" Text="..." Width="25px" />
           <asp:Button ID="BASDate" runat="server" Height="26px" Style="z-index: 111; left: 184px;
               position: absolute; top: 144px" Text="..." Width="25px" />
           <asp:Button ID="BAEDate1" runat="server" Height="26px" Style="z-index: 111; left: 304px;
               position: absolute; top: 112px" Text="..." Width="25px" />
           &nbsp;&nbsp;
           <asp:Button ID="BAEDate" runat="server" Height="26px" Style="z-index: 112; left: 304px;
               position: absolute; top: 144px" Text="..." Width="25px" />
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;&nbsp;
           <asp:DropDownList ID="DSTS" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 255; left: 456px; position: absolute; top: 35px"
               Width="200px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="0">�֩w��</asp:ListItem>
               <asp:ListItem Value="1">����</asp:ListItem>
               <asp:ListItem Value="2">����</asp:ListItem>
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DAppName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 456px; position: absolute;
               top: 64px" Width="200px"></asp:TextBox>
           <asp:TextBox ID="DFormNo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="20px" Style="z-index: 113; left: 456px; position: absolute;
               top: 92px" Width="200px"></asp:TextBox>
           <asp:Label ID="DCount" runat="server" Height="20px" Style="z-index: 114; left: 606px; position: absolute;
               top: 120px" Width="200px" ForeColor="Red"></asp:Label></div>
    </form>
</body>
</html>
