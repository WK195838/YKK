<%@ Page Language="vb" AutoEventWireup="false" Inherits="QCModListinqCommission" aspCompat="True" EnableEventValidation = "false"  CodeFile="QCModListinqCommission.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>資料修改依賴書調閱資料</title>
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
			<asp:imagebutton id="BExcel" style="Z-INDEX: 100; LEFT: 601px; POSITION: absolute; TOP: 135px" runat="server"
				ImageUrl="~/Images/msexcel.gif" Height="21px" Width="21px"></asp:imagebutton>
				<asp:datagrid id="DataGrid1" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 180px" runat="server"
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
                    <asp:BoundColumn DataField="Field7"></asp:BoundColumn>
                    <asp:BoundColumn DataField="Field8"></asp:BoundColumn>
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
			<asp:button id="Go" style="Z-INDEX: 121; LEFT: 551px; POSITION: absolute; TOP: 134px" runat="server"
				Height="24px" Width="40px" ForeColor="Blue" BackColor="Aqua" Text="Go"></asp:button>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/images/QAModing.jpg" Style="z-index: 99;
               left: 12px; position: absolute; top: 35px" />
           <asp:DropDownList ID="DSTS" runat="server" BackColor="Yellow" ForeColor="Blue" Height="20px"
               Style="z-index: 255; left: 682px; position: absolute; top: 100px" Width="120px">
               <asp:ListItem></asp:ListItem>
               <asp:ListItem Value="0">核定中</asp:ListItem>
               <asp:ListItem Value="1">完成</asp:ListItem>
               <asp:ListItem Value="2">取消</asp:ListItem>
           </asp:DropDownList>
           <asp:TextBox ID="DNo" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 682px; position: absolute; top: 66px;
               text-align: left" Width="112px"></asp:TextBox>
           <asp:TextBox ID="DCustomer" runat="server" BackColor="#FFFF80" BorderStyle="Groove"
               ForeColor="Black" Height="24px" Style="z-index: 112; left: 155px; position: absolute;
               top: 99px" Width="112px"></asp:TextBox>
           <asp:TextBox ID="DBuyer" runat="server" BackColor="#FFFF80" BorderStyle="Groove"
               ForeColor="Black" Height="24px" Style="z-index: 110; left: 415px; position: absolute;
               top: 100px" Width="112px"></asp:TextBox>
           <asp:TextBox ID="DModNo" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 156px; position: absolute; top: 135px;
               text-align: left" Width="112px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:TextBox ID="DName" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 415px; position: absolute; top: 66px;
               text-align: left" Width="112px"></asp:TextBox>
           <input id="BADate" runat="server" style="z-index: 152; left: 253px; width: 24px;
               position: absolute; top: 64px" type="button" value="..." />
           <asp:TextBox ID="DADate" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 155px; position: absolute; top: 66px;
               text-align: left" Width="96px"></asp:TextBox>
           <asp:TextBox ID="DPuller" runat="server" BackColor="#FFFF80" BorderStyle="None" ForeColor="Black"
               Height="20px" Style="z-index: 126; left: 417px; position: absolute; top: 136px;
               text-align: left" Width="112px"></asp:TextBox>
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
  	 
      </div>
    </form>
</body>
</html>
