<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default_Stock.aspx.vb" Inherits="_Default_Stock" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
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
    <form id="form1" runat="server">
    <div>
        &nbsp; &nbsp;FormNo:<asp:DropDownList ID="DropDownList1" runat="server">
            <asp:ListItem Value="001151">Item登錄單</asp:ListItem>
            <asp:ListItem Value="001152">ZIP登錄單</asp:ListItem>
            <asp:ListItem Value="001153">SLD登錄單</asp:ListItem>
            <asp:ListItem Value="001154">CH登錄單</asp:ListItem>
            <asp:ListItem Value="001155">SLD-工廠登錄單</asp:ListItem>
            <asp:ListItem Value="001161">單價承認委託單</asp:ListItem>
            <asp:ListItem Selected="True" Value="001101">案件委託</asp:ListItem>
        </asp:DropDownList>
        FormSNo<asp:TextBox ID="TextBox2" runat="server" Width="95px">0</asp:TextBox>Step<asp:TextBox
            ID="TextBox3" runat="server" Width="43px">1</asp:TextBox>
        SeqNo<asp:TextBox ID="TextBox4" runat="server" Width="66px">1</asp:TextBox>
        ApplyID<asp:TextBox ID="TextBox5" runat="server" Width="67px">it013</asp:TextBox>
        UserID<asp:TextBox ID="TextBox6" runat="server" Width="67px">it013</asp:TextBox>&nbsp;<br />
        &nbsp;<br />
        &nbsp;<br />
        &nbsp;&nbsp; &nbsp;<br />
        &nbsp; &nbsp;<br />
        &nbsp; &nbsp;<br />
        &nbsp; &nbsp;<asp:Button ID="Button4" runat="server" Text="清算申請01" style="z-index: 100; left: 16px; position: absolute; top: 112px" /><asp:Button ID="Button6" runat="server" Text="清算申請02" style="z-index: 100; left: 136px; position: absolute; top: 112px" /><asp:Button ID="Button8" runat="server" Text="出差清算申請主畫面" style="z-index: 105; left: 304px; position: absolute; top: 96px" /><br /><asp:Button ID="Button1" runat="server" Text="多筆棧板號" style="z-index: 101; left: 144px; position: absolute; top: 328px" />
        <br />
        <br /><asp:Button ID="Button7" runat="server" Text="出差申請02" style="z-index: 100; left: 120px; position: absolute; top: 64px" />
        <br />
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <asp:Button ID="Button26" runat="server" Text="入庫01" style="z-index: 102; left: 24px; position: absolute; top: 256px" /><asp:Button ID="Button27" runat="server" Text="出庫01" style="z-index: 103; left: 96px; position: absolute; top: 256px" />
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;<asp:Button ID="Button3" runat="server" Text="單筆棧板號" style="z-index: 104; left: 24px; position: absolute; top: 328px" />
        <asp:Button ID="Button2" runat="server" Text="出差申請" style="z-index: 105; left: 24px; position: absolute; top: 64px" />
        <asp:Button ID="Button5" runat="server" Style="z-index: 107; left: 24px; position: absolute;
            top: 208px" Text="出入庫主畫面" />
    </div>
        
        
    </form>
</body>
</html>
