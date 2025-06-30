<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

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
        &nbsp; &nbsp;<br /><asp:Button ID="Button3" runat="server" Text="品質分析申請01" style="z-index: 105; left: 32px; position: absolute; top: 312px" />
        <br /><asp:Button ID="Button1" runat="server" Text="品質分析申請01" style="z-index: 105; left: 40px; position: absolute; top: 64px" />
        <br /><asp:Button ID="Button7" runat="server" Text="品質分析申請02" style="z-index: 100; left: 192px; position: absolute; top: 64px" />
        &nbsp;<br />
        <br />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
        <asp:Button ID="Button2" runat="server" Text="首頁" style="z-index: 105; left: 24px; position: absolute; top: 128px" />
        &nbsp;&nbsp;
        <asp:Button ID="Button4" runat="server" Text="品質分析申請02" style="z-index: 105; left: 200px; position: absolute; top: 320px" />
    </div>
        
        
    </form>
</body>
</html>
