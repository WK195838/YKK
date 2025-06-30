<%@ Page Language="vb" AutoEventWireup="false" Inherits="CloseAccountSheet_03" aspCompat="True" EnableEventValidation = "false"  CodeFile="CloseAccountSheet_03.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>清算申請單</title>
    
		
		<style type="text/css">
        .TextUpper
        {
            text-transform:uppercase;
        }
</style>
 
		
</head>

<script language="javascript" type="text/javascript">
var hkey_root,hkey_path,hkey_key
hkey_root="HKEY_CURRENT_USER"
hkey_path="\\Software\\Microsoft\\Internet Explorer\\PageSetup\\"

function previewScreen(block){

// 設定網頁列印的頁首頁尾為空
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="header"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
hkey_key="footer"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
}catch(e){}

//只印列特定區域
var value = block.innerHTML;
var printPage = window.open("","printPage","");
printPage.document.open();
printPage.document.write("<OBJECT classid='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2' height=0 id=wc name=wc width=0></OBJECT>");
printPage.document.write("<HTML><head></head><BODY onload='javascript:wc.execwb(7,1);window.close()'>");
printPage.document.write("<PRE>");
printPage.document.write(value);
printPage.document.write("</PRE>");
printPage.document.close("</BODY></HTML>");
}
 
</script>


<body>


  	<form id="Form1" runat="server">
   	<input type="button" value="預覽列印 (橫式列印 85%）" onclick="previewScreen(block)" style="z-index: 134; left: 960px; width: 184px; position: absolute; top: 32px; height: 48px;" id="Button1"  />
          &nbsp;
   	<div id="block">   
  	 	
        	<FONT face="新細明體"></FONT><div>
               <asp:Label ID="DJobTitle1" runat="server" Font-Size="12pt" Style="z-index: 100; left: 520px;
                   position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DLocation" runat="server" Font-Size="12pt" Style="z-index: 101; left: 728px;
                   position: absolute; top: 216px"></asp:Label>
                &nbsp; &nbsp;
               <asp:Label ID="DSDate" runat="server" Font-Size="12pt" Style="z-index: 104; left: 168px;
                   position: absolute; top: 248px">2023/07/11</asp:Label>
               <asp:Label ID="DEDate" runat="server" Font-Size="12pt" Style="z-index: 105; left: 272px;
                   position: absolute; top: 248px"></asp:Label>
               <asp:Label ID="DSumAmt" runat="server" Font-Size="12pt" Style="z-index: 106; left: 1048px;
                   position: absolute; top: 296px"></asp:Label>           &nbsp; &nbsp;&nbsp;
             
               <asp:Label ID="DDivision2" runat="server" Font-Size="12pt" Style="z-index: 111; left: 168px;
                   position: absolute; top: 184px"></asp:Label>
               <asp:Label ID="DJobTitle2" runat="server" Font-Size="12pt" Style="z-index: 112; left: 520px;
                   position: absolute; top: 184px"></asp:Label>
               <asp:Label ID="DEmpID2" runat="server" Font-Size="12pt" Style="z-index: 113; left: 448px;
                   position: absolute; top: 184px"></asp:Label>
               <asp:Label ID="DDivisionCode2" runat="server" Font-Size="12pt" Style="z-index: 114;
                   left: 371px; position: absolute; top: 184px" Width="80px"></asp:Label>
               <asp:Label ID="DEmpName2" runat="server" Font-Size="12pt" Style="z-index: 115; left: 256px;
                   position: absolute; top: 184px" Width="112px"></asp:Label>
               <asp:Label ID="DDivision1" runat="server" Font-Size="12pt" Style="z-index: 116; left: 168px;
                   position: absolute; top: 144px"></asp:Label>
                &nbsp;
               <asp:Label ID="DEmpID1" runat="server" Font-Size="12pt" Style="z-index: 118; left: 448px;
                   position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DDivisionCode1" runat="server" Font-Size="12pt" Style="z-index: 119;
                   left: 376px; position: absolute; top: 144px" Width="80px"></asp:Label>
               <asp:Label ID="DEmpName1" runat="server" Font-Size="12pt" Style="z-index: 120; left: 256px;
                   position: absolute; top: 144px" Width="112px"></asp:Label>
               <asp:Label ID="DNo" runat="server" Font-Size="12pt" Style="z-index: 121; left: 728px;
                   position: absolute; top: 120px"></asp:Label>
               <asp:Label ID="DDate" runat="server" Font-Size="12pt" Style="z-index: 122; left: 168px;
                   position: absolute; top: 120px"></asp:Label>     &nbsp; &nbsp;&nbsp;
          <img id="IMG1" onclick="return IMG1_onclick()" src="images/CloseAccountSheet_03.jpg"
              style="z-index: 99; left: 24px; position: absolute; top: 8px; text-align: left" />
                <asp:Label ID="DQCNO" runat="server" Font-Size="12pt" Style="z-index: 121; left: 1024px;
                    position: absolute; top: 184px"></asp:Label>
                <asp:Label ID="DTripNo" runat="server" Font-Size="12pt" Style="z-index: 121; left: 728px;
                    position: absolute; top: 152px"></asp:Label>
                &nbsp; &nbsp;
                <asp:Label ID="DASDate" runat="server" Font-Size="12pt" Style="z-index: 104; left: 728px;
                    position: absolute; top: 248px">2023/07/11</asp:Label>
                <asp:Label ID="DAEDate" runat="server" Font-Size="12pt" Style="z-index: 105; left: 840px;
                    position: absolute; top: 248px"></asp:Label>
                <asp:Label ID="Label4" runat="server" Font-Size="12pt" Style="z-index: 105; left: 808px;
                    position: absolute; top: 248px">~</asp:Label>
                <asp:Label ID="Label1" runat="server" Font-Size="12pt" Style="z-index: 105; left: 248px;
                    position: absolute; top: 248px">~</asp:Label>
          <asp:Label ID="DDivision3" runat="server" Font-Size="12pt" Style="z-index: 127; left: 168px;
              position: absolute; top: 216px"></asp:Label>
          <asp:Label ID="DDivisionCode3" runat="server" Font-Size="12pt" Style="z-index: 128;
              left: 408px; position: absolute; top: 216px"></asp:Label>
                &nbsp; &nbsp;&nbsp;
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderStyle="Groove"
                    DataKeyNames="Unique_ID" Font-Size="9pt" PageSize="100" Style="z-index: 103;
                    left: 40px; position: absolute; top: 368px" Width="1176px">
                    <Columns>
                        <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Type" HeaderText="類別用途">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Pay" HeaderText="支付" >
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SDate" HeaderText="日期">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Currency" HeaderText="幣別">
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Money" DataFormatString="{0:N2}" HeaderText="金額">
                            <HeaderStyle Height="20px" />
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Days" HeaderText="數量/次數">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="80px" HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Rate" HeaderText="匯率">
                            <ItemStyle Width="100px" HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SumAmt" HeaderText="小計金額" DataFormatString="{0:N0}">
                            <ItemStyle HorizontalAlign="Right" Width="100px" />
                            <ControlStyle Width="120px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ExpenseNo" HeaderText="交際費單號" >
                            <ItemStyle Width="120px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Remark" HeaderText="備註">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="400px" />
                            <ControlStyle Width="200px" />
                        </asp:BoundField>
                    </Columns>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:GridView>
                <asp:TextBox ID="DObject" runat="server" AutoPostBack="True" BackColor="White" BorderStyle="Groove"
                    Font-Size="X-Small" ForeColor="Black" Height="32px" Style="z-index: 109; left: 720px;
                    position: absolute; top: 176px" TextMode="MultiLine" Width="184px"></asp:TextBox>
        
</div>
 
    </form>
</body>
</html>
