<%@ Page Language="vb" AutoEventWireup="false" Inherits="StockInSheet_04" aspCompat="True" EnableEventValidation = "false"  CodeFile="StockInSheet_04.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>經費申請單</title>
    	
		
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
 
 
     function RunExe(strFile)
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\StockInOut\\PGM\\Stock\\Stock_AutoPrint.exe ';
    try
    {
        var shellActiveXObject = new ActiveXObject("WScript.Shell");

        if ( !shellActiveXObject )
        {
            alert('Could not get reference to WScript.Shell');
        }

        shellActiveXObject.Run(executableFullPath + strFile , 1, false);
        shellActiveXObject = null;
    }
    catch (errorObject)
    {
        alert('Error:\n' + errorObject.message);
    }            
}


</script>

		

<body>

   <form id="form1" runat="server">
       &nbsp;&nbsp;
  	<div id="block">
  	
       <br />
       <br />
      
       <br />
          &nbsp;
       &nbsp;&nbsp;&nbsp;
       <br />
       <asp:Image ID="Image2" runat="server" ImageUrl="~/images/StockIn_04.jpg" Style="z-index: 99;
           left: 8px; position: absolute; top: 8px" />
          <asp:Label ID="DDepName" runat="server" Font-Size="12pt" Style="z-index: 105; left: 224px;
              position: absolute; top: 120px"></asp:Label>
          <asp:Label ID="DType" runat="server" Font-Size="12pt" Style="z-index: 105; left: 224px;
              position: absolute; top: 152px"></asp:Label>
          &nbsp;&nbsp;
          <asp:Label ID="DName" runat="server" Font-Size="12pt" Style="z-index: 114; left: 480px;
              position: absolute; top: 112px"></asp:Label>
          <asp:Label ID="DNo" runat="server" Style="z-index: 103; left: 480px; position: absolute;
              top: 88px"></asp:Label>
          <asp:Label ID="DDate" runat="server" Font-Size="12pt" Style="z-index: 104; left: 224px;
              position: absolute; top: 88px"></asp:Label>
          &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
          <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="White" BorderWidth="1px" CellPadding="4" Font-Size="9pt"
              Style="z-index: 105; left: 32px; position: absolute; top: 200px" Width="632px" BorderStyle="Ridge">
              <Columns>
                  <asp:BoundField DataField="StockNo" HeaderText="棧板號" />
                  <asp:BoundField DataField="ITEMCODE" HeaderText="ITEM" />
                  <asp:BoundField DataField="COLOR" HeaderText="COLOR" />
                  <asp:BoundField DataField="QTY" HeaderText="入庫數量" DataFormatString="{0:N0}" />
                  <asp:BoundField DataField="UNIT" HeaderText="單位" />
              </Columns>
              <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" HorizontalAlign="Center" />
              <HeaderStyle BackColor="White" ForeColor="Black" HorizontalAlign="Center" VerticalAlign="Middle" />
          </asp:GridView>
          <asp:Button ID="BPrint" runat="server" BackColor="#FFE0C0" Style="z-index: 104; left: 728px;
              position: absolute; top: 72px" Text="列印入庫單" Width="104px" />
         
      </div>
       
    </form>
</body>
</html>
