<%@ Page Language="vb" AutoEventWireup="false" Inherits="FundingSheet_03" aspCompat="True" EnableEventValidation = "false"  CodeFile="FundingSheet_03.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>�g�O�ӽг�</title>
    	
		
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

// �]�w�����C�L��������������
try{
var RegWsh = new ActiveXObject("WScript.Shell")
hkey_key="header"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
hkey_key="footer"
RegWsh.RegWrite(hkey_root+hkey_path+hkey_key,"")
}catch(e){}

//�u�L�C�S�w�ϰ�
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

  	<form id="Form1" runat="server" target="_blank" >
  	<input type="button" value="�w���C�L (��C�L 85%�^" onclick="previewScreen(block)"   style="z-index: 119; left: 944px; width: 184px; position: absolute; top: 40px; height: 48px;" id="Button1"  >
          <input id="File1" type="file" />
          &nbsp;
  	<div id="block">
  	
        <FONT face="�s�ө���"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<div>
               <img src="images/FundingSheet_03.jpg" style="z-index: 99; text-align:left; left: 16px; position: absolute;
                   top: -6px" id="IMG1" onclick="return IMG1_onclick()" />
             
               &nbsp;
               <asp:Label ID="DApplyAmt" runat="server"  Style="z-index: 100;text-align:right; left: 128px; position: absolute;  
                   top: 224px " BorderStyle="None" Width="80px" Font-Size="12pt" ></asp:Label>
               <asp:Label ID="DDebitAmt" runat="server" Style="z-index: 101;text-align:right; left: 432px; position: absolute;
                   top: 224px" Width="80px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DSumAmt" runat="server" Style="z-index: 102; text-align:right; left: 728px; position: absolute; 
                   top: 224px" Width="80px" Font-Size="12pt" ></asp:Label>
            
               <asp:Label ID="DNo" runat="server" Style="z-index: 103; left: 616px; position: absolute;
                   top: 48px"></asp:Label>
           
               <asp:Label ID="DDate" runat="server" Style="z-index: 104; left: 168px; position: absolute;
                   top: 48px" Font-Size="12pt"></asp:Label>
           
               <asp:Label ID="DDivision1" runat="server" Style="z-index: 105; left: 168px; position: absolute;
                   top: 80px" Font-Size="12pt"></asp:Label>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
               <asp:Label ID="DJobTitle1" runat="server" Style="z-index: 106; left: 632px; position: absolute;
                   top: 80px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DEmpID1" runat="server" Style="z-index: 107; left: 528px; position: absolute;
                   top: 80px" Font-Size="12pt"></asp:Label>
               &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
               <asp:Label ID="DDivisionCode1" runat="server" Style="z-index: 108; left: 392px; position: absolute;
                   top: 80px" Font-Size="12pt"></asp:Label>
           
               <asp:Label ID="DDivision2" runat="server" Style="z-index: 109; left: 168px; position: absolute;
                   top: 112px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DJobTitle2" runat="server" Style="z-index: 110; left: 632px; position: absolute;
                   top: 112px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DEmpID2" runat="server" Style="z-index: 111; left: 528px; position: absolute;
                   top: 112px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DDivisionCode2" runat="server" Style="z-index: 112; left: 392px; position: absolute;
                   top: 112px" Font-Size="12pt"></asp:Label>
               <asp:Label ID="DEmpName2" runat="server" Style="z-index: 113; left: 288px; position: absolute;
                   top: 112px" Font-Size="12pt"></asp:Label>
            
               <asp:Label ID="DEmpName1" runat="server" Style="z-index: 114; left: 288px; position: absolute;
                   top: 80px" Font-Size="12pt"></asp:Label>
  
             
               <!--************************************************************ 
            ** ����Ы� [BUTTON 2 ��]                               **     
            ************************************************************ -->
   
           </div>
        
           <asp:Label ID="DDivision3" runat="server" Style="z-index: 115; left: 168px; position: absolute;
               top: 152px" Font-Size="12pt"></asp:Label>
           <asp:Label ID="DDivisionCode3" runat="server" Style="z-index: 116; left: 400px; position: absolute;
               top: 152px" Font-Size="12pt"></asp:Label>
          &nbsp;&nbsp;
          <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" BorderColor="Black"
              Font-Size="12pt" Height="1px" Style="z-index: 117; left: 16px; position: absolute;
              top: 304px" Width="1304px">
              <Columns>
                  <asp:BoundField DataField="ExpItem" HeaderText="���O�γ~">
                      <HeaderStyle Height="20px" Width="80px" />
                      <ItemStyle Width="160px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="ADate" DataFormatString="{0:d}" HeaderText="���">
                      <HeaderStyle Height="20px" />
                      <ItemStyle Width="100px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="TaxType" HeaderText="�o������(�|�O)">
                      <HeaderStyle Height="20px" />
                      <ItemStyle Width="160px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="InvoiceNo" HeaderText="�o�����X�@�|�渹�X" >
                      <ItemStyle Width="160px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="GUINo" HeaderText="���Τ@�s��" >
                      <ItemStyle Width="120px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="NetAmt" DataFormatString="{0:N0}" HeaderText="�b�B">
                      <HeaderStyle Height="20px" />
                      <ItemStyle HorizontalAlign="Right" Width="80px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="TaxAmt" DataFormatString="{0:N0}" HeaderText="�|�B">
                      <HeaderStyle Height="20px" />
                      <ItemStyle HorizontalAlign="Right" Width="80px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="Amt" DataFormatString="{0:N0}" HeaderText="�`�B">
                      <HeaderStyle Height="20px" />
                      <ItemStyle HorizontalAlign="Right" Width="80px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="Content" HeaderText="���e">
                      <HeaderStyle Height="10px" />
                      <ItemStyle Width="200px" />
                  </asp:BoundField>
                  <asp:BoundField DataField="Remark" HeaderText="�Ƶ�">
                      <ItemStyle Width="150px" />
                  </asp:BoundField>
              </Columns>
              <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
          </asp:GridView>
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
         
      </div>
       
    </form>
</body>
</html>
