<%@ Page Language="vb" AutoEventWireup="false" Inherits="BusinessTripSheet_03" aspCompat="True" EnableEventValidation = "false"  CodeFile="BusinessTripSheet_03.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>�X�t�ӽг�</title>
    
		
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


  	<form id="Form1" runat="server">
   	<input type="button" value="�w���C�L (��C�L 85%�^" onclick="previewScreen(block)" style="z-index: 134; left: 984px; width: 184px; position: absolute; top: 16px; height: 48px;" id="Button1"  />
          &nbsp;
   	<div id="block">   
  	 	
        	<FONT face="�s�ө���"></FONT><div>
               <asp:Label ID="DJobTitle1" runat="server" Font-Size="12pt" Style="z-index: 100; left: 632px;
                   position: absolute; top: 112px"></asp:Label>
               <asp:Label ID="DLocation" runat="server" Font-Size="12pt" Style="z-index: 101; left: 280px;
                   position: absolute; top: 360px"></asp:Label>
               <asp:Label ID="DObject" runat="server" Font-Size="12pt" Style="z-index: 102; left: 48px;
                   position: absolute; top: 360px"></asp:Label>
               <asp:Label ID="DPassDate" runat="server" Font-Size="12pt" Style="z-index: 103; left: 896px;
                   position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DSDate" runat="server" Font-Size="12pt" Style="z-index: 104; left: 512px;
                   position: absolute; top: 360px"></asp:Label>
               <asp:Label ID="DEDate" runat="server" Font-Size="12pt" Style="z-index: 105; left: 640px;
                   position: absolute; top: 360px"></asp:Label>
               <asp:Label ID="DRemark" runat="server" Font-Size="12pt" Style="z-index: 106; left: 784px;
                   position: absolute; top: 360px"></asp:Label>           
               <asp:Label ID="DDivision4" runat="server" Font-Size="12pt" Style="z-index: 107; left: 184px;
                   position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DEmpID4" runat="server" Font-Size="12pt" Style="z-index: 108; left: 528px;
                   position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DDivisionCode4" runat="server" Font-Size="12pt" Style="z-index: 109;
                   left: 400px; position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DEmpName4" runat="server" Font-Size="12pt" Style="z-index: 110; left: 280px;
                   position: absolute; top: 144px"></asp:Label>
             
               <asp:Label ID="DDivision2" runat="server" Font-Size="12pt" Style="z-index: 111; left: 184px;
                   position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DJobTitle2" runat="server" Font-Size="12pt" Style="z-index: 112; left: 632px;
                   position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DEmpID2" runat="server" Font-Size="12pt" Style="z-index: 113; left: 528px;
                   position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DDivisionCode2" runat="server" Font-Size="12pt" Style="z-index: 114;
                   left: 400px; position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DEmpName2" runat="server" Font-Size="12pt" Style="z-index: 115; left: 280px;
                   position: absolute; top: 176px"></asp:Label>
               <asp:Label ID="DDivision1" runat="server" Font-Size="12pt" Style="z-index: 116; left: 184px;
                   position: absolute; top: 112px"></asp:Label>
               <asp:Label ID="DJobTitle4" runat="server" Font-Size="12pt" Style="z-index: 117; left: 632px;
                   position: absolute; top: 144px"></asp:Label>
               <asp:Label ID="DEmpID1" runat="server" Font-Size="12pt" Style="z-index: 118; left: 528px;
                   position: absolute; top: 112px"></asp:Label>
               <asp:Label ID="DDivisionCode1" runat="server" Font-Size="12pt" Style="z-index: 119;
                   left: 400px; position: absolute; top: 112px"></asp:Label>
               <asp:Label ID="DEmpName1" runat="server" Font-Size="12pt" Style="z-index: 120; left: 280px;
                   position: absolute; top: 112px"></asp:Label>
               <asp:Label ID="DNo" runat="server" Font-Size="12pt" Style="z-index: 121; left: 504px;
                   position: absolute; top: 80px"></asp:Label>
               <asp:Label ID="DDate" runat="server" Font-Size="12pt" Style="z-index: 122; left: 192px;
                   position: absolute; top: 80px"></asp:Label>     &nbsp;
               <asp:CheckBox ID="DChkPhone" runat="server" Style="z-index: 124;
               left: 896px; position: absolute; top: 112px" Width="40px" Text="��" Enabled="False" />      
               <asp:CheckBox ID="DchkVisa" runat="server" Style="z-index: 125;
               left: 896px; position: absolute; top: 144px" Width="88px" Text="��" Enabled="False" />          &nbsp;
          <img id="IMG1" onclick="return IMG1_onclick()" src="images/BusinessTripSheet_05.jpg"
              style="z-index: 99; left: 24px; position: absolute; top: 8px; text-align: left" />
                <asp:CheckBox ID="DchkInsurance" runat="server" Style="z-index: 135; left: 184px;
                    position: absolute; top: 240px" Text="���H�P�N�A�H�Ƴ��N�O�I�һݭӸ�]ID�B�X�ͤ���^���ѥx�_�`�ȳ��O�I�ӿ�H�F�S�����Ѫ̱N�L�k��O"
                    Width="560px" />
                &nbsp;
          <asp:Label ID="DDivision3" runat="server" Font-Size="12pt" Style="z-index: 127; left: 184px;
              position: absolute; top: 208px"></asp:Label>
          <asp:Label ID="DDivisionCode3" runat="server" Font-Size="12pt" Style="z-index: 128;
              left: 360px; position: absolute; top: 208px"></asp:Label>
          <asp:Label ID="DPhoneNo" runat="server" Font-Size="12pt" Style="z-index: 129; left: 1064px;
              position: absolute; top: 112px"></asp:Label>
          <asp:Label ID="DAirTickets" runat="server" Font-Size="12pt" Style="z-index: 130;
              left: 992px; position: absolute; top: 208px"></asp:Label>
          <asp:Label ID="DQCNO1" runat="server" Font-Size="12pt" Style="z-index: 131; left: 1056px;
              position: absolute; top: 88px"></asp:Label>
          <asp:Label ID="DQCNO" runat="server" Font-Size="12pt" Style="z-index: 132; left: 896px;
              position: absolute; top: 88px"></asp:Label>        
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="Unique_ID"
                    Font-Size="9pt" PageSize="100" Style="z-index: 103; left: 48px; position: absolute;
                    top: 448px" Width="1136px">
                    <Columns>
                        <asp:BoundField DataField="Unique_ID" HeaderText="ID">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Type" HeaderText="���O">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Appoint" HeaderText="�w��">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SDate" HeaderText="�X�o/�J��">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EDate" HeaderText="��F/�h�� ">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Days" HeaderText="�Ѽ�">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Currency" HeaderText="���O" />
                        <asp:BoundField DataField="Money" HeaderText="���B">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="50px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="FlyInf" HeaderText="��Z��T">
                            <ItemStyle Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SFly" HeaderText="�_�I" />
                        <asp:BoundField DataField="EFly" HeaderText="���I" />
                        <asp:BoundField DataField="HotelInf" HeaderText="������T" />
                        <asp:BoundField DataField="Remark" HeaderText="�Ƶ�">
                            <HeaderStyle Height="20px" />
                            <ItemStyle Width="200px" />
                        </asp:BoundField>
                    </Columns>
                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                </asp:GridView>
        
</div>
 
    </form>
</body>
</html>
