<%@ Page Language="vb" AutoEventWireup="false" Inherits="DASW_DISPOSALReason" aspCompat="True" CodeFile="DASW_DISPOSALReason.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>���o�B�z�ӽЪ�</title>
    	<script language="javascript" type="text/javascript">
			var wPop;
			//--Calendar------------------------------------
	function calendarPicker(strField)
			{
				window.open('DatePicker1.aspx?field=' + strField,'calendarPopup','width=250,height=190,resizable=yes');
			}

	function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}
			
			    	
    				
            function GetDISPOSALTYPE()
            {    
            window.open('DISPOSALTYPE.aspx?','','status=0,toolbar=0,width=300,height=270,top=10,resizable=yes');
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
//�u���J�p���I�P���			
 

function BlockNumber(e)
{

var key = window.event ? e.keyCode : e.which;

var keychar = String.fromCharCode(key);

reg = /[0-9]|\./;

return reg.test(keychar);

}

//�u���J���			
 

    
    //�d����
    function trans_amt(Name) {
var amt = document.getElementById(Name).value;
var amt_length = amt;
amt = amt.replace(/\,/g, "");
re = /(\d{1,3})(?=(\d{3})+(?:$|\D))/g;
n1 = amt.replace(re, "$1,");
document.getElementById(Name).value = n1;
//var sel = document.selection.createRange();
//sel.moveStart('character', amt_length.length);
//sel.collapse(true);
//sel.select();
}



// �������ޥ� html id�A�]�����A������������O ClientID�A��ClientID�P����h�������A���Q�{���X����
// �]���ɥi���ܪ����ǻ�����A�q�L DOM ���������������M�l����C
function CBL_SingleChoice(sender) 
{
    var container = sender.parentNode;        
    if(container.tagName.toUpperCase() == "TD") 
    { // table �����A�_�h��span����
        container = container.parentNode.parentNode; // �h��: <table><tr><td><input />
    }        
    var chkList = container.getElementsByTagName("input");
    var senderState = sender.checked;
    for(var i = 0; i < chkList.length; i++) 
    {
        chkList[i].checked = false;
    }     
    sender.checked = senderState;          
}



		</script>
		
 
  
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
  	 
  
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp; &nbsp;&nbsp;<br />
           <br />
           <br />
           &nbsp; &nbsp; &nbsp; &nbsp;
           <br />
           <asp:FileUpload ID="DFileUpload1" runat="server" style="z-index: 100; left: 156px; position: absolute; top: 33px" BackColor="Yellow" Height="27px" Width="365px" />
           &nbsp;&nbsp;<br />
           &nbsp;
           &nbsp; &nbsp; &nbsp;
           <asp:GridView ID="GridView1" runat="server" BorderColor="#CC9966" BorderStyle="None"
               BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True" Style="z-index: 101;
               left: 43px; position: absolute; top: 112px" Width="790px">
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
           &nbsp;
           <asp:RadioButtonList ID="rbHDR" runat="server" Style="z-index: 102; left: 984px;
               position: absolute; top: 42px" Visible="False">
               <asp:ListItem Selected="True" Text="Yes" Value="Yes"></asp:ListItem>
               <asp:ListItem Text="No" Value="No"></asp:ListItem>
           </asp:RadioButtonList>
           &nbsp;
           <asp:Button ID="DUpload" runat="server" Style="z-index: 103; left: 53px; position: absolute;
               top: 71px" Text="�w��" Width="47px" />
           <asp:Button ID="DInsert" runat="server" Enabled="False" Style="z-index: 104; left: 114px;
               position: absolute; top: 71px" Text="�W��" Width="47px" />
           <asp:GridView ID="GridView2" runat="server" Style="z-index: 105; left: 44px; position: absolute;
               top: 591px" Visible="False">
           </asp:GridView>
           &nbsp;&nbsp;&nbsp;
           <asp:Label ID="Label1" runat="server" Style="z-index: 107; left: 49px; position: absolute;
               top: 37px" Text="���o��]�ɮ�"></asp:Label>
       </div>
    </form>
</body>
</html> 
