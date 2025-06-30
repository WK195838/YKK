<%@ Page Language="vb" AutoEventWireup="false" Inherits="DASW_DISPOSALUpload01" aspCompat="True" CodeFile="DASW_DISPOSALUpload01.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <title>報廢處理申請表</title>
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
			
  function CopyNewColor(wformno)
{
        
    window.open('DASW_NOCopy.aspx?formno='+wformno,'','status=0,toolbar=0,width=700,height=650,top=10,resizable=yes');
   
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
//只能輸入小數點與整數			
 

function BlockNumber(e)
{

var key = window.event ? e.keyCode : e.which;

var keychar = String.fromCharCode(key);

reg = /[0-9]|\./;

return reg.test(keychar);

}

//只能輸入整數			
 

    
    //千分位
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



// 不直接引用 html id，因為伺服器控制項對應的是 ClientID，而ClientID與控制項層次有關，不利程式碼移植
// 因此盡可能選擇直接傳遞物件，通過 DOM 獲取相關的父控制項和子控制項。
function CBL_SingleChoice(sender) 
{
    var container = sender.parentNode;        
    if(container.tagName.toUpperCase() == "TD") 
    { // table 布局，否則為span布局
        container = container.parentNode.parentNode; // 層次: <table><tr><td><input />
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
  	 
  
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/Disposalupload_01.jpg" Style="z-index: 100;
               left: 30px; position: absolute; top: 39px" />
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
           <asp:TextBox ID="DAPPDepo" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 101; left: 368px; position: absolute;
               top: 160px" Width="136px"></asp:TextBox>
           <asp:TextBox ID="DAPPDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 102; left: 138px; position: absolute;
               top: 126px" Width="136px"></asp:TextBox>
           &nbsp;
           &nbsp;&nbsp; &nbsp;&nbsp;<br />
           <br />
           <br />
           &nbsp; &nbsp; &nbsp; &nbsp;
           <br />
           <asp:FileUpload ID="DFileUpload1" runat="server" style="z-index: 103; left: 138px; position: absolute; top: 193px" BackColor="Yellow" Height="27px" Width="365px" />
           &nbsp;<br />
           &nbsp;
           <asp:TextBox ID="DAppName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 104; left: 138px; position: absolute;
               top: 160px" Width="136px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp;
           <asp:GridView ID="GridView1" runat="server" BorderColor="#CC9966" BorderStyle="None"
               BorderWidth="1px" CellPadding="4" Font-Size="9pt" ShowFooter="True" Style="z-index: 105;
               left: 32px; position: absolute; top: 600px" Width="790px">
               <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
               <HeaderStyle BackColor="#CC6600" ForeColor="#FFFFCC" HorizontalAlign="Center" VerticalAlign="Middle" />
           </asp:GridView>
           &nbsp;
           <asp:RadioButtonList ID="rbHDR" runat="server" Style="z-index: 106; left: 984px;
               position: absolute; top: 42px" Visible="False">
               <asp:ListItem Selected="True" Text="Yes" Value="Yes"></asp:ListItem>
               <asp:ListItem Text="No" Value="No"></asp:ListItem>
           </asp:RadioButtonList>
           &nbsp;
           <asp:Button ID="DUpload" runat="server" Style="z-index: 107; left: 405px; position: absolute;
               top: 245px" Text="預覽" Width="47px" />
           <asp:Button ID="DInsert" runat="server" Enabled="False" Style="z-index: 108; left: 466px;
               position: absolute; top: 245px" Text="上傳" Width="47px" />
           <asp:GridView ID="GridView2" runat="server" Style="z-index: 109; left: 32px; position: absolute;
               top: 760px" Visible="False">
           </asp:GridView>
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 110; left: 984px; position: absolute; top: 110px"
               Visible="False" Width="120px"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DSMONTH" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 111; left: 371px; position: absolute;
               top: 128px" Width="136px"></asp:TextBox>
           <asp:TextBox ID="DNO1" runat="server" BackColor="Yellow" Style="z-index: 113; left: 552px;
               position: absolute; top: 126px"></asp:TextBox>
           <input id="BCopySheet" runat="server" style="z-index: 156; left: 712px; width: 24px;
               position: absolute; top: 128px" type="button" value="..." visible="true" />
           <asp:TextBox ID="DSNO" runat="server" BackColor="Yellow" Style="z-index: 113; left: 552px;
               position: absolute; top: 160px" Width="64px"></asp:TextBox>
       </div>
    </form>
</body>
</html> 
