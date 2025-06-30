<%@ Page Language="vb" AutoEventWireup="false" Inherits="DASW_DISPOSALSheet03" aspCompat="True" CodeFile="DASW_DISPOSALSheet03.aspx.vb"  enableEventValidation="false" %>
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
  	 
  
           &nbsp; &nbsp;
        <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/Disposal_04.jpg" style="z-index: 100; left: 34px; position: absolute; top: 18px" />
           <asp:DropDownList ID="DSTYPE" runat="server" AutoPostBack="True" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 104; left: 712px; position: absolute;
               top: 136px" Width="96px">
           </asp:DropDownList>
           <asp:TextBox ID="DDepoName" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" ReadOnly="True" Style="z-index: 101; left: 245px;
               position: absolute; top: 135px" Width="204px"></asp:TextBox>
           &nbsp;&nbsp;
           <asp:DropDownList ID="DDUTYDEPO" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 102; left: 249px; position: absolute; top: 185px"
               Width="105px" EnableTheming="True" AutoPostBack="True">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISPOSALRULE" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 103; left: 249px; position: absolute; top: 211px"
               Width="199px" AutoPostBack="True">
           </asp:DropDownList>
           <asp:DropDownList ID="DPLACE" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="20px" Style="z-index: 104; left: 584px; position: absolute; top: 187px"
               Width="221px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISPOSALREASON" runat="server" BackColor="Yellow"
               ForeColor="Blue" Height="20px" Style="z-index: 105; left: 249px; position: absolute;
               top: 159px" Width="199px">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Button ID="DSUM" runat="server" BackColor="#80FFFF" Height="72px" Style="z-index: 106;
               left: 645px; position: absolute; top: 755px" Text="計算合計" Width="164px" EnableViewState="False" CausesValidation="False" />
           &nbsp;&nbsp;
           <input id="DDISPOSALFILE1" runat="server" name="DISPOSALFILE1" style="z-index: 293;
               left: 984px; width: 270px; position: absolute; top: 709px; height: 26px; background-color: #ffff00"
               type="file" visible="false" />
           <asp:HyperLink ID="LDISPOSALFILE" runat="server" Font-Size="12pt" Height="8px"
               Style="z-index: 107; left: 249px; position: absolute; top: 910px" Target="_blank"
               Width="71px">報廢檔案</asp:HyperLink>
           &nbsp;&nbsp;&nbsp;&nbsp;
           <asp:TextBox ID="TextBox18" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 108; left: 55px;
               text-transform: uppercase; position: absolute; top: 493px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 109; left: 56px;
               text-transform: uppercase; position: absolute; top: 494px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 110; left: 57px;
               text-transform: uppercase; position: absolute; top: 701px" Width="78px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 111; left: 142px;
               text-transform: uppercase; position: absolute; top: 701px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE9" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 112; left: 250px;
               text-transform: uppercase; position: absolute; top: 701px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER9" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 113; left: 334px;
               text-transform: uppercase; position: absolute; top: 701px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD9" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 114; left: 417px;
               text-transform: uppercase; position: absolute; top: 700px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG9" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 115; left: 501px;
               text-transform: uppercase; position: absolute; top: 701px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE9" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 116; left: 563px;
               text-transform: uppercase; position: absolute; top: 701px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 117; left: 729px;
               text-transform: uppercase; position: absolute; top: 701px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 118; left: 729px;
               text-transform: uppercase; position: absolute; top: 728px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 119; left: 57px;
               text-transform: uppercase; position: absolute; top: 727px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 120; left: 142px;
               text-transform: uppercase; position: absolute; top: 727px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE10" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 121; left: 250px;
               text-transform: uppercase; position: absolute; top: 727px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DPIECE11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 122; left: 250px;
               text-transform: uppercase; position: absolute; top: 754px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DPIECE12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 123; left: 251px;
               text-transform: uppercase; position: absolute; top: 780px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DPIECE13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 124; left: 251px;
               text-transform: uppercase; position: absolute; top: 806px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DSIGN1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 125; left: 58px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 126; left: 58px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 127; left: 58px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 128; left: 58px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 129; left: 187px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 130; left: 309px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 131; left: 435px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 132; left: 562px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE9" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 133; left: 686px;
               text-transform: uppercase; position: absolute; top: 1014px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 134; left: 686px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 135; left: 686px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 136; left: 686px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN10" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 137; left: 562px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 138; left: 562px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 139; left: 562px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 140; left: 435px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 141; left: 435px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 142; left: 435px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 143; left: 309px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 144; left: 309px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGN8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 145; left: 309px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 146; left: 187px;
               text-transform: uppercase; position: absolute; top: 1039px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 147; left: 187px;
               text-transform: uppercase; position: absolute; top: 1065px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DSIGNDATE4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 148; left: 187px;
               text-transform: uppercase; position: absolute; top: 1091px" Width="116px"></asp:TextBox>
           <asp:TextBox ID="DMETER11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 149; left: 335px;
               text-transform: uppercase; position: absolute; top: 754px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 150; left: 417px;
               text-transform: uppercase; position: absolute; top: 753px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 151; left: 501px;
               text-transform: uppercase; position: absolute; top: 753px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DKG12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 152; left: 502px;
               text-transform: uppercase; position: absolute; top: 779px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DKG13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 153; left: 502px;
               text-transform: uppercase; position: absolute; top: 805px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DYARD12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 154; left: 418px;
               text-transform: uppercase; position: absolute; top: 779px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 155; left: 418px;
               text-transform: uppercase; position: absolute; top: 805px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 156; left: 336px;
               text-transform: uppercase; position: absolute; top: 780px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 157; left: 336px;
               text-transform: uppercase; position: absolute; top: 806px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER10" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 158; left: 334px;
               text-transform: uppercase; position: absolute; top: 727px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD10" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 159; left: 417px;
               text-transform: uppercase; position: absolute; top: 726px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG10" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 160; left: 501px;
               text-transform: uppercase; position: absolute; top: 727px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE10" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 161; left: 563px;
               text-transform: uppercase; position: absolute; top: 727px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DPRICE11" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 162; left: 561px;
               text-transform: uppercase; position: absolute; top: 756px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DPRICE12" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 163; left: 561px;
               text-transform: uppercase; position: absolute; top: 780px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DPRICE13" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 164; left: 561px;
               text-transform: uppercase; position: absolute; top: 805px" Width="80px"></asp:TextBox>
           &nbsp;
           <asp:DropDownList ID="DDISRULE10" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 165; left: 649px; position: absolute;
            top: 729px" Width="75px">
           </asp:DropDownList><asp:DropDownList ID="DPCNAME" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 166; left: 707px; position: absolute;
            top: 936px" Width="106px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISRULE9" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 167; left: 649px; position: absolute;
            top: 703px" Width="75px">
           </asp:DropDownList>
           <asp:TextBox ID="DITEMCODE7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 168; left: 57px;
               text-transform: uppercase; position: absolute; top: 650px" Width="78px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 169; left: 142px;
               text-transform: uppercase; position: absolute; top: 650px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE7" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 170; left: 250px;
               text-transform: uppercase; position: absolute; top: 650px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER7" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 171; left: 334px;
               text-transform: uppercase; position: absolute; top: 650px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD7" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 172; left: 417px;
               text-transform: uppercase; position: absolute; top: 649px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG7" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 173; left: 501px;
               text-transform: uppercase; position: absolute; top: 650px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE7" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 174; left: 562px;
               text-transform: uppercase; position: absolute; top: 650px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK7" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 175; left: 729px;
               text-transform: uppercase; position: absolute; top: 650px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 176; left: 57px;
               text-transform: uppercase; position: absolute; top: 676px" Width="78px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 177; left: 142px;
               text-transform: uppercase; position: absolute; top: 676px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE8" runat="server"  onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 178; left: 250px;
               text-transform: uppercase; position: absolute; top: 676px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER8" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 179; left: 334px;
               text-transform: uppercase; position: absolute; top: 676px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD8" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 180; left: 417px;
               text-transform: uppercase; position: absolute; top: 675px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG8" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 181; left: 501px;
               text-transform: uppercase; position: absolute; top: 676px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE8" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 182; left: 562px;
               text-transform: uppercase; position: absolute; top: 676px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK8" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 183; left: 729px;
               text-transform: uppercase; position: absolute; top: 676px" Width="80px"></asp:TextBox>
           <asp:DropDownList ID="DDISRULE8" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 184; left: 649px; position: absolute;
            top: 678px" Width="75px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISRULE7" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 185; left: 649px; position: absolute;
            top: 652px" Width="75px">
           </asp:DropDownList>
           <asp:TextBox ID="DITEMCODE3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 186; left: 56px;
               text-transform: uppercase; position: absolute; top: 546px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 187; left: 56px;
               text-transform: uppercase; position: absolute; top: 598px" Width="78px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 188; left: 141px;
               text-transform: uppercase; position: absolute; top: 598px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE5" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 189; left: 249px;
               text-transform: uppercase; position: absolute; top: 598px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER5" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 190; left: 333px;
               text-transform: uppercase; position: absolute; top: 598px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD5" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 191; left: 416px;
               text-transform: uppercase; position: absolute; top: 598px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG5" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 192; left: 500px;
               text-transform: uppercase; position: absolute; top: 598px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE5" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 193; left: 562px;
               text-transform: uppercase; position: absolute; top: 598px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK5" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 194; left: 729px;
               text-transform: uppercase; position: absolute; top: 598px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 195; left: 56px;
               text-transform: uppercase; position: absolute; top: 624px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 196; left: 141px;
               text-transform: uppercase; position: absolute; top: 624px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE6" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 197; left: 249px;
               text-transform: uppercase; position: absolute; top: 624px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER6" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 198; left: 333px;
               text-transform: uppercase; position: absolute; top: 624px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD6" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 199; left: 416px;
               text-transform: uppercase; position: absolute; top: 624px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG6" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 200; left: 500px;
               text-transform: uppercase; position: absolute; top: 625px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE6" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 201; left: 562px;
               text-transform: uppercase; position: absolute; top: 625px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK6" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 202; left: 729px;
               text-transform: uppercase; position: absolute; top: 625px" Width="80px"></asp:TextBox>
           <asp:DropDownList ID="DDISRULE6" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 203; left: 648px; position: absolute;
            top: 627px" Width="75px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISRULE5" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 204; left: 648px; position: absolute;
            top: 601px" Width="75px">
           </asp:DropDownList>
           <asp:TextBox ID="DITEMNAME3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 205; left: 141px;
               text-transform: uppercase; position: absolute; top: 546px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE3" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 206; left: 249px;
               text-transform: uppercase; position: absolute; top: 546px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER3" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 207; left: 333px;
               text-transform: uppercase; position: absolute; top: 546px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD3" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 208; left: 416px;
               text-transform: uppercase; position: absolute; top: 545px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG3" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 209; left: 500px;
               text-transform: uppercase; position: absolute; top: 546px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE3" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 210; left: 562px;
               text-transform: uppercase; position: absolute; top: 546px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK3" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 211; left: 729px;
               text-transform: uppercase; position: absolute; top: 546px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 212; left: 56px;
               text-transform: uppercase; position: absolute; top: 572px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 213; left: 141px;
               text-transform: uppercase; position: absolute; top: 572px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE4" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 214; left: 249px;
               text-transform: uppercase; position: absolute; top: 572px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER4" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 215; left: 333px;
               text-transform: uppercase; position: absolute; top: 572px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD4" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 216; left: 416px;
               text-transform: uppercase; position: absolute; top: 571px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG4" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 217; left: 500px;
               text-transform: uppercase; position: absolute; top: 572px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE4" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 218; left: 562px;
               text-transform: uppercase; position: absolute; top: 572px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK4" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 219; left: 729px;
               text-transform: uppercase; position: absolute; top: 572px" Width="80px"></asp:TextBox>
           <asp:DropDownList ID="DDISRULE4" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 220; left: 648px; position: absolute;
            top: 574px" Width="75px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISRULE3" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 221; left: 648px; position: absolute;
            top: 548px" Width="75px">
           </asp:DropDownList>
           &nbsp; &nbsp;
           &nbsp; &nbsp;
           <asp:TextBox ID="TextBox16" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 222; left: 416px;
               text-transform: uppercase; position: absolute; top: 493px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="TextBox17" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 223; left: 500px;
               text-transform: uppercase; position: absolute; top: 494px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="TextBox26" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 224; left: 562px;
               text-transform: uppercase; position: absolute; top: 494px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="TextBox27" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 225; left: 729px;
               text-transform: uppercase; position: absolute; top: 494px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="TextBox28" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 226; left: 56px;
               text-transform: uppercase; position: absolute; top: 520px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="TextBox29" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 227; left: 141px;
               text-transform: uppercase; position: absolute; top: 520px" Width="99px"></asp:TextBox>
           <asp:TextBox ID="TextBox30" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 228; left: 249px;
               text-transform: uppercase; position: absolute; top: 520px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="TextBox31" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 229; left: 333px;
               text-transform: uppercase; position: absolute; top: 520px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="TextBox32" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 230; left: 416px;
               text-transform: uppercase; position: absolute; top: 519px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="TextBox33" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 231; left: 500px;
               text-transform: uppercase; position: absolute; top: 520px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="TextBox34" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 232; left: 562px;
               text-transform: uppercase; position: absolute; top: 520px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="TextBox35" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="5" Style="z-index: 233; left: 729px;
               text-transform: uppercase; position: absolute; top: 520px" Width="80px"></asp:TextBox>
           <asp:DropDownList ID="DropDownList6" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 234; left: 648px; position: absolute;
            top: 522px" Width="75px" AutoPostBack="True">
           </asp:DropDownList>
           <asp:DropDownList ID="DropDownList8" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 235; left: 648px; position: absolute;
            top: 496px" Width="75px" AutoPostBack="True">
           </asp:DropDownList>
           <asp:TextBox ID="DITEMNAME1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 236; left: 141px;
               text-transform: uppercase; position: absolute; top: 494px" Width="99px" AutoPostBack="True"></asp:TextBox>
           &nbsp;
           <asp:TextBox ID="DMETER1" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 237; left: 334px;
               text-transform: uppercase; position: absolute; top: 495px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD1" runat="server" onKeypress="return BlockNumber(event);" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 238; left: 416px;
               text-transform: uppercase; position: absolute; top: 493px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG1" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 239; left: 500px;
               text-transform: uppercase; position: absolute; top: 494px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE1" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 240; left: 562px;
               text-transform: uppercase; position: absolute; top: 494px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 241; left: 729px;
               text-transform: uppercase; position: absolute; top: 494px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DITEMCODE2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="15" Style="z-index: 242; left: 56px;
               text-transform: uppercase; position: absolute; top: 520px" Width="78px"></asp:TextBox>
           <asp:TextBox ID="DITEMNAME2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="30" Style="z-index: 243; left: 141px;
               text-transform: uppercase; position: absolute; top: 520px" Width="99px" AutoPostBack="True"></asp:TextBox>
           <asp:TextBox ID="DPIECE2" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 244; left: 249px;
               text-transform: uppercase; position: absolute; top: 520px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DPIECE1" runat="server" onKeypress="return BlockNumber(event);"   BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 245; left: 250px; 
               text-transform: uppercase; position: absolute; top: 493px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DMETER2" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 246; left: 333px;
               text-transform: uppercase; position: absolute; top: 520px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DYARD2" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 247; left: 416px;
               text-transform: uppercase; position: absolute; top: 519px" Width="65px"></asp:TextBox>
           <asp:TextBox ID="DKG2" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 248; left: 500px;
               text-transform: uppercase; position: absolute; top: 520px" Width="37px"></asp:TextBox>
           <asp:TextBox ID="DPRICE2" runat="server" onKeypress="return BlockNumber(event);"  BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" MaxLength="10" Style="z-index: 249; left: 562px;
               text-transform: uppercase; position: absolute; top: 520px" Width="80px"></asp:TextBox>
           <asp:TextBox ID="DREMARK2" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 250; left: 729px;
               text-transform: uppercase; position: absolute; top: 520px" Width="80px"></asp:TextBox>
           <asp:DropDownList ID="DDISRULE2" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 251; left: 648px; position: absolute;
            top: 522px" Width="75px">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISRULE1" runat="server" BackColor="Yellow"
            ForeColor="Blue" Height="20px" Style="z-index: 252; left: 648px; position: absolute;
            top: 496px" Width="75px">
           </asp:DropDownList>
           <asp:TextBox ID="DCHINESEREASON" runat="server" BackColor="Yellow" ForeColor="Blue" Height="69px"
               Style="z-index: 253; left: 101px; position: absolute; top: 276px" TextMode="MultiLine"
               Width="707px"></asp:TextBox>
           &nbsp;
           <asp:CheckBox ID="DCUSTOMERTOLL" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 255;
               left: 154px; position: absolute; top: 77px" Text="向客戶取款" Width="103px" Font-Bold="True" />
           &nbsp;
           <asp:CheckBox ID="DMKTSIGN" runat="server" Checked="True" Font-Size="10pt" Style="z-index: 256;
               left: 52px; position: absolute; top: 77px" Text="需營業簽核" Width="93px" Font-Bold="True" />
           &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DDISPOSALREASON1" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 257; left: 456px; position: absolute;
               top: 160px" Width="346px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <input id="BSAVE" runat="server" name="Button1" onclick="Button('SAVE');" style="z-index: 292;
            left: 728px; width: 80px; position: absolute; top: 1160px; height: 28px" type="button"
            value="儲存" />
        <input id="BNG2" runat="server" name="Button1" onclick="Button('NG2');" style="z-index: 291;
            left: 568px; width: 80px; position: absolute; top: 1160px; height: 28px" type="button"
            value="NG2" />
        <input id="BNG1" runat="server" name="Button1" onclick="Button('NG1');" style="z-index: 290;
            left: 488px; width: 80px; position: absolute; top: 1160px; height: 28px" type="button"
            value="NG1" />
        &nbsp;&nbsp;
        <input id="BOK" runat="server" name="BOK" onclick="Button('OK');" style="z-index: 289;
            left: 648px; width: 80px; position: absolute; top: 1160px; height: 28px" type="button"
            value="OK"  />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp;
           <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 258; left: 464px; position: absolute;
               top: 136px" Width="112px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:TextBox ID="DAppName" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 259; left: 458px; position: absolute; top: 107px"
               Width="119px"></asp:TextBox>
           <asp:TextBox ID="DSignDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
               Font-Bold="False" ForeColor="Blue" Height="24px" Style="z-index: 260; left: 711px;
               position: absolute; top: 107px" Width="96px"></asp:TextBox>
        
          
           <asp:TextBox ID="DNo" runat="server" BackColor="Yellow" BorderStyle="Groove" ForeColor="Blue"
               Height="24px" Style="z-index: 261; left: 245px; position: absolute; top: 107px"
               Width="120px"></asp:TextBox>
           <asp:TextBox ID="DSales" runat="server" BackColor="Yellow" BorderStyle="Groove"
               ForeColor="Blue" Height="24px" Style="z-index: 262; left: 357px; position: absolute;
               top: 187px" Width="92px"></asp:TextBox>
           <asp:TextBox ID="DDUTYDEPOName" runat="server" BackColor="White" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 263; left: 152px; position: absolute;
               top: 44px" Width="77px"></asp:TextBox>
           <asp:TextBox ID="DDUTYDEPOID" runat="server" BackColor="White" BorderStyle="None"
               ForeColor="White" Height="24px" Style="z-index: 264; left: 62px; position: absolute;
               top: 45px" Width="77px"></asp:TextBox>
           <asp:TextBox ID="DDisposalYM" runat="server" BackColor="White" BorderStyle="None"
               Font-Bold="False" ForeColor="White" Height="24px" Style="z-index: 265; left: 904px;
               position: absolute; top: 176px" Width="129px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;
           &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
           &nbsp; 
           &nbsp;<br />
           &nbsp;
           &nbsp;
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
           <asp:TextBox ID="DJAPANREASON" runat="server" BackColor="Yellow" ForeColor="Blue"
               Height="69px" Style="z-index: 266; left: 99px; position: absolute; top: 356px"
               TextMode="MultiLine" Width="707px"></asp:TextBox>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp;&nbsp;<br />
        <br />
           <input id="DDISPOSALTYPE" runat="server" name="DDeliveryDate" style="z-index: 285;
               left: 582px; width: 196px; color: blue; border-top-style: groove; border-right-style: groove;
               border-left-style: groove; position: absolute; top: 211px; background-color: yellow;
               border-bottom-style: groove" type="text" />
           <asp:Button ID="BDISPOSALTYPE" runat="server" CausesValidation="False" Height="25px"
               Style="z-index: 272; left: 779px; position: absolute; top: 210px" Text="....."
               Width="28px" />
        <br />
           &nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
           &nbsp;&nbsp;<br />
                  
               
               
        <br />
           &nbsp;&nbsp;&nbsp;<br />
        <br />
           <asp:Button ID="BReason" runat="server" BackColor="#C0FFFF" Font-Bold="True" Font-Size="11pt"
               ForeColor="Fuchsia" Style="z-index: 273; left: 667px; position: absolute; top: 241px"
               Text="報廢原因參考" CausesValidation="False" />
           &nbsp;<br />
           &nbsp;
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
           &nbsp; &nbsp;
           &nbsp;&nbsp;
           <asp:Button ID="DDISPOSALFILE2" runat="server" Style="z-index: 274; left: 250px; position: absolute;
               top: 934px"  Text="開啟實物報廢附檔資料夾" Width="261px" CausesValidation="False" />
           &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBoxList ID="CheckASDate" runat="server" RepeatDirection="Horizontal" style="z-index: 275; left: 634px; position: absolute; top: 873px">
            <asp:ListItem Value="0" Text="不需要立會"></asp:ListItem>
            <asp:ListItem Value="1" Text="需立會"></asp:ListItem>
  
        </asp:CheckBoxList>
           <input id="DASDate" runat="server" name="DASDate" style="z-index: 286;
               left: 706px; width: 81px; color: blue; border-top-style: groove; border-right-style: groove;
               border-left-style: groove; position: absolute; top: 908px; background-color: yellow;
               border-bottom-style: groove" type="text" />
           <input id="BASDate" runat="server" style="z-index: 287; left: 788px; width: 24px; position: absolute;
               top: 908px" type="button" value="..." />
           <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
               ShowSummary="false" Style="z-index: 276; left: 56px; position: absolute; top: 1120px" />
           <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator"
               Style="z-index: 277; left: 214px; position: absolute; top: 1124px"></asp:CustomValidator>
           &nbsp;
           <asp:FileUpload ID="DFileUpload1" runat="server" BackColor="Yellow" Height="22px"
               Style="z-index: 278; left: 247px; position: absolute; top: 908px" Width="266px" />
           <asp:Button ID="DUpload" runat="server" Style="z-index: 279; left: 981px; position: absolute;
               top: 753px" Text="預覽" Width="47px" CausesValidation="False" Visible="False" />
           <asp:Button ID="DInsert" runat="server" Enabled="False" Style="z-index: 280; left: 1039px;
               position: absolute; top: 754px" Text="上傳" Width="47px" UseSubmitBehavior="False" Visible="False" />
           <asp:RadioButtonList ID="rbHDR" runat="server" Style="z-index: 281; left: 975px;
               position: absolute; top: 629px" Visible="False">
               <asp:ListItem Selected="True" Text="Yes" Value="Yes"></asp:ListItem>
               <asp:ListItem Text="No" Value="No"></asp:ListItem>
           </asp:RadioButtonList>
           <asp:GridView ID="GridView1" runat="server" Style="z-index: 282; left: 988px; position: absolute;
               top: 820px">
           </asp:GridView>
           &nbsp;&nbsp;
        
        </div>
    </form>
</body>
</html> 
