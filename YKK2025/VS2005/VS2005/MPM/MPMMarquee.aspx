<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MPMMarquee.aspx.vb" Inherits="MPMMarquee" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>工作記錄</title>
      <style type="text/css">
    </style>

    <script type="text/javascript">
	function calendarPicker(field1)
			{
				window.open('DatePicker.aspx?field1=' + field1,'calendarPopup','width=250,height=190,resizable=yes');
			}

 

    </script>
</head>
<body>
    <form id="Form1" runat="server">
    <div>
    
    <script type="text/javascript" src="jquery.js"></script> 
<script type="text/javascript" src="jquery.colorpicker.js"></script> &nbsp;
     
       <img src="Images/colorpicker.png" id="cp4" style="cursor:pointer; z-index: 113; left: 438px; position: absolute; top: 288px;"/>
          
                 <script type="text/javascript">
        


$("#cp4").colorpicker(

{
    
    
   fillcolor:true,
    success:function(o,color){
      
          $("#DFontColor").css("color",color);
         document.getElementById('DColor').value = color 
    }
      
 

}

)




  

  
  
   </script>

        &nbsp;
  <input id="DStartDate" runat="server" style="z-index: 114; left: 112px; width: 93px;
            color: black; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 249px; background-color: yellow; border-bottom-style: groove"
            type="text" /><input id="DEndDate" runat="server" style="z-index: 114; left: 331px; width: 89px;
            color: black; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 250px; background-color: yellow; border-bottom-style: groove"
            type="text" />
        <input id="DFontColor" runat="server" style="z-index: 114; left: 330px; width: 100px;
            color: black; border-top-style: groove; border-right-style: groove; border-left-style: groove;
            position: absolute; top: 285px; background-color: black; border-bottom-style: groove"
            type="text" value="文字的顏色" />
        &nbsp; &nbsp;
            
  
            &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;<asp:Image ID="Image1" runat="server" ImageUrl="~/Images/F_MPMMarquee1.jpg" style="z-index: 100; left: 13px; position: absolute; top: 16px" />
        <input id="BStartDate" runat="server" style="z-index: 144; left: 212px; width: 24px;
            position: absolute; top: 248px" type="button" value="..." />
        <input id="BEndDate" runat="server" style="z-index: 144; left: 429px; width: 24px;
            position: absolute; top: 247px" type="button" value="..." />
        <asp:DropDownList ID="DAction" runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px"
            Style="z-index: 101; left: 110px; position: absolute; top: 76px" Width="125px">
            <asp:ListItem Value="1">啟用</asp:ListItem>
            <asp:ListItem Value="0">關閉</asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DFontSize" runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px"
            Style="z-index: 102; left: 224px; position: absolute; top: 284px" Width="47px">
            <asp:ListItem>1</asp:ListItem>
            <asp:ListItem>2</asp:ListItem>
            <asp:ListItem>3</asp:ListItem>
            <asp:ListItem>4</asp:ListItem>
            <asp:ListItem>5</asp:ListItem>
            <asp:ListItem>6</asp:ListItem>
            <asp:ListItem>7</asp:ListItem>
        </asp:DropDownList><asp:DropDownList ID="DFontStyle" runat="server" BackColor="Yellow" ForeColor="Blue" Height="25px"
            Style="z-index: 102; left: 113px; position: absolute; top: 283px" Width="108px">
        </asp:DropDownList>
        &nbsp; &nbsp;&nbsp;
        <asp:DropDownList ID="DDep_Name" runat="server" BackColor="Silver" ForeColor="Black" Height="25px"
            Style="z-index: 103; left: 327px; position: absolute; top: 78px" Width="122px" Enabled="False">
        </asp:DropDownList>
        <asp:TextBox ID="DSubject" runat="server" BackColor="Yellow" BorderStyle="Groove"
            Font-Size="Smaller" ForeColor="Blue" Height="128px" MaxLength="240" Style="z-index: 104;
            left: 108px; position: absolute; top: 105px" TextMode="MultiLine" Width="340px"></asp:TextBox>
        &nbsp;
        <asp:Button ID="BSave" runat="server" Text="儲存" style="Z-INDEX: 105; LEFT: 355px; POSITION: absolute; TOP: 346px" Width="100px" /><asp:Button ID="BShow" runat="server" Text="效果展示" style="Z-INDEX: 105; LEFT: 249px; POSITION: absolute; TOP: 346px" Width="100px" /><asp:Button ID="BBack" runat="server" Text="回跑馬燈記錄一覽" style="Z-INDEX: 106; LEFT: 111px; POSITION: absolute; TOP: 346px" Width="134px" />
        &nbsp;
        <asp:TextBox ID="DAppUser" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 107; left: 111px; position: absolute;
            top: 314px" Width="119px"></asp:TextBox>
        &nbsp; &nbsp;
        &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DAppDate" runat="server" BackColor="Yellow" BorderStyle="Groove"
            ForeColor="Blue" Height="20px" Style="z-index: 108; left: 329px; position: absolute;
            top: 314px" Width="119px"></asp:TextBox>
        &nbsp; &nbsp;&nbsp;
        <asp:TextBox ID="DColor" runat="server" BackColor="White" BorderStyle="None" ForeColor="White"
            Height="20px" Style="z-index: 109; left: 6px; position: absolute; top: 511px"
            Width="92px"></asp:TextBox>
        </div>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <asp:Literal ID="obj" runat="server"></asp:Literal><br />
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
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           
    </form>
</body>
</html>
