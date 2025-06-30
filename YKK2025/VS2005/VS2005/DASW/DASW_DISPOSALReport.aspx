<%@ Page Language="vb" AutoEventWireup="false" Inherits="DASW_DISPOSALReport" aspCompat="True" EnableEventValidation = "false"  CodeFile="DASW_DISPOSALReport.aspx.vb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>報廢處理申請月總表</title>
 <script type="text/javascript">  
function SelectAllCheckboxesYES(spanChk)  
{  
    
     elm=document.forms[0];  
  
    for(i=0;i<elm.length;i++)  
        {          
            
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)  
            {  
         
            if(elm.elements[i].id.match('NODETAIL')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }           
              
                       
              if(elm.elements[i].id.match('YESDETAIL')!=null)
              {
              
               if(elm.elements[i].checked!=spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }     
              
         
             if(elm.elements[i].id.match('CheckNOHEAD')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }  
              
       
                       
          
                
           
                      
            }  
      

    }  
}  

function SelectAllCheckboxesNO(spanChk)  
{  


    elm=document.forms[0]; 
  
    for(i=0;i<elm.length;i++)  
        {         
        
          
       
                
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)           
            {  
            
             
             
            if(elm.elements[i].id.match('YESDETAIL')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }                 
                
                
              if(elm.elements[i].id.match('NODETAIL')!=null)
              {
               if(elm.elements[i].checked!=spanChk.checked)               
                 elm.elements[i].click();                     
                      
              
              }       
             
                
                 if(elm.elements[i].id.match('CheckYESHEAD')!=null)
              {
              
               if(elm.elements[i].checked==spanChk.checked)  
                 elm.elements[i].click();                                       
              
              }  
               
                
           
                      
            }  
      

    }  
}  


function CheckboxesYES(spanChk)  
{
   elm=document.forms[0]; 
  
    for(i=0;i<elm.length;i++)  
        {         
                
    
                
         if(elm[i].type=="checkbox" && elm[i].id!=spanChk.id)           
            {  
            
             
              if(elm.elements[i].id.match('NODETAIL')!=null)
              {
               if(elm.elements[i].checked==spanChk.checked)               
                 elm.elements[i].click();                     
                     
              
              }       
           
                
                       
                      
            }  
      

    }  
}

 function CheckboxesYES(spanChk)  
{
  
    var row = spanChk.id.split('_')[0] + '_' +spanChk.id.split('_')[1]; // 取得名稱的前兩段為識別列  
    var YES =document.getElementById(row + '_' +'CheckYESDETAIL');
    var NO =document.getElementById(row + '_' +'CheckNODETAIL');
    var Text =document.getElementById(row + '_' +'TextBox1');
  
 
    if (this.checked)  then 
    {
       NO.checked = false;
       Text.value = 'OK.';
   
    }
    
 }

 
 function CheckboxesNO(spanChk)  
{
 
   var row = spanChk.id.split('_')[0] + '_' +spanChk.id.split('_')[1]; // 取得名稱的前兩段為識別列  
    var YES =document.getElementById(row + '_' +'CheckYESDETAIL');
    var NO =document.getElementById(row + '_' +'CheckNODETAIL');
    var Text =document.getElementById(row + '_' +'TextBox1');
    if (this.checked)  then 
    {
       YES.checked = false;
       Text.value = 'NG.'; 
  
    }
    
    
}


</script>  
		
</head>
<body>

  	<form id="Form1" runat="server">
  	 <div>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:Label ID="Label2" runat="server" Style="z-index: 100; left: 137px; position: absolute;
               top: 18px" Text="廢棄準則"></asp:Label>
           <asp:Label ID="Label3" runat="server" Style="z-index: 101; left: 298px; position: absolute;
               top: 19px" Text="廢棄品分類"></asp:Label>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:ImageButton ID="BExcel" runat="server" Height="21px" ImageUrl="Images\msexcel.gif"
               Style="z-index: 102; left: 539px; position: absolute; top: 16px" Width="21px" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:HyperLink ID="LDISPOSALFILE" runat="server" Font-Size="12pt" Height="8px" Style="z-index: 102;
               left: 574px; position: absolute; top: 19px" Target="_blank" Width="238px">報廢附檔路徑</asp:HyperLink>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp;
           &nbsp;&nbsp;
           <asp:Button ID="Button1" runat="server" Style="z-index: 104; left: 480px; position: absolute;
               top: 16px" Text="查詢" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           <asp:DropDownList ID="DDisposalYM" runat="server" BackColor="Yellow" Style="z-index: 105;
               left: 47px; position: absolute; top: 16px" Width="81px" AutoPostBack="True">
           </asp:DropDownList><asp:DropDownList ID="DDISPOSALRULE" runat="server" BackColor="Yellow" Style="z-index: 106;
               left: 209px; position: absolute; top: 16px" Width="81px" AutoPostBack="True">
           </asp:DropDownList>
           <asp:DropDownList ID="DDISPOSALTYPE" runat="server" BackColor="Yellow" Style="z-index: 107;
               left: 388px; position: absolute; top: 17px" Width="81px" AutoPostBack="True">
           </asp:DropDownList>
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
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
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" Style="z-index: 110; left: 12px; position: absolute;
               top: 45px" AutoGenerateColumns="False" DataKeyNames="FormSno" ShowFooter="True" BackColor="White" BorderColor="#3366CC" BorderStyle="Solid" BorderWidth="1px" CellPadding="4" Width="1206px">
               <Columns>
                   <asp:BoundField DataField="NO" HeaderText="申請番號" >
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="appname" HeaderText="申請者">
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DUTYDEPO" HeaderText="責任部門" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="customertoll" HeaderText="向客戶請款">
                       <ItemStyle HorizontalAlign="Center" Width="85px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALTYPE" HeaderText="廢棄分類" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALREASON" HeaderText="申請原因" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALRULE" HeaderText="廢棄準則" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="STYPE" HeaderText="報廢形式">
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="PIECE" HeaderText="P" DataFormatString="{0:0.##}" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="METER" HeaderText="M" >
                       <ItemStyle Width="50px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="YARD" HeaderText="Y" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="KG" HeaderText="KG" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="PRICE" HeaderText="金額" DataFormatString="{0:0.##}" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="FormSno" HeaderText="FormSno" />
                   <asp:BoundField DataField="decidename" HeaderText="待簽核者">
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DisposalFile1" HeaderText="報廢明細檔">
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="OPURL" HeaderText="履歷" >
                       <ItemStyle Width="140px" />
                   </asp:BoundField>
               </Columns>
               <RowStyle BackColor="White" ForeColor="#003399" />
               <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
               <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
               <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
               <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="White" />
           </asp:GridView>
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
           <asp:Label ID="Label1" runat="server" Style="z-index: 111; left: 12px; position: absolute;
               top: 19px" Text="年月"></asp:Label>
  	 
      </div>
    </form>
</body>
</html>
