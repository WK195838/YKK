<%@ Page Language="vb" AutoEventWireup="false" Inherits="DASW_DISPOSALADMIN" aspCompat="True" EnableEventValidation = "false"  CodeFile="DASW_DISPOSALADMIN.aspx.vb" %>
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
        	<FONT face="新細明體"></FONT>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
           <asp:Button ID="BOK" runat="server" Text="簽核" style="z-index: 100; left: 20px; position: absolute; top: 9px" Font-Size="Medium" Width="120px" />
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
           &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
           <asp:GridView ID="GridView1" runat="server" Style="z-index: 101; left: 27px; position: absolute;
               top: 47px" AutoGenerateColumns="False" DataKeyNames="FormSno" ShowFooter="True" BackColor="White" BorderColor="#3366CC" BorderStyle="Solid" BorderWidth="1px" CellPadding="4" Width="1500px">
               <Columns>
                   <asp:TemplateField HeaderText="OK">
                       <HeaderTemplate>
                           &nbsp;OK
                             <asp:CheckBox ID="CheckYESHEAD" runat="server" onclick= "SelectAllCheckboxesYES(this);"  ToolTip="按一次全選，再按一次取消全選" />
                           / NG<asp:CheckBox ID="CheckNOHEAD" runat="server" onclick= "SelectAllCheckboxesNO(this);"  ToolTip="按一次全選，再按一次取消全選" />
                       </HeaderTemplate>
                       <ItemTemplate>
                           &nbsp;OK
                            <asp:CheckBox ID="CheckYESDETAIL" runat="server" onclick= "CheckboxesYES(this);"    />
                           / NG&nbsp;<asp:CheckBox ID="CheckNODETAIL" runat="server" onclick= "CheckboxesNO(this);"   />

                       </ItemTemplate>
                       <ItemStyle Width="140px" />
                   </asp:TemplateField>
                   <asp:BoundField DataField="NO" HeaderText="申請番號" >
                       <ItemStyle Width="80px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DUTYDEPO" HeaderText="責任部門" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="customertoll" HeaderText="向客戶請款">
                       <ItemStyle HorizontalAlign="Center" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALTYPE" HeaderText="廢棄分類" >
                       <ItemStyle Width="60px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALREASON" HeaderText="申請原因" >
                       <ItemStyle Width="100px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="DISPOSALRULE" HeaderText="廢棄準則" >
                       <ItemStyle Width="60px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="stype" HeaderText="報廢形式" />
                   <asp:BoundField DataField="PIECE" HeaderText="P" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="METER" HeaderText="M" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="YARD" HeaderText="Y" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="KG" HeaderText="KG" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:BoundField DataField="PRICE" HeaderText="金額" >
                       <ItemStyle Width="40px" />
                   </asp:BoundField>
                   <asp:TemplateField HeaderText="簽核說明">
                       <ItemTemplate>
                           <asp:TextBox ID="TextBox1" runat="server" TextMode="MultiLine" BackColor="GreenYellow" Width="108px"></asp:TextBox>
                       </ItemTemplate>
                       <ItemStyle Width="100px" />
                   </asp:TemplateField>
                   <asp:BoundField DataField="History" HeaderText="履歷" />
                   <asp:BoundField DataField="FormSno" HeaderText="FormSno" />
                   <asp:BoundField DataField="applyid" HeaderText="applyid" />
                   <asp:BoundField DataField="decideid" HeaderText="decideid" />
                   <asp:BoundField DataField="step" HeaderText="step" />
                   <asp:BoundField DataField="no" HeaderText="no" />
                   <asp:BoundField DataField="disposalfile1" HeaderText="報廢明細" />
               </Columns>
               <RowStyle BackColor="White" ForeColor="#003399" />
               <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
               <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
               <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
               <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="White" />
           </asp:GridView>
           &nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp;
  	 
      </div>
    </form>
</body>
</html>
