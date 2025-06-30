<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DTMW_NewCommission.aspx.vb" Inherits="DTMW_NewCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<head>
		<title>新委託</title>
		
	    <script language="javascript" type="text/javascript">	  
    function RunExe(strFile)
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\DTMW\\MIS\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
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

    function RunExcel()
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\DTMW\\MIS\\DTMWTOOLS\\DTMAutoPriority.exe ';
    try
    {
        var shellActiveXObject = new ActiveXObject("WScript.Shell");

        if ( !shellActiveXObject )
        {
            alert('Could not get reference to WScript.Shell');
        }

        shellActiveXObject.Run(executableFullPath, 1, false);
        shellActiveXObject = null;
    }
    catch (errorObject)
    {
        alert('Error:\n' + errorObject.message);
    }            
}




    	</script>
    	
    	
	
	</head>
	<body >
		<form id="Form1" method="post" runat="server">
			
				<asp:Label id="DSystemTitle" style="Z-INDEX: 100; LEFT: 16px; POSITION: absolute; TOP: 24px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="200px">新色依賴各項申請</asp:Label>
                <asp:DropDownList ID="DFormNo" runat="server" Style="z-index: 101; left: 15px;
                    position: absolute; top: 67px" Width="199px">
                </asp:DropDownList>
                <asp:TextBox ID="TextBox1" runat="server" Height="85px" ReadOnly="True" Style="z-index: 102;
                    left: 271px; position: absolute; top: 28px" TextMode="MultiLine" Width="250px">先選擇委託單之後點選需要的按鈕  新委託：發行新的委託單</asp:TextBox>
                <asp:Button ID="BNew" runat="server" Style="z-index: 103; left: 90px; position: absolute;
                    top: 112px" Text="新委託" Width="120px" />
                <asp:Button ID="BPrint" runat="server" Style="z-index: 104; left: 92px; position: absolute;
                    top: 155px" Text="列印新色依賴書" Width="120px" />
                    
              <a href="http://10.245.1.6/DTMW/IEQuestion.aspx" target="_blank"><img src="Images/q4.jpg"  title="常見的問題" alt="常見的問題"  style="z-index: 106; left: 273px; position: absolute; top: 145px"/></a><asp:Button ID="BDTMP" runat="server" Style="z-index: 104; left: 92px; position: absolute;
                    top: 200px" Text="DTM優先度維護" Width="120px" />
          
           
               
                 
              
            </form>
	</body>
</html>