<%@ Page Language="VB" AutoEventWireup="false" CodeFile="StockNewCommission.aspx.vb" Inherits="StockNewCommission" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<HEAD>
		<title>新委託</title>
		    <script language="javascript" type="text/javascript">	 
		    
//*********************************************************************
//CALL EXCEL
//*********************************************************************
   function RunExcelIN()
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\StockInOut\\PGM\\STOCK\\N2W_STOCKIN.exe ';
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

   function RunExcelOUT()
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\StockInOut\\PGM\\STOCK\\N2W_STOCKOUT.exe ';
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


   function RunExcelCHECK()
{      
    //要執行的程式如果要帶參數記得後面要空格而且參數中不可以有"\"
     
 //  var executableFullPath = 'D:\\db\\SalePrint.exe ';
   //  var executableFullPath ='\\\\10.245.1.61\\mis\\DTMWTOOLS\\DTMW_AutoPrint1.exe ';
    var executableFullPath ='\\\\10.245.1.6\\www$\\StockInOut\\PGM\\STOCK\\N2W_STOCKCHECK.exe ';
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
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                <asp:HyperLink ID="LFun04" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 101; left: 72px; position: absolute;
                    top: 312px" Target="_blank" Width="136px">出庫調閱資料</asp:HyperLink>
                <asp:HyperLink ID="LFun02" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 102; left: 72px; position: absolute;
                    top: 144px" Target="_blank" Width="80px">申請出庫</asp:HyperLink>
                <asp:HyperLink ID="LFun03" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 103; left: 72px; position: absolute;
                    top: 280px" Target="_blank" Width="136px">入庫調閱資料</asp:HyperLink>
                &nbsp; &nbsp; &nbsp;
                &nbsp; &nbsp;&nbsp;
                <asp:HyperLink ID="LFun01" runat="server" Enabled="False" Font-Size="12pt" Height="16px" Style="z-index: 104; left: 72px; position: absolute;
                    top: 112px" Target="_blank" Width="80px">申請入庫</asp:HyperLink>
                &nbsp;
                &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
				<asp:Label id="DSystemTitle" style="Z-INDEX: 105; LEFT: 40px; POSITION: absolute; TOP: 8px"
					runat="server" ForeColor="Navy" Font-Bold="True" Font-Size="16pt" Width="260px">發送出入庫各項申請</asp:Label>
                <asp:Image ID="Image1" runat="server" ImageUrl="~/images/StockDefault.jpg" Style="z-index: 99;
                    left: 32px; position: absolute; top: 40px" />
                <asp:Button ID="BPrint" runat="server" Style="z-index: 104; left: 160px; position: absolute;
                    top: 112px" Text="列印入庫單" Width="104px" BackColor="#FFE0C0" /><asp:Button ID="BStockIn" runat="server" Style="z-index: 104; left: 376px; position: absolute;
                    top: 104px" Text="待入庫明細" Width="120px" BackColor="#F4B084" /><asp:Button ID="BStockOut" runat="server" Style="z-index: 104; left: 376px; position: absolute;
                    top: 144px" Text="待出庫明細" Width="120px" BackColor="#F4B084" />
                &nbsp;&nbsp;
                <asp:TextBox ID="DITEMSUITABLEFile" runat="server" BorderStyle="None" BorderWidth="0px"
                    Height="16px" Style="font-size: 10px; z-index: 318; background: none transparent scroll repeat 0% 0%;
                    left: 584px; position: absolute; top: 72px" Width="116px"></asp:TextBox><asp:Button ID="BSTOCKCHECK" runat="server" Style="z-index: 104; left: 72px; position: absolute;
                    top: 344px" Text="在庫查詢" Width="120px" BackColor="#A9D08E" />
            </FONT></form>
	</body>
</HTML>