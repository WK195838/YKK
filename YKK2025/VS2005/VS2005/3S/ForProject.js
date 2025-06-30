//
//***********************************************************************************************
//** 功能按鈕點取
//**    設定各元件顯示內容
//** 
//***********************************************************************************************
//
function CheckAttribute(Action)
{
//alert("CheckAttribute-in");
//alert(Action);
    //JGS-準備資料
    if (Action == "JGS")
    {
        var xProJGS = document.getElementById('ProJGS');
        var xStsJGS = document.getElementById('StsJGS');
        var xBReset = document.getElementById('BReset');
        var xBJGS = document.getElementById('BJGS');

        xProJGS.style.left = '400px';
        xStsJGS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBJGS.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = "excel.exe " + xFilePath.value;

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
    //RIS-準備資料
    if (Action == "RIS")
    {
        var xProRIS = document.getElementById('ProRIS');
        var xStsRIS = document.getElementById('StsRIS');
        var xBReset = document.getElementById('BReset');
        var xBRIS = document.getElementById('BRIS');
        xProRIS.style.left = '266px';
        xStsRIS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBRIS.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = "excel.exe " + xFilePath.value;

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
    //BIS-產出BISheet
    if (Action == "BIS")
    {
        var xProBIS = document.getElementById('ProBIS');
        var xStsBIS = document.getElementById('StsBIS');
        var xBReset = document.getElementById('BReset');
        var xBBIS = document.getElementById('BBIS');

        xProBIS.style.left = '870px';
        xStsBIS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBBIS.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = xFilePath.value;

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
    //DGS-產出連結資料
    if (Action == "DGS")
    {
        var xProDGS = document.getElementById('ProDGS');
        var xStsDGS = document.getElementById('StsDGS');
        var xBReset = document.getElementById('BReset');
        var xBDGS = document.getElementById('BDGS');

        xProDGS.style.left = '870px';
        xStsDGS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBDGS.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = xFilePath.value;
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
    //RBS-產出報表資料
    if (Action == "RBS")
    {
        var xProRBS = document.getElementById('ProRBS');
        var xStsRBS = document.getElementById('StsRBS');
        var xBReset = document.getElementById('BReset');
        var xBRBS = document.getElementById('BRBS');

        xProRBS.style.left = '870px';
        xStsRBS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBRBS.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = xFilePath.value;

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
     if (Action == "TAX")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('xFilePath');
        var executableFullPath = "excel.exe " + xFilePath.value;

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
    //
//alert("CheckAttribute-out");
}
//
//***********************************************************************************************
//** 
//**
//***********************************************************************************************
//

       
