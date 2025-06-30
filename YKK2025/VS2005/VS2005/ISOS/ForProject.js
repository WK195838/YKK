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
    //
    //LISTPRICE
    if (Action == "LISTPRICEExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DLISTPRICEFile');
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
    //INQBUYERCOLOR
    if (Action == "INQBUYERCOLORExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DINQBUYERCOLORFile');
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
    //INQBUYERITEM
    if (Action == "INQBUYERITEMExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DINQBUYERITEMFile');
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
    //INQLISTPRICE
    if (Action == "INQLISTPRICEExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DINQLISTPRICEFile');
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
    //ISOS2FAS
    if (Action == "ISOS2FASExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DISOS2FASFile');
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
    //REPLACEITEM
    if (Action == "REPLACEITEMExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DREPLACEITEMFile');
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
    //REPLACECOLOR
    if (Action == "REPLACECOLORExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DREPLACECOLORFile');
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
    //SALESPRICE
    if (Action == "SALESPRICEExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DSALESPRICEFile');
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
    //INTEGRATEPRICE
    if (Action == "INTEGRATEPRICEExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DINTEGRATEPRICEFile');
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
    //BUYERITEM
    if (Action == "BUYERITEMExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DBUYERITEMFile');
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
    //BUYERCOLOR
    if (Action == "BUYERCOLORExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DBUYERCOLORFile');
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
    //POPUPIMAGE
    if (Action == "POPUPIMAGE")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPOPUPIMAGEFile');
        var executableFullPath =  xFilePath.value;

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
    //ITEMSUITABLE
    if (Action == "ITEMSUITABLEExcel")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DITEMSUITABLEFile');
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

//alert("CheckAttribute-out");
}

