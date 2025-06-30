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

    //--
    //--[SPS] IMPORT------------------------------------------
    if (Action == "Import")
    {
//alert("Import-in");
        var xProImport = document.getElementById('ProImport');
        var xStsImport = document.getElementById('StsImport');
        var xBReset = document.getElementById('BReset');
        var xBImport = document.getElementById('BImport');

        xProImport.style.left = '240px';
        xStsImport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBImport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathImport');
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
//alert("Import-out");
    }
    //--
    //--[SPS] DEMAND------------------------------------------
    if (Action == "Demand")
    {
//alert("Demand-in");
        var xProDemand = document.getElementById('ProDemand');
        var xStsDemand = document.getElementById('StsDemand');
        var xBReset = document.getElementById('BReset');
        var xBDemand = document.getElementById('BDemand');

        xProDemand.style.left = '240px';
        xStsDemand.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBDemand.style.visibility='hidden';        // hide
//alert("Demand-out");
    }
    //--
    //--[SPS] ActionPlan------------------------------------------
    if (Action == "ActionPlan")
    {
//alert("ActionPlan-in");
        var xProActionPlan = document.getElementById('ProActionPlan');
        var xStsActionPlan = document.getElementById('StsActionPlan');
        var xBReset = document.getElementById('BReset');
        var xBActionPlan = document.getElementById('BActionPlan');

        xProActionPlan.style.left = '240px';
        xStsActionPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBActionPlan.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathActionPlan');
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
//alert("ActionPlan-out");
    }
    //--
    //--[SPS] ActionPlanImport------------------------------------------
    if (Action == "ActionPlanImport")
    {
//alert("ActionPlan-in");
        var xProActionPlanImport = document.getElementById('ProActionPlanImport');
        var xStsActionPlanImport = document.getElementById('StsActionPlanImport');
        var xBReset = document.getElementById('BReset');
        var xBActionPlanImport = document.getElementById('BActionPlanImport');

        xProActionPlanImport.style.left = '688px';
        xStsActionPlanImport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBActionPlanImport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathActionPlanImport');
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
//alert("ActionPlan-out");
    }
    //--
    //--[SPS] ActionPlanYOC------------------------------------------
    if (Action == "ActionPlanYOC")
    {
//alert("ActionPlan-in");
        var xProActionPlanYOC = document.getElementById('ProActionPlanYOC');
        var xStsActionPlanYOC = document.getElementById('StsActionPlanYOC');
        var xBReset = document.getElementById('BReset');
        var xBActionPlanYOC = document.getElementById('BActionPlanYOC');

        xProActionPlanYOC.style.left = '688px';
        xStsActionPlanYOC.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBActionPlanYOC.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathActionPlanYOC');
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
//alert("ActionPlan-out");
    }
    //--
    //--[SPS] KPIFSheet------------------------------------------
    if (Action == "KPIFSheet")
    {
//alert("KPIFSheet-in");
        var xProKPIFSheet = document.getElementById('ProKPIFSheet');
        var xStsKPIFSheet = document.getElementById('StsKPIFSheet');
        var xBReset = document.getElementById('BReset');
        var xBKPIFSheet = document.getElementById('BKPIFSheet');

        xProKPIFSheet.style.left = '688px';
        xStsKPIFSheet.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBKPIFSheet.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathKPIFSheet');
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
//alert("KPIFSheet-out");
    }
    //--
    //--[SPS] WINGS------------------------------------------
    if (Action == "WINGS")
    {
//alert("KPIFSheet-in");
        var xProWINGS = document.getElementById('ProWINGS');
        var xStsWINGS = document.getElementById('StsWINGS');
        var xBReset = document.getElementById('BReset');
        var xBWINGS = document.getElementById('BWINGS');

        xProWINGS.style.left = '1192px';
        xStsWINGS.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBWINGS.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathWINGS');
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
//alert("KPIFSheet-out");
    }
    //--
    //--[SPS] PIL------------------------------------------
    if (Action == "PIL")
    {
//alert("PIL-in");
        var xProPIL = document.getElementById('ProPIL');
        var xStsPIL = document.getElementById('StsPIL');
        var xBReset = document.getElementById('BReset');
        var xBPIL = document.getElementById('BPIL');

        xProPIL.style.left = '1208px';
        xStsPIL.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBPIL.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPathPIL');
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
//alert("PIL-out");
    }
    //--
    //--[SPS] ChangeFinal------------------------------------------
    if (Action == "ChangeFinal")
    {
//alert("ChangeFinal-in");
        var xProChangeFinal = document.getElementById('ProChangeFinal');
        var xStsChangeFinal = document.getElementById('StsChangeFinal');
        var xBReset = document.getElementById('BReset');
        var xBChangeFinal = document.getElementById('BChangeFinal');

        xProChangeFinal.style.left = '1192px';
        xStsChangeFinal.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';         // hide
        xBChangeFinal.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
//        var xFilePath = document.getElementById('DPathChangeFinal');
//        var executableFullPath = "excel.exe " + xFilePath.value;

//        try
//        {
//            var shellActiveXObject = new ActiveXObject("WScript.Shell");

//            if ( !shellActiveXObject )
//            {
//                alert('Could not get reference to WScript.Shell');
//            }
//            shellActiveXObject.Run(executableFullPath, 1, false);
//            shellActiveXObject = null;
//        }
//        catch (errorObject)
//        {
//            alert('Error:\n' + errorObject.message);
//        }
//alert("ChangeFinal-out");
    }

    //--
    //--[SPS] JOY???------------------------------------------

}
//
//***********************************************************************************************
//** 
//**
//***********************************************************************************************
//

       
