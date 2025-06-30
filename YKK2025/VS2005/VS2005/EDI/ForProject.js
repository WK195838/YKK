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
    //----GESS--------------------------------------------------------------------------------------------
    //
    //CONVERT / EDI轉換
    if (Action == "GESSCONVERT")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath1');
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
    //MODIFY / EDI修改
    if (Action == "GESSMODIFY")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath2');
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
    //MAINTMST / 維護MST
    if (Action == "GESSMAINTMST")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath3');
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
    //INQ / 查詢修改履歷
    if (Action == "GESSINQHISTORY")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath4');
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
    //其它修改
    if (Action == "GESSOTHER")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath5');
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
    //----eKPI--------------------------------------------------------------------------------------------
    //
    //MST
    if (Action == "KPIMSTM")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath1');
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
    //MST展開
    if (Action == "KPIMSTE")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath2');
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
    //OVERFLOW-FCT
    if (Action == "KPIOFCT")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath3');
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
    //OVERFLOW-FCT(CAL)
    if (Action == "KPIOFCTCAL")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath7');
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
    //FINAL
    if (Action == "KPIFINAL")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath4');
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
    //REPORT
    if (Action == "KPIREPORT")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath5');
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
    //REPORT
    if (Action == "KPIADCHECK")
    {
        //--啟動程式        
        var xFilePath = document.getElementById('DFilePath6');
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
    //判定SELAY REASON
    if (Action == "KPIDELAYREASON")
    {
        var xProDecision = document.getElementById('ProDecision');

        xProDecision.style.left = '488px';
    }    
    //
    //----EOES專用--------------------------------------------------------------------------------------------
    //
    //TempLate
    if (Action == "Template")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath1');
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
    //Item
    if (Action == "Item")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath2');
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
    //Color
    if (Action == "Color")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath3');
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
    //----EDI專用--------------------------------------------------------------------------------------------
    //
    //準備客戶資料
    if (Action == "Excel")
    {
        var xProExcel = document.getElementById('ProExcel');
        var xStsExcel = document.getElementById('StsExcel');
        var xBReset = document.getElementById('BReset');
        var xBExcel = document.getElementById('BExcel');

        xProExcel.style.left = '20px';
        xStsExcel.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBExcel.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath');
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
    //匯入作業
    if (Action == "Import")
    {
        var xProImport = document.getElementById('ProImport');
        var xStsImport = document.getElementById('StsImport');
        var xBReset = document.getElementById('BReset');
        var xBImport = document.getElementById('BImport');

        xProImport.style.left = '20px';
        xStsImport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBImport.style.visibility='hidden';     // hide
    }
    //檢查資料
    if (Action == "DataCheck")
    {
        var xProDataCheck = document.getElementById('ProDataCheck');
        var xStsMakePONO = document.getElementById('StsMakePONO');
        var xStsCheckPONO = document.getElementById('StsCheckPONO');
        var xStsCheckGRPC = document.getElementById('StsCheckGRPC');
        var xStsCheckCompanyCode = document.getElementById('StsCheckCompanyCode');
        var xStsCheckKeepCode = document.getElementById('StsCheckKeepCode');
        var xStsCheckColorCode = document.getElementById('StsCheckColorCode');
        var xStsCheckItemCode = document.getElementById('StsCheckItemCode');
        var xStsCheckDuplicateData = document.getElementById('StsCheckDuplicateData');
        var xStsCheckPOStructure = document.getElementById('StsCheckPOStructure');
        var xBReset = document.getElementById('BReset');
        var xBDataCheck = document.getElementById('BDataCheck');

        xProDataCheck.style.left = '416px';
        xStsMakePONO.style.visibility='hidden';             // hide
        xStsCheckPONO.style.visibility='hidden';            // hide
        xStsCheckGRPC.style.visibility='hidden';            // hide
        xStsCheckCompanyCode.style.visibility='hidden';     // hide
        xStsCheckKeepCode.style.visibility='hidden';        // hide
        xStsCheckColorCode.style.visibility='hidden';       // hide
        xStsCheckItemCode.style.visibility='hidden';        // hide
        xStsCheckDuplicateData.style.visibility='hidden';   // hide
        xStsCheckPOStructure.style.visibility='hidden';     // hide
        xBReset.style.visibility='hidden';                  // hide
        xBDataCheck.style.visibility='hidden';              // hide
    }
    //Waves
    if (Action == "Waves")
    {
        var xProWaves = document.getElementById('ProWaves');
        var xStsWaves = document.getElementById('StsWaves');
        var xBReset = document.getElementById('BReset');
        var xBWaves = document.getElementById('BWaves');

        xProWaves.style.left = '732px';
        xStsWaves.style.visibility='hidden';        // hide
        xBReset.style.visibility='hidden';          // hide
        xBWaves.style.visibility='hidden';          // hide
    }
    //Price
    if (Action == "Price")
    {
        var xProSalesPrice = document.getElementById('ProSalesPrice');
        var xStsSalesPrice = document.getElementById('StsSalesPrice');
        var xBReset = document.getElementById('BReset');
        var xBSalesPrice = document.getElementById('BSalesPrice');

        xProSalesPrice.style.left = '732px';
        xStsSalesPrice.style.visibility='hidden';   // hide
        xBReset.style.visibility='hidden';          // hide
        xBSalesPrice.style.visibility='hidden';     // hide
    }
    //
    //Run Program
    if (Action == "Run")
    {
        var xProRun = document.getElementById('ProRun');
        var xStsRun = document.getElementById('StsRun');
        var xBReset = document.getElementById('BReset');
        var xBRun = document.getElementById('BRun');

        xProRun.style.left = '357px';
        xStsRun.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBRun.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath1');
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
    //
    //----FAS專用--------------------------------------------------------------------------------------------
    //
    //準備客戶資料
    if (Action == "FASExcel")
    {
        var xProExcel = document.getElementById('ProExcel');
        var xStsExcel = document.getElementById('StsExcel');
        var xBReset = document.getElementById('BReset');
        var xBExcel = document.getElementById('BExcel');

        xProExcel.style.left = '10px';
        xStsExcel.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBExcel.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath');
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
    //Item / Color 轉換
    if (Action == "Convert")
    {
        var xProConvert = document.getElementById('ProConvert');
        var xStsConvert = document.getElementById('StsConvert');
        var xBReset = document.getElementById('BReset');
        var xBConvert = document.getElementById('BConvert');

        xProConvert.style.left = '10px';
        xStsConvert.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBConvert.style.visibility='hidden';        // hide
    }
    //Forcast Plan展開
    if (Action == "FCTPlan")
    {
        var xProFCTPlan = document.getElementById('ProFCTPlan');
        var xStsFCTPlan = document.getElementById('StsFCTPlan');
        var xBReset = document.getElementById('BReset');
        var xBFCTPlan = document.getElementById('BFCTPlan');

        xProFCTPlan.style.left = '370px';
        xStsFCTPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBFCTPlan.style.visibility='hidden';        // hide
    }
    //Local Stock Plan展開
    if (Action == "LSPlan")
    {
        var xProLSPlan = document.getElementById('ProLSPlan');
        var xStsLSPlan = document.getElementById('StsLSPlan');
        var xBReset = document.getElementById('BReset');
        var xBLSPlan = document.getElementById('BLSPlan');

        xProLSPlan.style.left = '370px';
        xStsLSPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBLSPlan.style.visibility='hidden';        // hide
    }
    //Report
    if (Action == "Report")
    {
        var xProReport = document.getElementById('ProReport');
        var xStsReport = document.getElementById('StsReport');
        var xBReset = document.getElementById('BReset');
        var xBReport = document.getElementById('BReport');

        xProReport.style.left = '370px';
        xStsReport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBReport.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DReportFilePath');
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
    //Import LS Plan
    if (Action == "IPLSPlan")
    {
        var xProIPLSPlan = document.getElementById('ProIPLSPlan');
        var xStsIPLSPlan = document.getElementById('StsIPLSPlan');
        var xBReset = document.getElementById('BReset');
        var xBIPLSPlan = document.getElementById('BIPLSPlan');

        xProIPLSPlan.style.left = '730px';
        xStsIPLSPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBIPLSPlan.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath1');
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
    //Buyer Local Stock Plan展開
    if (Action == "BULSPlan")
    {
        var xProBULSPlan = document.getElementById('ProBULSPlan');
        var xStsBULSPlan = document.getElementById('StsBULSPlan');
        var xBReset = document.getElementById('BReset');
        var xBBULSPlan = document.getElementById('BBULSPlan');

        xProBULSPlan.style.left = '730px';
        xStsBULSPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBBULSPlan.style.visibility='hidden';        // hide
    }
    //Buyer Report
    if (Action == "BUReport")
    {
        var xProBUReport = document.getElementById('ProBUReport');
        var xStsBUReport = document.getElementById('StsBUReport');
        var xBReset = document.getElementById('BReset');
        var xBBUReport = document.getElementById('BBUReport');

        xProBUReport.style.left = '730px';
        xStsBUReport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBBUReport.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DReportFilePath1');
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
    //Buyer LS Plan最終確定
    if (Action == "LFLSPlan")
    {
        var xProLFLSPlan = document.getElementById('ProLFLSPlan');
        var xStsLFLSPlan = document.getElementById('StsLFLSPlan');
        var xBReset = document.getElementById('BReset');
        var xBLFLSPlan = document.getElementById('BLFLSPlan');

        xProLFLSPlan.style.left = '1090px';
        xStsLFLSPlan.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBLFLSPlan.style.visibility='hidden';        // hide
    }
    //EDI 轉換
    if (Action == "EPEDI")
    {
        var xProEPEDI = document.getElementById('ProEPEDI');
        var xStsEPEDI = document.getElementById('StsEPEDI');
        var xBReset = document.getElementById('BReset');
        var xBEPEDI = document.getElementById('BEPEDI');

        xProEPEDI.style.left = '1090px';
        xStsEPEDI.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';          // hide
        xBEPEDI.style.visibility='hidden';        // hide
    }
    //EDI Report
    if (Action == "EDIReport")
    {
        var xProEDIReport = document.getElementById('ProEDIReport');
        var xStsEDIReport = document.getElementById('StsEDIReport');
        var xBReset = document.getElementById('BReset');
        var xBEDIReport = document.getElementById('BEDIReport');

        xProEDIReport.style.left = '1090px';
        xStsEDIReport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBEDIReport.style.visibility='hidden';      // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DReportFilePath2');
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
    //PLAN Report
    if (Action == "PLANReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DPLANReportFilePath');
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
    //FCT2ACT Report
    if (Action == "FCT2ACTReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFCT2ACTReportFilePath');
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
    //FCT2ACTMM Report
    if (Action == "FCT2ACTMMReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFCT2ACTMMReportFilePath');
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
    //FCT2ACTVDP Report
    if (Action == "FCT2ACTVDPReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFCT2ACTVDPReportFilePath');
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
    //MAT2SEA Report
    if (Action == "MAT2SEAReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DMAT2SEAReportFilePath');
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
    //MAT2MON Report
    if (Action == "MAT2MONReport")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DMAT2MONReportFilePath');
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
    //----VENDOR FAS專用--------------------------------------------------------------------------------------------
    //
    //備料INPUT
    if (Action == "LSInput")
    {
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DLSInputFilePath');
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
    //----FAS BUYER FCT 000001/000016 專用--------------------------------------------------------------------------------------------
    //
    //BY FCT 格式轉換
    if (Action == "BYFMTChange")
    {
        var xProBYFMTChange = document.getElementById('ProBYFMTChange');
        var xStsBYFMTChange = document.getElementById('StsBYFMTChange');
        var xBReset = document.getElementById('BReset');
        var xBBYFMTChange = document.getElementById('BBYFMTChange');
        var xAtGoImport = document.getElementById('AtGoImport');
        var xAtGoReport = document.getElementById('AtGoReport');

        xProBYFMTChange.style.left = '10px';
        xStsBYFMTChange.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYFMTChange.style.visibility='hidden';        // hide
        
        if ( !xAtGoImport.checked && !xAtGoReport.checked )
        {
            //--啟動Excel程式        
            var xFilePath = document.getElementById('DFilePath');
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
    }
    //準備 BY FCT 資料
    if (Action == "BYImport")
    {
        var xProBYImport = document.getElementById('ProBYImport');
        var xStsBYImport = document.getElementById('StsBYImport');
        var xBReset = document.getElementById('BReset');
        var xBBYImport = document.getElementById('BBYImport');
        var xAtNotFirst = document.getElementById('AtNotFirst');

        xProBYImport.style.left = '10px';
        xStsBYImport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYImport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        if ( !xAtNotFirst.checked )
        {
            var xFilePath = document.getElementById('DFilePath1');
        }
        if ( xAtNotFirst.checked )
        {
            var xFilePath = document.getElementById('DFilePath2');
        }
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
    //檢查資料
    if (Action == "BYDataCheck")
    {
        var xProBYDataCheck = document.getElementById('ProBYDataCheck');
        var xStsBYDataCheck = document.getElementById('StsBYDataCheck');
        var xBReset = document.getElementById('BReset');
        var xBBYDataCheck = document.getElementById('BBYDataCheck');

        xProBYDataCheck.style.left = '370px';
        xStsBYDataCheck.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBBYDataCheck.style.visibility='hidden';     // hide
    }
    //報表列出
    if (Action == "BYReport")
    {
        var xProBYReport = document.getElementById('ProBYReport');
        var xStsBYReport = document.getElementById('StsBYReport');
        var xBReset = document.getElementById('BReset');
        var xBBYReport = document.getElementById('BBYReport');

        xProBYReport.style.left = '370px';
        xStsBYReport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYReport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DReportFilePath');
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
    //轉FAS BYFCT
    if (Action == "BYConvert")
    {
        var xProBYConvert = document.getElementById('ProBYConvert');
        var xStsBYConvert = document.getElementById('StsBYConvert');
        var xBReset = document.getElementById('BReset');
        var xBBYConvert = document.getElementById('BBYConvert');

        xProBYConvert.style.left = '730px';
        xStsBYConvert.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';      // hide
        xBBYConvert.style.visibility='hidden';     // hide
    }
    //BYFCT FAS報表列出
    if (Action == "BYFASReport")
    {
        var xProBYFASReport = document.getElementById('ProBYFASReport');
        var xStsBYFASReport = document.getElementById('StsBYFASReport');
        var xBReset = document.getElementById('BReset');
        var xBBYFASReport = document.getElementById('BBYFASReport');

        xProBYFASReport.style.left = '730px';
        xStsBYFASReport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYFASReport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DReportFilePath1');
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
    //----FAS BUYER FCT NIKE專用--------------------------------------------------------------------------------------------
    //
    //準備 BY FCT 資料
    if (Action == "NIKEBYImport")
    {
        var xProBYImport = document.getElementById('ProBYImport');
        var xStsBYImport = document.getElementById('StsBYImport');
        //var xBReset = document.getElementById('BReset');
        //var xBBYImport = document.getElementById('BBYImport');

        xProBYImport.style.left = '10px';
        xStsBYImport.style.visibility='hidden';      // hide
        //xBReset.style.visibility='hidden';           // hide
        //xBBYImport.style.visibility='hidden';        // hide

        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath');
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
    //----FAS BUYER FCT COLUMBIA專用--------------------------------------------------------------------------------------------
    //
    //BY FCT 格式轉換
    if (Action == "CBBYFMTChange")
    {
        var xProBYFMTChange = document.getElementById('ProBYFMTChange');
        var xStsBYFMTChange = document.getElementById('StsBYFMTChange');
        var xBReset = document.getElementById('BReset');
        var xBBYFMTChange = document.getElementById('BBYFMTChange');

        xProBYFMTChange.style.left = '10px';
        xStsBYFMTChange.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYFMTChange.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath');
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
    //準備 BY FCT 資料
    if (Action == "CBBYImport")
    {
        var xProBYImport = document.getElementById('ProBYImport');
        var xStsBYImport = document.getElementById('StsBYImport');
        var xBReset = document.getElementById('BReset');
        var xBBYImport = document.getElementById('BBYImport');

        xProBYImport.style.left = '10px';
        xStsBYImport.style.visibility='hidden';      // hide
        xBReset.style.visibility='hidden';              // hide
        xBBYImport.style.visibility='hidden';        // hide
        
        //--啟動Excel程式        
        var xFilePath = document.getElementById('DFilePath1');
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
//
//***********************************************************************************************
//** Action CheckBox 點取
//**    設定CheckBox元件顯示內容
//** 
//***********************************************************************************************
//
function ActionCheck(Action)
{
//alert("ActionCheck-in");
    var Report = document.getElementById('AtReport');
    var IPLSPlan = document.getElementById('AtIPLSPlan');
    var BULSPlan = document.getElementById('AtBULSPlan');
    var Convert = document.getElementById('AtConvert');
    var FCTPlan = document.getElementById('AtFCTPlan');
    var LSPlan = document.getElementById('AtLSPlan');
    var LFLSPlan = document.getElementById('AtLFLSPlan');
    var EPEDI = document.getElementById('AtEPEDI');

    Report.checked = false;
    IPLSPlan.checked = false;
    BULSPlan.checked = false;
    Convert.checked = false;
    FCTPlan.checked = false;
    LSPlan.checked = false;
    LFLSPlan.checked = false;
    EPEDI.checked = false;
     
    if (Action == "AtReport")
    {
        Report.checked = true;
    }
    if (Action == "AtIPLSPlan")
    {
        IPLSPlan.checked = true;
    }
    if (Action == "AtBULSPlan")
    {
        BULSPlan.checked = true;
    }
    if (Action == "AtConvert")
    {
        Convert.checked = true;
    }
    if (Action == "AtFCTPlan")
    {
        FCTPlan.checked = true;
    }
    if (Action == "AtLSPlan")
    {
        LSPlan.checked = true;
    }
    if (Action == "AtLFLSPlan")
    {
        LFLSPlan.checked = true;
    }
    if (Action == "AtEPEDI")
    {
        EPEDI.checked = true;
    }
//alert("ActionCheck-out");
}
//
//***********************************************************************************************
//** Action CheckBox 點取 BYFCT
//**    設定CheckBox元件顯示內容
//** 
//***********************************************************************************************
//
function ActionCheckBY(Action)
{
//alert("ActionCheck-in");
    var Import = document.getElementById('AtImport');
    var Report = document.getElementById('AtReport');
     
    if (Action == "AtImport")
    {
        Report.checked = false;
    }
    if (Action == "AtReport")
    {
        Import.checked = false;
    }
//alert("ActionCheck-out");
}


//
//***********************************************************************************************
//** 
//**
//***********************************************************************************************
//

       
