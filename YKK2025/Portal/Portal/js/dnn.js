var dnn;	//should really make this m_dnn... but want to treat like namespace

var DNN_HIGHLIGHT_COLOR = '#9999FF';
var COL_DELIMITER = String.fromCharCode(18);
var ROW_DELIMITER = String.fromCharCode(17);

if (typeof(__dnn_m_aNamespaces) == 'undefined')	//include in each DNN ClientAPI namespace file for dependency loading
	var __dnn_m_aNamespaces = new Array();

//NameSpace DNN
	function __dnn()
	{
		this.apiversion = .1;
		this.pns = '';
		this.ns = 'dnn';
		this.diagnostics = null;
		this.vars = null;
		this.dependencies = new Array();
		this.isLoaded = false;
	}

	__dnn.prototype.getVars = function()
	{
		if (this.vars == null)
		{
			this.vars = new Array();
			var oCtl = dnn.dom.getById('__dnnVariable');
			if (oCtl != null)
			{
				var aryItems = oCtl.value.split(ROW_DELIMITER);
				for (var i=0; i<aryItems.length; i++)
				{
					var aryItem = aryItems[i].split(COL_DELIMITER);
					
					if (aryItem.length == 2)
						this.vars[aryItem[0]] = aryItem[1];
				}
			}
		}
		return this.vars;	
	}

	__dnn.prototype.getVar = function(sKey)
	{
		return this.getVars()[sKey];
	}

	__dnn.prototype.setVar = function(sKey, sVal)
	{			
		if (this.vars == null)
			this.getVars();			
		this.vars[sKey] = sVal;
		var oCtl = dnn.dom.getById('__dnnVariable');
		if (oCtl == null)
		{
			oCtl = dnn.dom.createElement('INPUT');
			oCtl.type = 'hidden';
			oCtl.id = '__dnnVariable';
			dnn.dom.appendChild(dnn.dom.getByTagName("body")[0], oCtl);		
		}
		var sVals = '';
		var sKey
		for (sKey in this.vars)
		{
			sVals += ROW_DELIMITER + sKey + COL_DELIMITER + this.vars[sKey];
		}
		oCtl.value = sVals;
		return true;
	}

	__dnn.prototype.callPostBack = function(sAction)
	{
		var sPostBack = dnn.getVar('__dnn_postBack');
		var sData = '';
		if (sPostBack.length > 0)
		{
			sData += sAction;
			for (var i=1; i<arguments.length; i++)
			{
				var aryParam = arguments[i].split('=');
				sData += COL_DELIMITER + aryParam[0] + COL_DELIMITER + aryParam[1];
			}
			eval(sPostBack.replace('[DATA]', sData));
		}
	}

	__dnn.prototype.dependenciesLoaded = function()
	{
		return true;
	}

	__dnn.prototype.loadNamespace = function ()
	{
		if (this.isLoaded == false)
		{
			if (this.dependenciesLoaded())
			{
				dnn = this; 
				this.isLoaded = true;
				this.loadDependencies(this.pns, this.ns);
			}
		}	
	}

	__dnn.prototype.loadDependencies = function (sPNS, sNS)
	{
		for (var i=0; i<__dnn_m_aNamespaces.length; i++)
		{
			for (var iDep=0; i<__dnn_m_aNamespaces[i].dependencies.length; i++)
			{
				if (__dnn_m_aNamespaces[i].dependencies[iDep] == sPNS + (sPNS.length>0 ? '.': '') + sNS)
					__dnn_m_aNamespaces[i].loadNamespace();
			}
		}
	}


	//--- dnn.dom
		function dnn_dom()
		{
			this.pns = 'dnn';
			this.ns = 'dom';
			this.dependencies = 'dnn'.split(',');
			this.isLoaded = false;
		}

		dnn_dom.prototype.getById = function (sID, oCtl)
		{
			if (oCtl == null)
				oCtl = document;
			if (oCtl.getElementById) //(dnn.dom.browser.isType(dnn.dom.browser.InternetExplorer) == false)
				return oCtl.getElementById(sID);
			else
				return oCtl.all(sID);
		}

		dnn_dom.prototype.getByTagName = function (sTag, oCtl)
		{
			if (oCtl == null)
				oCtl = document;
			if (oCtl.getElementsByTagName) //(dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
				return oCtl.getElementsByTagName(sTag);
			else if (oCtl.all.tags)
				return oCtl.all.tags(sTag);
			else
				return null;
		}

		dnn_dom.prototype.createElement = function (sTagName) 
		{
			if (document.createElement) 
				return document.createElement(sTagName);
			else 
				return null;
		}
		
		dnn_dom.prototype.isNonTextNode = function (oNode)
		{
			return (oNode.nodeType != 3 && oNode.nodeType != 8); //exclude nodeType of Text (Netscape/Mozilla) issue!
		}
		
		dnn_dom.prototype.getNonTextNode = function (oNode)
		{
			if (this.isNonTextNode(oNode))	
				return oNode;
			
			while (oNode != null && this.isNonTextNode(oNode))
			{
				oNode = this.getSibling(oNode, 1);
			}
			return oNode;
		}

		dnn_dom.prototype.getSibling = function (oCtl, iOffset)
		{
			if (oCtl != null && oCtl.parentNode != null)
			{
				for (var i=0; i<oCtl.parentNode.childNodes.length; i++)
				{
					if (oCtl.parentNode.childNodes[i].id == oCtl.id)
					{
						if (oCtl.parentNode.childNodes[i + iOffset] != null)
							return oCtl.parentNode.childNodes[i + iOffset];
					}
				}
			}
		}

		dnn_dom.prototype.appendChild = function (oParent, oChild) 
		{
			if (oParent.appendChild) 
				return oParent.appendChild(oChild);
			else 
				return null;
		}

		dnn_dom.prototype.removeChild = function (oChild) 
		{
			if (oChild.parentNode.removeChild) 
				return oChild.parentNode.removeChild(oChild);
			else 
				return null;
		}


		dnn_dom.prototype.setCookie = function (sName, sVal, iDays, sPath, sDomain, bSecure) 
		{
			var sExpires;
			if (iDays)
			{
				sExpires = new Date();
				sExpires.setTime(sExpires.getTime()+(iDays*24*60*60*1000));
			}
			document.cookie = sName + "=" + escape(sVal) + ((sExpires) ? "; expires=" + sExpires : "") + 
				((sPath) ? "; path=" + sPath : "") + ((sDomain) ? "; domain=" + sDomain : "") + ((bSecure) ? "; secure" : "");
			
			if (document.cookie.length > 0)
				return true;
				
		}
		dnn_dom.prototype.getCookie = function (sName) 
		{
			var sCookie = " " + document.cookie;
			var sSearch = " " + sName + "=";
			var sStr = null;
			var iOffset = 0;
			var iEnd = 0;
			if (sCookie.length > 0) 
			{
				iOffset = sCookie.indexOf(sSearch);
				if (iOffset != -1) 
				{
					iOffset += sSearch.length;
					iEnd = sCookie.indexOf(";", iOffset)
					if (iEnd == -1) 
						iEnd = sCookie.length;
					sStr = unescape(sCookie.substring(iOffset, iEnd));
				}
			}
			return(sStr);
		}

		dnn_dom.prototype.deleteCookie = function (sName, sPath, sDomain) 
		{
			if (this.getCookie(sName)) 
			{
				this.setCookie(sName, '', -1, sPath, sDomain);
				return true;
			}
			return false;
		}

		dnn_dom.prototype.dependenciesLoaded = function()
		{
			return (typeof(dnn) != 'undefined');
		}

		dnn_dom.prototype.loadNamespace = function ()
		{
			if (this.isLoaded == false)
			{
				if (this.dependenciesLoaded())
				{
					dnn.dom = this; 
					this.isLoaded = true;
					dnn.loadDependencies(this.pns, this.ns);
				}
			}	
		}


			//--- dnn.dom.browser
			function dnn_dom_browser()
			{
				this.pns = 'dnn.dom';
				this.ns = 'browser';
				this.dependencies = 'dnn,dnn.dom'.split(',');
				this.isLoaded = false;
				this.InternetExplorer = 'ie';
				this.Netscape = 'ns';
				this.Mozilla = 'mo';
				this.Opera = 'op';
				this.Safari = 'safari';
				this.Konqueror = 'kq';
				
				//Please offer a better solution if you have one!
				var sType;
				var agt=navigator.userAgent.toLowerCase();

				if (agt.toLowerCase().indexOf('konqueror') != -1) 
					sType = this.Konqueror;
				else if (agt.toLowerCase().indexOf('opera') != -1) 
					sType = this.Opera;
				else if (agt.toLowerCase().indexOf('netscape') != -1) 
					sType = this.Netscape;
				else if (agt.toLowerCase().indexOf('msie') != -1)
					sType = this.InternetExplorer;
				else if (agt.toLowerCase().indexOf('safari') != -1)
					sType = 'safari';
				
				if (sType == null)
					sType = this.Mozilla;  
				
				this.type = sType;
				this.version = parseFloat(navigator.appVersion);
				
				var sAgent = navigator.userAgent.toLowerCase();
				if (this.type == this.InternetExplorer)
				{
					var temp=navigator.appVersion.split("MSIE");
					this.version=parseFloat(temp[1]);
				}
				if (this.type == this.Netscape)
				{
					var temp=sAgent.split("netscape");
					this.version=parseFloat(temp[1].split("/")[1]);	
				}

				//this.majorVersion = null;
				//this.minorVersion = null;
			}
			
			dnn_dom_browser.prototype.toString = function ()
			{
				return this.type + ' ' + this.version;
			}
			
			dnn_dom_browser.prototype.isType = function ()
			{
				for (var i=0; i<arguments.length; i++)
				{
					if (dnn.dom.browser.type == arguments[i])
						return true;
				}
				return false;
			}

			dnn_dom_browser.prototype.dependenciesLoaded = function()
			{
				return (typeof(dnn) != 'undefined' && typeof(dnn.dom) != 'undefined');
			}

			dnn_dom_browser.prototype.loadNamespace = function ()
			{
				if (this.isLoaded == false)
				{
					if (this.dependenciesLoaded())
					{
						dnn.dom.browser = this; 
						this.isLoaded = true;
						dnn.loadDependencies(this.pns, this.ns);
					}
				}	
			}
			
			//--- End dnn.dom.browser
			

	dnn_dom.prototype.attachEvent = function (oCtl, sType, fHandler) 
	{
		if (dnn.dom.browser.isType(dnn.dom.browser.InternetExplorer) == false)
		{
			var sName = sType.substring(2);
			oCtl.addEventListener(sName, function (evt) {dnn.dom.event = new dnn_dom_event(evt, evt.target); return fHandler();}, false);
		}
		else
			oCtl.attachEvent(sType, function () {dnn.dom.event = new dnn_dom_event(window.event, window.event.srcElement); return fHandler();});
		return true;
	}
	
	function dnn_dom_event(e, srcElement)
	{
		this.object = e;
		this.srcElement = srcElement;
	}
			
	//--- End dnn.dom


//--- End dnn

//load namespaces
__dnn_m_aNamespaces[__dnn_m_aNamespaces.length] = new dnn_dom_browser();
__dnn_m_aNamespaces[__dnn_m_aNamespaces.length] = new dnn_dom();
__dnn_m_aNamespaces[__dnn_m_aNamespaces.length] = new __dnn();
for (var i=__dnn_m_aNamespaces.length-1; i>=0; i--)
	__dnn_m_aNamespaces[i].loadNamespace();
