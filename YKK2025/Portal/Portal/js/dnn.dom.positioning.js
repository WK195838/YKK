if (typeof(__dnn_m_aNamespaces) == 'undefined')	//include in each DNN ClientAPI namespace file for dependency loading
	var __dnn_m_aNamespaces = new Array();

function dnn_dom_positioning()
{
	this.pns = 'dnn.dom';
	this.ns = 'positioning';
	this.dragCtr=null;
	this.dragCtrDims=null;
	this.dependencies = 'dnn,dnn.dom'.split(',');
	this.isLoaded = false;
	
}

dnn_dom_positioning.prototype.dims = function (eSrc)
{
	var bHidden = (eSrc.style.display == 'none');
	
	if (bHidden)
		eSrc.style.display = "";
	this.w = dnn.dom.positioning.elementWidth(eSrc);
	this.h = dnn.dom.positioning.elementHeight(eSrc);
	this.t = dnn.dom.positioning.elementTop(eSrc);
	this.l = dnn.dom.positioning.elementLeft(eSrc);
	this.r = this.l + this.w;
	this.b = this.t + this.h;
	
	if (bHidden)
		eSrc.style.display = "none";
}


dnn_dom_positioning.prototype.elementTop = function (eSrc)
{
	var iTop = 0;
	var eParent;
	eParent = eSrc;
	
	while (eParent.tagName.toUpperCase() != "BODY")
	{
		//Safari incorrectly calculates the TR tag to be at the top of the table, so try and get child TD tag to use for measurement
		if (dnn.dom.browser.type == dnn.dom.browser.Safari && eParent.tagName.toUpperCase() == 'TR' && dnn.dom.getByTagName('TD', eParent).length)
			eParent = dnn.dom.getByTagName('TD', eParent)[0];


		iTop += eParent.offsetTop;
		eParent = eParent.offsetParent;
		if (eParent == null)
			break;
	}
	if (eParent != null && (dnn.dom.browser.type == dnn.dom.browser.Safari || dnn.dom.browser.type == dnn.dom.browser.Konqueror)) 
		iTop += eParent.offsetTop;
	
	return iTop;
}

dnn_dom_positioning.prototype.elementLeft = function (eSrc)
{	
	var iLeft = 0;
	var eParent;
	eParent = eSrc;
	while (eParent.tagName.toUpperCase() != "BODY")
	{
		iLeft += eParent.offsetLeft;
		eParent = eParent.offsetParent;
		if (eParent == null)
			break;
	}
	if (eParent != null && (dnn.dom.browser.type == dnn.dom.browser.Safari || dnn.dom.browser.type == dnn.dom.browser.Konqueror)) 
		iLeft += eParent.offsetLeft;
	
	return iLeft;
}

dnn_dom_positioning.prototype.elementHeight = function (eSrc)
{	
	if (eSrc.offsetHeight == null || eSrc.offsetHeight == 0)
	{
		if (eSrc.offsetParent.offsetHeight == null || eSrc.offsetParent.offsetHeight == 0)
		{
			if (eSrc.offsetParent.offsetParent != null)
				return eSrc.offsetParent.offsetParent.offsetHeight; //needed for Konqueror
			else
				return 0;
		}
		else
			return eSrc.offsetParent.offsetHeight;
	}
	else
		return eSrc.offsetHeight;
}

dnn_dom_positioning.prototype.elementWidth = function (eSrc)
{
	if (eSrc.offsetWidth == null || eSrc.offsetWidth == 0)
	{
		if (eSrc.offsetParent.offsetWidth == null || eSrc.offsetParent.offsetWidth == 0)
		{
			if (eSrc.offsetParent.offsetParent != null)
				return eSrc.offsetParent.offsetParent.offsetWidth; //needed for Konqueror
			else
				return 0;
		}
		else
			return eSrc.offsetParent.offsetWidth

	}
	else
		return eSrc.offsetWidth;
}

dnn_dom_positioning.prototype.bodyScrollTop = function()
{
	//if ('|ie|op|mo|ns|'.indexOf('|' + dnn.() + '|') != -1)
	if (dnn.dom.browser.isType(dnn.dom.browser.InternetExplorer, dnn.dom.browser.Opera, dnn.dom.browser.Mozilla, dnn.dom.browser.Netscape))
	{
		if (document.body.scrollTop != null)
			return document.body.scrollTop;
	}
	return 0;
}

dnn_dom_positioning.prototype.bodyScrollLeft = function ()
{
	//if ('|op|'.indexOf('|' + spm_browserType() + '|') != -1)
	if (dnn.dom.browser.isType(dnn.dom.browser.InternetExplorer, dnn.dom.browser.Opera, dnn.dom.browser.Mozilla, dnn.dom.browser.Netscape))
	{
		if (document.body.scrollLeft != null)
			return document.body.scrollLeft;
	}
	return 0;
}

dnn_dom_positioning.prototype.viewPortWidth = function()
{
	if(window.innerWidth)	// supported in Mozilla, Opera, and Safari
		return window.innerWidth;
	if(window.document.documentElement.clientWidth) // supported in standards mode of IE, but not in any other mode
		return document.documentElement.clientWidth;

	return window.document.body.clientWidth;	// supported in quirks mode, older versions of IE, and mac IE (anything else).
}

dnn_dom_positioning.prototype.viewPortHeight = function()
{
    if(window.innerHeight)	// supported in Mozilla, Opera, and Safari
		return window.innerHeight;
    if(window.document.documentElement.clientHeight)	// supported in standards mode of IE, but not in any other mode
		return document.documentElement.clientHeight;
    
    return window.document.body.clientHeight;	// supported in quirks mode, older versions of IE, and mac IE (anything else).
}


dnn_dom_positioning.prototype.elementOverlapScore = function (oDims1, oDims2)
{		
	var iLeftScore = 0;
	var iTopScore = 0;
	if (oDims1.l <= oDims2.l && oDims2.l <= oDims1.r)	//if left of content fits between panel borders
		iLeftScore += (oDims1.r < oDims2.r ? oDims1.r : oDims2.r) - oDims2.l;	//set score based off left of content to closest right border
	if (oDims2.l <= oDims1.l && oDims1.l <= oDims2.r)	//if left of panel fits between content borders
		iLeftScore += (oDims2.r < oDims1.r ? oDims2.r : oDims1.r)  - oDims1.l; //set score based off left of panel to closest right border
	if (oDims1.t <= oDims2.t && oDims2.t <= oDims1.b)	//if top of content fits between panel borders
		iTopScore += (oDims1.b < oDims2.b ? oDims1.b : oDims2.b)  - oDims2.t;	//set score based off top of content to closest bottom border
	if (oDims2.t <= oDims1.t && oDims1.t <= oDims2.b)	//if top of panel fits between content borders
		iTopScore += (oDims2.b < oDims1.b ? oDims2.b : oDims1.b) -  - oDims1.t; //set score based off top of panel to closest bottom border
	
	return iLeftScore * iTopScore;
}

function __isWindows2003()	//quickfix for bug in Windows2003 IE - when an item is set to position=relative it doubles the top position
{
	return navigator.userAgent.toLowerCase().indexOf('nt 5.2') > -1;
}

dnn_dom_positioning.prototype.enableDragAndDrop = function(oContainer, oTitle, sDragCompleteEvent, sDragOverEvent)
{

	dnn.dom.attachEvent(document.body, 'onmousemove', __dnn_bodyMouseMove);
	dnn.dom.attachEvent(document.body, 'onmouseup', __dnn_bodyMouseUp);
	dnn.dom.attachEvent(oTitle, 'onmousedown', __dnn_containerMouseDown);
	
	if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
		oTitle.style.cursor = 'hand';
	else
		oTitle.style.cursor = 'pointer';
	
	if (__isWindows2003() == false)
	{
		if (oContainer.style.position == null || oContainer.style.position.length == 0)
			oContainer.style.position = 'relative';
	}
	
	if (oContainer.id.length == 0)
		oContainer.id = oTitle.id + '__dnnCtr';
		
	oTitle.contID = oContainer.id;
	if (sDragCompleteEvent != null)
		oTitle.dragComplete = sDragCompleteEvent;
	if (sDragOverEvent != null)
		oTitle.dragOver = sDragOverEvent;
	
	return true;
}

dnn_dom_positioning.prototype.dragContainer = function(oCtl)
{
	var iNewLeft=0;
	var iNewTop=0;
	var e = dnn.dom.event.object;
	var oCont = dnn.dom.getById(oCtl.contID);
	var oTitle = dnn.dom.positioning.dragCtr;
	//var oDims = dnn.dom.positioning.dragCtrDims;
	var iScrollTop = this.bodyScrollTop();
	var iScrollLeft = this.bodyScrollLeft();

	if (oCtl.startLeft == null)
		oCtl.startLeft = e.clientX - this.elementLeft(oCont) + iScrollLeft;

	if (oCtl.startTop == null)
		oCtl.startTop = e.clientY - this.elementTop(oCont) + iScrollTop;

	if (oCont.style.position == 'relative' || __isWindows2003())
		oCont.style.position = 'absolute';
	
	iNewLeft = e.clientX - oCtl.startLeft + iScrollLeft;
	iNewTop = e.clientY - oCtl.startTop + iScrollTop;

	if (iNewLeft > this.elementWidth(document.forms[0]))// this.viewPortWidth() + iScrollLeft)
		iNewLeft = this.elementWidth(document.forms[0]);//this.viewPortWidth() + iScrollLeft;
	
	if (iNewTop > this.elementHeight(document.forms[0])) //this.viewPortHeight() + iScrollTop)
		iNewTop = this.elementHeight(document.forms[0]);//this.viewPortHeight() + iScrollTop;
	
	oCont.style.left = iNewLeft;
	oCont.style.top = iNewTop;

	if (oTitle != null && oTitle.dragOver != null)
		eval(oCtl.dragOver);
}

dnn_dom_positioning.prototype.dependenciesLoaded = function()
{
	return (typeof(dnn) != 'undefined' && typeof(dnn.dom) != 'undefined');
}

dnn_dom_positioning.prototype.loadNamespace = function ()
{
	if (this.isLoaded == false)
	{		
		if (this.dependenciesLoaded())
		{
			dnn.dom.positioning = this; 
			this.isLoaded = true;
			dnn.loadDependencies(this.pns, this.ns);
		}
	}	
}

function __dnn_cancelEvent()
{
	return false;
}

function __dnn_containerMouseDown()
{
	oCtl = dnn.dom.event.srcElement;
	while (oCtl.contID == null)
	{
		oCtl = oCtl.parentNode;
		if (oCtl.tagName.toUpperCase() == 'BODY')
			return;
	}
	dnn.dom.attachEvent(oCtl, 'onselectstart', __dnn_cancelEvent);
	
	dnn.dom.positioning.dragCtr = oCtl;	//assumption is we can only drag one thing at a time
	oCtl.startTop = null;
	oCtl.startLeft = null;

	var oCont = dnn.dom.getById(oCtl.contID);
	dnn.dom.positioning.dragCtrDims = new dnn.dom.positioning.dims(oCont);	//store now so we aren't continually calculating
	
	if (oCont.getAttribute('_b') == null)
	{
		oCont.setAttribute('_b', oCont.style.backgroundColor); 
		oCont.setAttribute('_z', oCont.style.zIndex); 
		oCont.setAttribute('_w', oCont.style.width); 
		oCont.setAttribute('_d', oCont.style.border); 
		oCont.style.zIndex = 9999;
		oCont.style.backgroundColor = DNN_HIGHLIGHT_COLOR;
		oCont.style.border = '4px outset ' + DNN_HIGHLIGHT_COLOR;
		oCont.style.width = dnn.dom.positioning.elementWidth(oCont);
		if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
			oCont.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=80)';
	}
}


function __dnn_bodyMouseUp()
{
	var oCtl = dnn.dom.positioning.dragCtr;
	if (oCtl != null && oCtl.dragComplete != null)
	{
		eval(oCtl.dragComplete);

		var oCont = dnn.dom.getById(oCtl.contID);

		oCont.style.backgroundColor = oCont.getAttribute('_b'); 
		oCont.style.zIndex = oCont.getAttribute('_z'); 
		oCont.style.width = oCont.getAttribute('_w'); 
		oCont.style.border = oCont.getAttribute('_d'); 
		oCont.setAttribute('_b', null); 
		oCont.setAttribute('_z', null); 
		if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
			oCont.style.filter = null;
		
	}
	
	dnn.dom.positioning.dragCtr = null;
}

function __dnn_bodyMouseMove()
{
	if (dnn.dom.positioning.dragCtr != null)
		dnn.dom.positioning.dragContainer(dnn.dom.positioning.dragCtr);
}

__dnn_m_aNamespaces[__dnn_m_aNamespaces.length] = new dnn_dom_positioning();
for (var i=__dnn_m_aNamespaces.length-1; i>=0; i--)
	__dnn_m_aNamespaces[i].loadNamespace();
