//Currently only supports Netscape, Mozilla, and Internet Explorer
if (typeof(__dnn_m_aNamespaces) == 'undefined')	//include in each DNN ClientAPI namespace file for dependency loading
	var __dnn_m_aNamespaces = new Array();

function dnn_xml()
{
	this.pns = 'dnn';
	this.ns = 'xml';
	this.dependencies = 'dnn'.split(',');
	this.isLoaded = false;
}

dnn_xml.prototype.createDocument = function()
{
	if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
	{
		var o = new ActiveXObject('MSXML.DOMDocument');
		o.async = false;
		return new dnn_xml_document(o);
	}
	else
		return new dnn_xml_document(document.implementation.createDocument("", "", null));
}

dnn_xml.prototype.dependenciesLoaded = function()
{
	return (typeof(dnn) != 'undefined');
}

dnn_xml.prototype.loadNamespace = function ()
{
	if (this.isLoaded == false)
	{
		if (this.dependenciesLoaded())
		{
			dnn.xml = this; 
			this.isLoaded = true;
			dnn.loadDependencies(this.pns, this.ns);
			__dnn_xml_init();
		}
	}	
}


function dnn_xml_document(oDoc)
{
	this._doc = oDoc;
	//if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
	//{
		this.childNodes = this._doc.childNodes;
	//}
}

function __dnn_xml_init()
{
	//Netscape / Mozilla implementation of loadXML
	if (dnn.dom.browser.type == dnn.dom.browser.InternetExplorer)
	{
		dnn_xml_document.prototype.loadXML = function (sXML) 
		{
			return this._doc.loadXML(sXML);
		}
	}

	if (dnn.dom.browser.type != dnn.dom.browser.Opera && typeof(Document) != 'undefined')
	{
		dnn_xml_document.prototype.loadXML = function (sXML) 
		{
			// parse the string to a new doc
			var oDoc = (new DOMParser()).parseFromString(sXML, "text/xml");
			    
			// remove all initial children
			while (this._doc.hasChildNodes())
				this._doc.removeChild(this._doc.lastChild);

			// insert and import nodes
			for (var i = 0; i < oDoc.childNodes.length; i++) 
				this._doc.appendChild(this._doc.importNode(oDoc.childNodes[i], true));
		}

		function _Node_getXML() 
		{
			//create a new XMLSerializer
			var oXMLSerializer = new XMLSerializer;
			    
			//get the XML string
			var sXML = oXMLSerializer.serializeToString(this);
			    
			//return the XML string
			return sXML;
		}
		Node.prototype.__defineGetter__("xml", _Node_getXML);
	}
}

dnn_xml_document.prototype.applyStyle = function (sXSL)
{
	if (typeof(this._doc.transformNode) == 'undefined')
		return null;
		
	var oXML = dnn.xml.createDocument();
	oXML.loadXML(sXSL);
	return this._doc.transformNode(oXML._doc);
}

__dnn_m_aNamespaces[__dnn_m_aNamespaces.length] = new dnn_xml();
for (var i=__dnn_m_aNamespaces.length-1; i>=0; i--)
	__dnn_m_aNamespaces[i].loadNamespace();
