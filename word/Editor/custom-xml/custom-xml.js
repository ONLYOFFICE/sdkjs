/*
 * (c) Copyright Ascensio System SIA 2010-2023
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function()
{
	/**
	 * @param {AscWord.CustomXmlManager} xmlManager
	 * @param {string} [itemId=null]
	 * @param {CustomXmlPrefixMappings} [nsManager=null]
	 * @param {CustomXmlContent} [content]
	 *
	 * Класс представляющий CustomXML
	 * @constructor
	 */
	function CustomXml(xmlManager, itemId, nsManager, content)
	{
		this.Id			= AscCommon.g_oIdCounter.Get_NewId();
		this.Parent		= xmlManager;
		this.itemId		= itemId ? itemId : AscCommon.CreateGUID();
		this.content	= null;
		this.nsManager	= (nsManager && nsManager instanceof CustomXmlPrefixMappings)
			? nsManager
			: new CustomXmlPrefixMappings(nsManager);

		this.addContentByXMLString(content);
		this.m_aCustomXmlData = '';

		AscCommon.g_oTableId.Add(this, this.Id);
	}
	CustomXml.prototype.Copy = function ()
	{
		let strXml = this.getText();
		let oCopy = new CustomXml(
			this.Parent,
			this.itemId,
			this.nsManager,
			undefined
		);

		oCopy.addContentByXMLString(strXml);

		return oCopy;
	};
	CustomXml.prototype.Delete = function ()
	{
		if (this.Parent)
			return this.Parent.deleteExactXml(this.itemId);
		
		return false;
	};
	CustomXml.prototype.Get_Id = function ()
	{
		return this.Id;
	};
	CustomXml.prototype.GetId = function()
	{
		return this.Id;
	};
	CustomXml.prototype.Refresh_RecalcData = function(Data)
	{
		// Ничего не делаем (если что просто будет перерисовка)
	};
	CustomXml.prototype.Write_ToBinary2 = function(Writer)
	{
		Writer.WriteLong(AscDFH.historyitem_type_CustomXml);

		// String : Id
		// Long   : Количество элементов
		// Array of Strings : массив с Id элементов

		Writer.WriteString2(this.Id);
		Writer.WriteString2(this.itemId);
		this.nsManager.Write_ToBinary2(Writer);
	};
	CustomXml.prototype.Read_FromBinary2 = function(Reader)
	{
		// String : Id
		// Long   : Количество элементов
		// Array of Strings : массив с Id элементов
		this.Id = Reader.GetString2();
		this.itemId = Reader.GetString2();
		this.Parent = editor.WordControl.m_oLogicDocument.getCustomXmlManager();
		this.nsManager.Read_FromBinary2(Reader);
	};
	CustomXml.prototype.Write_ToBinary = function(Writer)
	{
		this.Write_ToBinary2(Writer);
	};
	CustomXml.prototype.Read_FromBinary = function (Reader)
	{
		this.Read_FromBinary2(Reader);
	};
	CustomXml.prototype.writeContent = function(strPrevXml, strCustomXml)
	{
		const BINARY_PART_HISTORY_LIMIT = 1048576;
		const maxLen = Math.max(strCustomXml.length, strPrevXml.length);
		const oldParts = [];
		const newParts = [];
		const amountOfParts = Math.ceil(maxLen / BINARY_PART_HISTORY_LIMIT);
		for (let i = 0; i < amountOfParts; i += 1) {
			oldParts.push(strPrevXml.slice(i * BINARY_PART_HISTORY_LIMIT, (i + 1) * BINARY_PART_HISTORY_LIMIT));
			newParts.push(strCustomXml.slice(i * BINARY_PART_HISTORY_LIMIT, (i + 1) * BINARY_PART_HISTORY_LIMIT));
		}

		AscCommon.History.Add(new AscDFH.CChangesCustomXmlContentStart(this, null, null, false));
		for (let i = 0; i < amountOfParts; i += 1) {
			AscCommon.History.Add(new AscDFH.CChangesCustomXmlContentPart(this, oldParts[i], newParts[i], false));
		}
		AscCommon.History.Add(new AscDFH.CChangesCustomXmlContentEnd(this, null, null, false));
		this.m_aBinaryData = strCustomXml;
	};
	/**
	 * Set UID of CustomXML by given data
	 * @param itemId {string}
	 */
	CustomXml.prototype.setItemId = function (itemId)
	{
		this.itemId = itemId;
	};
	/**
	 * Add given uri to CustomXMl uri list
	 * @param {string} prefix
	 * @param {string} ns
	 */
	CustomXml.prototype.addNamespace = function(prefix, ns)
	{
		if (!prefix)
			prefix = "defaultNamespace";

		this.nsManager.addNamespace(prefix, ns);
		return true;
	};
	/**
	 * Get CustomXML data by string
	 * @return {string}
	 */
	CustomXml.prototype.getText = function ()
	{
		if (this.content)
		return this.content.getStringFromBuffer();
		else
			return "";
	};
	/**
	 * Find url in uri array
	 * @return {boolean}
	 */
	CustomXml.prototype.checkUrl = function (str)
	{
		if (!str)
			return false;

		return this.nsManager.getPrefix(str) !== undefined;
	}
	/**
	 * Add content of CustomXML
	 * @param arrData {array}
	 */
	CustomXml.prototype.addContent = function (arrData)
	{
		let customXml = fromUtf8(arrData);
		let startPos = customXml.indexOf("<");
		if (-1 !== startPos)
			customXml = customXml.slice(customXml.indexOf("<")); // Skip "L"
		
		this.addContentByXMLString(customXml);
	};
	CustomXml.prototype.addContentByXMLString = function(strCustomXml)
	{
		if (strCustomXml === undefined)
			return;
		
		if (strCustomXml instanceof CustomXmlContent)
			strCustomXml = strCustomXml.getStringFromBuffer();

		this.content = CustomXmlCreateContent(strCustomXml, this);
	};
	CustomXml.prototype.findElementByXPath = function (xpath) {
		function getDescendantsByTagName(node, tagName) {
			let result = [];

			function recurse(current) {
				for (let i = 0; i < current.childNodes.length; i++)
				{
					let child = current.childNodes[i];
					let nameParts = child.nodeName.split(':');
					let pureName = nameParts.length > 1 ? nameParts[1] : nameParts[0];

					if (pureName === tagName || tagName === '*')
						result.push(child);

					if (child.childNodes && child.childNodes.length)
						recurse(child);
				}
			}

			recurse(node);
			return result;
		}

		function parseSteps(xpath) {
			let steps = [];

			const marker = '#DOUBLE_SLASH#';
			let temp = xpath.replace(/\/\//g, function() {return '/' + marker + '/'});
			let parts = temp.split('/').filter(function(el) {return el !== ""});
			let deepNext = false;

			for (let i = 0; i < parts.length; i++) {
				let part = parts[i];
				if (part === marker) {
					deepNext = true;
					continue;
				}
				steps.push({part: part, deepNext: deepNext});
				deepNext = false;
			}
			return steps;
		}

		function findMatchingNodes(elements, steps) {
			
			if (steps.length === 0)
				return elements;
			
			let step = steps[0];
			let part = step.part;
			let deep = step.deepNext;
			let restSteps = steps.slice(1);

			let tagName = null;
			let index = null;

			if (part[0] === '@')
			{
				let attrName = part.slice(1);
				return elements.map(function (el) {
					if (el.getAttribute(attrName))
						return el
				});
			}

			if (part === '*')
			{
				tagName = '*';
			}
			else
			{
				let bracketOpen = part.indexOf('[');
				let bracketClose = part.indexOf(']');
			
				if (bracketOpen !== -1 && bracketClose !== -1 && bracketClose > bracketOpen)
				{
					let rawTag		= part.slice(0, bracketOpen);
					let rawIndex	= part.slice(bracketOpen + 1, bracketClose);
					tagName			= rawTag.includes(':') ? rawTag.split(':')[1] : rawTag;
					index			= parseInt(rawIndex, 10) - 1;
				}
				else
				{
					tagName = part.includes(':') ? part.split(':')[1] : part;
				}
			}

			let matched = [];

			for (let i = 0; i < elements.length; i++)
			{
				let el = elements[i];
				let candidates = deep
					? getDescendantsByTagName(el, tagName)
					: el.childNodes.filter(function(child) {
						let nameParts	= child.nodeName.split(':');
						let pureName	= nameParts.length > 1 ? nameParts[1] : nameParts[0];
						return tagName === '*' || pureName === tagName;
					});

				if (index !== null)
				{
					if (index >= 0 && index < candidates.length)
					{
						matched.push(candidates[index]);
					}
				}
				else
				{
					matched = matched.concat(candidates);
				}
			}

			return findMatchingNodes(matched, restSteps);
		}

		let steps = parseSteps(xpath);
		let result = findMatchingNodes([this.content], steps);
		return result;
	};
	CustomXml.prototype.deleteAttribute = function(xPath, name)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				return el.deleteAttribute(name);
			}
			return false;
		}, this);
	};
	CustomXml.prototype.insertAttribute = function(xPath, name, value)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				if (!el.attributes[name])
					return el.setAttribute(name, value);
			}
			return false;
		}, this);
	};
	CustomXml.prototype.updateAttribute = function(xPath, name, value)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				if (el.attributes[name])
					return el.setAttribute(name, value);
			}
			return false;
		}, this);
	};
	CustomXml.prototype.getAttribute = function(xPath, name)
	{
		let nodes = this.findElementByXPath(xPath);
		if (nodes.length)
		{
			let el = nodes[0];
			if (el.attributes[name])
				return el.attributes[name];
		}
		return null;
	};
	CustomXml.prototype.deleteElement = function (xPath)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				return el.delete();
			}
			return false;
		}, this);
	};
	CustomXml.prototype.updateElement = function(xPath, xmlStr)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				el.setXml(xmlStr);
				return true;
			}
			return false;
		}, this);
	};
	CustomXml.prototype.insertElement = function (xPath, xmlStr, index)
	{
		return this.Change(function (){
			let nodes = this.findElementByXPath(xPath);
			if (nodes.length)
			{
				let el = nodes[0];
				el.addElement(xmlStr, index);
				return true;
			}
			return false;
		}, this);
	};
	CustomXml.prototype.beforeChange = function ()
	{
		this.lastContent = this.Copy();
	};
	CustomXml.prototype.afterChange = function ()
	{
		let strLast			= this.lastContent.getText();
		let strCurrentData	= this.getText();

		if (strLast !== strCurrentData)
			this.writeContent(strLast, strCurrentData);

		this.lastContent = undefined;
	};
	CustomXml.prototype.Change = function (f, oThis)
	{
		this.beforeChange();
		let data = f.apply(oThis);
		this.afterChange();
		return data;
	};
	CustomXml.prototype.setNamespaceUri = function(ns)
	{
		this.nsManager.setNamespaceUri(ns);
	};
	CustomXml.prototype.getNamespaceUri = function()
	{
		return this.nsManager.getNamespaceUri();
	};
	CustomXml.prototype.getAllNamespaces = function()
	{
		return Object.keys(this.nsManager.urls);
	};

	/**
	 * @constructor
	 */
	function CustomXmlContent(oParentNode, oNodeName, xml)
	{
		this.parentNode = oParentNode;
		this.nodeName = oNodeName ? oNodeName : "";
		this.childNodes = [];
		this.attributes = {};
		this.text = "";
		this.xmlQuestionHeader = null;
		this.xml = xml;

		this.getNamespace = function ()
		{
			const parts = this.nodeName.split(":");
			let prefix = parts.length === 2 ? parts[0] : null;

			if (prefix)
				return this.xml.nsManager.getNamespace(prefix);

			return "";
		}
		this.getNodeName = function ()
		{
			return this.nodeName;
		};
		this.getText = function ()
		{
			return this.text;
		};
		this.getXPath = function ()
		{
			let parent		= this.getParent();
			let parentText	= parent ? parent.getXPath() + "/" : "/";
			let curNodeName = this.getNodeName();
			let count		= parent ? parent.getPosOfChilderNode(this) : null;
			let textOfCount = count !== null ? "[" + count + "]" : "";

			if (!curNodeName)
				return '';

			return parentText + curNodeName + textOfCount;
		};
		this.getChildNodesCount = function()
		{
			return this.childNodes.length;
		};
		this.getPosOfChilderNode = function(childNode)
		{
			if (!childNode || !(childNode instanceof CustomXmlContent) || this.childNodes.length <= 1)
				return null;

			let childNameNodes = this.childNodes.filter(function(node){return node.nodeName === childNode.nodeName});

			if (childNameNodes.length <= 1)
				return null;

			for (let i = 0; i < childNameNodes.length; i++)
			{
				let node = childNameNodes[i];
				if (node === childNode)
					return i + 1;
			}
		};
		this.deleteChild = function (childNode)
		{
			this.childNodes = this.childNodes.filter(function (item) {return item !== childNode});
		};
		this.delete = function ()
		{
			this.parentNode.deleteChild(this);
		};
		this.addAttribute = function (name, value)
		{
			this.attributes[name] = value;
		};
		this.getAttribute = function (name)
		{
			return this.attributes[name];
		};
		this.addContent = function(name)
		{
			let newItem = new CustomXmlContent(this, name, this.xml);

			this.childNodes.push(newItem);
			return newItem;
		};
		this.addElement = function(xmlStr, index)
		{
			let newItem = new CustomXmlContent(this, null, this.xml);
			newItem.setXml(xmlStr);

			if (index !== undefined)
				this.childNodes.splice(index, 0, newItem);
			else
				this.childNodes.push(newItem);
		};
		this.setAttribute = function (attribute, value)
		{
			this.attributes[attribute] = value;
		};
		this.deleteAttribute = function(name)
		{
			if (this.attributes[name])
			{
				delete this.attributes[name];
				return true;
			}

			return false;
		};
		this.getParent = function ()
		{
			if (this.parentNode)
				return this.parentNode;

			return null;
		};
		this.addTextContent = function (text)
		{
			if (text !== "")
				this.text += text;
		};
		this.setTextContent = function (str)
		{
			this.text = str;
			return true;
		};
		this.getInnerText = function ()
		{
			let result = [];

			function GetText(node)
			{
				if (node.text)
					result.push(node.text);

				for (let i = 0; i < node.childNodes.length; i++)
				{
					GetText(node.childNodes[i]);
				}
			}

			GetText(this);

			return result.join("");
		}
		this.getBuffer = function ()
		{
			let writer = new AscCommon.CMemory();

			function Write(content)
			{
				if (content.nodeName === "" && content.text === "" && content.childNodes.length === 0)
				{
					writer.WriteXmlString("");
					return;
				}

				let current;
				if (!content.nodeName)
				{
					if (content.xmlQuestionHeader !== null)
						writer.WriteXmlString(content.xmlQuestionHeader + "\n");

					current = content.childNodes[0];
				}
				else
				{
					current = content;
				}

				writer.WriteXmlNodeStart(current.nodeName);

				let atr = Object.keys(current.attributes)

				for (let i = 0; i < atr.length; i++)
				{
					let cur = atr[i];
					writer.WriteXmlAttributeStringEncode(cur, current.attributes[cur]);
				}

				writer.WriteXmlAttributesEnd();

				for (let i = 0; i < current.childNodes.length; i++)
				{
					Write(current.childNodes[i]);
				}

				if (current.text)
					writer.WriteXmlString(current.text.toString().trim());

				writer.WriteXmlNodeEnd(current.nodeName);
			}

			Write(this);
			return writer;
		};
		this.getStringFromBuffer = function (isOnlyInner)
		{
			let buffer	= this.getBuffer(isOnlyInner);
			let str		= fromUtf8(buffer.GetData());
			str			= str.replaceAll("&quot;", "\"");
			str			= str.replaceAll("&amp;", "&");
			return str;
		};
		this.setXml = function(strXml)
		{
			let content = CustomXmlCreateContent(strXml, this.xml);
			let data = content.childNodes[0];

			this.nodeName	= data.nodeName;
			this.childNodes	= data.childNodes;
			this.attributes	= data.attributes;
			this.text		= data.text;

			if (data.xmlQuestionHeader)
				this.xmlQuestionHeader = data.xmlQuestionHeader;
		};
		this.Write_ToBinary2 = function (Writer)
		{
			Writer.WriteString2( this.getStringFromBuffer() );
		};
		this.Read_FromBinary2 = function (Reader)
		{
			let oContent = AscWord.CustomXmlCreateContent(Reader.GetString2(), this.xml);

			this.parentNode			= oContent.parentNode;
			this.name				= oContent.name;
			this.childNodes			= oContent.childNodes;
			this.attributes			= oContent.attributes;
			this.textContent		= oContent.textContent;
			this.xmlQuestionHeader	= oContent.xmlQuestionHeader;
		};
	}

	// TODO: Временно вынес метод сюда, потом перенести надо будет
	// разница с AscCommon.UTF8ArrayToString, что тут на 0-символе не останавливаемся
	function fromUtf8(buffer, start, len)
	{
		if (undefined === start)
			start = 0;
		if (undefined === len)
			len = buffer.length;
		
		var result = "";
		var index  = start;
		var end = start + len;
		while (index < end)
		{
			var u0 = buffer[index++];
			if (!(u0 & 128))
			{
				result += String.fromCharCode(u0);
				continue;
			}
			var u1 = buffer[index++] & 63;
			if ((u0 & 224) == 192)
			{
				result += String.fromCharCode((u0 & 31) << 6 | u1);
				continue;
			}
			var u2 = buffer[index++] & 63;
			if ((u0 & 240) == 224)
				u0 = (u0 & 15) << 12 | u1 << 6 | u2;
			else
				u0 = (u0 & 7) << 18 | u1 << 12 | u2 << 6 | buffer[index++] & 63;
			if (u0 < 65536)
				result += String.fromCharCode(u0);
			else
			{
				var ch = u0 - 65536;
				result += String.fromCharCode(55296 | ch >> 10, 56320 | ch & 1023);
			}
		}
		return result;
	}
	function CustomXmlCreateContent(strCustomXml, xml)
	{
		function getPrefix(xmlnsAttrName) {
			if (xmlnsAttrName.startsWith("xmlns:"))
				return xmlnsAttrName.slice(6);
			return null;
		}

		let nXmlHeaderStart	= strCustomXml.indexOf('<?', 0);
		let nXmlHeaderEnd	= strCustomXml.indexOf('?>', nXmlHeaderStart);
		let strXmlHeader 	= null;
		if (nXmlHeaderStart !== -1 && nXmlHeaderEnd !== -1)
		{
			strXmlHeader = strCustomXml.substring(nXmlHeaderStart, nXmlHeaderEnd + "?>".length);
			strCustomXml = strCustomXml.substring(nXmlHeaderEnd + '?>'.length, strCustomXml.length);
		}

		let oStax			= new StaxParser(strCustomXml),
			rootContent		= new CustomXmlContent(null, null, xml);

		if (strXmlHeader !== null)
			rootContent.xmlQuestionHeader = strXmlHeader;

		while (oStax.Read())
		{
			switch (oStax.GetEventType()) {
				case EasySAXEvent.CHARACTERS:
					rootContent.addTextContent(oStax.text);
					break;
				case EasySAXEvent.END_ELEMENT:
					rootContent = rootContent.getParent();
					break;
				case EasySAXEvent.START_ELEMENT:
					let name = oStax.GetName();
					let childElement = rootContent.addContent(name);

					while (oStax.MoveToNextAttribute())
					{
						let attributeName = oStax.GetName();
						let attributeValue = oStax.GetValue();

						if (attributeName === 'xmlns' && rootContent.xml)
							rootContent.xml.addNamespace("defaultNamespace", attributeValue);

						let prefix = getPrefix(attributeName);
						if (prefix && rootContent.xml)
							rootContent.xml.addNamespace(prefix, attributeValue);

						childElement.addAttribute(attributeName, attributeValue);
					}
					rootContent = childElement;break;
			}
		}
		return rootContent;
	}
	function CustomXmlPrefixMappings()
	{
		this.urls = {};
		this.prefix = {};
		this.namespaceUri = "";

		this.setNamespaceUri = function(ns)
		{
			this.namespaceUri = ns;
		};
		this.getNamespaceUri = function()
		{
			return this.namespaceUri;
		};
		this.addNamespace = function (prefix, ns)
		{
			if (ns !== "" && this.namespaceUri === "")
				this.namespaceUri = ns;

			if (prefix && ns)
			{
				let prevPrefix = this.urls[ns];

				this.urls[ns] = prefix;

				if (prevPrefix)
					delete this.prefix[prevPrefix];

				this.prefix[prefix] = ns;
			}
		};
		this.getNamespace = function(prefix)
		{
			return this.prefix[prefix];
		};
		this.getPrefix = function(urls)
		{
			return this.urls[urls];
		};
		this.Write_ToBinary2 = function(Writer)
		{
			Writer.WriteString2(this.namespaceUri);

			let urls = Object.keys(this.urls);
			let Count = urls.length;
			Writer.WriteLong(Count);

			for (let Index = 0; Index < Count; Index++)
				Writer.WriteString2(urls[Index]);

			Writer.WriteLong(Count);
			for (let Index = 0; Index < Count; Index++)
				Writer.WriteString2(this.urls[urls[Index]]);
		};
		this.Read_FromBinary2 = function(Reader)
		{
			this.namespaceUri = Reader.GetString2()
			let url = [];

			let Count = Reader.GetLong();
			for (let Index = 0; Index < Count; Index++)
			{
				url.push(Reader.GetString2());
				this.urls[url[url.length - 1]] = "";
			}
			
			Count = Reader.GetLong();
			for (let Index = 0; Index < Count; Index++)
			{
				let prefix = Reader.GetString2()
				this.prefix[prefix] = url[Index];
				this.urls[url[Index]] = prefix;
			}
		};
	}

	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CustomXml					= CustomXml;
	AscWord.CustomXmlContent			= CustomXmlContent;
	AscWord.CustomXmlCreateContent		= CustomXmlCreateContent;
})();
