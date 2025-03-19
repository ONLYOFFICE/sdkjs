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
	 * @param {array} [uri]
	 * @param {string} [itemId]
	 * @param {CustomXmlContent} [content]
	 * @param [oContentLink]
	 *
	 * Класс представляющий CustomXML
	 * @constructor
	 */
	function CustomXml(oParent, itemId, uri, content)
	{
		this.Id			= AscCommon.g_oIdCounter.Get_NewId();
		this.Parent		= oParent;
		this.uri		= uri ? uri : [];
		this.itemId		= itemId ? itemId : null;

		if (content instanceof CustomXmlContent)
			content		= content.getStringFromBuffer();
		else if (!content)
			content		= "";

		this.addContentByXMLString(content);

		if (this.itemId === null)
			this.setItemId();

		AscCommon.g_oTableId.Add(this, this.Id);
	}
	CustomXml.prototype.Copy = function ()
	{
		let strXml = this.getText();
		let oCopy = new CustomXml(
			this.Parent,
			this.itemId,
			this.uri,
			undefined,
		);

		oCopy.addContentByXMLString(strXml);

		return oCopy;
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
		Writer.WriteLong(AscDFH.historyitem_type_CustomXML_Add);

		// String : Id
		// Long   : Количество элементов
		// Array of Strings : массив с Id элементов

		Writer.WriteString2(this.Id);
		Writer.WriteString2(this.itemId);
		Writer.WriteString2(this.getText());

		var Count = this.uri.length;
		Writer.WriteLong(Count);

		for (var Index = 0; Index < Count; Index++)
			Writer.WriteString2(this.uri[Index]);
	};
	CustomXml.prototype.Read_FromBinary2 = function(Reader)
	{
		// String : Id
		// Long   : Количество элементов
		// Array of Strings : массив с Id элементов
		this.Id = Reader.GetString2();
		this.itemId = Reader.GetString2();
		this.addContentByXMLString(Reader.GetString2());
		this.Parent = editor.WordControl.m_oLogicDocument.getCustomXmlManager();

		var Count = Reader.GetLong();
		for (var Index = 0; Index < Count; Index++)
			this.uri.push(Reader.GetString2());
	};
	CustomXml.prototype.Write_ToBinary = function(Writer)
	{
		this.Write_ToBinary2(Writer);
	};
	CustomXml.prototype.Read_FromBinary = function (Reader) {
		this.Read_FromBinary2(Reader);
	};
	/**
	 * Set UID of CustomXML
	 * @param {uId} itemId
	 */
	CustomXml.prototype.setItemId = function ()
	{
		this.itemId = AscCommon.CreateGUID();
	};
	/**
	 * Set UID of CustomXML by given data
	 * @param itemUId {string}
	 */
	CustomXml.prototype.setItemIdManually = function (itemUId)
	{
		this.itemId = itemUId;
	};
	/**
	 * Add given uri to CustomXMl uri list
	 * @param uri {string}
	 */
	CustomXml.prototype.addUri = function(uri)
	{
		if (undefined !== uri && uri !== "")
			this.uri.push(uri);
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

		for (let i = 0; i < this.uri.length; i++)
		{
			if (str.includes(this.uri[i]))
				return true;
		}
		return false;
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
	CustomXml.prototype.addContentByXMLString = function (strCustomXml)
	{
		this.content = CustomXmlCreateContent(strCustomXml);
	};
	CustomXml.prototype.findElementByXPath = function (xpath)
	{
		let arrParts		= xpath.split('/');
		let currentElement	= this.content;

		arrParts.shift(); // Убираем пустой первый элемент

		for (let i = 0; i < arrParts.length; i++)
		{
			let namespaceAndTag,
				index,
				tagName,
				part = arrParts[i];

			if (part.includes("@"))
			{
				let strAttributeName		= part.slice(1);
				return {
					content: currentElement,
					attribute: strAttributeName,
				};
			}
			else if (part.includes("["))
			{
				namespaceAndTag				= part.split('[')[0];
				let partBeforeCloseBracket	= part.split(']')[0];
				index						= partBeforeCloseBracket.slice(-1) - 1;
			}
			else
			{
				namespaceAndTag				= part;
				index						= 0;
			}

			tagName = namespaceAndTag.includes(":")
				? namespaceAndTag.split(':')[1]
				: namespaceAndTag;

			let matchingChildren = currentElement.content.filter(function (child) {
				let arr = child.name.split(":");

				if (arr.length > 1)
					return arr[1] === tagName;
				else
					return arr[0] === tagName;
			});

			if (matchingChildren.length <= index)
				break; // Элемент не найден

			currentElement = matchingChildren[index];
		}

		return currentElement;
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
			AscCommon.History.Add(new CChangesCustomXMLChange(this, this.lastContent.content, this.content));

		this.lastContent = undefined;
	};
	CustomXml.prototype.addDataByPath = function (xPath, nodes, text, attributes)
	{
		this.beforeChange();

		let oParentNode = this.findElementByXPath(xPath);
		if (oParentNode)
			this.addData(oParentNode, nodes, text, attributes);

		this.afterChange();
	};
	CustomXml.prototype.addData = function (oParentNode, nodes, text, attributes)
	{
		//todo
	};

	/**
	 * @constructor
	 */
	function CustomXmlContent(parent, name)
	{
		this.parent				= parent;
		this.name				= name ? name : "";
		this.content			= [];
		this.attribute			= {};
		this.textContent		= "";
		this.xmlQuestionHeader	= null;

		this.addAttribute = function (name, value)
		{
			this.attribute[name] = value;
		};
		this.getAttribute = function (name)
		{
			return this.attribute[name];
		};
		this.addContent = function (name)
		{
			let newItem = new CustomXmlContent(this, name);

			this.content.push(newItem);
			return newItem;
		};
		this.getParent = function ()
		{
			if (this.parent)
				return this.parent;

			return null;
		};
		this.addTextContent = function (text)
		{
			if (text !== "")
				this.textContent += text;
		};
		this.setTextContent = function (str)
		{
			this.textContent = str;
		};
		this.setAttribute = function (attribute, value)
		{
			this.attribute[attribute] = value;
		};
		this.getBuffer = function ()
		{
			let writer = new AscCommon.CMemory();

			function Write(content)
			{
				if (content.name === "" && content.textContent === "" && content.content.length === 0)
				{
					writer.WriteXmlString("");
					return;
				}

				let current = null;

				if (!content.name)
				{
					if (content.xmlQuestionHeader !== null)
						writer.WriteXmlString(content.xmlQuestionHeader + "\n");

					current = content.content[0];
				}
				else
				{
					current = content;
				}

				writer.WriteXmlNodeStart(current.name);

				let atr = Object.keys(current.attribute)

				for (let i = 0; i < atr.length; i++)
				{
					let cur = atr[i];
					writer.WriteXmlAttributeStringEncode(cur, current.attribute[cur]);
				}

				writer.WriteXmlAttributesEnd();

				for (let i = 0; i < current.content.length; i++)
				{
					Write(current.content[i]);
				}

				if (current.textContent)
					writer.WriteXmlString(current.textContent.toString().trim());

				writer.WriteXmlNodeEnd(current.name);
			}

			Write(this);
			return writer;
		};
		this.getStringFromBuffer = function ()
		{
			let buffer	= this.getBuffer();
			let str		= fromUtf8(buffer.GetData());
			str			= str.replaceAll("&quot;", "\"");
			str			= str.replaceAll("&amp;", "&");
			return str;
		};

		this.Write_ToBinary2 = function (Writer)
		{
			Writer.WriteString2( this.getStringFromBuffer() );
		};
		this.Read_FromBinary2 = function (Reader)
		{
			let oContent = AscWord.CustomXmlCreateContent(Reader.GetString2());

			this.parent				= oContent.parent;
			this.name				= oContent.name;
			this.content			= oContent.content;
			this.attribute			= oContent.attribute;
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

	function CustomXmlCreateContent (strCustomXml)
	{
		let nXmlHeaderStart	= strCustomXml.indexOf('<?', 0);
		let nXmlHeaderEnd	= strCustomXml.indexOf('?>', nXmlHeaderStart);
		let strXmlHeader 	= null;
		if (nXmlHeaderStart !== -1 && nXmlHeaderEnd !== -1)
		{
			strXmlHeader = strCustomXml.substring(nXmlHeaderStart, nXmlHeaderEnd + "?>".length);
			strCustomXml = strCustomXml.substring(nXmlHeaderEnd + '?>'.length, strCustomXml.length);
		}

		let oStax			= new StaxParser(strCustomXml),
			rootContent		= new CustomXmlContent(null, null);

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
						childElement.addAttribute(attributeName, attributeValue);
					}

					rootContent = childElement;
					break;
			}
		}

		return rootContent;
	}

	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CustomXml				= CustomXml;
	AscWord.CustomXmlContent		= CustomXmlContent;
	AscWord.CustomXmlCreateContent	= CustomXmlCreateContent;
})();
