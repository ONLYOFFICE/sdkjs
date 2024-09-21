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

(function(window)
{
	/**
	 * @param {array} [uri]
	 * @param {string} [itemId]
	 * @param {CustomXMLContent} [content]
	 * @param [oContentLink]
	 *
	 * Класс представляющий CustomXML
	 * @constructor
	 */
	function CustomXml(uri, itemId, content, oContentLink)
	{
		this.uri				= uri ? uri : [];
		this.itemId				= itemId ? itemId : "";
		this.content			= content ? content : null;
		this.oContentLink		= oContentLink ? oContentLink : null;
	}

	/**
	 * Get CustomXML data by string
	 * @return {string}
	 */
	CustomXml.prototype.getText = function ()
	{
		return this.content.getStringFromBuffer();
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
		let strContent		= "".fromUtf8(arrData),
			strCustomXml	= strContent.slice(strContent.indexOf("<"), strContent.length); // Skip "L"

		this.addContentByXMLString(strCustomXml);
	};
	CustomXml.prototype.addContentByXMLString = function (strCustomXml)
	{
		let oStax			= new StaxParser(strCustomXml),
			rootContent		= new CustomXMLContent(null);

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

		this.content = rootContent;
	}

	function CustomXMLContent(parent, name)
	{
		this.parent			= parent;
		this.name			= name ? name : "";
		this.content		= [];
		this.attribute		= {};
		this.textContent	= "";

		this.addAttribute = function (name, value)
		{
			this.attribute[name] = value;
		};
		this.addContent = function (name)
		{
			let newItem = new CustomXMLContent(this, name);

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
			let nTab = 0;

			function Write(content)
			{
				let current = null;

				if (!content.name)
				{
					writer.WriteXmlString("\x00<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
					current = content.content[0];
				}
				else
				{
					current = content;
				}

				for (let i = 0; i < nTab; i++)
				{
					writer.WriteXmlString("	");
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
					nTab++;
					if (i === 0)
						writer.WriteXmlString("\n");
					let curContent = current.content[i];
					Write(curContent);
					nTab--;
					writer.WriteXmlString("\n");
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
			let str		= AscCommon.UTF8ArrayToString(buffer.data, 1);
			str			= str.replaceAll("&quot;", "\"");
			str			= str.replaceAll("&amp;", "&");
			return str;
		};
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CustomXml = CustomXml;

})(window);
