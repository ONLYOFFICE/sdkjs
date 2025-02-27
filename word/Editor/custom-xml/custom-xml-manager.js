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
	 * Класс представляющий менеджер CustomXMLs
	 * @param {AscWord.CDocument} document
	 * @constructor
	 */
	function CustomXmlManager(document)
	{
		this.document	= document;
		this.xml		= [];
	}
	CustomXmlManager.prototype.add = function(customXml)
	{
		// TODO: Надо будет сделать этот метод с сохранением в историю, когда
		//       будем реализовывать возможность добавления таких xml во время работы
		this.xml.push(customXml);
	};
	CustomXmlManager.prototype.getCount = function()
	{
		return this.xml.length;
	};
	CustomXmlManager.prototype.getCustomXml = function(index)
	{
		return this.xml[index];
	};

	/**
	 * Find element/attribute of CustomXMl by xpath string
	 * @param root {AscWord.CustomXmlContent}
	 * @param xpath {string}
	 * @return {{attribute: string, content: AscWord.CustomXmlContent}}
	 */
	CustomXmlManager.prototype.findElementByXPath = function (root, xpath)
	{
		let arrParts		= xpath.split('/');
		let currentElement	= root;

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

		return {
			content: currentElement,
			attribute: undefined,
		};
	};
	/**
	 * Get custom xml data of content control by data binding property
	 * @param dataBinding {window.AscWord.DataBinding}
	 * @param oContentLink {CBlockLevelSdt | CInlineLevelSdt}
	 * @return {string | undefined}
	 */
	CustomXmlManager.prototype.getContentByDataBinding = function(dataBinding, oContentLink)
	{
		for (let i = 0; i < this.xml.length; ++i)
		{
			let customXml			= this.xml[i];
			customXml.oContentLink	= oContentLink;

			if (dataBinding.storeItemID === customXml.itemId || customXml.checkUrl(dataBinding.prefixMappings))
			{
				let xPath			= dataBinding.xpath;
				let oFindEl			= this.findElementByXPath(customXml.content, xPath);
				let content			= oFindEl.content;
				let strAttribute	= oFindEl.attribute;

				return (undefined !== strAttribute)
					? content.attribute[strAttribute]
					: content.textContent;
			}
		}
	};
	/**
	 * Set custom xml data of content control by data binding property
	 * @param {window.AscWord.DataBinding} dataBinding
	 * @param data {string}
	 */
	CustomXmlManager.prototype.setContentByDataBinding = function (dataBinding, data)
	{
		for (let i = 0; i < this.xml.length; ++i)
		{
			let customXml = this.xml[i];

			if (dataBinding.storeItemID === customXml.itemId)
			{
				let xPath			= dataBinding.xpath;
				let oFindEl			= this.findElementByXPath(customXml.content, xPath);
				let oContent		= oFindEl.content;
				let strAttribute	= oFindEl.attribute;
				
				if (strAttribute)
					oContent.setAttribute(strAttribute, data);
				else
					oContent.setTextContent(data);
			}
		}
	};
	/**
	 * Set custom xml data of content control by data binding property
	 * @param {AscWord.CBlockLevelSdt | AscWord.CInlineLevelSdt} contentControl
	 */
	CustomXmlManager.prototype.updateDataBinding = function(contentControl)
	{
		if (!this.isSupported())
			return;
		
		let dataBinding = contentControl.getDataBinding();
		if (!dataBinding)
			return;
		
		if (contentControl.Pr.Text || contentControl.IsComboBox() || contentControl.IsDropDownList() || contentControl.IsDatePicker())
			this.setContentByDataBinding(dataBinding, contentControl.GetInnerText());
		else if (contentControl.IsCheckBox())
			this.setContentByDataBinding(dataBinding, contentControl.IsCheckBoxChecked() ? "true" : "false");
		else if (contentControl instanceof AscWord.CBlockLevelSdt)
			this.updateRichTextCustomXML(contentControl);
		else
			this.setContentByDataBinding(dataBinding, contentControl.GetInnerText());
	};
	/**
	 * Write linear xml data of content control in CustomXML
	 * @param oCC {CBlockLevelSdt}
	 */
	CustomXmlManager.prototype.updateRichTextCustomXML = function (oCC)
	{
		function replaceSubstring(originalString, startPoint, endPoint, insertionString)
		{
			if (startPoint < 0 || endPoint >= originalString.length || startPoint > endPoint)
				return originalString;

			const prefix	= originalString.substring(0, startPoint);
			const suffix	= originalString.substring(endPoint + 1);

			return prefix + insertionString + suffix;
		}

		AscCommon.ExecuteNoHistory(function() {
			let doc 						= new AscWord.CDocument(null, false);
			let oSdtContent					= oCC.GetContent().Copy();
			let jsZlib						= new AscCommon.ZLib();
			
			let copyContent = [];
			for (let i = 0, count = oSdtContent.GetElementsCount(); i < count; ++i)
			{
				copyContent.push(oSdtContent.GetElement(i).Copy());
			}
			
			doc.ReplaceContent(copyContent);
			jsZlib.create();
			doc.toZip(jsZlib, new AscCommon.XmlWriterContext(AscCommon.c_oEditorId.Word));

			let openDoc						= new AscCommon.openXml.OpenXmlPackage(jsZlib, null);
			let outputUString				= "<?xml version=\"1.0\" standalone=\"yes\"?><?mso-application progid=\"Word.Document\"?><pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">";
			let arrPath						= jsZlib.getPaths();

			arrPath.forEach(function(path)
			{
				if ((path === "_rels/.rels" || path === "word/document.xml" || path === "word/_rels/document.xml.rels") && !path.includes("glossary"))
				{
					let ctfBytes	= jsZlib.getFile(path);
					let ctfText		= AscCommon.UTF8ArrayToString(ctfBytes, 0, ctfBytes.length);
					let type		= openDoc.getContentType(path);

					if (path === "word/_rels/document.xml.rels")
					{
						let text = '';
						let arrRelationships = openDoc.getRelationships();

						for (let i = 0; i < arrRelationships.length; i++)
						{
							let relation	= arrRelationships[i];
							let relId		= relation.relationshipId;
							let relType		= relation.relationshipType;
							let relTarget	= relation.target;

							if (i === 0)
							{
								relType		= relType.replace("relationships\/officeDocument", "relationships\/styles");
								relTarget	= relTarget.replace("word/document.xml", "styles.xml");
							}

							text			+= "<Relationship Id=\"" + relId + "\" Type=\"" + relType + "\" Target=\"" + relTarget + "\"/>"
						}

						let nStart	= ctfText.indexOf("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">", 0) + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">".length;
						let nEnd	= ctfText.indexOf("</Relationships>", nStart) - 1;
						ctfText		= replaceSubstring(ctfText, nStart, nEnd, text);
					}

					outputUString += "<pkg:part pkg:name=\"/" + path + "\" " +
						"pkg:contentType=\"" + type +"\"> " +
						"<pkg:xmlData>" + ctfText.replace("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "").replace("\n", "") + "</pkg:xmlData></pkg:part>"
				}
			});

			outputUString	= outputUString.replace("pkg:contentType=\"application/xml\"", "pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"");
			outputUString	+= "</pkg:package>";
			outputUString	= outputUString.replaceAll("<", "&lt;");
			outputUString	= outputUString.replaceAll(">", "&gt;");
			this.setContentByDataBinding(oCC.Pr.DataBinding, outputUString);
		}, this.document, this, []);
	};
	/**
	 * Proceed linear xml from CustomXMl attribute or element for fill content control
	 * @param strLinearXML {string}
	 * @return {[]} Return array of CC content
	 */
	CustomXmlManager.prototype.proceedLinearXMl = function (strLinearXML)
	{
		let drawingDocument = this.document.GetDrawingDocument();
		let docContent = AscCommon.ExecuteNoHistory(function(){
			strLinearXML = strLinearXML.replaceAll("&lt;", "<");
			strLinearXML = strLinearXML.replaceAll("&gt;", ">");
			strLinearXML = strLinearXML.replaceAll("<?xml version=\"1.0\" standalone=\"yes\"?>", "");
			strLinearXML = strLinearXML.replaceAll("<?mso-application progid=\"Word.Document\"?>", "");
			
			// при записи в атрибут больше проблем, изменить подход если в будущем еще будут проблемы c html entry
			strLinearXML = strLinearXML.replaceAll("&#xA;", "");
			strLinearXML = strLinearXML.replaceAll("&amp;", "&");
			strLinearXML = strLinearXML.replaceAll("&quot;", "\"");
			strLinearXML = strLinearXML.replaceAll("&#039;", "'");
			
			let zLib = new AscCommon.ZLib;
			zLib.create();
			zLib.addFile('[Content_Types].xml', AscCommon.Utf8.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
				'<Default Extension="wmf" ContentType="image/x-wmf"/>' +
				'<Default Extension="png" ContentType="image/png"/>' +
				'<Default Extension="jpeg" ContentType="image/jpeg"/>' +
				'<Default Extension="xml" ContentType="application/xml"/>' +
				'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
				'<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>' +
				'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
				'<Override PartName="/word/fontTable.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' +
				'<Override PartName="/word/styles.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
				'<Override PartName="/word/document.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
				'<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
				'<Override PartName="/word/endnotes.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>' +
				'<Override PartName="/word/webSettings.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>' +
				'<Override PartName="/word/glossary/webSettings.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>' +
				'<Override PartName="/word/glossary/fontTable.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>' +
				'<Override PartName="/word/glossary/settings.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
				'<Override PartName="/docProps/app.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
				'<Override PartName="/word/footnotes.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' +
				'<Override PartName="/word/settings.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
				'<Override PartName="/word/glossary/styles.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
				'<Override PartName="/word/glossary/document.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>' +
				'<Override PartName="/customXml/itemProps1.xml"' +
				'ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>' +
				'</Types>'));
			
			let nPos = 0;
			while (true)
			{
				let nStartPos = nPos = strLinearXML.indexOf('<pkg:part', nPos);
				
				if (!nStartPos || nStartPos === -1)
					break;
				
				let nEndPos = nPos = strLinearXML.indexOf('</pkg:part>', nStartPos);
				let strText = strLinearXML.substring(nStartPos, nEndPos);
				
				let nPosStartName = strText.indexOf('name="', 0) + 'name="'.length;
				let nPosEndName   = strText.indexOf('"', nPosStartName);
				let name          = strText.substring(nPosStartName, nPosEndName);
				
				let nDataStartPos = strText.indexOf('<pkg:xmlData>', 0);
				let nDataEndPos;
				
				if (nDataStartPos !== -1)
				{
					nDataStartPos = nDataStartPos + '<pkg:xmlData>'.length;
					nDataEndPos   = strText.indexOf('</pkg:xmlData>', nDataStartPos);
				}
				else
				{
					nDataStartPos = strText.indexOf('<pkg:binaryData>', 0);
					if (nDataStartPos !== -1)
						nDataStartPos += '<pkg:binaryData>'.length;
					nDataEndPos = strText.indexOf('</pkg:binaryData>', nDataStartPos);
				}
				
				if (nStartPos === -1 || nEndPos === -1)
					continue;
				
				let data = strText.substring(nDataStartPos, nDataEndPos).trim();
				
				if (name[0] === "/")
					name = name.substring(1, name.length);
				
				zLib.addFile(name, AscCommon.Utf8.encode(data));
			}
			
			let arr              = zLib.save();
			let Doc              = new AscWord.CDocument(drawingDocument, false);
			let xmlParserContext = new AscCommon.XmlParserContext();
			let jsZlib           = new AscCommon.ZLib();
			
			xmlParserContext.DrawingDocument = drawingDocument;
			
			if (!jsZlib.open(arr))
				return [];
			
			let oBinaryFileReader	= new AscCommonWord.BinaryFileReader(Doc, {});
			oBinaryFileReader.PreLoadPrepare(undefined, false);
			
			Doc.fromZip(jsZlib, xmlParserContext, oBinaryFileReader.oReadResult);
			//очищать pptx_content_loader не надо чтобы не было проблем с вызовом внутри ReadPPTXElement и т.к. открываем zip
			//лучше уйти от глобального pptx_content_loader
			oBinaryFileReader.PostLoadPrepare(xmlParserContext, false);
			jsZlib.close();
			return Doc.Content;
		}, this.document, this, []);
		
		let resultContent = [];
		for (let i = 0; i < docContent.length; ++i)
		{
			resultContent.push(docContent[i].Copy(null));
		}

		return resultContent;
	};
	/**
	 * Get CustomXML text
	 * @param oCustomXMl {AscWord.CustomXml}
	 * @return {string}
	 */
	CustomXmlManager.prototype.getCustomXMLString = function(oCustomXMl)
	{
		return oCustomXMl.getText();
	};
	CustomXmlManager.prototype.isSupported = function()
	{
		return window['Asc'] && window['Asc']['Addons'] && true === window['Asc']['Addons']['ooxml'];
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.CustomXmlManager = CustomXmlManager;
	
})();
