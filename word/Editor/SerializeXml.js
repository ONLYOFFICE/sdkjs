/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

(function(window, undefined) {
	let openXml = AscCommon.openXml;
	let StaxParser = AscCommon.StaxParser;
	let CT_BoolW = window['AscCommon'].CT_BoolW;
	let CT_StringW = window['AscCommon'].CT_StringW;
	let CT_IntW = window['AscCommon'].CT_IntW;
	let CT_UIntW = window['AscCommon'].CT_UIntW;
	let CT_DoubleW = window['AscCommon'].CT_DoubleW;
	function CT_StringW15() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_StringW15, AscCommon.CT_String, 0);
	CT_StringW15.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeStringEncode("w15:val", val);
	};
	function CT_StringW14() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_StringW14, AscCommon.CT_String, 0);
	CT_StringW14.prototype.writeAttrVal = function(writer, val) {
		writer.WriteXmlNullableAttributeStringEncode("w14:val", val);
	};
	function CT_BoolW14() {
		AscCommon.CT_ComplexType.call(this);
	}
	AscFormat.InitClass(CT_BoolW14, AscCommon.CT_Bool, 0);
	CT_BoolW14.prototype.writeAttrVal = function(writer, val) {
		if (!val) {
			writer.WriteXmlNullableAttributeBool("w14:val", val);
		}
	};
	let CComments = window['AscCommon'].CComments;
	let CCommentData = window['AscCommon'].CCommentData;
	let CParagraphBookmark = window['AscCommonWord'].CParagraphBookmark;
	let CT_TrackChange = window['AscCommonWord'].CT_TrackChange;

//document
	CDocument.prototype.fromZip = function(zip, context, oReadResult) {
		context.oReadResult = oReadResult;
		context.zip = zip;

		let reader;
		let doc = new openXml.OpenXmlPackage(zip, null);

		let appPart = doc.getPartByRelationshipType(openXml.Types.extendedFileProperties.relationType);
		if (appPart) {
			let appContent = appPart.getDocumentContent();
			reader = new StaxParser(appContent, appPart, context);
			this.App = new AscCommon.CApp();
			this.App.fromXml(reader, true);
		}

		let corePart = doc.getPartByRelationshipType(openXml.Types.coreFileProperties.relationType);
		if (corePart) {
			let coreContent = corePart.getDocumentContent();
			reader = new StaxParser(coreContent, corePart, context);
			this.Core = new AscCommon.CCore();
			this.Core.fromXml(reader, true);
		}

		let customPrPart = doc.getPartByRelationshipType(openXml.Types.customFileProperties.relationType);
		if (customPrPart) {
			let customPrPartContent = customPrPart.getDocumentContent();
			reader = new StaxParser(customPrPartContent, customPrPart, context);
			this.CustomProperties = new AscCommon.CCustomProperties();
			this.CustomProperties.fromXml(reader, true);
		}

		let documentPart = doc.getPartByRelationshipType(openXml.Types.mainDocument.relationType);
		if (documentPart) {
			let themePart = documentPart.getPartByRelationshipType(openXml.Types.theme.relationType);
			if (themePart) {
				let themePartContent = themePart.getDocumentContent();
				reader = new StaxParser(themePartContent, themePart, context);
				this.theme = new AscFormat.CTheme();
				this.theme.fromXml(reader, true);
			}
			this.Comments.ReadFromXml(documentPart, context);
			let stylesPart = documentPart.getPartByRelationshipType(openXml.Types.styles.relationType);
			if (stylesPart) {
				let contentStyles = stylesPart.getDocumentContent();
				reader = new StaxParser(contentStyles, stylesPart, context);
				this.Styles.fromXml(reader);
			}
			let numberingPart = documentPart.getPartByRelationshipType(openXml.Types.numbering.relationType);
			if (numberingPart) {
				let numberingContent = numberingPart.getDocumentContent();
				reader = new StaxParser(numberingContent, numberingPart, context);
				this.Numbering.fromXml(reader);
			}
			let footnotesPart = documentPart.getPartByRelationshipType(openXml.Types.footnotes.relationType);
			if (footnotesPart) {
				let footnotesContent = footnotesPart.getDocumentContent();
				reader = new StaxParser(footnotesContent, footnotesPart, context);
				this.Footnotes.fromXml(reader);
			}
			let endnotesPart = documentPart.getPartByRelationshipType(openXml.Types.endnotes.relationType);
			if (endnotesPart) {
				let endnotesContent = endnotesPart.getDocumentContent();
				reader = new StaxParser(endnotesContent, endnotesPart, context);
				this.Endnotes.fromXml(reader);
			}
			let settingsPart = documentPart.getPartByRelationshipType(openXml.Types.documentSettings.relationType);
			if (settingsPart) {
				let settingsContent = settingsPart.getDocumentContent();
				reader = new StaxParser(settingsContent, settingsPart, context);
				this.Settings.fromXml(reader, this);
			}

			var customXmlParts = documentPart.getPartsByRelationshipType(openXml.Types.customXml.relationType);
			if (customXmlParts) {
				for (let i = 0; i < customXmlParts.length; i++) {
					var customXmlPart = customXmlParts[i];
					var customXml = customXmlPart.getDocumentContent();
					var customXmlPropsPart = customXmlPart.getPartByRelationshipType(openXml.Types.customXmlProps.relationType);
					var customXmlProps = customXmlPropsPart && customXmlPropsPart.getDocumentContent();

					if (-1 !== customXmlProps.indexOf("http://schemas.onlyoffice.com/settingsCustom")) {
						reader = new StaxParser(customXml, customXmlPart, context);
						this.Settings.fromXml(reader, this);
					} else {
						//todo read and save
						var custom = {Uri: [], ItemId: null, Content: null, item: customXml, itemProps: customXmlProps};
						this.CustomXmls.push(custom);
					}
				}
			}

			let glossaryPart = documentPart.getPartByRelationshipType(openXml.Types.glossaryDocument.relationType);
			if (glossaryPart) {
				let glossaryDocument = this.GetGlossaryDocument()
				context.oReadResult = new AscCommonWord.DocReadResult(glossaryDocument);

				let glossaryContent = glossaryPart.getDocumentContent();
				reader = new StaxParser(glossaryContent, glossaryPart, context);
				this.GetGlossaryDocument().fromXml(reader);

				context.oReadResult = oReadResult;
			}

			let contentDocument = documentPart.getDocumentContent();
			reader = new StaxParser(contentDocument, documentPart, context);
			this.fromXml(reader, oReadResult.DocumentContent);
		}
	};
	CDocument.prototype.toZip = function(zip, context) {
		var memory = new AscCommon.CMemory();
		memory.context = context;
		context.document = this;
		context.docSaveParams = new DocSaveParams();

		var filePart = new AscCommon.openXml.OpenXmlPackage(zip, memory);

		if (this.Core) {
			var corePart = filePart.addPart(AscCommon.openXml.Types.coreFileProperties);
			corePart.part.setDataXml(this.Core, memory);
			memory.Seek(0);
		}

		if (this.App) {
			var appPart = filePart.addPart(AscCommon.openXml.Types.extendedFileProperties);
			appPart.part.setDataXml(this.App, memory);
			memory.Seek(0);
		}

		var docPart = filePart.addPart(AscCommon.openXml.Types.mainDocument);

		var stylesPart = docPart.part.addPart(AscCommon.openXml.Types.styles);
		stylesPart.part.setDataXml(this.Styles, memory);
		memory.Seek(0);

		if (!this.Numbering.IsEmpty()) {
			var numberingPart = docPart.part.addPart(AscCommon.openXml.Types.numbering);
			numberingPart.part.setDataXml(this.Numbering, memory);
			memory.Seek(0);
		}

		memory.WriteXmlString(AscCommonWord.g_sXmlWebSettings);
		var sampleData = memory.GetDataUint8();
		var webSettingsPart = docPart.part.addPart(AscCommon.openXml.Types.webSettings);
		webSettingsPart.part.setData(sampleData);
		memory.Seek(0);

		var themePart = docPart.part.addPart(AscCommon.openXml.Types.theme);
		themePart.part.setDataXml(this.theme, memory);
		memory.Seek(0);

		this.Comments.WriteToXml(memory, docPart);

		memory.WriteXmlString(AscCommonWord.g_sXmlFonts);
		sampleData = memory.GetDataUint8();
		var fontsPart = docPart.part.addPart(AscCommon.openXml.Types.fontTable);
		fontsPart.part.setData(sampleData);
		memory.Seek(0);

		let footnotesPart = docPart.part.addPart(AscCommon.openXml.Types.footnotes);
		footnotesPart.part.setDataXml(this.Footnotes, memory);
		memory.Seek(0);

		let endnotesPart = docPart.part.addPart(AscCommon.openXml.Types.endnotes);
		endnotesPart.part.setDataXml(this.Endnotes, memory);
		memory.Seek(0);

		var settingsPart = docPart.part.addPart(AscCommon.openXml.Types.documentSettings);
		settingsPart.part.setDataXml(this.Settings, memory);
		memory.Seek(0);

		docPart.part.setDataXml(this, memory);
		memory.Seek(0);

		var glossaryDocument = this.GetGlossaryDocument();
		if (glossaryDocument) {
			var glossaryPart = docPart.part.addPart(AscCommon.openXml.Types.glossaryDocument);
			glossaryPart.part.setDataXml(glossaryDocument, memory);
			memory.Seek(0);
		}
	};
	CDocument.prototype.fromXml = function(reader, Content) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("document" !== name && "wordDocument" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("document" === name || "wordDocument" === name) {
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				name = reader.GetNameNoNS();
				if ("background" === name) {

				} else if ("body" === name) {
					this.fromXmlDocContent(reader, Content, this.DrawingDocument, this);
				}
			}
		}
	};
	CDocument.prototype.fromXmlDocContent = function(reader, Content, DrawingDocument, Parent) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			CDocument.prototype.fromXmlDocContentElem(reader, name, Content, DrawingDocument, Parent);
		}
	};
	CDocument.prototype.fromXmlDocContentElem = function(reader, name, Content, DrawingDocument, Parent) {
		let LogicDocument = DrawingDocument && DrawingDocument.m_oLogicDocument;
		var res = null, newItem = null;
		let elem, oReadResult = reader.context.oReadResult;
		switch (name) {
			case "altChunk":
				break;
			case "bookmarkStart":
				elem = new CParagraphBookmark(true);
				elem.fromXml(reader);
				oReadResult.addBookmarkStart(null, elem, true);
				break;
			case "bookmarkEnd":
				elem = new CParagraphBookmark(false);
				elem.fromXml(reader);
				oReadResult.addBookmarkEnd(oReadResult.lastPar, elem, true);
				break;
			case "moveFromRangeStart" : {
				oReadResult.readMoveRangeStartXml(oReadResult, reader, null, true);
				break;
			}
			case "moveFromRangeEnd" : {
				oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
				break;
			}
			case "moveToRangeStart" : {
				oReadResult.readMoveRangeStartXml(oReadResult, reader, null, false);
				break;
			}
			case "moveToRangeEnd" : {
				oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
				break;
			}
			case "commentRangeStart":
				break;
			case "commentRangeEnd":
				break;
			case "customXmlDelRangeStart":
				break;
			case "customXmlDelRangeEnd":
				break;
			case "customXmlInsRangeStart":
				break;
			case "customXmlInsRangeEnd":
				break;
			case "customXmlMoveFromRangeStart":
				break;
			case "customXmlMoveFromRangeEnd":
				break;
			case "customXmlMoveToRangeStart":
				break;
			case "customXmlMoveToRangeEnd":
				break;
			case "del":
				break;
			case "ins":
				break;
			case "moveFrom":
				break;
			case "moveTo":
				break;
			// case "oMath":
			// 	break;
			// case "oMathPara":
			// 	break;
			case "p":
				elem = new Paragraph(DrawingDocument, Parent);
				oReadResult.addToNextPar(elem);
				elem.fromXml(reader);
				if (reviewtype_Common === elem.GetReviewType() || oReadResult.checkReadRevisions()) {
					elem.Correct_Content();
					Content.push(elem);
				}
				break;
			case "permStart":
				break;
			case "permEnd":
				break;
			case "proofErr":
				break;
			case "sdt" : {
				newItem = new AscCommonWord.CBlockLevelSdt(LogicDocument, Parent);
				break;
			}
			case "sectPr":
				if (Parent.SectPr) {
					Parent.SectPr.fromXml(reader);
				}
				break;
			case "tbl":
				var table = new CTable(DrawingDocument, Parent, true, 0, 0, []);
				table.Set_TableStyle2(null);
				table.fromXml(reader);
				if (table.Get_RowsCount() > 0) {
					oReadResult.aTableCorrect.push(table);
					res = table;
					Content.push(table);
				}
				break;
		}
		if (newItem) {
			newItem.fromXml(reader);
			res = newItem;
			Content.push(newItem);
		}
		return res;
	};
	CDocument.prototype.toXml = function(writer) {
		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart("w:document");
		writer.WriteXmlString(AscCommonWord.g_sXmlDocumentNamespaces);
		writer.WriteXmlAttributesEnd();
		if (this.Background) {
			//this.Background.toXml(writer, "w:background");
		}
		writer.WriteXmlNodeStart("w:body");
		writer.WriteXmlAttributesEnd();
		this.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		this.SectPr.toXml(writer, "w:sectPr");
		writer.WriteXmlNodeEnd("w:body");
		writer.WriteXmlNodeEnd("w:document");
	};
	CDocument.prototype.toXmlDocContentElem = function(writer, item) {
		switch (item.GetType()) {
			case type_Paragraph:
				item.toXml(writer, "w:p");
				writer.context.docSaveParams.WriteRunRevisionMove(item, function(runRevisionMove) {
					WiteMoveRangeXml(writer, runRevisionMove);
				});
				break;
			case type_Table:
				item.toXml(writer, "w:tbl");
				break;
			case type_BlockLevelSdt:
				item.toXml(writer, "w:sdt");
				break;
		}
	};
	CGlossaryDocument.prototype.fromXml = function(reader) {
		let rels = reader.rels;
		let context = reader.context;
		let stylesPart = rels.getPartByRelationshipType(openXml.Types.styles.relationType);
		if (stylesPart) {
			let contentStyles = stylesPart.getDocumentContent();
			let reader = new StaxParser(contentStyles, stylesPart, context);
			this.Styles.fromXml(reader);
		}
		let numberingPart = rels.getPartByRelationshipType(openXml.Types.numbering.relationType);
		if (numberingPart) {
			let numberingContent = numberingPart.getDocumentContent();
			let reader = new StaxParser(numberingContent, numberingPart, context);
			this.Numbering.fromXml(reader);
		}
		let footnotesPart = rels.getPartByRelationshipType(openXml.Types.footnotes.relationType);
		if (footnotesPart) {
			let footnotesContent = footnotesPart.getDocumentContent();
			let reader = new StaxParser(footnotesContent, footnotesPart, context);
			this.Footnotes.fromXml(reader);
		}
		let endnotesPart = rels.getPartByRelationshipType(openXml.Types.endnotes.relationType);
		if (endnotesPart) {
			let endnotesContent = endnotesPart.getDocumentContent();
			let reader = new StaxParser(endnotesContent, endnotesPart, context);
			this.Endnotes.fromXml(reader);
		}
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("glossaryDocument" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("glossaryDocument" === name) {
			reader.ReadNextNode();//docParts
			var depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				name = reader.GetNameNoNS();
				if ("docPart" === name) {
					var docPart = new CDocPart(this);
					docPart.fromXml(reader);
					this.AddDocPart(docPart);
				}
			}
		}
	};
	CGlossaryDocument.prototype.toXml = function(writer) {
		let glossaryPart = writer.context.part;
		var stylesPart = glossaryPart.addPart(AscCommon.openXml.Types.styles);
		stylesPart.part.setDataXml(this.Styles, writer);
		writer.Seek(0);

		if (!this.Numbering.IsEmpty()) {
			var numberingPart = glossaryPart.addPart(AscCommon.openXml.Types.numbering);
			numberingPart.part.setDataXml(this.Numbering, writer);
			writer.Seek(0);
		}

		let footnotesPart = glossaryPart.addPart(AscCommon.openXml.Types.footnotes);
		footnotesPart.part.setDataXml(this.Footnotes, writer);
		writer.Seek(0);

		let endnotesPart = glossaryPart.addPart(AscCommon.openXml.Types.endnotes);
		endnotesPart.part.setDataXml(this.Endnotes, writer);
		writer.Seek(0);

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart("w:glossaryDocument");
		writer.WriteXmlString(AscCommonWord.g_sXmlDocumentNamespaces);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeStart("w:docParts");
		writer.WriteXmlAttributesEnd();
		for (var sId in this.DocParts) {
			if(this.DocParts.hasOwnProperty(sId)) {
				this.DocParts[sId].toXml(writer, "w:docPart");
			}
		}
		writer.WriteXmlNodeEnd("w:docParts");
		writer.WriteXmlNodeEnd("w:glossaryDocument");
	};
	CDocPart.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "docPartPr" : {
					this.Pr.fromXml(reader);
					break;
				}
				case "docPartBody" : {
					var Content = [];
					CDocument.prototype.fromXmlDocContent(reader, Content, this.DrawingDocument, this);
					if (Content.length > 0) {
						this.ReplaceContent(Content);
					}
					break;
				}
			}
		}
	};
	CDocPart.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Pr, "w:docPartPr");
		writer.WriteXmlNodeStart("w:docPartBody");
		writer.WriteXmlAttributesEnd();
		this.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		writer.WriteXmlNodeEnd("w:docPartBody");
		writer.WriteXmlNodeEnd(name);
	};
	CDocPartPr.prototype.fromXml = function(reader) {
		let t = this;
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "name" : {
					this.Name = CT_StringW.prototype.toVal(reader, this.Name);
					break;
				}
				case "style" : {
					this.Style = CT_StringW.prototype.toVal(reader, this.Style);
					break;
				}
				case "category" : {
					this.Category = new CDocPartCategory();
					this.Category.fromXml(reader);
					break;
				}
				case "types" : {
					reader.readXmlArray("type", function() {
						t.Types = fromXml_ST_DocPartType(CT_StringW.prototype.toVal(reader, t.Types), t.Types);
					});
					break;
				}
				case "behaviors" : {
					reader.readXmlArray("behavior", function() {
						t.Behaviors = fromXml_ST_DocPartBehavior(CT_StringW.prototype.toVal(reader, t.Behaviors), t.Behaviors);
					});
					break;
				}
				case "description" : {
					this.Description = CT_StringW.prototype.toVal(reader, this.Description);
					break;
				}
				case "guid" : {
					this.GUID = CT_StringW.prototype.toVal(reader, this.GUID);
					break;
				}
			}
		}
	};
	CDocPartPr.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Name), "w:name");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Style), "w:style");
		writer.WriteXmlNullable(this.Category, "w:category");
		if (null !== this.Types) {
			writer.WriteXmlNodeStart("w:types");
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_DocPartType(this.Types)), "w:type");
			writer.WriteXmlNodeEnd("w:types");
		}
		if (null !== this.Behaviors) {
			writer.WriteXmlNodeStart("w:behaviors");
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_DocPartBehavior(this.Behaviors)), "w:behavior");
			writer.WriteXmlNodeEnd("w:behaviors");
		}
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Description), "w:description");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.GUID), "w:guid");
		writer.WriteXmlNodeEnd(name);
	};
	CDocPartCategory.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "name" : {
					this.Name = CT_StringW.prototype.toVal(reader, this.Name);
					break;
				}
				case "gallery" : {
					this.Gallery = fromXml_ST_DocPartGallery(CT_StringW.prototype.toVal(reader, this.Name), this.Name);
					break;
				}
			}
		}
	};
	CDocPartCategory.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Name), "w:name");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_DocPartGallery(this.Gallery)), "w:gallery");
		writer.WriteXmlNodeEnd(name);
	};
	CTable.prototype.fromXml = function(reader) {
		let elem, t = this, oReadResult = reader.context.oReadResult;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "bookmarkStart":
					elem = new CParagraphBookmark(true);
					elem.fromXml(reader);
					oReadResult.addBookmarkStart(null, elem, true);
					break;
				case "bookmarkEnd":
					elem = new CParagraphBookmark(false);
					elem.fromXml(reader);
					oReadResult.addBookmarkEnd(oReadResult.lastPar, elem, true);
					break;
				case "moveFromRangeStart" :
					oReadResult.readMoveRangeStartXml(oReadResult, reader, null, true);
					break;
				case "moveFromRangeEnd" :
					oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
					break;
				case "moveToRangeStart" :
					oReadResult.readMoveRangeStartXml(oReadResult, reader, null, false);
					break;
				case "moveToRangeEnd" :
					oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
					break;
				case "commentRangeStart" : {
					break;
				}
				case "commentRangeEnd" : {
					break;
				}
				case "customXmlInsRangeStart" : {
					break;
				}
				case "customXmlInsRangeEnd" : {
					break;
				}
				case "customXmlDelRangeStart" : {
					break;
				}
				case "customXmlDelRangeEnd" : {
					break;
				}
				case "customXmlMoveFromRangeStart" : {
					break;
				}
				case "customXmlMoveFromRangeEnd" : {
					break;
				}
				case "customXmlMoveToRangeStart" : {
					break;
				}
				case "customXmlMoveToRangeEnd" : {
					break;
				}
				case "tblPr" : {
					this.Pr = new CTablePr();
					this.Pr.fromXml(reader, this);
					break;
				}
				case "tblGrid" : {
					var tblGrid = new CT_TblGrid();
					tblGrid.fromXml(reader);
					tblGrid.toTable(this);
					break;
				}
				case "tr" : {
					var row = this.private_AddRow(this.Content.length, 0);
					row.fromXml(reader);
					break;
				}
				case "customXml" : {
					break;
				}
				case "sdt" : {
					let sdt = new CT_SdtRow(this);
					sdt.fromXml(reader);
					break;
				}
				case "todo_EG_RunLevelElts" : {
					break;
				}
			}
		}
	};
	CTable.prototype.toXml = function(writer, name) {
		var tblGrid = new CT_TblGrid();
		tblGrid.fromTable(this);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlArray(this.bookmarkStart, "w:bookmarkStart");
		// writer.WriteXmlArray(this.bookmarkEnd, "w:bookmarkEnd");
		writer.WriteXmlArray(this.moveFromRangeStart, "w:moveFromRangeStart");
		writer.WriteXmlArray(this.moveFromRangeEnd, "w:moveFromRangeEnd");
		writer.WriteXmlArray(this.moveToRangeStart, "w:moveToRangeStart");
		writer.WriteXmlArray(this.moveToRangeEnd, "w:moveToRangeEnd");
		writer.WriteXmlArray(this.commentRangeStart, "w:commentRangeStart");
		writer.WriteXmlArray(this.commentRangeEnd, "w:commentRangeEnd");
		writer.WriteXmlArray(this.customXmlInsRangeStart, "w:customXmlInsRangeStart");
		writer.WriteXmlArray(this.customXmlInsRangeEnd, "w:customXmlInsRangeEnd");
		writer.WriteXmlArray(this.customXmlDelRangeStart, "w:customXmlDelRangeStart");
		writer.WriteXmlArray(this.customXmlDelRangeEnd, "w:customXmlDelRangeEnd");
		writer.WriteXmlArray(this.customXmlMoveFromRangeStart, "w:customXmlMoveFromRangeStart");
		writer.WriteXmlArray(this.customXmlMoveFromRangeEnd, "w:customXmlMoveFromRangeEnd");
		writer.WriteXmlArray(this.customXmlMoveToRangeStart, "w:customXmlMoveToRangeStart");
		writer.WriteXmlArray(this.customXmlMoveToRangeEnd, "w:customXmlMoveToRangeEnd");
		if (null !== this.Pr) {
			this.Pr.toXml(writer, "w:tblPr", this);
		}
		tblGrid.toXml(writer, "w:tblGrid");
		writer.WriteXmlArray(this.Content, "w:tr");
		writer.WriteXmlArray(this.customXml, "w:customXml");
		writer.WriteXmlArray(this.sdt, "w:sdt");
		writer.WriteXmlArray(this.todo_EG_RunLevelElts, "w:todo_EG_RunLevelElts");
		writer.WriteXmlNodeEnd(name);
	};
	CTablePr.prototype.fromXml = function(reader, opt_table) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "tblStyle" : {
					if (opt_table) {
						elem = CT_StringW.prototype.toVal(reader, undefined);
						if (elem) {
							reader.context.oReadResult.tableStyles.push({pPr: opt_table, style: elem});
						}
					}
					break;
				}
				case "tblpPr" : {
					if (opt_table) {
						elem = new CT_TblPPr();
						elem.fromXml(reader);
						elem.toTable(opt_table);
					}
					break;
				}
				// case "tblOverlap" : {
				// 	this.TblOverlap = new CT_TblOverlap();
				// 	this.TblOverlap.fromXml(reader);
				// 	break;
				// }
				// case "bidiVisual" : {
				// 	this.BidiVisual = new CT_BoolW();
				// 	this.BidiVisual.fromXml(reader);
				// 	break;
				// }
				case "tblStyleRowBandSize" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.TableStyleRowBandSize = elem.getVal(undefined);
					break;
				}
				case "tblStyleColBandSize" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.TableStyleColBandSize = elem.getVal(undefined);
					break;
				}
				case "tblW" : {
					this.TableW = new CTableMeasurement(tblwidth_Auto, 0);
					this.TableW.fromXml(reader);
					break;
				}
				case "jc" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Jc = fromXml_ST_JcTable(elem.getVal(undefined));
					break;
				}
				case "tblCellSpacing" : {
					var tblCellSpacing = new CTableMeasurement(tblwidth_Auto, 0);
					tblCellSpacing.fromXml(reader);
					if (tblwidth_Mm === tblCellSpacing.Type) {
						//different understanding of TableCellSpacing with Word
						this.TableCellSpacing = 2 * tblCellSpacing.W;
					}
					break;
				}
				case "tblInd" : {
					var tblInd = new CTableMeasurement(tblwidth_Auto, 0);
					tblInd.fromXml(reader);
					if (tblwidth_Mm === tblInd.Type) {
						this.TableInd = tblInd.W;
					}
					break;
				}
				case "tblBorders" : {
					var pBdr = new CT_Bdr();
					pBdr.fromXml(reader);
					pBdr.toObj(this.TableBorders);
					break;
				}
				case "shd" : {
					this.Shd = new CDocumentShd();
					this.Shd.fromXml(reader);
					break;
				}
				case "tblLayout" : {
					elem = new CT_TblLayoutType();
					elem.fromXml(reader);
					if (null !== elem.type) {
						this.TableLayout = elem.type;
					}
					break;
				}
				case "tblCellMar" : {
					this.TableCellMar = {Bottom: undefined, Left: undefined, Right: undefined, Top: undefined};
					elem = new CT_TblCellMar();
					elem.fromXml(reader);
					elem.toObj(this.TableCellMar);
					break;
				}
				case "tblLook" : {
					if (opt_table) {
						elem = new AscCommon.CTableLook();
						elem.SetDefault();
						elem.fromXml(reader);
						opt_table.Set_TableLook(elem);
					}
					break;
				}
				case "tblCaption" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.TableCaption = elem.getVal(undefined);
					break;
				}
				case "tblDescription" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.TableDescription = elem.getVal(undefined);
					break;
				}
				case "tblPrChange" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						this.SetPrChange(trackChange.tblPrChange, trackChange.ReviewInfo);
					}
					break;
				}
			}
		}
	};
	CTablePr.prototype.toXml = function(writer, name, opt_table) {
		var TableStyle;
		if (opt_table) {
			TableStyle = CT_StringW.prototype.fromVal(opt_table.Get_TableStyle());
		}
		var TblpPr;
		if (opt_table && !opt_table.Inline) {
			TblpPr = new CT_TblPPr();
			TblpPr.fromTable(opt_table);
		}
		var TableStyleRowBandSize = CT_IntW.prototype.fromVal(this.TableStyleRowBandSize);
		var TableStyleColBandSize = CT_IntW.prototype.fromVal(this.TableStyleColBandSize);
		var Jc = CT_StringW.prototype.fromVal(toXml_ST_JcTable(this.Jc));
		var TableCellSpacing;
		if (undefined !== this.TableCellSpacing) {
			TableCellSpacing = new CTableMeasurement(tblwidth_Mm, this.TableCellSpacing / 2);
		}
		var TableInd;
		if (undefined !== this.TableInd) {
			TableInd = new CTableMeasurement(tblwidth_Mm, this.TableInd);
		}
		var TableBorders = new CT_Bdr();
		TableBorders.fromObj(this.TableBorders);
		if (TableBorders.isEmpty()) {
			TableBorders = null;
		}
		var TableLayout;
		if (undefined !== this.TableLayout) {
			TableLayout = new CT_TblLayoutType();
			TableLayout.type = this.TableLayout;
		}
		var TableCellMar;
		if (this.TableCellMar) {
			TableCellMar = new CT_TblCellMar();
			TableCellMar.fromObj(this.TableCellMar);
			if (TableCellMar.isEmpty()) {
				TableCellMar = null;
			}
		}
		var TableCaption = CT_StringW.prototype.fromVal(this.TableCaption);
		var TableDescription = CT_StringW.prototype.fromVal(this.TableDescription);
		let trackChange;
		if(null != this.PrChange && this.ReviewInfo) {
			trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, this.ReviewInfo);
			trackChange.tblPrChange = this.PrChange;
		}
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(TableStyle, "w:tblStyle");
		writer.WriteXmlNullable(TblpPr, "w:tblpPr");
		// writer.WriteXmlNullable(this.TblOverlap, "w:tblOverlap");
		// writer.WriteXmlNullable(this.BidiVisual, "w:bidiVisual");
		writer.WriteXmlNullable(TableStyleRowBandSize, "w:tblStyleRowBandSize");
		writer.WriteXmlNullable(TableStyleColBandSize, "w:tblStyleColBandSize");
		writer.WriteXmlNullable(this.TableW, "w:tblW");
		writer.WriteXmlNullable(Jc, "w:jc");
		writer.WriteXmlNullable(TableCellSpacing, "w:tblCellSpacing");
		writer.WriteXmlNullable(TableInd, "w:tblInd");
		writer.WriteXmlNullable(TableBorders, "w:tblBorders");
		writer.WriteXmlNullable(this.Shd, "w:shd");
		writer.WriteXmlNullable(TableLayout, "w:tblLayout");
		writer.WriteXmlNullable(TableCellMar, "w:tblCellMar");
		if (opt_table) {
			writer.WriteXmlNullable(opt_table.Get_TableLook(), "w:tblLook");
		}
		writer.WriteXmlNullable(TableCaption, "w:tblCaption");
		writer.WriteXmlNullable(TableDescription, "w:tblDescription");
		writer.WriteXmlNullable(trackChange, "w:tblPrChange");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_TblGrid() {
		this.gridCol = [];
		this.tblGridChange = null;
		return this;
	}

	CT_TblGrid.prototype.fromTable = function(table) {
		if (table.TableGrid) {
			for (var i = 0; i < table.TableGrid.length; ++i) {
				var elem = new CT_TblGridCol();
				elem.w = table.TableGrid[i];
				this.gridCol.push(elem);
			}
		}
		if (table.TableGridChange) {
			this.tblGridChange = new CT_TblGrid();
			for (var i = 0; i < table.TableGridChange.length; ++i) {
				var elem = new CT_TblGridCol();
				elem.w = table.TableGridChange[i];
				this.tblGridChange.gridCol.push(elem);
			}
		}
	};
	CT_TblGrid.prototype.toTable = function(table) {
		let tableGrid = this.gridCol.map(function(elem) {
			return elem.w;
		});
		table.SetTableGrid(tableGrid);
		if (table.SetTableGridChange && this.tblGridChange) {
			let tableGridChange = this.tblGridChange.gridCol.map(function(elem) {
				return elem.w;
			});
			table.SetTableGridChange(tableGridChange);
		}
	};
	CT_TblGrid.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "gridCol" : {
					var elem = new CT_TblGridCol();
					elem.fromXml(reader);
					this.gridCol.push(elem);
					break;
				}
				case "tblGridChange" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						this.tblGridChange = trackChange.tblGridChange;
					}
					break;
				}
			}
		}
	};
	CT_TblGrid.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlArray(this.gridCol, "w:gridCol");
		if (this.tblGridChange) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, null);
			trackChange.ReviewInfo = null;
			trackChange.tblGridChange = this.tblGridChange;
			writer.WriteXmlNullable(trackChange, "w:tblGridChange");
		}
		writer.WriteXmlNodeEnd(name);
	};
	function CT_TblGridCol() {
		this.w = null;
		return this;
	}

	CT_TblGridCol.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "w": {
					this.w =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
			}
		}
	};
	CT_TblGridCol.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_TblGridCol.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeIntWithKoef("w:w", this.w, g_dKoef_mm_to_twips);
		writer.WriteXmlAttributesEnd(true);
	};
	AscCommon.CTableLook.prototype.readAttr = function(reader) {
		var nLook, FirstRow, LastRow, FirstColumn, LastColumn, NoHBand, NoVBand;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					nLook = reader.GetValueUInt(nLook, 16);
					if (undefined !== nLook) {
						FirstColumn = 0 != (nLook & 0x0080);
						FirstRow = 0 != (nLook & 0x0020);
						LastColumn = 0 != (nLook & 0x0100);
						LastRow = 0 != (nLook & 0x0040);
						NoHBand = 0 != (nLook & 0x0200);
						NoVBand = 0 != (nLook & 0x0400);
					}
					break;
				}
				case "firstRow": {
					FirstRow = reader.GetValueBool();
					break;
				}
				case "lastRow": {
					LastRow = reader.GetValueBool();
					break;
				}
				case "firstColumn": {
					FirstColumn = reader.GetValueBool();
					break;
				}
				case "lastColumn": {
					LastColumn = reader.GetValueBool();
					break;
				}
				case "noHBand": {
					NoHBand = reader.GetValueBool();
					break;
				}
				case "noVBand": {
					NoVBand = reader.GetValueBool();
					break;
				}
			}
		}
		this.Set(!!FirstColumn, !!FirstRow, !!LastColumn, !!LastRow, !NoHBand, !NoVBand);
	};
	AscCommon.CTableLook.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscCommon.CTableLook.prototype.toXml = function(writer, name) {
		var nLook = 0;
		if(this.IsFirstCol())
			nLook |= 0x0080;
		if(this.IsFirstRow())
			nLook |= 0x0020;
		if(this.IsLastCol())
			nLook |= 0x0100;
		if(this.IsLastRow())
			nLook |= 0x0040;
		if(!this.IsBandHor())
			nLook |= 0x0200;
		if(!this.IsBandVer())
			nLook |= 0x0400;
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:val", AscCommon.Int16ToHex(nLook));
		writer.WriteXmlNullableAttributeBool("w:firstRow", this.IsFirstRow());
		writer.WriteXmlNullableAttributeBool("w:lastRow", this.IsLastRow());
		writer.WriteXmlNullableAttributeBool("w:firstColumn", this.IsFirstCol());
		writer.WriteXmlNullableAttributeBool("w:lastColumn", this.IsLastCol());
		writer.WriteXmlNullableAttributeBool("w:noHBand", !this.IsBandHor());
		writer.WriteXmlNullableAttributeBool("w:noVBand", !this.IsBandVer());
		writer.WriteXmlAttributesEnd(true);
	};
	CTableRow.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "rsidRPr": {
					break;
				}
				case "rsidR": {
					break;
				}
				case "rsidDel": {
					break;
				}
				case "rsidTr": {
					break;
				}
			}
		}
	};
	CTableRow.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let elem, oReadResult = reader.context.oReadResult;
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "tblPrEx" : {
					break;
				}
				case "trPr" : {
					var rowPr = new CTableRowPr();
					rowPr.fromXml(reader, this);
					this.Set_Pr(rowPr);
					break;
				}
				case "tc" : {
					var cell = this.Add_Cell(this.Get_CellsCount(), this, null, false);
					cell.fromXml(reader);
					break;
				}
				case "customXml" : {
					break;
				}
				case "sdt" : {
					let sdt = new CT_SdtCell(this);
					sdt.fromXml(reader);
					break;
				}
				case "bookmarkStart":
					elem = new CParagraphBookmark(true);
					elem.fromXml(reader);
					oReadResult.addBookmarkStart(null, elem, true);
					break;
				case "bookmarkEnd":
					elem = new CParagraphBookmark(false);
					elem.fromXml(reader);
					oReadResult.addBookmarkEnd(oReadResult.lastPar, elem, true);
					break;
				case "moveFromRangeStart" :
					oReadResult.readMoveRangeStartXml(oReadResult, reader, null, true);
					break;
				case "moveFromRangeEnd" :
					oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
					break;
				case "moveToRangeStart" :
					oReadResult.readMoveRangeStartXml(oReadResult, reader, null, false);
					break;
				case "moveToRangeEnd" :
					oReadResult.readMoveRangeEndXml(oReadResult, reader, oReadResult.lastPar, true);
					break;
				case "todo_EG_RunLevelElts" : {
					break;
				}
			}
		}
	};
	CTableRow.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.tblPrEx, "w:tblPrEx");
		if (this.Pr) {
			this.Pr.toXml(writer, "w:trPr", this);
		}
		writer.WriteXmlArray(this.Content, "w:tc");
		// writer.WriteXmlArray(this.customXml, "w:customXml");
		// writer.WriteXmlArray(this.sdt, "w:sdt");
		// writer.WriteXmlArray(this.todo_EG_RunLevelElts, "w:todo_EG_RunLevelElts");
		writer.WriteXmlNodeEnd(name);
	};
	CTableRowPr.prototype.fromXml = function(reader, row) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				// case "cnfStyle" : {
				// 	this.CnfStyle = new CT_Cnf();
				// 	this.CnfStyle.fromXml(reader);
				// 	break;
				// }
				// case "divId" : {
				// 	this.DivId = new CT_IntW();
				// 	this.DivId.fromXml(reader);
				// 	break;
				// }
				case "gridBefore" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.GridBefore = elem.getVal(undefined);
					break;
				}
				case "gridAfter" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.GridAfter = elem.getVal(undefined);
					break;
				}
				case "wBefore" : {
					this.WBefore = new CTableMeasurement(tblwidth_Auto, 0);
					this.WBefore.fromXml(reader);
					break;
				}
				case "wAfter" : {
					this.WAfter = new CTableMeasurement(tblwidth_Auto, 0);
					this.WAfter.fromXml(reader);
					break;
				}
				case "cantSplit" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.CantSplit = elem.getVal(undefined);
					break;
				}
				case "trHeight" : {
					this.Height = new CTableRowHeight(0, Asc.linerule_AtLeast);
					this.Height.fromXml(reader);
					break;
				}
				case "tblHeader" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.TableHeader = elem.getVal(undefined);
					break;
				}
				case "tblCellSpacing" : {
					var tblCellSpacing = new CTableMeasurement(tblwidth_Auto, 0);
					tblCellSpacing.fromXml(reader);
					if (tblwidth_Mm === tblCellSpacing.Type) {
						//different understanding of TableCellSpacing with Word
						this.TableCellSpacing = 2 * tblCellSpacing.W;
					}
					break;
				}
				case "jc" : {
					var Jc = new CT_StringW();
					Jc.fromXml(reader);
					this.Jc = fromXml_ST_JcTable(Jc.getVal(undefined));
					break;
				}
				// case "hidden" : {
				// 	this.Hidden = new CT_BoolW();
				// 	this.Hidden.fromXml(reader);
				// 	break;
				// }
				case "del" : {
					if (row) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						row.SetReviewTypeWithInfo(reviewtype_Remove, trackChange.ReviewInfo);
					}
					break;
				}
				case "ins" : {
					if (row && reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						row.SetReviewTypeWithInfo(reviewtype_Add, trackChange.ReviewInfo);
					}
					break;
				}
				case "trPrChange" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						this.SetPrChange(trackChange.trPrChange, trackChange.ReviewInfo);
					}
					break;
				}
			}
		}
	};
	CTableRowPr.prototype.toXml = function(writer, name, row) {
		var GridBefore = CT_IntW.prototype.fromVal(this.GridBefore);
		var GridAfter = CT_IntW.prototype.fromVal(this.GridAfter);
		var CantSplit = CT_BoolW.prototype.fromVal(this.CantSplit);
		var TableHeader = CT_BoolW.prototype.fromVal(this.TableHeader);
		var TableCellSpacing;
		if (undefined !== this.TableCellSpacing) {
			TableCellSpacing = new CTableMeasurement(tblwidth_Mm, this.TableCellSpacing / 2);
		}
		var Jc = CT_StringW.prototype.fromVal(toXml_ST_JcTable(this.Jc));

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.CnfStyle, "w:cnfStyle");
		// writer.WriteXmlNullable(this.DivId, "w:divId");
		writer.WriteXmlNullable(GridBefore, "w:gridBefore");
		writer.WriteXmlNullable(GridAfter, "w:gridAfter");
		writer.WriteXmlNullable(this.WBefore, "w:wBefore");
		writer.WriteXmlNullable(this.WAfter, "w:wAfter");
		writer.WriteXmlNullable(CantSplit, "w:cantSplit");
		writer.WriteXmlNullable(this.Height, "w:trHeight");
		writer.WriteXmlNullable(TableHeader, "w:tblHeader");
		writer.WriteXmlNullable(TableCellSpacing, "w:tblCellSpacing");
		writer.WriteXmlNullable(Jc, "w:jc");
		// writer.WriteXmlNullable(this.Hidden, "w:hidden");
		if (this.PrChange && this.ReviewInfo) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, this.ReviewInfo);
			trackChange.trPrChange = this.PrChange;
			writer.WriteXmlNullable(trackChange, "w:trPrChange");
		}
		if (row && reviewtype_Common !== row.GetReviewType()) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, row.GetReviewInfo());
			var ReviewType = row.GetReviewType();
			if (reviewtype_Add === ReviewType)
				writer.WriteXmlNullable(trackChange, "w:ins");
			else if (reviewtype_Remove === ReviewType)
				writer.WriteXmlNullable(trackChange, "w:del");
		}
		writer.WriteXmlNodeEnd(name);
	};
	CTableMeasurement.prototype.readAttr = function(reader) {
		var w;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "w": {
					w = reader.GetValueInt();
					break;
				}
				case "type": {
					this.Type = fromXml_ST_TblWidth(reader.GetValue(), tblwidth_Auto);
					break;
				}
			}
		}
		this.SetValueByType(w);
	};
	CTableMeasurement.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CTableMeasurement.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:w", this.GetValueByType());
		writer.WriteXmlNullableAttributeString("w:type", toXml_ST_TblWidth(this.Type));
		writer.WriteXmlAttributesEnd(true);
	};
	CTableRowHeight.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					this.Value =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
				case "hRule": {
					this.HRule = fromXml_ST_HeightRule(reader.GetValue());
					break;
				}
			}
		}
	};
	CTableRowHeight.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CTableRowHeight.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeIntWithKoef("w:val", this.Value, AscCommonWord.g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:hRule", toXml_ST_HeightRule(this.HRule));
		writer.WriteXmlAttributesEnd(true);
	};
	CTableCell.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					break;
				}
			}
		}
	};
	CTableCell.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var Content = [];
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			var newItem = CDocument.prototype.fromXmlDocContentElem(reader, name, Content, this.Content.DrawingDocument, this.Content);
			if (!newItem) {
				switch (name) {
					case "tcPr" : {
						var cellPr = new CTableCellPr();
						cellPr.fromXml(reader);
						this.Set_Pr(cellPr);
						break;
					}
				}
			}
		}
		if (Content.length > 0) {
			this.Content.ReplaceContent(Content);
		}
	};
	CTableCell.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeString("id", this.id);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Pr, "w:tcPr");
		this.Content.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		writer.WriteXmlNodeEnd(name);
	};
	CTableCellPr.prototype.fromXml = function(reader) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				// case "cnfStyle" : {
				// 	this.CnfStyle = new CT_Cnf();
				// 	this.CnfStyle.fromXml(reader);
				// 	break;
				// }
				case "tcW" : {
					this.TableCellW = new CTableMeasurement(tblwidth_Auto, 0);
					this.TableCellW.fromXml(reader);
					break;
				}
				case "gridSpan" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.GridSpan = elem.getVal(undefined);
					break;
				}
				case "hmerge" :
				case "hMerge" : {
					this.HMerge = fromXml_ST_Merge(CT_StringW.prototype.toVal(reader), vmerge_Continue);
					break;
				}
				case "vmerge" :
				case "vMerge" : {
					this.VMerge = fromXml_ST_Merge(CT_StringW.prototype.toVal(reader), vmerge_Continue);
					break;
				}
				case "tcBorders" : {
					var pBdr = new CT_Bdr();
					pBdr.fromXml(reader);
					pBdr.toObj(this.TableCellBorders);
					break;
				}
				case "shd" : {
					this.Shd = new CDocumentShd();
					this.Shd.fromXml(reader);
					break;
				}
				case "noWrap" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.NoWrap = elem.getVal(undefined);
					break;
				}
				case "tcMar" : {
					this.TableCellMar = {Bottom: undefined, Left: undefined, Right: undefined, Top: undefined};
					elem = new CT_TblCellMar();
					elem.fromXml(reader);
					elem.toObj(this.TableCellMar);
					break;
				}
				case "textDirection" : {
					this.TextDirection = fromXml_ST_TextDirection(CT_StringW.prototype.toVal(reader, this.TextDirection));
					break;
				}
				// case "tcFitText" : {
				// 	this.TcFitText = new CT_BoolW();
				// 	this.TcFitText.fromXml(reader);
				// 	break;
				// }
				case "vAlign" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.VAlign = fromXml_ST_VerticalJc(elem.getVal(undefined));
					break;
				}
				// case "hideMark" : {
				// 	this.HideMark = new CT_BoolW();
				// 	this.HideMark.fromXml(reader);
				// 	break;
				// }
				// case "headers" : {
				// 	this.Headers = new CT_Headers();
				// 	this.Headers.fromXml(reader);
				// 	break;
				// }
				case "cellIns" : {
					// this.CellIns = new CT_TrackChange();
					// this.CellIns.fromXml(reader);
					break;
				}
				case "cellDel" : {
					// this.CellDel = new CT_TrackChange();
					// this.CellDel.fromXml(reader);
					break;
				}
				// case "cellMerge" : {
				// 	this.CellMerge = new CT_CellMergeTrackChange();
				// 	this.CellMerge.fromXml(reader);
				// 	break;
				// }
				case "tcPrChange" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						this.SetPrChange(trackChange.tcPrChange, trackChange.ReviewInfo);
					}
					break;
				}
			}
		}
	};
	CTableCellPr.prototype.toXml = function(writer, name) {
		var GridSpan = CT_IntW.prototype.fromVal(this.GridSpan);
		var HMerge = CT_StringW.prototype.fromVal(toXml_ST_Merge(this.HMerge));
		var VMerge = CT_StringW.prototype.fromVal(toXml_ST_Merge(this.VMerge));
		var TableCellBorders = new CT_Bdr();
		TableCellBorders.fromObj(this.TableCellBorders);
		if (TableCellBorders.isEmpty()) {
			TableCellBorders = null;
		}
		var NoWrap = CT_BoolW.prototype.fromVal(this.NoWrap);
		var TableCellMar;
		if (this.TableCellMar) {
			TableCellMar = new CT_TblCellMar();
			TableCellMar.fromObj(this.TableCellMar);
			if (TableCellMar.isEmpty()) {
				TableCellMar = null;
			}
		}
		var TextDirection = CT_StringW.prototype.fromVal(toXml_ST_TextDirection(this.TextDirection));
		var VAlign = CT_StringW.prototype.fromVal(toXml_ST_VerticalJc(this.VAlign));

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.CnfStyle, "w:cnfStyle");
		writer.WriteXmlNullable(this.TableCellW, "w:tcW");
		writer.WriteXmlNullable(GridSpan, "w:gridSpan");
		writer.WriteXmlNullable(HMerge, "w:hMerge");
		writer.WriteXmlNullable(VMerge, "w:vMerge");
		writer.WriteXmlNullable(TableCellBorders, "w:tcBorders");
		writer.WriteXmlNullable(this.Shd, "w:shd");
		writer.WriteXmlNullable(NoWrap, "w:noWrap");
		writer.WriteXmlNullable(TableCellMar, "w:tcMar");
		writer.WriteXmlNullable(TextDirection, "w:textDirection");
		// writer.WriteXmlNullable(this.TcFitText, "w:tcFitText");
		writer.WriteXmlNullable(VAlign, "w:vAlign");
		// writer.WriteXmlNullable(this.HideMark, "w:hideMark");
		// writer.WriteXmlNullable(this.Headers, "w:headers");
		// writer.WriteXmlNullable(this.CellIns, "w:cellIns");
		// writer.WriteXmlNullable(this.CellDel, "w:cellDel");
		// writer.WriteXmlNullable(this.CellMerge, "w:cellMerge");
		if (this.PrChange && this.ReviewInfo) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, this.ReviewInfo);
			trackChange.tcPrChange = this.PrChange;
			writer.WriteXmlNullable(trackChange, "w:tcPrChange");
		}
		writer.WriteXmlNodeEnd(name);
	};
//Paragraph
	CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem = function (reader, name) {
		let paragraph = this.GetParagraph();
		let elem, oReadResult = reader.context.oReadResult, context = reader.context;
		switch (name) {
			case "bdo":
				elem = new CT_BdoDirContentRun();
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				this.ConcatContent(elem.Content);
				break;
			case "bookmarkStart":
				elem = new CParagraphBookmark(true);
				elem.fromXml(reader);
				oReadResult.addBookmarkStart(this, elem, true);
				break;
			case "bookmarkEnd":
				elem = new CParagraphBookmark(false);
				elem.fromXml(reader);
				oReadResult.addBookmarkEnd(this, elem, true);
				break;
			case "moveFromRangeStart" :
				oReadResult.readMoveRangeStartXml(oReadResult, reader, this, true);
				break;
			case "moveFromRangeEnd" :
				oReadResult.readMoveRangeEndXml(oReadResult, reader, this, true);
				break;
			case "moveToRangeStart" :
				oReadResult.readMoveRangeStartXml(oReadResult, reader, this, false);
				break;
			case "moveToRangeEnd" :
				oReadResult.readMoveRangeEndXml(oReadResult, reader, this, true);
				break;
			case "commentRangeStart" : {
				elem = new CT_MarkupRange();
				elem.fromXml(reader);
				let commentData = context.commentDataById[elem.id];
				if (commentData) {
					this.AddToContent(this.GetElementsCount(), new AscCommon.ParaComment(true, commentData.Get_Id()));
				}
				break;
			}
			case "commentRangeEnd" : {
				elem = new CT_MarkupRange();
				elem.fromXml(reader);
				let commentData = context.commentDataById[elem.id];
				if (commentData) {
					this.AddToContent(this.GetElementsCount(), new AscCommon.ParaComment(false, commentData.Get_Id()));
				}
				break;
			}
			// case "customXmlDelRangeEnd":
			// 	break;
			// case "customXmlDelRangeStart":
			// 	break;
			// case "customXmlInsRangeEnd":
			// 	break;
			// case "customXmlInsRangeStart":
			// 	break;
			// case "customXmlMoveFromRangeEnd":
			// 	break;
			// case "customXmlMoveFromRangeStart":
			// 	break;
			// case "customXmlMoveToRangeEnd":
			// 	break;
			// case "customXmlMoveToRangeStart":
			// 	break;
			case "dir":
				elem = new CT_BdoDirContentRun();
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				this.ConcatContent(elem.Content);
				break;
			case "fldSimple":
				elem = new ParaField(fieldtype_UNKNOWN);
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				let pageField = elem.GetRunWithPageField(paragraph);
				if (pageField) {
					this.AddToContentToEnd(pageField);
				} else {
					reader.context.oReadResult.logicDocument.Register_Field(elem);
					this.ConcatContent(elem.Content);
				}
				break;
			case "hyperlink":
				elem = new ParaHyperlink();
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				elem.Check_Content();
				this.AddToContent(this.GetElementsCount(), elem);
				break;
			case "oMath":
				elem = new ParaMath();
				this.AddToContentToEnd(elem);
				elem.fromXml(reader);
				break;
			case "oMathPara":
				elem = new AscCommon.CT_OMathPara();
				elem.fromXml(reader, this);
				break;
			// case "permEnd":
			// 	break;
			// case "permStart":
			// 	break;
			// case "proofErr":
			// 	break;
			case "r":
				elem = new ParaRun(paragraph, para_Math_Content === this.Type);
				elem.fromXml(reader);
				this.AddToContent(this.GetElementsCount(), elem);

				if (elem.GetElementsCount() > Asc.c_dMaxParaRunContentLength && !(this instanceof CInlineLevelSdt && this.IsForm())) {
					oReadResult.runsToSplit.push(elem);
				}
				break;
			case "sdt" : {
				elem = new AscCommonWord.CInlineLevelSdt();
				elem.RemoveFromContent(0, elem.GetElementsCount());
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				if (elem.IsEmpty())
					elem.ReplaceContentWithPlaceHolder();
				this.AddToContent(this.GetElementsCount(), elem);
				break;
			}
			case "smartTag":
				elem = new CT_SmartTagRun();
				elem.SetParagraph(paragraph);
				elem.fromXml(reader);
				this.ConcatContent(elem.Content);
				break;
			case "del": {
				if (oReadResult.checkReadRevisions()) {
					let trackChange = new CT_TrackChange();
					trackChange.paragraphContent = this;
					let startPos = this.GetElementsCount();
					trackChange.fromXml(reader);
					let endPos = this.GetElementsCount();
					for (let i = startPos; i < endPos; ++i) {
						oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Remove, trackChange.ReviewInfo);
					}
				}
				break;
			}
			case "ins": {
				let trackChange = new CT_TrackChange();
				trackChange.paragraphContent = this;
				let startPos = this.GetElementsCount();
				trackChange.fromXml(reader);
				let endPos = this.GetElementsCount();
				if (reader.context.oReadResult.checkReadRevisions()) {
					for (let i = startPos; i < endPos; ++i) {
						oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Add, trackChange.ReviewInfo);
					}
				}
				break;
			}
			case "moveFrom": {
				if (reader.context.oReadResult.checkReadRevisions()) {
					let trackChange = new CT_TrackChange();
					trackChange.ReviewInfo.SetMove(Asc.c_oAscRevisionsMove.MoveFrom);
					trackChange.paragraphContent = this;
					let startPos = this.GetElementsCount();
					trackChange.fromXml(reader);
					let endPos = this.GetElementsCount();
					for (let i = startPos; i < endPos; ++i) {
						oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Remove, trackChange.ReviewInfo);
					}
				}
				break;
			}
			case "moveTo": {
				let trackChange = new CT_TrackChange();
				trackChange.ReviewInfo.SetMove(Asc.c_oAscRevisionsMove.MoveTo);
				trackChange.paragraphContent = this;
				let startPos = this.GetElementsCount();
				trackChange.fromXml(reader);
				let endPos = this.GetElementsCount();
				if (reader.context.oReadResult.checkReadRevisions()) {
					for (let i = startPos; i < endPos; ++i) {
						oReadResult.setNestedReviewType(this.GetElement(i), reviewtype_Add, trackChange.ReviewInfo);
					}
				}
				break;
			}
		}
	};
	CParagraphContentWithParagraphLikeContent.prototype.toXml = function(writer) {
		var t = this;
		let context = writer.context;
		this.Content.forEach(function(item, index) {
			switch (item.Type) {
				case para_Run:
					let reviewType = item.GetReviewType();
					if (reviewtype_Common !== reviewType) {
						context.docSaveParams.writeNestedReviewType(reviewType, item.GetReviewInfo(), function(reviewType, reviewInfo, delText, fCallback){
							let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, reviewInfo);
							trackChange.writeCallback = function(){fCallback(delText);};
							if (reviewtype_Remove === reviewType) {
								if (reviewInfo.IsMovedFrom()) {
									writer.WriteXmlNullable(trackChange, "w:moveFrom");
								} else {
									delText = true;
									writer.WriteXmlNullable(trackChange, "w:del");
								}
							} else if (reviewtype_Add === reviewType) {
								if (reviewInfo.IsMovedTo()) {
									writer.WriteXmlNullable(trackChange, "w:moveTo");
								} else {
									writer.WriteXmlNullable(trackChange, "w:ins");
								}
							}
						}, function(delText) {
							item.toXml(writer, "w:r", delText);
						});
					} else {
						item.toXml(writer, "w:r", false);
					}
					break;
				case para_Field:
					item.toXml(writer, "w:fldSimple");
					break;
				case para_Hyperlink:
					item.toXml(writer, "w:hyperlink");
					break;
				case para_Comment:
					let commentData = context.commentDataById[item.CommentId];
					if (commentData) {
						commentData.toXmlRef(writer, item.Start);
					}
					break;
				case para_Math:
					if (t.CheckMathPara(index)) {
						let mathPara = new AscCommon.CT_OMathPara();
						mathPara.initMathParaPr(item);
						mathPara.toXml(writer, "m:oMathPara", item);
					} else {
						item.toXml(writer, "m:oMath");
					}
					break;
				case para_InlineLevelSdt:
					item.toXml(writer, "w:sdt");
					break;
				case para_Bookmark:
					item.toXml(writer, item.IsStart() ? "w:bookmarkStart" : "w:bookmarkEnd");
					break;
				case para_RevisionMove:
					WiteMoveRangeXml(writer, item);
					break;
			}
		});
	};
	CParagraphContentWithParagraphLikeContent.prototype.fromXml = function(reader) {
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			this.fromXmlElem(reader, reader.GetNameNoNS());
		}
	};
	Paragraph.prototype.fromXmlElem = CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem;
	Paragraph.prototype.toXml = CParagraphContentWithParagraphLikeContent.prototype.toXml;
	Paragraph.prototype.fromXml = CParagraphContentWithParagraphLikeContent.prototype.fromXml;
	Paragraph.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				// case "rsidRPr": {
				// 	this.rsidRPr = reader.GetValueInt(this.rsidRPr, 16);
				// 	break;
				// }
				// case "rsidR": {
				// 	this.rsidR = reader.GetValueInt(this.rsidR, 16);
				// 	break;
				// }
				// case "rsidDel": {
				// 	this.rsidDel = reader.GetValueInt(this.rsidDel, 16);
				// 	break;
				// }
				// case "rsidP": {
				// 	this.rsidP = reader.GetValueInt(this.rsidP, 16);
				// 	break;
				// }
				// case "rsidRDefault": {
				// 	this.rsidRDefault = reader.GetValueInt(this.rsidRDefault, 16);
				// 	break;
				// }
				case "paraId": {
					this.ParaId = reader.GetValueInt(null, 16);
					break;
				}
				// case "textId": {
				// 	this.textId = reader.GetValueInt(this.textId, 16);
				// 	break;
				// }
			}
		}
	};
	Paragraph.prototype.fromXml = function(reader) {
		let t = this, oReadResult = reader.context.oReadResult;
		this.readAttr(reader);
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			if ("pPr" === name) {
				this.Pr = new CParaPr();
				this.Pr.fromXml(reader, this);
			} else {
				CParagraphContentWithParagraphLikeContent.prototype.fromXmlElem.call(this, reader, name);
			}
		}
	};
	Paragraph.prototype.toXml = function(writer, name) {
		var t = this;
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w14:paraId", AscCommon.Int32ToHexOrNull(this.ParaId));
		writer.WriteXmlAttributesEnd();
		if (this.Pr) {
			this.Pr.toXml(writer, "w:pPr", this, true);
		}
		CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
		writer.WriteXmlNodeEnd(name);
	};
	CParaPr.prototype.fromXml = function(reader, opt_paragraph) {
		var depth = reader.GetDepth(), oReadResult = reader.context.oReadResult;
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "pStyle" : {
					this.PStyle = CT_StringW.prototype.toVal(reader, this.PStyle);
					oReadResult.paraStyles.push({pPr: this, style: this.PStyle});
					break;
				}
				case "keepNext" : {
					this.KeepNext = CT_BoolW.prototype.toVal(reader, this.KeepNext);
					break;
				}
				case "keepLines" : {
					this.KeepLines = CT_BoolW.prototype.toVal(reader, this.KeepLines);
					break;
				}
				case "pageBreakBefore" : {
					this.PageBreakBefore = CT_BoolW.prototype.toVal(reader, this.PageBreakBefore);
					break;
				}
				case "framePr" : {
					this.FramePr = new CFramePr();
					this.FramePr.fromXml(reader);
					break;
				}
				case "widowControl" : {
					this.WidowControl = CT_BoolW.prototype.toVal(reader, this.WidowControl);
					break;
				}
				case "numPr" : {
					let numPr = new CNumPr();
					numPr.Set(undefined, undefined);
					numPr.fromXml(reader);
					if (null != numPr.NumId || null != numPr.Lvl) {
						if (null != numPr.NumId) {
							oReadResult.paraNumPrs.push(numPr);
						}
						this.NumPr = numPr;
					}
					break;
				}
				case "suppressLineNumbers" : {
					this.SuppressLineNumbers = CT_BoolW.prototype.toVal(reader, this.SuppressLineNumbers);
					break;
				}
				case "pBdr" : {
					var pBdr = new CT_Bdr();
					pBdr.fromXml(reader);
					pBdr.toObj(this.Brd);
					break;
				}
				case "shd" : {
					this.Shd = new CDocumentShd();
					this.Shd.fromXml(reader);
					break;
				}
				case "tabs" : {
					this.Tabs = new CParaTabs();
					var subDepth = reader.GetDepth();
					while (reader.ReadNextSiblingNode(subDepth)) {
						if ("tab" === reader.GetNameNoNS()) {
							var elem = new CParaTab();
							elem.fromXml(reader);
							if (elem.IsValid()) {
								this.Tabs.Add(elem);
							}
						}
					}
					break;
				}
				// case "suppressAutoHyphens" : {
				// 	this.Suppressautohyphens = new CT_BoolW();
				// 	this.Suppressautohyphens.fromXml(reader);
				// 	break;
				// }
				// case "kinsoku" : {
				// 	this.Kinsoku = new CT_BoolW();
				// 	this.Kinsoku.fromXml(reader);
				// 	break;
				// }
				// case "wordWrap" : {
				// 	this.Wordwrap = new CT_BoolW();
				// 	this.Wordwrap.fromXml(reader);
				// 	break;
				// }
				// case "overflowPunct" : {
				// 	this.Overflowpunct = new CT_BoolW();
				// 	this.Overflowpunct.fromXml(reader);
				// 	break;
				// }
				// case "topLinePunct" : {
				// 	this.Toplinepunct = new CT_BoolW();
				// 	this.Toplinepunct.fromXml(reader);
				// 	break;
				// }
				// case "autoSpaceDE" : {
				// 	this.Autospacede = new CT_BoolW();
				// 	this.Autospacede.fromXml(reader);
				// 	break;
				// }
				// case "autoSpaceDN" : {
				// 	this.Autospacedn = new CT_BoolW();
				// 	this.Autospacedn.fromXml(reader);
				// 	break;
				// }
				// case "bidi" : {
				// 	this.Bidi = new CT_BoolW();
				// 	this.Bidi.fromXml(reader);
				// 	break;
				// }
				// case "adjustRightInd" : {
				// 	this.Adjustrightind = new CT_BoolW();
				// 	this.Adjustrightind.fromXml(reader);
				// 	break;
				// }
				// case "snapToGrid" : {
				// 	this.Snaptogrid = new CT_BoolW();
				// 	this.Snaptogrid.fromXml(reader);
				// 	break;
				// }
				case "spacing" : {
					this.Spacing = new CParaSpacing();
					this.Spacing.fromXml(reader);
					break;
				}
				case "ind" : {
					this.Ind = new CParaInd();
					this.Ind.fromXml(reader);
					break;
				}
				case "contextualSpacing" : {
					this.ContextualSpacing = CT_BoolW.prototype.toVal(reader, this.ContextualSpacing);
					break;
				}
				// case "mirrorIndents" : {
				// 	this.Mirrorindents = new CT_BoolW();
				// 	this.Mirrorindents.fromXml(reader);
				// 	break;
				// }
				// case "suppressOverlap" : {
				// 	this.Suppressoverlap = new CT_BoolW();
				// 	this.Suppressoverlap.fromXml(reader);
				// 	break;
				// }
				case "jc" : {
					this.Jc = fromXml_ST_Jc1(CT_StringW.prototype.toVal(reader, this.Jc), this.Jc);
					break;
				}
				// case "textDirection" : {
				// 	this.Textdirection = new CT_TextDirection();
				// 	this.Textdirection.fromXml(reader);
				// 	break;
				// }
				// case "textAlignment" : {
				// 	this.Textalignment = new CT_TextAlignment();
				// 	this.Textalignment.fromXml(reader);
				// 	break;
				// }
				// case "textboxTightWrap" : {
				// 	this.Textboxtightwrap = new CT_TextboxTightWrap();
				// 	this.Textboxtightwrap.fromXml(reader);
				// 	break;
				// }
				case "outlineLvl" : {
					this.OutlineLvl = CT_IntW.prototype.toVal(reader, this.OutlineLvl);
					break;
				}
				// case "divId" : {
				// 	this.Divid = new CT_IntW();
				// 	this.Divid.fromXml(reader);
				// 	break;
				// }
				// case "cnfStyle" : {
				// 	this.Cnfstyle = new CT_Cnf();
				// 	this.Cnfstyle.fromXml(reader);
				// 	break;
				// }
				case "rPr" : {
					if (opt_paragraph) {
						let EndRun = opt_paragraph.GetParaEndRun();
						opt_paragraph.TextPr.Value = new CTextPr();
						opt_paragraph.TextPr.Value.fromXml(reader, EndRun);
					}
					break;
				}
				case "sectPr" : {
					if (opt_paragraph) {
						elem = new CSectionPr(opt_paragraph.LogicDocument);
						elem.fromXml(reader);
						opt_paragraph.Set_SectionPr(elem);
					}
					break;
				}
				case "pPrChange" : {
					if (oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						this.SetPrChange(trackChange.pPrChange, trackChange.ReviewInfo);
					}
					break;
				}
			}
		}
	};
	CParaPr.prototype.toXml = function(writer, name, opt_paragraph, opt_ignoreEmpty) {
		var PStyle = CT_StringW.prototype.fromVal(this.PStyle);
		var KeepNext = CT_BoolW.prototype.fromVal(this.KeepNext);
		var KeepLines = CT_BoolW.prototype.fromVal(this.KeepLines);
		var PageBreakBefore = CT_BoolW.prototype.fromVal(this.PageBreakBefore);
		var WidowControl = CT_BoolW.prototype.fromVal(this.WidowControl);
		var SuppressLineNumbers = CT_BoolW.prototype.fromVal(this.SuppressLineNumbers);
		var pBdr = new CT_Bdr();
		pBdr.fromObj(this.Brd);
		if (pBdr.isEmpty()) {
			pBdr = null;
		}
		var Spacing = this.Spacing.Is_Empty() ? null : this.Spacing;
		var Ind = this.Ind.Is_Empty() ? null : this.Ind;
		var ContextualSpacing = CT_BoolW.prototype.fromVal(this.ContextualSpacing);
		var OutlineLvl = CT_IntW.prototype.fromVal(this.OutlineLvl);
		var JcStr = toXml_ST_Jc1(this.Jc);
		var Jc = JcStr ? CT_StringW.prototype.fromVal(JcStr) : null;

		let startPos = writer.GetCurPosition();
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		let headerPos = writer.GetCurPosition();
		writer.WriteXmlNullable(PStyle, "w:pStyle");
		writer.WriteXmlNullable(KeepNext, "w:keepNext");
		writer.WriteXmlNullable(KeepLines, "w:keepLines");
		writer.WriteXmlNullable(PageBreakBefore, "w:pageBreakBefore");
		writer.WriteXmlNullable(this.FramePr, "w:framePr");
		writer.WriteXmlNullable(WidowControl, "w:widowControl");
		writer.WriteXmlNullable(this.NumPr, "w:numPr");//todo
		writer.WriteXmlNullable(SuppressLineNumbers, "w:suppressLineNumbers");
		writer.WriteXmlNullable(pBdr, "w:pBdr");
		writer.WriteXmlNullable(this.Shd, "w:shd");
		writer.WriteXmlArray(this.Tabs && this.Tabs.Tabs, "w:tabs", "w:tab");//todo
		// writer.WriteXmlNullable(this.Suppressautohyphens, "w:suppressAutoHyphens");
		// writer.WriteXmlNullable(this.Kinsoku, "w:kinsoku");
		// writer.WriteXmlNullable(this.Wordwrap, "w:wordWrap");
		// writer.WriteXmlNullable(this.Overflowpunct, "w:overflowPunct");
		// writer.WriteXmlNullable(this.Toplinepunct, "w:topLinePunct");
		// writer.WriteXmlNullable(this.Autospacede, "w:autoSpaceDE");
		// writer.WriteXmlNullable(this.Autospacedn, "w:autoSpaceDN");
		// writer.WriteXmlNullable(this.Bidi, "w:bidi");
		// writer.WriteXmlNullable(this.Adjustrightind, "w:adjustRightInd");
		// writer.WriteXmlNullable(this.Snaptogrid, "w:snapToGrid");
		writer.WriteXmlNullable(Spacing, "w:spacing");
		writer.WriteXmlNullable(Ind, "w:ind");
		writer.WriteXmlNullable(ContextualSpacing, "w:contextualSpacing");
		// writer.WriteXmlNullable(this.Mirrorindents, "w:mirrorIndents");
		// writer.WriteXmlNullable(this.Suppressoverlap, "w:suppressOverlap");
		writer.WriteXmlNullable(Jc, "w:jc");
		// writer.WriteXmlNullable(this.Textdirection, "w:textDirection");
		// writer.WriteXmlNullable(this.Textalignment, "w:textAlignment");
		// writer.WriteXmlNullable(this.Textboxtightwrap, "w:textboxTightWrap");
		writer.WriteXmlNullable(OutlineLvl, "w:outlineLvl");
		// writer.WriteXmlNullable(this.Divid, "w:divId");
		// writer.WriteXmlNullable(this.Cnfstyle, "w:cnfStyle");
		if (opt_paragraph) {
			var EndRun = opt_paragraph.GetParaEndRun();
			opt_paragraph.TextPr.Value.toXml(writer, "w:rPr", EndRun, EndRun);
			writer.WriteXmlNullable(opt_paragraph.SectPr, "w:sectPr");
		}
		if (this.PrChange && this.ReviewInfo) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, this.ReviewInfo);
			trackChange.pPrChange = this.PrChange;
			writer.WriteXmlNullable(trackChange, "w:pPrChange");
		}
		//todo
		if (opt_ignoreEmpty && headerPos === writer.GetCurPosition()) {
			writer.Seek(startPos);
		} else {
			writer.WriteXmlNodeEnd(name);
		}
	};
	CFramePr.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "dropCap": {
					this.DropCap = fromXml_ST_DropCap(reader.GetValue());
					break;
				}
				case "lines": {
					this.Lines = reader.GetValueInt(this.Lines);
					break;
				}
				case "w": {
					this.W =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				case "h": {
					this.H =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				case "vSpace": {
					this.VSpace =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				case "hSpace": {
					this.HSpace =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				case "wrap": {
					this.Wrap = fromXml_ST_Wrap(reader.GetValue());
					break;
				}
				case "hAnchor": {
					this.HAnchor = fromXml_ST_HAnchor(reader.GetValue());
					break;
				}
				case "vAnchor": {
					this.VAnchor = fromXml_ST_VAnchor(reader.GetValue());
					break;
				}
				case "x": {
					var x = AscCommon.universalMeasureToPt(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_pt, null);
					if (null !== x) {
						if (-4 === x * AscCommonWord.g_dKoef_pt_to_twips) {
							this.XAlign = Asc.c_oAscXAlign.Center;
						} else {
							this.X = x * AscCommonWord.g_dKoef_pt_to_mm;
						}
					}
					break;
				}
				case "xAlign": {
					this.XAlign = fromXml_ST_XAlign(reader.GetValue());
					break;
				}
				case "y": {
					var y = AscCommon.universalMeasureToPt(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_pt, null);
					if (null !== y) {
						if (-4 === y * AscCommonWord.g_dKoef_pt_to_twips) {
							this.YAlign = Asc.c_oAscYAlign.Top;
						} else {
							this.Y = y * AscCommonWord.g_dKoef_pt_to_mm;
						}
					}
					break;
				}
				case "yAlign": {
					this.YAlign = fromXml_ST_YAlign(reader.GetValue());
					break;
				}
				case "hRule": {
					this.HRule = fromXml_ST_HeightRule(reader.GetValue());
					break;
				}
				// case "anchorLock": {
				// 	this.anchorLock = reader.GetValueDecodeXml();
				// 	break;
				// }
			}
		}
	};
	CFramePr.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CFramePr.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:dropCap", toXml_ST_DropCap(this.DropCap));
		writer.WriteXmlNullableAttributeIntWithKoef("w:lines", this.Lines, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeIntWithKoef("w:w", this.W, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeIntWithKoef("w:h", this.H, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeIntWithKoef("w:vSpace", this.VSpace, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeIntWithKoef("w:hSpace", this.HSpace, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:wrap", toXml_ST_Wrap(this.Wrap));
		writer.WriteXmlNullableAttributeString("w:hAnchor", toXml_ST_HAnchor(this.HAnchor));
		writer.WriteXmlNullableAttributeString("w:vAnchor", toXml_ST_VAnchor(this.VAnchor));
		writer.WriteXmlNullableAttributeIntWithKoef("w:x", this.X, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:xAlign", toXml_ST_XAlign(this.XAlign));
		writer.WriteXmlNullableAttributeIntWithKoef("w:y", this.Y, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:yAlign", toXml_ST_YAlign(this.YAlign));
		writer.WriteXmlNullableAttributeString("w:hRule", toXml_ST_HeightRule(this.HRule));
		// writer.WriteXmlNullableAttributeString("w:anchorLock", this.anchorLock);
		writer.WriteXmlAttributesEnd(true);
	};
	CNumPr.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ilvl" : {
					this.Lvl = CT_IntW.prototype.toVal(reader, this.Lvl);
					break;
				}
				case "numId" : {
					this.NumId = CT_IntW.prototype.toVal(reader, this.NumId);
					break;
				}
				// case "ins" : {
				// 	this.ins = new CT_TrackChange();
				// 	this.ins.fromXml(reader);
				// 	break;
				// }
			}
		}
	};
	CNumPr.prototype.toXml = function(writer, name) {
		let Lvl = CT_IntW.prototype.fromVal(this.Lvl);
		let NumId = CT_IntW.prototype.fromVal(writer.context.oNumIdMap[this.NumId] || 0);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(Lvl, "w:ilvl");
		writer.WriteXmlNullable(NumId, "w:numId");
		// writer.WriteXmlNullable(this.ins, "w:ins");
		writer.WriteXmlNodeEnd(name);
	};
	CDocumentBorder.prototype.readAttr = function(reader) {
		var themeColor = new CT_Color("color", "themeColor", "themeTint", "themeShade");
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			switch (name) {
				case "val": {
					this.Value = fromXml_ST_Border(reader.GetValue());
					break;
				}
				case "sz": {
					this.setSizeIn8Point(reader.GetValueUInt());
					break;
				}
				case "space": {
					this.setSpaceInPoint(reader.GetValueUInt());
					break;
				}
// 				case "shadow": {
// 					this.Shadow = reader.GetValueDecodeXml();
// 					break;
// 				}
// 				case "frame": {
// 					this.Frame = reader.GetValueDecodeXml();
// 					break;
// 				}
				default:
					themeColor.readAttrElem(reader, name);
					break;
			}
		}
		this.Color = themeColor.getColor(0, 0, 0);
		this.Unifill = themeColor.getUnifill();
	};
	CDocumentBorder.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CDocumentBorder.prototype.toXml = function(writer, name) {
		var themeColor = new CT_Color("color", "themeColor", "themeTint", "themeShade");
		themeColor.setColor(this.Color);
		themeColor.setUniFill(this.Unifill);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:val", toXml_ST_Border(this.Value));
		writer.WriteXmlNullableAttributeUInt("w:sz", this.getSizeIn8Point());
		writer.WriteXmlNullableAttributeUInt("w:space", this.getSpaceInPoint());
		themeColor.toXmlElems(writer);
// 		writer.WriteXmlNullableAttributeStringEncode("w:shadow", this.Shadow);
// 		writer.WriteXmlNullableAttributeStringEncode("w:frame", this.Frame);
		writer.WriteXmlAttributesEnd(true);
	};
	CDocumentShd.prototype.readAttr = function(reader) {
		this.Value = undefined;
		var themeColor = new CT_Color("color", "themeColor", "themeTint", "themeShade");
		var themeFill = new CT_Color("fill", "themeFill", "themeFillTint", "themeFillShade");
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			if ("val" === name) {
				this.Value = fromXml_ST_Shd(reader.GetValue());
			} else if (!themeColor.readAttrElem(reader, name)) {
				themeFill.readAttrElem(reader, name)
			}
		}
		this.Color = themeColor.getColor(255, 255, 255);
		this.Unifill = themeColor.getUnifill();
		this.Fill = themeFill.getColor(255, 255, 255);
		this.themeFill = themeFill.getUnifill();
		//если нет Value, но есть цвет, то Value по умолчанию ShdClear(Тарифы May,01,2016.docx, Тарифы_на_комплексное_обслуживание_клиен.docx)
		if (undefined === this.Value) {
			if (this.Color || this.Fill || this.Unifill || this.ThemeFill) {
				this.Value = Asc.c_oAscShdClear;
			} else {
				this.Value = Asc.c_oAscShdNil;
			}
		}
		// TODO: Как только будем нормально воспринимать CDocumentShd.Color = undefined, убрать отсюда проверку (!oShd.Color)
		if (!this.Color) {
			this.Color = new AscCommonWord.CDocumentColor(255, 255, 255, true);
		}
	};
	CDocumentShd.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CDocumentShd.prototype.toXml = function(writer, name) {
		var themeColor = new CT_Color("color", "themeColor", "themeTint", "themeShade");
		themeColor.setColor(this.Color);
		themeColor.setUniFill(this.Unifill);
		var themeFill = new CT_Color("fill", "themeFill", "themeFillTint", "themeFillShade");
		themeFill.setColor(this.Fill);
		themeColor.setUniFill(this.themeFill);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:val", toXml_ST_Shd(this.Value));
		themeColor.toXmlElems(writer);
		themeFill.toXmlElems(writer);
		writer.WriteXmlAttributesEnd(true);
	};
	CParaTab.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					this.Value = fromXml_ST_TabJc(reader.GetValue(), this.Value);
					break;
				}
				case "leader": {
					this.Leader = fromXml_ST_TabTlc(reader.GetValue());
					break;
				}
				case "pos": {
					this.Pos =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, undefined);
					break;
				}
			}
		}
	};
	CParaTab.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CParaTab.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:val", toXml_ST_TabJc(this.Value));
		writer.WriteXmlNullableAttributeString("w:leader", toXml_ST_TabTlc(this.Leader));
		writer.WriteXmlNullableAttributeIntWithKoef("w:pos", this.Pos, g_dKoef_mm_to_twips);
		writer.WriteXmlAttributesEnd(true);
	};
	CParaSpacing.prototype.readAttr = function(reader) {
		var linePt = null;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "before": {
					this.Before =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				// case "beforeLines": {
				// 	this.Beforelines = reader.GetValueInt();
				// 	break;
				// }
				case "beforeAutospacing": {
					this.BeforeAutoSpacing = reader.GetValueBool();
					break;
				}
				case "after": {
					this.After =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				// case "afterLines": {
				// 	this.Afterlines = reader.GetValueInt();
				// 	break;
				// }
				case "afterAutospacing": {
					this.AfterAutoSpacing = reader.GetValueBool();
					break;
				}
				case "line": {
					linePt =
						AscCommon.universalMeasureToPt(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_pt, undefined);
					break;
				}
				case "lineRule": {
					this.LineRule = fromXml_ST_LineSpacingRule(reader.GetValue());
					break;
				}
			}
		}
		if (null !== linePt) {
			this.SetLineTwips(linePt * AscCommonWord.g_dKoef_pt_to_twips);
		}
	};
	CParaSpacing.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CParaSpacing.prototype.toXml = function(writer, name) {
		var lineKoef = Asc.linerule_Auto === this.LineRule ? 240 : AscCommonWord.g_dKoef_mm_to_twips;

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeIntWithKoef("w:before", this.Before, AscCommonWord.g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:beforeLines", this.Beforelines);
		writer.WriteXmlNullableAttributeBool("w:beforeAutospacing", this.BeforeAutoSpacing);
		writer.WriteXmlNullableAttributeIntWithKoef("w:after", this.After, AscCommonWord.g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:afterLines", this.Afterlines);
		writer.WriteXmlNullableAttributeBool("w:afterAutospacing", this.AfterAutoSpacing);
		writer.WriteXmlNullableAttributeIntWithKoef("w:line", this.Line, lineKoef);
		writer.WriteXmlNullableAttributeString("w:lineRule", toXml_ST_LineSpacingRule(this.LineRule));
		writer.WriteXmlAttributesEnd(true);
	};
	CParaInd.prototype.readAttr = function(reader) {
		var Hanging = undefined;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "start":
				case "left": {
					this.Left =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				// case "startChars": {
				// 	this.Startchars = reader.GetValueInt();
				// 	break;
				// }
				case "right":
				case "end": {
					this.Right =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				// case "endChars": {
				// 	this.Endchars = reader.GetValueInt();
				// 	break;
				// }
				case "hanging": {
					Hanging =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					if (undefined !== Hanging) {
						this.FirstLine = -Hanging;
					}
					break;
				}
				// case "hangingChars": {
				// 	this.Hangingchars = reader.GetValueInt();
				// 	break;
				// }
				case "firstLine": {
					this.FirstLine =
						AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm,
							undefined);
					break;
				}
				// case "firstLineChars": {
				// 	this.Firstlinechars = reader.GetValueInt();
				// 	break;
				// }
			}
		}
	};
	CParaInd.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CParaInd.prototype.toXml = function(writer, name) {
		var Hanging = null;
		var FirstLine = this.FirstLine;
		if (FirstLine < 0) {
			Hanging = Math.abs(FirstLine);
			FirstLine = null;
		}
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeIntWithKoef("w:left", this.Left, g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:startChars", this.Startchars);
		writer.WriteXmlNullableAttributeIntWithKoef("w:right", this.Right, g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:endChars", this.Endchars);
		writer.WriteXmlNullableAttributeIntWithKoef("w:hanging", Hanging, g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:hangingChars", this.Hangingchars);
		writer.WriteXmlNullableAttributeIntWithKoef("w:firstLine", FirstLine, g_dKoef_mm_to_twips);
		// writer.WriteXmlNullableAttributeNumber("w:firstLineChars", this.Firstlinechars);
		writer.WriteXmlAttributesEnd(true);
	};
	ParaHyperlink.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "tgtFrame": {
					break;
				}
				case "tooltip": {
					this.ToolTip = reader.GetValueDecodeXml();
					break;
				}
				case "docLocation": {
					break;
				}
				case "history": {
					break;
				}
				case "anchor": {
					this.Anchor = reader.GetValueDecodeXml();
					break;
				}
				case "id": {
					let id = reader.GetValueDecodeXml();
					let rel = reader.rels.getRelationship(id);
					if (rel) {
						this.Value = rel.target;
					}
					break;
				}
			}
		}
	};
	ParaHyperlink.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		CParagraphContentWithParagraphLikeContent.prototype.fromXml.call(this, reader);
	};
	ParaHyperlink.prototype.toXml = function(writer, name) {
		let id = null;
		if (this.Value) {
			id = writer.context.part.addRelationship(AscCommon.openXml.Types.hyperlink.relationType, this.Value, openXml.TargetMode.external);
		}

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("r:id", id);
		writer.WriteXmlNullableAttributeStringEncode("w:tooltip", this.Tooltip);
		writer.WriteXmlNullableAttributeStringEncode("w:anchor", this.Anchor);
		// writer.WriteXmlNullableAttributeStringEncode("w:tgtFrame", this.tgtFrame);
		// writer.WriteXmlNullableAttributeStringEncode("w:docLocation", this.docLocation);
		writer.WriteXmlNullableAttributeBool("w:history", true);
		writer.WriteXmlAttributesEnd();
		CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
		writer.WriteXmlNodeEnd(name);
	};
	ParaRun.prototype.readDrawing = function(reader) {
		var oParagraph = this.GetParagraph();
		var drawing = new ParaDrawing(0, 0, null, oParagraph.Parent.DrawingDocument, oParagraph.Parent, oParagraph);
		drawing.fromXml(reader);
		if (drawing.GraphicObj) {
			let newItem = drawing;
			let oParaDrawing = drawing;
			if (null != oParaDrawing.SimplePos)
				oParaDrawing.setSimplePos(oParaDrawing.SimplePos.Use, oParaDrawing.SimplePos.X, oParaDrawing.SimplePos.Y);
			if (null != oParaDrawing.Extent)
				oParaDrawing.setExtent(oParaDrawing.Extent.W, oParaDrawing.Extent.H);
			if (null != oParaDrawing.wrappingPolygon)
				oParaDrawing.addWrapPolygon(oParaDrawing.wrappingPolygon);
			// if (oDrawing.ParaMath)
			// 	oParaDrawing.Set_ParaMath(oDrawing.ParaMath);

			if (oParaDrawing.GraphicObj) {

				// if (oParaDrawing.GraphicObj.setLocks && graphicFramePr.locks > 0) {
				// 	oParaDrawing.GraphicObj.setLocks(graphicFramePr.locks);
				// }
				if (oParaDrawing.GraphicObj.getObjectType() !== AscDFH.historyitem_type_ChartSpace &&
					oParaDrawing.GraphicObj.getObjectType() !== AscDFH.historyitem_type_SmartArt)
				{
					if (!oParaDrawing.GraphicObj.spPr) {
						oParaDrawing.GraphicObj = null;
					}
				}
				if (AscCommon.isRealObject(oParaDrawing.docPr) && oParaDrawing.docPr.isHidden) {
					oParaDrawing.GraphicObj = null;
				}
				let oGrObject = oParaDrawing.GraphicObj;
				if (oGrObject) {
					if (oGrObject.bEmptyTransform) {
						var oXfrm = new AscFormat.CXfrm();
						oXfrm.setOffX(0);
						oXfrm.setOffY(0);
						//oXfrm.setChOffX(0);
						//oXfrm.setChOffY(0);
						oXfrm.setExtX(oParaDrawing.Extent.W);
						oXfrm.setExtY(oParaDrawing.Extent.H);
						//oXfrm.setChExtX(oParaDrawing.Extent.W);
						//oXfrm.setChExtY(oParaDrawing.Extent.H);
						oXfrm.setParent(oParaDrawing.GraphicObj.spPr);
						oGrObject.checkEmptySpPrAndXfrm(oXfrm);
						delete oParaDrawing.GraphicObj.bEmptyTransform;
					}
					if (drawing_Anchor === oParaDrawing.DrawingType && typeof AscCommon.History.RecalcData_Add === "function")//TODO некорректная проверка typeof
						AscCommon.History.RecalcData_Add({
							Type: AscDFH.historyitem_recalctype_Flow,
							Data: oParaDrawing
						});
				}
			}
			return oParaDrawing.GraphicObj ? newItem : null;
		}
		return null;
	};
	ParaRun.prototype.readVMLDrawing = function(reader, name) {
		let oVMLConverter = new AscFormat.CVMLToDrawingMLConverter();
		return oVMLConverter.createParaDrawingMLFromVMLNode(reader, name, this.Paragraph);
	};
	ParaRun.prototype.fromXml = function(reader, opt_paragraphContent) {
		let oReadResult = reader.context.oReadResult;
		let footnotes = oReadResult.footnotes;
		let endnotes = oReadResult.endnotes;
		let isMathRun = this.IsMathRun();
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			var newItem = null;
			switch (name) {
				case "annotationRef":
					break;
				case "br":
					newItem = new AscWord.CRunBreak(AscWord.break_Line);
					newItem.fromXml(reader);
					//reinit because of logic in constructor
					newItem = new AscWord.CRunBreak(newItem.BreakType, newItem.Clear);
					break;
				case "commentReference":
					break;
				case "contentPart":
					break;
				case "continuationSeparator":
					newItem = new AscWord.CRunContinuationSeparator();
					break;
				case "cr":
					newItem = new AscWord.CRunBreak(AscWord.break_Line);
					break;
				case "dayLong":
					break;
				case "dayShort":
					break;
				case "delInstrText":
					this.AddInstrText(reader.GetTextDecodeXml(), -1);
					break;
				case "delText":
					this.AddText(reader.GetTextDecodeXml(), -1);
					break;
				case "drawing":
					newItem = this.readDrawing(reader);
					break;
				case "AlternateContent":
					let oRun = this;
					let elem = new CT_XmlNode(function(reader, name) {
						if(!newItem) {
							if ("Choice" === name || "Fallback" === name) {
								let elem = new CT_XmlNode(function(reader, name) {
									if ("drawing" === name) {
										newItem = oRun.readDrawing(reader);
									}
									return true;
								});
								elem.fromXml(reader);
								return elem;
							}
						}
						return true;
					});
					elem.fromXml(reader);
					break;
				case "endnoteRef":
					newItem = new AscWord.CRunEndnoteRef(null);
					break;
				case "endnoteReference":
					let ednRef = new CT_FtnEdnRef();
					ednRef.fromXml(reader);
					let endnote = endnotes[ednRef.id];
					if (endnote) {
						oReadResult.logicDocument.Endnotes.AddEndnote(endnote.content);
						newItem = new AscWord.CRunEndnoteReference(endnote.content, ednRef.customMarkFollows);
					}
					break;
				case "fldChar":
					newItem = new ParaFieldChar();
					newItem.fromXml(reader);
					break;
				case "footnoteRef":
					newItem = new AscWord.CRunFootnoteRef(null);
					break;
				case "footnoteReference":
					let ftnRef = new CT_FtnEdnRef();
					ftnRef.fromXml(reader);
					let footnote = footnotes[ftnRef.id];
					if (footnote) {
						oReadResult.logicDocument.Footnotes.AddFootnote(footnote.content);
						newItem = new AscWord.CRunFootnoteReference(footnote.content, ftnRef.customMarkFollows);
					}
					break;
				case "instrText":
					this.AddInstrText(reader.GetTextDecodeXml(), -1);
					break;
				case "lastRenderedPageBreak":
					break;
				case "monthLong":
					break;
				case "monthShort":
					break;
				case "noBreakHyphen":
					newItem = AscWord.CreateNonBreakingHyphen();
					break;
				case "object":
					newItem = this.readVMLDrawing(reader, name);
					break;
				case "pgNum":
					newItem = new AscWord.CRunPageNum();
					break;
				case "pict":
					newItem = this.readVMLDrawing(reader, name);
					break;
				case "ptab":
					break;
				case "rPr":
					if (para_Math_Run === this.Type) {
						let fullName = reader.GetName();
						//todo if another ns
						if ("m:rPr" === fullName) {
							let mrPr = new CMPrp();
							mrPr.fromXml(reader);
							this.Set_MathPr(mrPr);
							continue;
						} else if ("a:rPr" === fullName) {
							this.Pr.fromDrawingML(reader);
							continue;
						}
					}
					this.Pr = new CTextPr();
					this.Pr.fromXml(reader, this);
					break;
				case "ruby":
					break;
				case "separator":
					newItem = new AscWord.CRunSeparator();
					break;
				case "softHyphen":
					break;
				case "sym":
					let sym = new CT_Sym();
					sym.fromXml(reader);
					if (!isMathRun && null !== sym.char) {
						//todo split here
						if (!oReadResult.runsToSplitBySym[this.Id]) {
							oReadResult.runsToSplitBySym[this.Id] = {run: this, syms: []};
						}
						oReadResult.runsToSplitBySym[this.Id].syms.push({pos: this.GetElementsCount(), sym: sym});
						this.AddText(String.fromCharCode(sym.char), -1);
					}
					break;
				case "t":
					this.AddText(reader.GetTextDecodeXml(), -1);
					break;
				case "tab":
					newItem = new AscWord.CRunTab();
					break;
				case "yearLong":
					break;
				case "yearShort":
					break;
				case "del" : {
					let trackChange = new CT_TrackChange();
					trackChange.run = this;
					trackChange.fromXml(reader);
					this.SetReviewTypeWithInfo(reviewtype_Remove, trackChange.ReviewInfo, false);
					break;
				}
				case "ins" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.run = this;
						trackChange.fromXml(reader);
						this.SetReviewTypeWithInfo(reviewtype_Add, trackChange.ReviewInfo, false);
					}
					break;
				}
			}
			if (newItem) {
				if (!isMathRun) {
					this.AddToContentToEnd(newItem, false);
				} else if (opt_paragraphContent) {
					let oNewRun = new ParaRun(opt_paragraphContent.GetParagraph());
					oNewRun.AddToContentToEnd(newItem, false);
					opt_paragraphContent.AddToContentToEnd(oNewRun);
				}
			}
		}
	};
	ParaRun.prototype.toXml = function(writer, name, delText, startIndex) {
		var t = this;
		let pageFieldIndex = -1;
		startIndex = startIndex || 0;
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		var reviewType = this.GetReviewType();
		if (para_Math_Run === this.Type && reviewtype_Common !== reviewType) {
			writer.context.docSaveParams.writeNestedReviewType(reviewType, this.GetReviewInfo(), function(reviewType, reviewInfo, delText, fCallback){
				let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, reviewInfo);
				trackChange.writeCallback = function(){fCallback(delText);};
				if (reviewtype_Remove === reviewType) {
					if (reviewInfo.IsMovedFrom()) {
						writer.WriteXmlNullable(trackChange, "w:moveFrom");
					} else {
						delText = true;
						writer.WriteXmlNullable(trackChange, "w:del");
					}
				} else if (reviewtype_Add === reviewType) {
					if (reviewInfo.IsMovedTo()) {
						writer.WriteXmlNullable(trackChange, "w:moveTo");
					} else {
						writer.WriteXmlNullable(trackChange, "w:ins");
					}
				}
			}, function(delText) {
				pageFieldIndex = t.toXmlContent(writer, startIndex, delText);
			});
		} else {
			pageFieldIndex = t.toXmlContent(writer, startIndex, delText);
		}
		writer.WriteXmlNodeEnd(name);
		if (pageFieldIndex > 0) {
			let paraField, num;
			let pageNum = this.Content[pageFieldIndex];
			if (pageNum.Type == para_PageCount) {
				paraField = new ParaField(fieldtype_PAGECOUNT, "\\* MERGEFORMAT");
				num = pageNum.GetPageCountValue();

			} else {
				paraField = new ParaField(fieldtype_PAGENUM, "\\* MERGEFORMAT");
				num = pageNum.GetPageNumValue();
			}
			let oldContent = this.Content;
			this.Content = [];
			this.AddText(num.toString());
			paraField.Content.push(this);
			paraField.toXml(writer, "w:fldSimple");
			this.Content = oldContent;

			this.toXml(writer, name, delText, pageFieldIndex + 1);
		}
	};
	ParaRun.prototype.toXmlContent = function(writer, startIndex, delText) {
		if (this.MathPrp && !this.MathPrp.Is_Empty()) {
			this.MathPrp.toXml(writer, "m:rPr");
		}
		if (this.Pr) {
			//todo
			if (this.Paragraph.bFromDocument) {
				this.Pr.toXml(writer, "w:rPr", this, undefined, true);
			} else {
				this.Pr.toDrawingML(writer, "a:rPr");
			}
		}
		let footnoteIdToIndex;
		let endnoteIdToIndex;
		if(writer.context.docSaveParams) {
			footnoteIdToIndex = writer.context.docSaveParams.footnoteIdToIndex;
			endnoteIdToIndex = writer.context.docSaveParams.endnoteIdToIndex;
		}
		let index;
		let textNs = "w:";
		let textName = delText ? "delText" : "t";
		let textInstrName = delText ? "delInstrText" : "instrText";
		if (para_Math_Run === this.Type) {
			textNs = "m:";
			textName = "t";
		}

		let sCurText = "";
		let sCurInstrText = "";
		for (let i = startIndex; i < this.Content.length; ++i)
		{
			var item = this.Content[i];
			if (para_Text === item.Type || para_Space === item.Type || para_Math_Text === item.Type) {
				sCurInstrText = writeRunText(writer, textNs, textInstrName, sCurInstrText);
			} else if (para_InstrText === item.Type) {
				sCurText = writeRunText(writer, textNs, textName, sCurText);
			} else {
				sCurText = writeRunText(writer, textNs, textName, sCurText);
				sCurInstrText = writeRunText(writer, textNs, textInstrName, sCurInstrText);
			}
			switch ( item.Type ) {
				case para_Text:
					if (item.IsNoBreakHyphen()) {
						sCurText = writeRunText(writer, textNs, textName, sCurText);
						writer.WriteXmlNodeStart("w:noBreakHyphen");
						writer.WriteXmlAttributesEnd(true);
					} else {
						sCurText += AscCommon.encodeSurrogateChar(item.Value);
					}
					break;
				case para_Math_Ampersand:
					sCurText += "&";
					break;
				case para_Math_BreakOperator:
				case para_Math_Text:
					sCurText += AscCommon.encodeSurrogateChar(item.value);
					break;
				case para_Space:
					sCurText += AscCommon.encodeSurrogateChar(item.Value);
					break;
				case para_Tab:
					writer.WriteXmlNodeStart("w:tab");
					writer.WriteXmlAttributesEnd(true);
					break;
				case para_NewLine:
					item.toXml(writer, "w:br");
					break;
				case para_Separator:
					writer.WriteXmlNodeStart("w:separator");
					writer.WriteXmlAttributesEnd(true);
					break;
				case para_ContinuationSeparator:
					writer.WriteXmlNodeStart("w:continuationSeparator");
					writer.WriteXmlAttributesEnd(true);
					break;
				case para_FootnoteRef:
					writer.WriteXmlNodeStart("w:footnoteRef");
					writer.WriteXmlAttributesEnd(true);
					break;
				case para_FootnoteReference:
					index = footnoteIdToIndex && footnoteIdToIndex[item.Footnote.GetId()];
					if (undefined !== index) {
						let ftnRef = new CT_FtnEdnRef();
						ftnRef.fromVal(index, item.CustomMark);
						ftnRef.toXml(writer, "w:footnoteReference");
					}
					break;
				case para_EndnoteRef:
					writer.WriteXmlNodeStart("w:endnoteRef");
					writer.WriteXmlAttributesEnd(true);
					break;
				case para_EndnoteReference:
					index = endnoteIdToIndex && endnoteIdToIndex[item.Footnote.GetId()];
					if (undefined !== index) {
						let endRef = new CT_FtnEdnRef();
						endRef.fromVal(index, item.CustomMark);
						endRef.toXml(writer, "w:endnoteReference");
					}
					break;
				case para_FieldChar:
					item.toXml(writer, "w:fldChar");
					break;
				case para_InstrText:
					sCurInstrText += AscCommon.encodeSurrogateChar(item.Value);
					break;
				case para_Drawing:
					item.toXml(writer, "w:drawing");
					break;
				case para_PageNum:
				case para_PageCount:
					return i;
			}
		}
		sCurText = writeRunText(writer, textNs, textName, sCurText);
		sCurInstrText = writeRunText(writer, textNs, textInstrName, sCurInstrText);
		return -1;
	};
	CTextPr.prototype.fromXml = function(reader, run) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "rStyle" : {
					this.RStyle = CT_StringW.prototype.toVal(reader, this.RStyle);
					reader.context.oReadResult.runStyles.push({pPr: this, style: this.RStyle});
					break;
				}
				case "rFonts" : {
					this.RFonts = new CRFonts();
					this.RFonts.fromXml(reader);
					break;
				}
				case "b" : {
					this.Bold = CT_BoolW.prototype.toVal(reader, this.Bold);
					break;
				}
				case "bCs" : {
					this.BoldCS = CT_BoolW.prototype.toVal(reader, this.BoldCS);
					break;
				}
				case "i" : {
					this.Italic = CT_BoolW.prototype.toVal(reader, this.Italic);
					break;
				}
				case "iCs" : {
					this.ItalicCS = CT_BoolW.prototype.toVal(reader, this.ItalicCS);
					break;
				}
				case "caps" : {
					this.Caps = CT_BoolW.prototype.toVal(reader, this.Caps);
					break;
				}
				case "smallCaps" : {
					this.SmallCaps = CT_BoolW.prototype.toVal(reader, this.SmallCaps);
					break;
				}
				case "strike" : {
					this.Strikeout = CT_BoolW.prototype.toVal(reader, this.Strikeout);
					break;
				}
				case "dstrike" : {
					this.DStrikeout = CT_BoolW.prototype.toVal(reader, this.DStrikeout);
					break;
				}
				// case "outline" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Outline.push(elem);
				// 	break;
				// }
				// case "shadow" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Shadow.push(elem);
				// 	break;
				// }
				// case "emboss" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Emboss.push(elem);
				// 	break;
				// }
				// case "imprint" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Imprint.push(elem);
				// 	break;
				// }
				// case "noProof" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Noproof.push(elem);
				// 	break;
				// }
				// case "snapToGrid" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Snaptogrid.push(elem);
				// 	break;
				// }
				case "vanish" : {
					this.Vanish = CT_BoolW.prototype.toVal(reader, this.Vanish);
					break;
				}
				// case "webHidden" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Webhidden.push(elem);
				// 	break;
				// }
				case "color" : {
					elem = new CT_Color("val", "themeColor", "themeTint", "themeShade");
					elem.fromXml(reader);
					this.Color = elem.getColor(0, 0, 0);
					this.Unifill = elem.getUnifill();
					break;
				}
				case "spacing" : {
					this.Spacing = AscCommon.universalMeasureToMm(CT_StringW.prototype.toVal(reader, this.Spacing), AscCommonWord.g_dKoef_twips_to_mm, this.Spacing);
					break;
				}
				// case "w" : {
				// 	elem = new CT_TextScale();
				// 	elem.fromXml(reader);
				// 	this.W.push(elem);
				// 	break;
				// }
				// case "kern" : {
				// 	elem = new CT_HpsMeasure();
				// 	elem.fromXml(reader);
				// 	this.Kern.push(elem);
				// 	break;
				// }
				case "position" : {
					this.Position = AscCommon.universalMeasureToMm(CT_StringW.prototype.toVal(reader, this.Position), AscCommonWord.g_dKoef_pt_to_mm / 2, this.Position);
					break;
				}
				case "sz" : {
					elem = CT_UIntW.prototype.toVal(reader, undefined);
					if (undefined !== elem) {
						this.FontSize = elem / 2;
					}
					break;
				}
				case "szCs" : {
					elem = CT_UIntW.prototype.toVal(reader, undefined);
					if (undefined !== elem) {
						this.FontSizeCS = elem / 2;
					}
					break;
				}
				case "highlight" : {
					elem = CT_StringW.prototype.toVal(reader, undefined);
					let rgb = AscFormat.map_hightlight[elem];
					if ("none" === elem) {
						this.HighLight = highlight_None;
					} else if (undefined !== rgb) {
						this.HighLight = new CDocumentColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
					}
					break;
				}
				case "u" : {
					elem = new CT_Underline();
					elem.fromXml(reader);
					this.Underline = elem.Val;
					break;
				}
				// case "effect" : {
				// 	elem = new CT_TextEffect();
				// 	elem.fromXml(reader);
				// 	this.Effect.push(elem);
				// 	break;
				// }
				// case "bdr" : {
				// 	elem = new CT_Border();
				// 	elem.fromXml(reader);
				// 	this.Bdr.push(elem);
				// 	break;
				// }
				case "shd" : {
					this.Shd = new CDocumentShd();
					this.Shd.fromXml(reader);
					break;
				}
				// case "fitText" : {
				// 	elem = new CT_FitText();
				// 	elem.fromXml(reader);
				// 	this.Fittext.push(elem);
				// 	break;
				// }
				case "vertAlign" : {
					this.VertAlign = fromXml_ST_VerticalAlignRun(CT_StringW.prototype.toVal(reader, this.VertAlign), this.VertAlign);
					break;
				}
				case "rtl" : {
					this.RTL = CT_BoolW.prototype.toVal(reader, this.RTL);
					break;
				}
				case "cs" : {
					this.CS = CT_BoolW.prototype.toVal(reader, this.CS);
					break;
				}
				// case "em" : {
				// 	elem = new CT_Em();
				// 	elem.fromXml(reader);
				// 	this.Em.push(elem);
				// 	break;
				// }
				case "lang" : {
					this.Lang = new CLang();
					this.Lang.fromXml(reader);
					break;
				}
				// case "eastAsianLayout" : {
				// 	elem = new CT_EastAsianLayout();
				// 	elem.fromXml(reader);
				// 	this.Eastasianlayout.push(elem);
				// 	break;
				// }
				// case "specVanish" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Vanish = elem.getVal(undefined);
				// 	break;
				// }
				// case "oMath" : {
				// 	elem = new CT_BoolW();
				// 	elem.fromXml(reader);
				// 	this.Omath.push(elem);
				// 	break;
				// }
				case "del" : {
					let trackChange = new CT_TrackChange();
					trackChange.fromXml(reader);
					if (run) {
						run.SetReviewTypeWithInfo(reviewtype_Remove, trackChange.ReviewInfo, false);
					}
					break;
				}
				case "ins" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						if (run) {
							run.SetReviewTypeWithInfo(reviewtype_Add, trackChange.ReviewInfo, false);
						}
					}
					break;
				}
				case "moveFrom" : {
					let trackChange = new CT_TrackChange();
					trackChange.ReviewInfo.SetMove(Asc.c_oAscRevisionsMove.MoveFrom);
					trackChange.fromXml(reader);
					if (run) {
						run.SetReviewTypeWithInfo(reviewtype_Remove, trackChange.ReviewInfo, false);
					}
					break;
				}
				case "moveTo" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.ReviewInfo.SetMove(Asc.c_oAscRevisionsMove.MoveTo);
						trackChange.fromXml(reader);
						if (run) {
							run.SetReviewTypeWithInfo(reviewtype_Add, trackChange.ReviewInfo, false);
						}
					}
					break;
				}
				case "rPrChange" : {
					if (reader.context.oReadResult.checkReadRevisions()) {
						let trackChange = new CT_TrackChange();
						trackChange.fromXml(reader);
						if (run) {
							run.SetPrChange(trackChange.rPrChange, trackChange.ReviewInfo);
						}
					}
					break;
				}
				case "textOutline" : {
					let oLn = new AscFormat.CLn();
					oLn.fromXml(reader);
					this.TextOutline = oLn;
					break;
				}
				case "textFill" : {
					let oThis = this;
					let oNode = new CT_XmlNode(function (reader, name) {
						if(AscFormat.CUniFill.prototype.isFillName(name)) {
							oThis.TextFill = new AscFormat.CUniFill();
							oThis.TextFill.fromXml(reader, name);
							if(AscFormat.isRealNumber(oThis.TextFill.transparent))
							{
								oThis.TextFill.transparent = 255 - oThis.TextFill.transparent;
							}
						}
						return true;
					});
					oNode.fromXml(reader);
					break;
				}
					//c_oSerProp_rPrType.Del
					//c_oSerProp_rPrType.Ins
					//c_oSerProp_rPrType.MoveFrom
					//c_oSerProp_rPrType.MoveTo
			}
		}
	};
	CTextPr.prototype.toXml = function(writer, name, run, EndRun, opt_ignoreEmpty) {
		var RFonts = this.RFonts && this.RFonts.Is_Empty() ? null : this.RFonts;
		var Color = new CT_Color("val", "themeColor", "themeTint", "themeShade");
		Color.setColor(this.Color);
		Color.setUniFill(this.Unifill);
		if (Color.isEmpty()) {
			Color = null;
		}
		var HighLight;
		if (undefined !== this.HighLight) {
			HighLight = CT_StringW.prototype.fromVal(highlight_None !== this.HighLight ? this.HighLight.ToHexColor() : "none");
		}
		var Underline;
		if (undefined !== this.Underline) {
			Underline = new CT_Underline();
			Underline.Val = this.Underline;
		}
		var Lang = this.Lang && this.Lang.Is_Empty() ? null : this.Lang;

		let startPos = writer.GetCurPosition();
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		let headerPos = writer.GetCurPosition();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.RStyle), "w:rStyle");
		writer.WriteXmlNullable(RFonts, "w:rFonts");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Bold), "w:b");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.BoldCS), "w:bCs");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Italic), "w:i");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.ItalicCS), "w:iCs");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Caps), "w:caps");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.SmallCaps), "w:smallCaps");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Strikeout), "w:strike");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.DStrikeout), "w:dstrike");
		// writer.WriteXmlNullable(this.Outline, "w:outline");
		// writer.WriteXmlNullable(this.Shadow, "w:shadow");
		// writer.WriteXmlNullable(this.Emboss, "w:emboss");
		// writer.WriteXmlNullable(this.Imprint, "w:imprint");
		// writer.WriteXmlNullable(this.NoProof, "w:noProof");
		// writer.WriteXmlNullable(this.SnapToGrid, "w:snapToGrid");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Vanish), "w:vanish");
		// writer.WriteXmlNullable(this.WebHidden, "w:webHidden");
		writer.WriteXmlNullable(Color, "w:color");
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.Spacing, g_dKoef_mm_to_twips), "w:spacing");
		// writer.WriteXmlNullable(this.W, "w:w");
		// writer.WriteXmlNullable(this.Kern, "w:kern");
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.Position, g_dKoef_mm_to_pt * 2), "w:position");
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.FontSize, 2), "w:sz");
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.FontSizeCS, 2), "w:szCs");
		writer.WriteXmlNullable(HighLight, "w:highlight");
		writer.WriteXmlNullable(Underline, "w:u");
		// writer.WriteXmlNullable(this.Effect, "w:effect");
		// writer.WriteXmlNullable(this.Bdr, "w:bdr");
		writer.WriteXmlNullable(this.Shd, "w:shd");
		// writer.WriteXmlNullable(this.FitText, "w:fitText");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_VerticalAlignRun(this.VertAlign)), "w:vertAlign");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.RTL), "w:rtl");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.CS), "w:cs");
		// writer.WriteXmlNullable(this.Em, "w:em");
		writer.WriteXmlNullable(Lang, "w:lang");
		if(this.TextFill) {
			writer.WriteXmlString("<w14:textFill>");
			let nOldDocType = writer.context.docType;
			writer.context.docType = AscFormat.XMLWRITER_DOC_TYPE_WORDART;
			this.TextFill.toXml(writer);
			writer.context.docType = nOldDocType;
			writer.WriteXmlString("</w14:textFill>");
		}
		if(this.TextOutline) {
			let nOldDocType = writer.context.docType;
			writer.context.docType = AscFormat.XMLWRITER_DOC_TYPE_WORDART;
			this.TextOutline.toXml(writer, "w14:textOutline");
			writer.context.docType = nOldDocType;
		}
		// writer.WriteXmlNullable(this.EastAsianLayout, "w:eastAsianLayout");
		// writer.WriteXmlNullable(Vanish, "w:specVanish");
		// writer.WriteXmlNullable(this.OMath, "w:oMath");
		if (EndRun) {
			if (run.ReviewInfo && reviewtype_Common !== EndRun.GetReviewType()) {
				let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, EndRun.ReviewInfo);
				let ReviewType = EndRun.GetReviewType();
				if (reviewtype_Remove === ReviewType) {
					if (EndRun.ReviewInfo.IsMovedFrom()) {
						writer.WriteXmlNullable(trackChange, "w:moveFrom");
					} else {
						writer.WriteXmlNullable(trackChange, "w:del");
					}
				} else if (reviewtype_Add === ReviewType) {
					if (EndRun.ReviewInfo.IsMovedTo()) {
						writer.WriteXmlNullable(trackChange, "w:moveTo");
					} else {
						writer.WriteXmlNullable(trackChange, "w:ins");
					}
				}
			}
		}
		if (run && run.Pr.PrChange && run.Pr.ReviewInfo) {
			let trackChange = new CT_TrackChange(writer.context.docSaveParams.trackRevisionId++, run.Pr.ReviewInfo);
			trackChange.rPrChange = run.Pr.PrChange;
			writer.WriteXmlNullable(trackChange, "w:rPrChange");
		}
		//todo
		if (opt_ignoreEmpty && headerPos === writer.GetCurPosition()) {
			writer.Seek(startPos);
		} else {
			writer.WriteXmlNodeEnd(name);
		}
	};
	CRFonts.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "hint": {
					this.Hint = fromXml_ST_Hint(reader.GetValue());
					break;
				}
				case "ascii": {
					this.Ascii = {Name: reader.GetValueDecodeXml(), Index: -1};
					break;
				}
				case "hAnsi": {
					this.HAnsi = {Name: reader.GetValueDecodeXml(), Index: -1};
					break;
				}
				case "eastAsia": {
					this.EastAsia = {Name: reader.GetValueDecodeXml(), Index: -1};
					break;
				}
				case "cs": {
					this.CS = {Name: reader.GetValueDecodeXml(), Index: -1};
					break;
				}
				case "asciiTheme": {
					this.AsciiTheme = reader.GetValue();
					break;
				}
				case "hAnsiTheme": {
					this.HAnsiTheme = reader.GetValue();
					break;
				}
				case "eastAsiaTheme": {
					this.EastAsiaTheme = reader.GetValue();
					break;
				}
				case "cstheme": {
					this.CSTheme = reader.GetValue();
					break;
				}
			}
		}
	};
	CRFonts.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CRFonts.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:hint", toXml_ST_Hint(this.Hint));
		writer.WriteXmlNullableAttributeStringEncode("w:ascii", this.Ascii && this.Ascii.Name);
		writer.WriteXmlNullableAttributeString("w:asciiTheme", this.AsciiTheme);
		writer.WriteXmlNullableAttributeStringEncode("w:eastAsia", this.EastAsia && this.EastAsia.Name);
		writer.WriteXmlNullableAttributeString("w:eastAsiaTheme", this.EastAsiaTheme);
		writer.WriteXmlNullableAttributeStringEncode("w:hAnsi", this.HAnsi && this.HAnsi.Name);
		writer.WriteXmlNullableAttributeString("w:hAnsiTheme", this.HAnsiTheme);
		writer.WriteXmlNullableAttributeStringEncode("w:cs", this.CS && this.CS.Name);
		writer.WriteXmlNullableAttributeString("w:cstheme", this.CSTheme);
		writer.WriteXmlAttributesEnd(true);
	};
	CLang.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					this.Val = Asc.g_oLcidNameToIdMap[reader.GetValue()];
					break;
				}
				case "eastAsia": {
					this.EastAsia = Asc.g_oLcidNameToIdMap[reader.GetValue()];
					break;
				}
				case "bidi": {
					this.Bidi = Asc.g_oLcidNameToIdMap[reader.GetValue()];
					break;
				}
			}
		}
	};
	CLang.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CLang.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:val", Asc.g_oLcidIdToNameMap[this.Val]);
		writer.WriteXmlNullableAttributeStringEncode("w:eastAsia", Asc.g_oLcidIdToNameMap[this.EastAsia]);
		writer.WriteXmlNullableAttributeStringEncode("w:bidi", Asc.g_oLcidIdToNameMap[this.Bidi]);
		writer.WriteXmlAttributesEnd(true);
	};
	AscWord.CRunBreak.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					this.BreakType = fromXml_ST_BrType(reader.GetValue(), this.BreakType);
					break;
				}
				case "clear": {
					break;
				}
			}
		}
	};
	AscWord.CRunBreak.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscWord.CRunBreak.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:type", toXml_ST_BrType(this.BreakType));
		// writer.WriteXmlNullableAttributeString("w:clear", toXml_ST_BrClear(this.clear));
		writer.WriteXmlAttributesEnd(true);
	};
	ParaField.prototype.readAttr = function(reader) {
		let instr;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "instr": {
					let instr = reader.GetValueDecodeXml();
					var oParser = new CFieldInstructionParser();
					let Instruction = oParser.GetInstructionClass(instr);
					oParser.InitParaFieldArguments(Instruction.Type, instr, this);
					break;
				}
				case "fldLock": {
					break;
				}
				case "dirty": {
					break;
				}
			}
		}
	};
	ParaField.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		CParagraphContentWithParagraphLikeContent.prototype.fromXml.call(this, reader);
	};
	ParaField.prototype.toXml = function(writer, name) {
		let instr = this.GetInstructionLine();
		if (!instr) {
			return;
		}
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:instr", instr);
		// writer.WriteXmlNullableAttributeBool("w:fldLock", this.fldLock);
		// writer.WriteXmlNullableAttributeBool("w:dirty", this.dirty);
		writer.WriteXmlAttributesEnd();
		CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
		writer.WriteXmlNodeEnd(name);
	};
	ParaFieldChar.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "fldCharType": {
					let charType = fromXml_ST_FldCharType(reader.GetValue(), this.CharType);
					this.Init(charType, reader.context.oReadResult.logicDocument);
					break;
				}
				case "fldLock": {
					break;
				}
				case "dirty": {
					break;
				}
			}
		}
	};
	ParaFieldChar.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "ffData" : {
					break;
				}
			}
		}
	};
	ParaFieldChar.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:fldCharType", toXml_ST_FldCharType(this.CharType));
		// writer.WriteXmlNullableAttributeBool("w:fldLock", this.fldLock);
		// writer.WriteXmlNullableAttributeBool("w:dirty", this.dirty);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.ffData, "w:ffData");
		writer.WriteXmlNodeEnd(name);
	};
	CSectionPr.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			// switch (reader.GetNameNoNS()) {
			// 	case "rsidRPr": {
			// 		this.RsidRPr = reader.GetValueByte();
			// 		break;
			// 	}
			// 	case "rsidDel": {
			// 		this.RsidDel = reader.GetValueByte();
			// 		break;
			// 	}
			// 	case "rsidR": {
			// 		this.RsidR = reader.GetValueByte();
			// 		break;
			// 	}
			// 	case "rsidSect": {
			// 		this.RsidSect = reader.GetValueByte();
			// 		break;
			// 	}
			// }
		}
	};
	CSectionPr.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			var name = reader.GetNameNoNS();
			switch (name) {
				case "headerReference" :
				case "footerReference" : {
					var type = name === "headerReference" ? AscCommon.hdrftr_Header : AscCommon.hdrftr_Footer;
					var hdrFtrRef = new CT_HdrFtrRef();
					hdrFtrRef.fromXml(reader);
					if (hdrFtrRef.Id) {
						var rel = reader.rels.getRelationship(hdrFtrRef.Id);
						var hdrFtrPart = reader.rels.pkg.getPartByUri(rel.targetFullName);
						if (hdrFtrPart) {
							var contentHdrFtr = hdrFtrPart.getDocumentContent();
							var hdrFtrReader = new StaxParser(contentHdrFtr, hdrFtrPart, reader.context);
							var hdrFtr = new CHeaderFooter(this.LogicDocument.HdrFtr, this.LogicDocument, this.LogicDocument.DrawingDocument, type);
							hdrFtr.fromXml(hdrFtrReader);
							if(AscCommon.hdrftr_Header === type) {
								switch(hdrFtrRef.Type) {
									case "even":
										this.Set_Header_Even(hdrFtr);
										break;
									case "default":
										this.Set_Header_Default(hdrFtr);
										break;
									case "first":
										this.Set_Header_First(hdrFtr);
										break;
								}
							} else {
								switch(hdrFtrRef.Type) {
									case "even":
										this.Set_Footer_Even(hdrFtr);
										break;
									case "default":
										this.Set_Footer_Default(hdrFtr);
										break;
									case "first":
										this.Set_Footer_First(hdrFtr);
										break;
								}
							}
						}
					}
					break;
				}
				case "footnotePr" : {
					this.FootnotePr.fromXml(reader, true);
					break;
				}
				case "endnotePr" : {
					this.EndnotePr.fromXml(reader, false);
					break;
				}
				case "type" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Type = fromXml_ST_SectionMark(elem.getVal(undefined));
					break;
				}
				case "pgSz" : {
					elem = new CSectionPageSize();
					elem.fromXml(reader);
					this.SetPageSize(elem.W, elem.H);
					this.SetOrientation(elem.Orient);
					break;
				}
				case "pgMar" : {
					this.PageMargins.fromXml(reader);
					break;
				}
				// case "paperSrc" : {
				// 	this.PaperSrc = new CT_PaperSource();
				// 	this.PaperSrc.fromXml(reader);
				// 	break;
				// }
				case "pgBorders" : {
					this.Borders.fromXml(reader);
					break;
				}
				case "lnNumType" : {
					elem = new CSectionLnNumType(1, undefined, undefined, undefined);
					elem.fromXml(reader);
					this.SetLineNumbers(elem.CountBy, elem.Distance, elem.Start, elem.Restart);
					break;
				}
				case "pgNumType" : {
					elem = new CSectionPageNumType();
					elem.fromXml(reader);
					this.Set_PageNum_Start(elem.Start);
					break;
				}
				case "cols" : {
					elem = new CSectionColumns(this);
					elem.fromXml(reader);
					this.Set_Columns_EqualWidth(elem.EqualWidth);
					this.Set_Columns_Num(elem.Num);
					this.Set_Columns_Sep(elem.Sep);
					this.Set_Columns_Space(elem.Space);
					if (elem.Cols.length > 0) {
						elem.Cols.forEach(function(elem) {
							this.Set_Columns_Col(this.Columns.Cols.length, elem.W, elem.Space);
						}, this);
					}
					break;
				}
				// case "formProt" : {
				// 	this.FormProt = new CT_BoolW();
				// 	this.FormProt.fromXml(reader);
				// 	break;
				// }
				// case "vAlign" : {
				// 	this.VAlign = new CT_VerticalJc();
				// 	this.VAlign.fromXml(reader);
				// 	break;
				// }
				// case "noEndnote" : {
				// 	this.NoEndnote = new CT_BoolW();
				// 	this.NoEndnote.fromXml(reader);
				// 	break;
				// }
				case "titlePg" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.Set_TitlePage(elem.getVal(true));
					break;
				}
				// case "textDirection" : {
				// 	this.TextDirection = new CT_TextDirection();
				// 	this.TextDirection.fromXml(reader);
				// 	break;
				// }
				// case "bidi" : {
				// 	this.Bidi = new CT_BoolW();
				// 	this.Bidi.fromXml(reader);
				// 	break;
				// }
				case "rtlGutter" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.SetGutterRTL(elem.getVal(undefined));
					break;
				}
				// case "docGrid" : {
				// 	this.DocGrid = new CT_DocGrid();
				// 	this.DocGrid.fromXml(reader);
				// 	break;
				// }
				// case "printerSettings" : {
				// 	this.PrinterSettings = new CT_Rel();
				// 	this.PrinterSettings.fromXml(reader);
				// 	break;
				// }
				case "sectPrChange" : {
					// this.SectPrChange = new CT_SectPrChange();
					// this.SectPrChange.fromXml(reader);
					break;
				}
			}
		}
	};
	CSectionPr.prototype.toXml = function(writer, name) {
		var HdrFtrTypes = ["default", "default", "first", "first", "even", "even"];
		var HdrFtrs = [this.HeaderDefault, this.FooterDefault, this.HeaderFirst, this.FooterFirst, this.HeaderEven, this.FooterEven];
		var HdrFtrRefs = [null, null, null, null, null, null];
		HdrFtrs.forEach(function(elem, index) {
			if (elem) {
				var partType = 0 === (index % 2) ? AscCommon.openXml.Types.header : AscCommon.openXml.Types.footer;
				var drawingPart = writer.context.part.addPart(partType);
				drawingPart.part.setDataXml(elem, writer);
				var HeaderReferenceDef = new CT_HdrFtrRef();
				HeaderReferenceDef.Id = drawingPart.rId;
				HeaderReferenceDef.Type = HdrFtrTypes[index];
				HdrFtrRefs[index] = HeaderReferenceDef;
			}
		});
		var Type = CT_StringW.prototype.fromVal(toXml_ST_SectionMark(this.Type));
		var TitlePage = CT_BoolW.prototype.fromVal(this.TitlePage);
		var GutterRTL = CT_BoolW.prototype.fromVal(this.GutterRTL);

		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeByte("w:rsidRPr", this.RsidRPr);
		// writer.WriteXmlNullableAttributeByte("w:rsidDel", this.RsidDel);
		// writer.WriteXmlNullableAttributeByte("w:rsidR", this.RsidR);
		// writer.WriteXmlNullableAttributeByte("w:rsidSect", this.RsidSect);
		writer.WriteXmlAttributesEnd();
		HdrFtrRefs.forEach(function(elem, index) {
			var hdrFtrName = 0 === (index % 2) ? "w:headerReference" : "w:footerReference";
			writer.WriteXmlNullable(elem, hdrFtrName);
		});
		writer.WriteXmlNullable(this.FootnotePr, "w:footnotePr");
		writer.WriteXmlNullable(this.EndnotePr, "w:endnotePr");
		writer.WriteXmlNullable(Type, "w:type");
		writer.WriteXmlNullable(this.PageSize, "w:pgSz");
		writer.WriteXmlNullable(this.PageMargins, "w:pgMar");
		// writer.WriteXmlNullable(this.PaperSrc, "w:paperSrc");
		writer.WriteXmlNullable(this.Borders, "w:pgBorders");
		writer.WriteXmlNullable(this.LnNumType, "w:lnNumType");
		writer.WriteXmlNullable(this.PageNumType, "w:pgNumType");
		writer.WriteXmlNullable(this.Columns, "w:cols");
		// writer.WriteXmlNullable(this.FormProt, "w:formProt");
		// writer.WriteXmlNullable(this.VAlign, "w:vAlign");
		// writer.WriteXmlNullable(this.NoEndnote, "w:noEndnote");
		writer.WriteXmlNullable(TitlePage, "w:titlePg");
		// writer.WriteXmlNullable(this.TextDirection, "w:textDirection");
		// writer.WriteXmlNullable(this.Bidi, "w:bidi");
		writer.WriteXmlNullable(GutterRTL, "w:rtlGutter");
		// writer.WriteXmlNullable(this.DocGrid, "w:docGrid");
		// writer.WriteXmlNullable(this.PrinterSettings, "w:printerSettings");
		// writer.WriteXmlNullable(this.SectPrChange, "w:sectPrChange");
		writer.WriteXmlNodeEnd(name);
	};
	CSectionPageSize.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "w": {
					this.W = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.W);
					break;
				}
				case "h": {
					this.H = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.H);
					break;
				}
				case "orient": {
					this.Orient = fromXml_ST_PageOrientation(reader.GetValue());
					break;
				}
				// case "code": {
				// 	this.Code = reader.GetValueInt();
				// 	break;
				// }
			}
		}
	};
	CSectionPageSize.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CSectionPageSize.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:w", this.W, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:h", this.H, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:orient", toXml_ST_PageOrientation(this.Orient));
		// writer.WriteXmlNullableAttributeInt("w:code", this.Code);
		writer.WriteXmlAttributesEnd(true);
	};
	CSectionPageMargins.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "top": {
					this.Top = AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Top);
					break;
				}
				case "right": {
					this.Right = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Right);
					break;
				}
				case "bottom": {
					this.Bottom = AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Bottom);
					break;
				}
				case "left": {
					this.Left = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Left);
					break;
				}
				case "header": {
					this.Header = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Header);
					break;
				}
				case "footer": {
					this.Footer = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Footer);
					break;
				}
				case "gutter": {
					this.Gutter = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Gutter);
					break;
				}
			}
		}
	};
	CSectionPageMargins.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CSectionPageMargins.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeIntWithKoef("w:top", this.Top, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:right", this.Right, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeIntWithKoef("w:bottom", this.Bottom, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:left", this.Left, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:header", this.Header, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:footer", this.Footer, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:gutter", this.Gutter, g_dKoef_mm_to_twips);
		writer.WriteXmlAttributesEnd(true);
	};
	CSectionBorders.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "zOrder": {
					this.ZOrder = fromXml_ST_PageBorderZOrder(reader.GetValue());
					break;
				}
				case "display": {
					this.Display = fromXml_ST_PageBorderDisplay(reader.GetValue());
					break;
				}
				case "offsetFrom": {
					this.OffsetFrom = fromXml_ST_PageBorderOffset(reader.GetValue());
					break;
				}
			}
		}
	};
	CSectionBorders.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "top" : {
					this.Top.fromXml(reader);
					break;
				}
				case "left" : {
					this.Left.fromXml(reader);
					break;
				}
				case "bottom" : {
					this.Bottom.fromXml(reader);
					break;
				}
				case "right" : {
					this.Right.fromXml(reader);
					break;
				}
			}
		}
	};
	CSectionBorders.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:zOrder", toXml_ST_PageBorderZOrder(this.ZOrder));
		writer.WriteXmlNullableAttributeString("w:display", toXml_ST_PageBorderDisplay(this.Display));
		writer.WriteXmlNullableAttributeString("w:offsetFrom", toXml_ST_PageBorderOffset(this.OffsetFrom));
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Top, "w:top");
		writer.WriteXmlNullable(this.Left, "w:left");
		writer.WriteXmlNullable(this.Bottom, "w:bottom");
		writer.WriteXmlNullable(this.Right, "w:right");
		writer.WriteXmlNodeEnd(name);
	};
	CSectionPageNumType.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				// case "fmt": {
				// 	this.Fmt = fromXml_ST_NumberFormat(reader.GetValue());
				// 	break;
				// }
				case "start": {
					this.Start =  reader.GetValueInt(this.Start);
					break;
				}
				// case "chapStyle": {
				// 	this.ChapStyle = reader.GetValueInt();
				// 	break;
				// }
				// case "chapSep": {
				// 	this.ChapSep = fromXml_ST_ChapterSep(reader.GetValue());
				// 	break;
				// }
			}
		}
	};
	CSectionPageNumType.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CSectionPageNumType.prototype.toXml = function(writer, name) {
		var Start = -1 !== this.Start ? this.Start : undefined;
		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeString("w:fmt", toXml_ST_NumberFormat(this.Fmt));
		writer.WriteXmlNullableAttributeInt("w:start", Start);
		// writer.WriteXmlNullableAttributeInt("w:chapStyle", this.ChapStyle);
		// writer.WriteXmlNullableAttributeString("w:chapSep", toXml_ST_ChapterSep(this.ChapSep));
		writer.WriteXmlAttributesEnd(true);
	};
	CSectionColumns.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "equalWidth": {
					this.EqualWidth = reader.GetValueBool();
					break;
				}
				case "space": {
					this.Space = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Space);
					break;
				}
				case "num": {
					this.Num = reader.GetValueInt(this.Num);
					break;
				}
				case "sep": {
					this.Sep = reader.GetValueBool();
					break;
				}
			}
		}
	};
	CSectionColumns.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "col" : {
					elem = new CSectionColumn();
					elem.fromXml(reader);
					this.Cols.push(elem);
					break;
				}
			}
		}
	};
	CSectionColumns.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("w:equalWidth", this.EqualWidth);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:space", this.Space, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeInt("w:num", this.Num);
		writer.WriteXmlNullableAttributeBool("w:sep", this.Sep);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlArray(this.Cols, "w:col");
		writer.WriteXmlNodeEnd(name);
	};
	CSectionColumn.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "w": {
					this.W = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.W);
					break;
				}
				case "space": {
					this.Space = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Space);
					break;
				}
			}
		}
	};
	CSectionColumn.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CSectionColumn.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:w", this.W, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:space", this.Space, g_dKoef_mm_to_twips);
		writer.WriteXmlAttributesEnd(true);
	};
	CSectionLnNumType.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "countBy": {
					this.CountBy = reader.GetValueInt(this.CountBy);
					break;
				}
				case "start": {
					this.Start = reader.GetValueInt(this.Start);
					break;
				}
				case "distance": {
					this.Distance = AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, this.Distance);
					break;
				}
				case "restart": {
					this.Restart = fromXml_ST_LineNumberRestart(reader.GetValue(), this.Restart);
					break;
				}
			}
		}
	};
	CSectionLnNumType.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CSectionLnNumType.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:countBy", this.CountBy);
		writer.WriteXmlNullableAttributeInt("w:start", this.Start);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:distance", this.Distance, g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:restart", toXml_ST_LineNumberRestart(this.Restart));
		writer.WriteXmlAttributesEnd(true);
	};
	CFootnotePr.prototype.fromXml = function(reader, isFootnotePr, noteRefs) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "pos" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					if (isFootnotePr) {
						this.Pos = fromXml_ST_FtnPos(reader.GetValue(), this.Pos);
					} else {
						this.Pos = fromXml_ST_EdnPos(reader.GetValue(), this.Pos);
					}
					break;
				}
				case "numFmt" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.NumFormat = fromXml_ST_NumberFormat(reader.GetValue(), this.NumFormat);
					break;
				}
				case "numStart" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.NumStart = elem.getVal(undefined);
					break;
				}
				case "numRestart" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.NumRestart = fromXml_ST_RestartNumber(reader.GetValue(), this.NumRestart);
					break;
				}
				case "footnote" : {
					let ednRef = new CT_FtnEdnRef();
					ednRef.fromXml(reader);
					if(noteRefs) {
						noteRefs.push(ednRef.id);
					}
					break;
				}
				case "endnote" : {
					let ednRef = new CT_FtnEdnRef();
					ednRef.fromXml(reader);
					if(noteRefs) {
						noteRefs.push(ednRef.id);
					}
					break;
				}
			}
		}
	};
	CFootnotePr.prototype.toXml = function(writer, name, isFootnotePr, controller) {
		let elemName = isFootnotePr ? "w:footnote": "w:endnote";
		var Pos;
		if(isFootnotePr) {
			Pos = CT_StringW.prototype.fromVal(toXml_ST_FtnPos(this.Pos));
		} else {
			Pos = CT_StringW.prototype.fromVal(toXml_ST_EdnPos(this.Pos));
		}
		var NumFmt = CT_StringW.prototype.fromVal(toXml_ST_NumberFormat(this.NumFmt));
		var NumStart = CT_IntW.prototype.fromVal(this.NumStart);
		var NumRestart = CT_StringW.prototype.fromVal(toXml_ST_RestartNumber(this.NumRestart));

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(Pos, "w:pos");
		writer.WriteXmlNullable(NumFmt, "w:numFmt");
		writer.WriteXmlNullable(NumStart, "w:numStart");
		writer.WriteXmlNullable(NumRestart, "w:numRestart");
		if (controller) {
			let index = -1;
			if (controller.Separator) {
				writer.WriteXmlString('<' + elemName + ' w:id="' + index + '"/>');
				index++;
			}
			if (controller.ContinuationSeparator) {
				writer.WriteXmlString('<' + elemName + ' w:id="' + index + '"/>');
				index++;
			}
			if (controller.ContinuationNotice) {
				writer.WriteXmlString('<' + elemName + ' w:id="' + index + '"/>');
				index++;
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	if(typeof CParagraphBookmark !== "undefined") {
		CParagraphBookmark.prototype.readAttr = function(reader) {
			while (reader.MoveToNextAttribute()) {
				switch (reader.GetNameNoNS()) {
					// case "colFirst": {
					// 	this.colFirst = reader.GetValueInt(this.colFirst);
					// 	break;
					// }
					// case "colLast": {
					// 	this.colLast = reader.GetValueInt(this.colLast);
					// 	break;
					// }
					case "name": {
						this.BookmarkName = reader.GetValueDecodeXml();
						break;
					}
					case "id": {
						this.BookmarkId = reader.GetValueInt(this.id);
						break;
					}
				}
			}
		};
		CParagraphBookmark.prototype.fromXml = function(reader) {
			this.readAttr(reader);
			reader.ReadTillEnd();
		};
		CParagraphBookmark.prototype.toXml = function(writer, name) {
			writer.WriteXmlNodeStart(name);
			// writer.WriteXmlNullableAttributeInt("w:colFirst", this.colFirst);
			// writer.WriteXmlNullableAttributeInt("w:colLast", this.colLast);
			writer.WriteXmlNullableAttributeInt("w:id", this.BookmarkId);
			writer.WriteXmlNullableAttributeStringEncode("w:name", this.BookmarkName);
			writer.WriteXmlAttributesEnd(true);
		};
	}
//styles
	CStyles.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("styles" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("styles" === name) {
			var elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "docDefaults" : {
						var DocDefaults = new CT_DocDefaults();
						DocDefaults.fromXml(reader);
						reader.context.oReadResult.DefpPr = DocDefaults.PPrDefault;
						reader.context.oReadResult.DefrPr = DocDefaults.RPrDefault;
						break;
					}
					// case "latentStyles" : {
					// 	this.LatentStyles = new CT_LatentStyles();
					// 	this.LatentStyles.fromXml(reader);
					// 	break;
					// }
					case "style" : {
						var oNewId = {id: null, def: false};
						elem = new CStyle(null, null, null, null);
						elem.fromXml(reader, oNewId);
						reader.context.oReadResult.styles[oNewId.id] = {style: elem, param: oNewId};
						break;
					}
				}
			}
		}
	};
	CStyles.prototype.toXml = function(writer) {
		var DocDefaults = new CT_DocDefaults();
		DocDefaults.RPrDefault = this.Default.TextPr;
		DocDefaults.PPrDefault = this.Default.ParaPr;

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart("w:styles");
		writer.WriteXmlString(AscCommonWord.g_sXmlStylesNamespaces);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(DocDefaults, "w:docDefaults");
		writer.WriteXmlString(AscCommonWord.g_sXmlStylesLatentStyles);
		for (var id in this.Style) {
			if (this.Style.hasOwnProperty(id)) {
				var style = this.Style[id];
				var addition = {id: id, def: this.Is_StyleDefaultOOXML(style.Name)};
				style.toXml(writer, "w:style", addition);
			}
		}
		this.Style.forEach(function(elem) {

		});
		writer.WriteXmlNodeEnd("w:styles");
	};
	function CT_DocDefaults() {
		this.RPrDefault = null;
		this.PPrDefault = null;
		return this;
	}

	CT_DocDefaults.prototype.fromXml = function(reader) {
		var elem, subDepth, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "rPrDefault" : {
					this.RPrDefault = new CTextPr();
					subDepth = reader.GetDepth();
					while (reader.ReadNextSiblingNode(subDepth)) {
						if ("rPr" === reader.GetNameNoNS()) {
							this.RPrDefault.fromXml(reader);
						}
					}
					break;
				}
				case "pPrDefault" : {
					this.PPrDefault = new CParaPr();
					subDepth = reader.GetDepth();
					while (reader.ReadNextSiblingNode(subDepth)) {
						if ("pPr" === reader.GetNameNoNS()) {
							this.PPrDefault.fromXml(reader);
						}
					}
					break;
				}
			}
		}
	};
	CT_DocDefaults.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNodeStart("w:rPrDefault");
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.RPrDefault, "w:rPr");
		writer.WriteXmlNodeEnd("w:rPrDefault");
		writer.WriteXmlNodeStart("w:pPrDefault");
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.PPrDefault, "w:pPr");
		writer.WriteXmlNodeEnd("w:pPrDefault");
		writer.WriteXmlNodeEnd(name);
	};
	CStyle.prototype.readAttr = function(reader, opt_addition) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					this.Set_Type(fromXml_ST_StyleType(reader.GetValue()) || styletype_Paragraph);
					break;
				}
				case "styleId": {
					if (opt_addition) {
						opt_addition.id = reader.GetValueDecodeXml();
					}
					break;
				}
				case "default": {
					if (opt_addition) {
						opt_addition.def = reader.GetValueBool();
					}
					break;
				}
				case "customStyle": {
					this.SetCustom(reader.GetValueBool());
					break;
				}
			}
		}
	};
	CStyle.prototype.fromXml = function(reader, opt_addition) {
		this.readAttr(reader, opt_addition);
		var elem, name, depth = reader.GetDepth(), oReadResult = reader.context.oReadResult;
		let t = this;
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "name" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					name = elem.getVal(name);
					break;
				}
				// case "aliases" : {
				// 	this.Aliases = new CT_StringW();
				// 	this.Aliases.fromXml(reader);
				// 	break;
				// }
				case "basedOn" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Set_BasedOn(elem.getVal(undefined));
					break;
				}
				case "next" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Set_Next(elem.getVal(undefined));
					break;
				}
				case "link" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Set_Link(elem.getVal(undefined));
					break;
				}
				// case "autoRedefine" : {
				// 	this.AutoRedefine = new CT_BoolW();
				// 	this.AutoRedefine.fromXml(reader);
				// 	break;
				// }
				case "hidden" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.Set_QFormat(elem.getVal(undefined));
					break;
				}
				case "uiPriority" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.Set_UiPriority(elem.getVal(undefined));
					break;
				}
				case "semiHidden" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.Set_SemiHidden(elem.getVal(undefined));
					break;
				}
				case "unhideWhenUsed" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.Set_UnhideWhenUsed(elem.getVal(undefined));
					break;
				}
				case "qFormat" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.Set_QFormat(elem.getVal(undefined));
					break;
				}
				// case "locked" : {
				// 	this.Locked = new CT_BoolW();
				// 	this.Locked.fromXml(reader);
				// 	break;
				// }
				// case "personal" : {
				// 	this.Personal = new CT_BoolW();
				// 	this.Personal.fromXml(reader);
				// 	break;
				// }
				// case "personalCompose" : {
				// 	this.PersonalCompose = new CT_BoolW();
				// 	this.PersonalCompose.fromXml(reader);
				// 	break;
				// }
				// case "personalReply" : {
				// 	this.PersonalReply = new CT_BoolW();
				// 	this.PersonalReply.fromXml(reader);
				// 	break;
				// }
				// case "rsid" : {
				// 	this.Rsid = new CT_LongHexNumber();
				// 	this.Rsid.fromXml(reader);
				// 	break;
				// }
				case "pPr" : {
					this.ParaPr = new CParaPr();
					this.ParaPr.fromXml(reader);
					break;
				}
				case "rPr" : {
					this.TextPr = new CTextPr();
					this.TextPr.fromXml(reader);
					break;
				}
				case "tblPr" : {
					this.TablePr = new CTablePr();
					this.TablePr.fromXml(reader);
					break;
				}
				case "trPr" : {
					this.TableRowPr = new CTableRowPr();
					this.TableRowPr.fromXml(reader);
					break;
				}
				case "tcPr" : {
					this.TableCellPr = new CTableCellPr();
					this.TableCellPr.fromXml(reader);
					break;
				}
				case "tblStylePr" : {
					let opt_addition = {type: null};
					elem = new CTableStylePr();
					elem.fromXml(reader, opt_addition);
					fromXml_ST_TblStyleOverrideType(this, elem, opt_addition.type);
					break;
				}
			}
		}
		if (!name) {
			name = "StGen" + oReadResult.styleGenIndex++;
		}
		this.Set_Name(name);
	};
	CStyle.prototype.toXml = function(writer, name, opt_addition) {
		var Name = CT_StringW.prototype.fromVal(this.Name);
		var BasedOn = CT_StringW.prototype.fromVal(this.BasedOn);
		var Next = CT_StringW.prototype.fromVal(this.Next);
		var Link = CT_StringW.prototype.fromVal(this.Link);
		var hidden = CT_BoolW.prototype.fromVal(this.hidden);
		var uiPriority = CT_IntW.prototype.fromVal(this.uiPriority);
		var semiHidden = CT_BoolW.prototype.fromVal(this.semiHidden);
		var unhideWhenUsed = CT_BoolW.prototype.fromVal(this.unhideWhenUsed);
		var qFormat = CT_BoolW.prototype.fromVal(this.qFormat);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:type", toXml_ST_StyleType(this.Type));
		writer.WriteXmlNullableAttributeStringEncode("w:styleId", opt_addition && opt_addition.id);
		writer.WriteXmlAttributeBoolIfTrue("w:default", opt_addition && opt_addition.def);
		writer.WriteXmlAttributeBoolIfTrue("w:customStyle", this.IsCustom());
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(Name, "w:name");
		// writer.WriteXmlNullable(this.Aliases, "w:aliases");
		writer.WriteXmlNullable(BasedOn, "w:basedOn");
		writer.WriteXmlNullable(Next, "w:next");
		writer.WriteXmlNullable(Link, "w:link");
		// writer.WriteXmlNullable(this.AutoRedefine, "w:autoRedefine");
		writer.WriteXmlNullable(hidden, "w:hidden");
		writer.WriteXmlNullable(uiPriority, "w:uiPriority");
		writer.WriteXmlNullable(semiHidden, "w:semiHidden");
		writer.WriteXmlNullable(unhideWhenUsed, "w:unhideWhenUsed");
		writer.WriteXmlNullable(qFormat, "w:qFormat");
		// writer.WriteXmlNullable(this.Locked, "w:locked");
		// writer.WriteXmlNullable(this.Personal, "w:personal");
		// writer.WriteXmlNullable(this.PersonalCompose, "w:personalCompose");
		// writer.WriteXmlNullable(this.PersonalReply, "w:personalReply");
		// writer.WriteXmlNullable(this.Rsid, "w:rsid");
		writer.WriteXmlNullable(this.ParaPr, "w:pPr");
		writer.WriteXmlNullable(this.TextPr, "w:rPr");
		if (styletype_Table == this.Type) {
			writer.WriteXmlNullable(this.TablePr, "w:tblPr");
			writer.WriteXmlNullable(this.TableRowPr, "w:trPr");
			writer.WriteXmlNullable(this.TableCellPr, "w:tcPr");
			toXml_ST_TblStyleOverrideType(writer, this);
		}
		writer.WriteXmlNodeEnd(name);
	};
	CTableStylePr.prototype.readAttr = function(reader, opt_addition) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					opt_addition.type = reader.GetValue();
					break;
				}
			}
		}
	};
	CTableStylePr.prototype.fromXml = function(reader, opt_addition) {
		this.readAttr(reader, opt_addition);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "pPr" : {
					this.ParaPr = new CParaPr();
					this.ParaPr.fromXml(reader);
					break;
				}
				case "rPr" : {
					this.TextPr = new CTextPr();
					this.TextPr.fromXml(reader);
					break;
				}
				case "tblPr" : {
					this.TablePr = new CTablePr();
					this.TablePr.fromXml(reader);
					break;
				}
				case "trPr" : {
					this.TableRowPr = new CTableRowPr();
					this.TableRowPr.fromXml(reader);
					break;
				}
				case "tcPr" : {
					this.TableCellPr = new CTableCellPr();
					this.TableCellPr.fromXml(reader);
					break;
				}
			}
		}
	};
	CTableStylePr.prototype.toXml = function(writer, name, type) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:type", type);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.ParaPr, "w:pPr");
		writer.WriteXmlNullable(this.TextPr, "w:rPr");
		writer.WriteXmlNullable(this.TablePr, "w:tblPr");
		writer.WriteXmlNullable(this.TableRowPr, "w:trPr");
		writer.WriteXmlNullable(this.TableCellPr, "w:tcPr");
		writer.WriteXmlNodeEnd(name);
	};
//numbering
	CNumbering.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("numbering" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		var aNumsMap = {}, additional, oReadResult = reader.context.oReadResult;
		if ("numbering" === name) {
			var elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					// case "numPicBullet" : {
					// 	elem = new CT_NumPicBullet();
					// 	elem.fromXml(reader);
					// 	this.NumPicBullet.push(elem);
					// 	break;
					// }
					case "abstractNum" : {
						additional = {aNumId: null};
						elem = new CAbstractNum();
						elem.fromXml(reader, additional);
						if (null !== additional.aNumId) {
							aNumsMap[additional.aNumId] = elem;
						}
						break;
					}
					case "num" : {
						additional = {numId: null, aNumId: null};
						elem = new CNum(this);
						elem.fromXml(reader, additional);
						var ANum = aNumsMap[additional.aNumId];
						if (ANum) {
							elem.SetAbstractNumId(ANum.GetId());
							oReadResult.logicDocument.Numbering.AddAbstractNum(ANum);
						}
						if (null !== additional.numId) {
							oReadResult.logicDocument.Numbering.AddNum(elem);
							oReadResult.numToNumClass[additional.numId] = elem;
						}
						break;
					}
					// case "numIdMacAtCleanup" : {
					// 	this.NumIdMacAtCleanup = new CT_IntW();
					// 	this.NumIdMacAtCleanup.fromXml(reader);
					// 	break;
					// }
				}
			}
		}
	};
	CNumbering.prototype.toXml = function(writer) {
		var i, aNumsMap = {}, aNumsIndex = 0, numsIndex = 1, name = "w:numbering";
		var oNumIdMap = writer.context.oNumIdMap;
		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlNumberingNamespaces);
		writer.WriteXmlAttributesEnd();

		// writer.WriteXmlArray(this.NumPicBullet, "w:numPicBullet");
		for (i in this.AbstractNum) {
			if (this.AbstractNum.hasOwnProperty(i)) {
				var aNum = this.AbstractNum[i];
				aNum.toXml(writer, "w:abstractNum", aNumsIndex);
				aNumsMap[aNum.GetId()] = aNumsIndex++;
			}
		}
		for (i in this.Num) {
			if (this.Num.hasOwnProperty(i)) {
				var Num = this.Num[i];
				var aNumId = aNumsMap[Num.AbstractNumId];
				Num.toXml(writer, "w:num", numsIndex, aNumId);
				oNumIdMap[Num.GetId()] = numsIndex++;
			}
		}
		// writer.WriteXmlNullable(this.NumIdMacAtCleanup, "w:numIdMacAtCleanup");
		writer.WriteXmlNodeEnd(name);
	};
	CAbstractNum.prototype.readAttr = function(reader, additional) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "abstractNumId":
				case "listDefId": {
					additional.aNumId = reader.GetValueInt();
					break;
				}
			}
		}
	};
	CAbstractNum.prototype.fromXml = function(reader, additional) {
		var oReadResult = reader.context.oReadResult;
		this.readAttr(reader, additional);
		let elem, index = 0, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			//todo alter
			switch (reader.GetNameNoNS()) {
				// case "nsid" : {
				// 	this.Nsid = new CT_LongHexNumber();
				// 	this.Nsid.fromXml(reader);
				// 	break;
				// }
				// case "multiLevelType" :
				// case "plt" : {
				// 	this.MultiLevelType = new CT_MultiLevelType();
				// 	this.MultiLevelType.fromXml(reader);
				// 	break;
				// }
				// case "tmpl" : {
				// 	this.Tmpl = new CT_LongHexNumber();
				// 	this.Tmpl.fromXml(reader);
				// 	break;
				// }
				// case "name" : {
				// 	this.Name = new CT_StringW();
				// 	this.Name.fromXml(reader);
				// 	break;
				// }
				case "styleLink" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					if (elem.val) {
						oReadResult.styleLinks.push({pPr: this, style: elem.val});
					}
					break;
				}
				case "numStyleLink" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					if (elem.val) {
						oReadResult.numStyleLinks.push({pPr: this, style: elem.val});
					}
					break;
				}
				case "lvl" : {
					if (index < this.Lvl.length) {
						let additionalLvl = {ilvl: index};
						elem = this.Lvl[index].Copy();
						//сбрасываем свойства
						elem.ParaPr = new CParaPr();
						elem.TextPr = new CTextPr();
						elem.fromXml(reader, additionalLvl);
						index = additionalLvl.ilvl || index;
						this.Lvl[index] = elem;
						index++;
					}
					break;
				}
			}
		}
	};
	CAbstractNum.prototype.toXml = function(writer, name, abstractNumId) {
		var MultiLevelType = CT_StringW.prototype.fromVal("hybridMultilevel");
		var StyleLink = CT_StringW.prototype.fromVal(this.StyleLink);
		var NumStyleLink = CT_StringW.prototype.fromVal(this.NumStyleLink);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:abstractNumId", abstractNumId);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.Nsid, "w:nsid");
		writer.WriteXmlNullable(MultiLevelType, "w:multiLevelType");
		// writer.WriteXmlNullable(this.Tmpl, "w:tmpl");
		// writer.WriteXmlNullable(this.Name, "w:name");
		writer.WriteXmlNullable(StyleLink, "w:styleLink");
		writer.WriteXmlNullable(NumStyleLink, "w:numStyleLink");
		for (var i = 0; i < this.Lvl.length; ++i) {
			this.Lvl[i].toXml(writer, "w:lvl", i);
		}
		writer.WriteXmlNodeEnd(name);
	};
	CNumberingLvl.prototype.readAttr = function(reader, additional) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "ilvl": {
					additional.ilvl = reader.GetValueInt();
					break;
				}
				// case "tplc": {
				// 	this.Tplc = reader.GetValueByte();
				// 	break;
				// }
				// case "tentative": {
				// 	this.Tentative = reader.GetValueBool();
				// 	break;
				// }
			}
		}
	};
	CNumberingLvl.prototype.fromXml = function(reader, additional) {
		var oReadResult = reader.context.oReadResult;
		this.readAttr(reader, additional);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "start" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.Start = elem.getVal(this.Start);
					break;
				}
				case "numFmt" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					this.SetFormat(fromXml_ST_NumberFormat(elem.attributes["val"], this.Format));
					break;
				}
				case "lvlRestart" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.LvlRestart = elem.getVal(this.LvlRestart);
					break;
				}
				case "pStyle" : {
					elem = CT_StringW.prototype.toVal(reader);
					if (elem) {
						oReadResult.lvlStyles.push({pPr: this, style: elem});
					}
					break;
				}
				case "isLgl" : {
					elem = new CT_BoolW();
					elem.fromXml(reader);
					this.IsLgl = elem.getVal(this.IsLgl);
					break;
				}
				case "legacy" : {
					this.Legacy = new CNumberingLvlLegacy();
					this.Legacy.fromXml(reader);
					break;
				}
				case "suff" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Suff = fromXml_ST_LevelSuffix(elem.getVal(), this.Suff);
					break;
				}
				case "lvlText" : {
					this.SetLvlTextFormat(additional.ilvl, CT_StringW.prototype.toVal(reader, ""));
					break;
				}
				// case "lvlPicBulletId" : {
				// 	this.LvlPicBulletId = new CT_IntW();
				// 	this.LvlPicBulletId.fromXml(reader);
				// 	break;
				// }
				case "lvlJc" : {
					elem = new CT_StringW();
					elem.fromXml(reader);
					this.Jc = fromXml_ST_Jc1(elem.getVal(), this.Jc);
					break;
				}
				case "pPr" : {
					this.ParaPr.fromXml(reader);
					break;
				}
				case "rPr" : {
					this.TextPr.fromXml(reader);
					break;
				}
			}
		}
	};
	CNumberingLvl.prototype.toXml = function(writer, name, ilvl) {
		var Start = CT_IntW.prototype.fromVal(this.Start);
		var Format = CT_StringW.prototype.fromVal(toXml_ST_NumberFormat(this.Format));
		var LvlRestart = CT_IntW.prototype.fromVal(this.LvlRestart);
		var PStyle = CT_StringW.prototype.fromVal(this.PStyle);
		var IsLgl = CT_BoolW.prototype.fromVal(this.IsLgl);
		var Suff = CT_StringW.prototype.fromVal(toXml_ST_LevelSuffix(this.Suff));
		var LvlText = CT_StringW.prototype.fromVal(this.GetLvlTextFormat());
		var Jc = CT_StringW.prototype.fromVal(toXml_ST_Jc1(this.Jc));

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:ilvl", ilvl);
		// writer.WriteXmlNullableAttributeByte("w:tplc", this.Tplc);
		// writer.WriteXmlNullableAttributeBool("w:tentative", this.Tentative);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(Start, "w:start");
		writer.WriteXmlNullable(Format, "w:numFmt");
		writer.WriteXmlNullable(LvlRestart, "w:lvlRestart");
		writer.WriteXmlNullable(PStyle, "w:pStyle");
		writer.WriteXmlNullable(IsLgl, "w:isLgl");
		writer.WriteXmlNullable(this.Legacy, "w:legacy");
		writer.WriteXmlNullable(Suff, "w:suff");
		writer.WriteXmlNullable(LvlText, "w:lvlText");
		// writer.WriteXmlNullable(this.LvlPicBulletId, "w:lvlPicBulletId");
		writer.WriteXmlNullable(Jc, "w:lvlJc");
		writer.WriteXmlNullable(this.ParaPr, "w:pPr");
		writer.WriteXmlNullable(this.TextPr, "w:rPr");
		writer.WriteXmlNodeEnd(name);
	};
	CNumberingLvlLegacy.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "legacy": {
					this.Legacy = reader.GetValueBool();
					break;
				}
				case "legacyIndent": {
					this.Indent = AscCommon.universalMeasureToTwips(reader.GetValue(), 1, this.Indent);
					break;
				}
				case "legacySpace": {
					this.Space = AscCommon.universalMeasureToUnsignedTwips(reader.GetValue(), 1, this.Space);
					break;
				}
			}
		}
	};
	CNumberingLvlLegacy.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CNumberingLvlLegacy.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("w:legacy", this.Legacy);
		writer.WriteXmlNullableAttributeUInt("w:legacyIndent", this.Indent);
		writer.WriteXmlNullableAttributeUInt("w:legacySpace", this.Space);
		writer.WriteXmlAttributesEnd(true);
	};
	CNum.prototype.readAttr = function(reader, additional) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "numId": {
					additional.numId = reader.GetValueInt();
					break;
				}
			}
		}
	};
	CNum.prototype.fromXml = function(reader, additional) {
		this.readAttr(reader, additional);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "abstractNumId" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					additional.aNumId = elem.getVal(additional.aNumId);
					break;
				}
				case "lvlOverride" : {
					elem = new CLvlOverride();
					elem.StartOverride = undefined;
					elem.fromXml(reader);
					this.SetLvlOverride(elem.NumberingLvl, elem.Lvl, elem.StartOverride);
					break;
				}
			}
		}
	};
	CNum.prototype.toXml = function(writer, name, numId, aNumId) {
		var AbstractNumId = CT_IntW.prototype.fromVal(aNumId);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:numId", numId);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(AbstractNumId, "w:abstractNumId");
		for (var i = 0; i < this.LvlOverride.length; ++i) {
			if (this.LvlOverride[i]) {
				this.LvlOverride[i].toXml(writer, "w:lvlOverride", i);
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	CLvlOverride.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "ilvl": {
					this.Lvl = reader.GetValueInt();
					break;
				}
			}
		}
	};
	CLvlOverride.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "startOverride" : {
					elem = new CT_IntW();
					elem.fromXml(reader);
					this.StartOverride = elem.getVal(this.StartOverride);
					break;
				}
				case "lvl" : {
					let additionalLvl = {ilvl: 0};
					this.NumberingLvl = new CNumberingLvl();
					this.NumberingLvl.fromXml(reader, additionalLvl);
					break;
				}
			}
		}
	};
	CLvlOverride.prototype.toXml = function(writer, name) {
		var StartOverride = CT_IntW.prototype.fromVal(this.StartOverride);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:ilvl", this.Lvl);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(StartOverride, "w:startOverride");
		writer.WriteXmlNullable(this.NumberingLvl, "w:lvl");
		writer.WriteXmlNodeEnd(name);
	};
//Header/Footer
	CHeaderFooter.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("hdr" !== name && "ftr" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("hdr" === name || "ftr" === name) {
			var Content = [];
			CDocument.prototype.fromXmlDocContent(reader, Content, this.DrawingDocument, this.Content);
			if (Content.length > 0) {
				this.Content.ReplaceContent(Content);
			}
		}
	};
	CHeaderFooter.prototype.toXml = function(writer) {
		var name = this.Type === AscCommon.hdrftr_Header ? "w:hdr" : "w:ftr";

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlHeaderFooterNamespaces);
		writer.WriteXmlAttributesEnd();
		this.Content.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		writer.WriteXmlNodeEnd(name);
	};
//footnote/endnote
	CFootnotesController.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("footnotes" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("footnotes" === name) {
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "footnote" : {
						elem = new CFootEndnote(this);
						elem.fromXml(reader, reader.context.oReadResult.footnotes);
						break;
					}
				}
			}
		}
	};
	CFootnotesController.prototype.toXml = function(writer) {
		let name = "w:footnotes";

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlFootnotesEndnotesNamespaces);
		writer.WriteXmlAttributesEnd();
		let index = -1;
		if (this.Separator) {
			this.Separator.toXml(writer, "w:footnote", index++, 3);
		}
		if (this.ContinuationSeparator) {
			this.ContinuationSeparator.toXml(writer, "w:footnote", index++, 1);
		}
		if (this.ContinuationNotice) {
			this.ContinuationSeparator.toXml(writer, "w:footnote", index++, 0);
		}
		for(let id in this.Footnote) {
			if(this.Footnote.hasOwnProperty(id)) {
				writer.context.docSaveParams.footnoteIdToIndex[id] = index;
				this.Footnote[id].toXml(writer, "w:footnote", index);
				index++;
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	CEndnotesController.prototype.fromXml = function(reader) {
		var name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("endnotes" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("endnotes" === name) {
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "endnote" : {
						elem = new CFootEndnote(this);
						elem.fromXml(reader, reader.context.oReadResult.endnotes);
						break;
					}
				}
			}
		}
	};
	CEndnotesController.prototype.toXml = function(writer) {
		let name = "w:endnotes";

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlFootnotesEndnotesNamespaces);
		writer.WriteXmlAttributesEnd();
		let index = -1;
		if (this.Separator) {
			this.Separator.toXml(writer, "w:endnote", index++, 3);
		}
		if (this.ContinuationSeparator) {
			this.ContinuationSeparator.toXml(writer, "w:endnote", index++, 1);
		}
		if (this.ContinuationNotice) {
			this.ContinuationSeparator.toXml(writer, "w:endnote", index++, 0);
		}
		for(let id in this.Endnote) {
			if(this.Endnote.hasOwnProperty(id)) {
				writer.context.docSaveParams.endnoteIdToIndex[id] = index;
				this.Endnote[id].toXml(writer, "w:endnote", index);
				index++;
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	CFootEndnote.prototype.readAttr = function(reader, note) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					note.type = fromXml_ST_FtnEdn(reader.GetValue(), note.type);
					break;
				}
				case "id": {
					note.id = reader.GetValueInt(note.id);
					break;
				}
			}
		}
	};
	CFootEndnote.prototype.fromXml = function(reader, notes) {
		let note = {id: null, type: null, content: this};
		this.readAttr(reader, note);
		var Content = [];
		CDocument.prototype.fromXmlDocContent(reader, Content, this.DrawingDocument, this);
		if (Content.length > 0) {
			this.ReplaceContent(Content);
		}
		if (null !== note.id && null !== note.content) {
			notes[note.id] = note;
		}
	};
	CFootEndnote.prototype.toXml = function(writer, name, id, type) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:type", toXml_ST_FtnEdn(type));
		writer.WriteXmlNullableAttributeInt("w:id", id);
		writer.WriteXmlAttributesEnd();
		this.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		writer.WriteXmlNodeEnd(name);
	};
//drawing
	ParaDrawing.prototype.fromXml = function(reader) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "anchor" : {
					this.Set_DrawingType(drawing_Anchor);
					elem = new CT_Anchor(this);
					elem.fromXml(reader);
					break;
				}
				case "inline" : {
					this.Set_DrawingType(drawing_Inline);
					elem = new CT_Inline(this);
					elem.fromXml(reader);
					break;
				}
			}
		}
	};
	ParaDrawing.prototype.GetVmlMainProps = function() {
		let sMainCSS = "";
		let sMainNodes = "";
		let sMainAttributes = "";

		let oParaDrawing = this;
		let dKoefMMToPT = 72.0 / 25.4;
		let oDistance = oParaDrawing.Distance;
		let oExtent = oParaDrawing.Extent;
		let fAddDistanceToCSS = function () {
			if (oDistance.L !== null)
				sMainCSS += ("mso-wrap-distance-left:" + (dKoefMMToPT * oDistance.L).toFixed(2) + "pt;");
			if (oDistance.T !== null)
				sMainCSS += ("mso-wrap-distance-top:" + (dKoefMMToPT * oDistance.T).toFixed(2) + "pt;");
			if (oDistance.R !== null)
				sMainCSS += ("mso-wrap-distance-right:" + (dKoefMMToPT * oDistance.R).toFixed(2) + "pt;");
			if (oDistance.B !== null)
				sMainCSS += ("mso-wrap-distance-bottom:" + (dKoefMMToPT * oDistance.B).toFixed(2) + "pt;");
		};
		let fAddExtentToCSS = function() {
			if (oExtent) {
				sMainCSS += ("width:" + (dKoefMMToPT * oExtent.W).toFixed(2) + "pt;");
				sMainCSS += ("height:" + (dKoefMMToPT * oExtent.H).toFixed(2) + "pt;");
			}
		};
		if (oParaDrawing.IsInline())
		{
			fAddDistanceToCSS();
			fAddExtentToCSS();
		}
		else
		{
			sMainCSS += ("position:absolute;");
			fAddDistanceToCSS();

			if (oParaDrawing.RelativeHeight !== null)
			{
				let z_index = oParaDrawing.RelativeHeight;

				if (oParaDrawing.behindDoc)
				{
					z_index = -z_index;
				}
				sMainCSS += ("z-index:" + z_index + ";");
			}

			if (oParaDrawing.AllowOverlap !== null) {
				sMainCSS += ("o:allowoverlap:" + (oParaDrawing.AllowOverlap  ? "true;" : "false;"));
			}
			if (oParaDrawing.LayoutInCell !== null) {
				sMainCSS += ("o:allowincell:" + (oParaDrawing.LayoutInCell  ? "true;" : "false;"));
			}

			let oPositionH = oParaDrawing.PositionH;
			if (oPositionH) {
				if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.Character)
					sMainCSS += ("mso-position-horizontal-relative:char;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.Page)
					sMainCSS += ("mso-position-horizontal-relative:page;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.Margin)
					sMainCSS += ("mso-position-horizontal-relative:margin;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.LeftMargin)
					sMainCSS += ("mso-position-horizontal-relative:left-margin-area;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.RightMargin)
					sMainCSS += ("mso-position-horizontal-relative:right-margin-area;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.InsideMargin)
					sMainCSS += ("mso-position-horizontal-relative:inner-margin-area;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.OutsideMargin)
					sMainCSS += ("mso-position-horizontal-relative:outer-margin-area;");
				else if (oPositionH.RelativeFrom === Asc.c_oAscRelativeFromH.Column)
					sMainCSS += ("mso-position-horizontal-relative:text;");

				if (!oPositionH.Align) {
					sMainCSS += ("margin-left:" + (dKoefMMToPT * oPositionH.Value).toFixed(2) + "pt;");
					sMainCSS += ("mso-position-horizontal:absolute;");
				}
				else {
					switch (oPositionH.Value) {
						case Asc.c_oAscAlignH.Center:
						{
							sMainCSS += ("mso-position-horizontal:center;");
							break;
						}
						case Asc.c_oAscAlignH.Inside:
						{
							sMainCSS += ("mso-position-horizontal:inside;");
							break;
						}
						case Asc.c_oAscAlignH.Outside:
						{
							sMainCSS += ("mso-position-horizontal:outside;");
							break;
						}
						case Asc.c_oAscAlignH.Left:
						{
							sMainCSS += ("mso-position-horizontal:left;");
							break;
						}
						case Asc.c_oAscAlignH.Right:
						{
							sMainCSS += ("mso-position-horizontal:right;");
							break;
						}
					}
				}
			}

			let oPositionV = oParaDrawing.PositionV;
			if (oPositionV)
			{
				if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.Margin)
					sMainCSS += ("mso-position-vertical-relative:margin;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.Paragraph)
					sMainCSS += ("mso-position-vertical-relative:text;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.Page)
					sMainCSS += ("mso-position-vertical-relative:page;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.TopMargin)
					sMainCSS += ("mso-position-vertical-relative:top-margin-area;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.BottomMargin)
					sMainCSS += ("mso-position-vertical-relative:bottom-margin-area;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.InsideMargin)
					sMainCSS += ("mso-position-vertical-relative:inner-margin-area;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.OutsideMargin)
					sMainCSS += ("mso-position-vertical-relative:outer-margin-area;");
				else if (oPositionV.RelativeFrom === Asc.c_oAscRelativeFromV.Line)
					sMainCSS += ("mso-position-vertical-relative:line;");

				if (!oPositionV.Align)
				{
					sMainCSS += ("margin-top:" + (dKoefMMToPT * oPositionV.Value).toFixed(2) + "pt;");
					sMainCSS += ("mso-position-vertical:absolute");
				}
				else
				{

					switch (oPositionV.Value)
					{
						case c_oAscAlignV.Bottom:
						{
							sMainCSS += ("mso-position-vertical:bottom;");
							break;
						}
						case c_oAscAlignV.Outside:
						{
							sMainCSS += ("mso-position-vertical:outside;");
							break;
						}
						case c_oAscAlignV.Center:
						{
							sMainCSS += ("mso-position-vertical:center;");
							break;
						}

						case c_oAscAlignV.Inside:
						{
							sMainCSS += ("mso-position-vertical:inside;");
							break;
						}
						case c_oAscAlignV.Top:
						{
							sMainCSS += ("mso-position-vertical:top;");
							break;
						}
					}
				}
			}
			fAddExtentToCSS();
			if(oParaDrawing.wrappingType === WRAPPING_TYPE_NONE) {

			}
			else if(oParaDrawing.wrappingType === WRAPPING_TYPE_SQUARE) {

				sMainNodes += "<w10:wrap type=\"square\"/>";
			}
			else if(oParaDrawing.wrappingType === WRAPPING_TYPE_TOP_AND_BOTTOM) {
				sMainNodes += "<w10:wrap type=\"topAndBottom\"/>";
			}
			else if(oParaDrawing.wrappingType === WRAPPING_TYPE_TIGHT) {
				sMainNodes += "<w10:wrap type=\"tight\"/>";
			}
			else if(oParaDrawing.wrappingType === WRAPPING_TYPE_THROUGH) {
				sMainNodes += "<w10:wrap type=\"through\"/>";
			}
			if(oParaDrawing.wrappingType === WRAPPING_TYPE_TIGHT ||
				oParaDrawing.wrappingType === WRAPPING_TYPE_THROUGH) {
				let oWrapPolygon = oParaDrawing.wrappingPolygon;
				let dWrapKoef = 100000.0 / 21600.0;
				if(oWrapPolygon) {
					sMainAttributes += " wrapcoords=\"";
					let aPoints = oWrapPolygon.relativeArrPoints;
					let nCountP = aPoints.length;
					for (let i = 0; i < nCountP; ++i)  {
						let oPoint = aPoints[i];
						let nX = (dWrapKoef * oPoint.x + 0.5) >> 0;
						let nY = (dWrapKoef * oPoint.y + 0.5) >> 0;

						sMainAttributes += (nX + " " + nY);

						if (i < (nCountP - 1))
							sMainAttributes += " ";
					}
					sMainAttributes += "\"";
				}
			}
		}
		return {sMainCSS: sMainCSS, sMainNodes: sMainNodes, sMainAttributes: sMainAttributes};
	};
	ParaDrawing.prototype.toXml = function(writer, name) {
		let oGraphic = this.GraphicObj;
		if(!oGraphic) {
			return;
		}
		if(oGraphic.isOleObject() || oGraphic.isSignatureLine()) {
			oGraphic.toXml(writer);
			return;
		}
		if(drawing_Inline === this.DrawingType) {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlAttributesEnd();
			var anchor = new CT_Inline(this);
			anchor.toXml(writer, "wp:inline");
			writer.WriteXmlNodeEnd(name);
		} else {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlAttributesEnd();
			var anchor = new CT_Anchor(this);
			anchor.toXml(writer, "wp:anchor");
			writer.WriteXmlNodeEnd(name);
		}
	};
	if(typeof CWrapPolygon !== "undefined") {

		CWrapPolygon.prototype.readAttr = function(reader) {
			while (reader.MoveToNextAttribute()) {
				switch (reader.GetNameNoNS()) {
					case "edited": {
						this.setEdited(reader.GetValueBool());
						break;
					}
				}
			}
		};
		CWrapPolygon.prototype.fromXml = function(reader) {
			this.readAttr(reader);
			var elem, depth = reader.GetDepth();
			var points = [];
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "start" : {
						elem = new CT_XmlNode();
						elem.fromXml(reader);
						points.unshift({x: parseInt(elem.attributes["x"]), y: parseInt(elem.attributes["y"])});
						break;
					}
					case "lineTo" : {
						elem = new CT_XmlNode();
						elem.fromXml(reader);
						points.push({x: parseInt(elem.attributes["x"]), y: parseInt(elem.attributes["y"])});
						break;
					}
				}
			}
			this.setArrRelPoints(points);
		};
		CWrapPolygon.prototype.toXml = function(writer, name) {
			var elem;
			writer.WriteXmlNodeStart(name);
			//всегда пишем Edited == true потому что наш контур отличается от word.
			writer.WriteXmlNullableAttributeBool("edited", true);
			writer.WriteXmlAttributesEnd();
			if (this.relativeArrPoints.length > 0) {
				elem = new CT_XmlNode();
				elem.attributes["x"] = this.relativeArrPoints[0].x;
				elem.attributes["y"] = this.relativeArrPoints[0].y;
				writer.WriteXmlNullable(elem, "wp:start");
				for (var i = 1; i < this.relativeArrPoints.length; ++i) {
					elem = new CT_XmlNode();
					elem.attributes["x"] = this.relativeArrPoints[i].x;
					elem.attributes["y"] = this.relativeArrPoints[i].y;
					writer.WriteXmlNullable(elem, "wp:lineTo");
				}
			}
			writer.WriteXmlNodeEnd(name);
		};
		CPolygonPoint.prototype.readAttr = function(reader) {
			while (reader.MoveToNextAttribute()) {
				switch (reader.GetNameNoNS()) {
					case "x": {
						this.x = AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_emu_to_mm, this.x);
						break;
					}
					case "y": {
						this.y = AscCommon.universalMeasureToMm(reader.GetValue(), AscCommonWord.g_dKoef_emu_to_mm, this.y);
						break;
					}
				}
			}
		};
		CPolygonPoint.prototype.fromXml = function(reader) {
			this.readAttr(reader);
			reader.ReadTillEnd();
		};
		CPolygonPoint.prototype.toXml = function(writer, name) {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlNullableAttributeIntWithKoef("x", this.x, AscCommonWord.g_dKoef_mm_to_emu);
			writer.WriteXmlNullableAttributeIntWithKoef("y", this.y, AscCommonWord.g_dKoef_mm_to_emu);
			writer.WriteXmlAttributesEnd(true);
		};
	}
//settings
	if(typeof AscWord.CDocumentSettings !== "undefined") {
		AscWord.CDocumentSettings.prototype.fromXml = function(reader, doc) {
			let name;
			if (!reader.ReadNextNode()) {
				return;
			}
			name = reader.GetNameNoNS();
			if ("settings" !== name) {
				if (!reader.ReadNextNode()) {
					return;
				}
			}
			name = reader.GetNameNoNS();
			if ("settings" !== name) {
				return;
			}

			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "writeProtection" : {
						this.WriteProtection = new CDocProtect();
						this.WriteProtection.fromXml(reader);
						break;
					}
					case "mirrorMargins" : {
						this.MirrorMargins = CT_BoolW.prototype.toVal(reader, this.MirrorMargins);
						break;
					}
					case "gutterAtTop" : {
						this.GutterAtTop = CT_BoolW.prototype.toVal(reader, this.GutterAtTop);
						break;
					}
					case "trackRevisions" : {
						reader.context.oReadResult.TrackRevisions = CT_BoolW.prototype.toVal(reader, this.TrackRevisions);
						break;
					}
					case "documentProtection" : {
						this.DocumentProtection = new CDocProtect();
						this.DocumentProtection.fromXml(reader);
						break;
					}
					case "defaultTabStop" : {
						let def = reader.context.oReadResult.defaultTabStop;
						reader.context.oReadResult.defaultTabStop = AscCommon.universalMeasureToMm(CT_StringW.prototype.toVal(reader, def), AscCommonWord.g_dKoef_twips_to_mm, def);
						break;
					}
					case "evenAndOddHeaders" : {
						if (doc) {
							doc.Set_DocumentEvenAndOddHeaders(CT_BoolW.prototype.toVal(reader, true));
						}
						break;
					}
					case "footnotePr" : {
						if (doc) {
							doc.Footnotes.FootnotePr.fromXml(reader, true, reader.context.oReadResult.footnoteRefs);
						}
						break;
					}
					case "endnotePr" : {
						if (doc) {
							doc.Endnotes.EndnotePr.fromXml(reader, false, reader.context.oReadResult.endnoteRefs);
						}
						break;
					}
					case "compat" : {
						this.fromXmlCompat(reader);
						break;
					}
					case "mathPr" : {
						elem = new CMathPropertiesSettings();
						elem.fromXml(reader);
						this.MathSettings.SetPr(elem);
						break;
					}
					case "clrSchemeMapping" : {
						if (doc) {
							elem = new AscFormat.ClrMap();
							elem.fromXmlWord(reader);
						}
						break;
					}
					case "decimalSymbol" : {
						this.DecimalSymbol = CT_StringW.prototype.toVal(reader, this.DecimalSymbol);
						break;
					}
					case "listSeparator" : {
						this.ListSeparator = CT_StringW.prototype.toVal(reader, this.ListSeparator);
						break;
					}
					//non standard
					case "SdtGlobalColor" : {
						if (doc) {
							elem = new CT_Color("val", "themeColor", "themeTint", "themeShade");
							elem.fromXml(reader);
							let color = elem.getColor(0, 0, 0);
							if (!color.Auto) {
								doc.SetSdtGlobalColor(color.r, color.g, color.b);
							}
						}
						break;
					}
					case "SdtGlobalShowHighlight" : {
						if (doc) {
							doc.SetSdtGlobalShowHighlight(CT_BoolW.prototype.toVal(reader));
						}
						break;
					}
					case "SpecialFormsHighlight" : {
						if (doc) {
							elem = new CT_Color("val", "themeColor", "themeTint", "themeShade");
							elem.fromXml(reader);
							let color = elem.getColor(0, 0, 0);
							if (color.Auto) {
								doc.SetSpecialFormsHighlight(null);
							} else {
								doc.SetSpecialFormsHighlight(color.r, color.g, color.b);
							}
						}
						break;
					}

				}
			}
		};
		AscWord.CDocumentSettings.prototype.toXml = function(writer) {
			let name = 'w:settings';
			let doc = writer.context.document;

			writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlString(AscCommonWord.g_sXmlSettingsNamespaces);
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(this.WriteProtection, "w:writeProtection");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.MirrorMargins), "w:mirrorMargins");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.GutterAtTop), "w:gutterAtTop");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.TrackRevisions), "w:trackRevisions");
			writer.WriteXmlNullable(this.DocumentProtection, "w:documentProtection");
			writer.WriteXmlNullable(CT_UIntW.prototype.fromVal(AscCommonWord.Default_Tab_Stop, g_dKoef_mm_to_twips), "w:defaultTabStop");
			if (doc) {
				doc.Footnotes.FootnotePr.toXml(writer, "w:footnotePr", true, doc.Footnotes);
				doc.Endnotes.EndnotePr.toXml(writer, "w:endnotePr", false, doc.Endnotes);
			}
			this.toXmlCompat(writer, "w:compat");
			if (this.MathSettings && this.MathSettings.GetPr) {
				writer.WriteXmlNullable(this.MathSettings.GetPr(), "w:mathPr");
			}
			if (doc) {
				doc.clrSchemeMap.toXmlWord(writer, "w:clrSchemeMapping");
			}
			writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.DecimalSymbol), "w:decimalSymbol");
			writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.ListSeparator), "w:listSeparator");
			writer.WriteXmlNodeEnd(name);
		};
		AscWord.CDocumentSettings.prototype.fromXmlCompat = function(reader) {
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "splitPgBreakAndParaMark" : {
						this.SplitPageBreakAndParaMark = CT_BoolW.prototype.toVal(reader, this.SplitPageBreakAndParaMark);
						break;
					}
					case "doNotExpandShiftReturn" : {
						this.DoNotExpandShiftReturn = CT_BoolW.prototype.toVal(reader, this.DoNotExpandShiftReturn);
						break;
					}
					case "balanceSingleByteDoubleByteWidth" : {
						this.BalanceSingleByteDoubleByteWidth = CT_BoolW.prototype.toVal(reader, this.BalanceSingleByteDoubleByteWidth);
						break;
					}
					case "ulTrailSpace" : {
						this.UlTrailSpace = CT_BoolW.prototype.toVal(reader, this.UlTrailSpace);
						break;
					}
					case "useFELayout" : {
						this.UseFELayout = CT_BoolW.prototype.toVal(reader, this.UseFELayout);
						break;
					}
					case "compatSetting" : {
						elem = new CT_CompatSetting();
						elem.fromXml(reader);
						if ("compatibilityMode" === elem.name && "http://schemas.microsoft.com/office/word" === elem.uri) {
							this.CompatibilityMode = parseInt(elem.val);
						}
						break;
					}
				}
			}
		};
		AscWord.CDocumentSettings.prototype.toXmlCompat = function(writer, name) {
			let CompatibilityMode = false === writer.context.docSaveParams.isCompatible ? AscCommon.document_compatibility_mode_Word15 : this.CompatibilityMode;
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.SplitPageBreakAndParaMark), "w:splitPgBreakAndParaMark");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.DoNotExpandShiftReturn), "w:doNotExpandShiftReturn");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.BalanceSingleByteDoubleByteWidth), "w:balanceSingleByteDoubleByteWidth");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.UlTrailSpace), "w:ulTrailSpace");
			writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.UseFELayout), "w:useFELayout");
			CT_CompatSetting.prototype.toXmlValues(writer, "w:compatSetting", "compatibilityMode", "http://schemas.microsoft.com/office/word", CompatibilityMode.toString());
			CT_CompatSetting.prototype.toXmlValues(writer, "w:compatSetting", "overrideTableStyleFontSizeAndJustification", "http://schemas.microsoft.com/office/word", "1");
			CT_CompatSetting.prototype.toXmlValues(writer, "w:compatSetting", "enableOpenTypeFeatures", "http://schemas.microsoft.com/office/word", "1");
			CT_CompatSetting.prototype.toXmlValues(writer, "w:compatSetting", "doNotFlipMirrorIndents", "http://schemas.microsoft.com/office/word", "1");
			writer.WriteXmlNodeEnd(name);
		};
	}
	if(typeof CDocProtect !== "undefined") {
		CDocProtect.prototype.readAttr = function(reader) {
			while (reader.MoveToNextAttribute()) {
				switch (reader.GetNameNoNS()) {
					case "algorithmName": {
						this.algorithmName = fromXml_ST_CryptAlgoritmName(reader.GetValueDecodeXml(), this.algorithmName);
						break;
					}
					case "algIdExt": {
						this.algIdExt = reader.GetValueDecodeXml();
						break;
					}
					case "algIdExtSource": {
						this.algIdExtSource = reader.GetValueDecodeXml();
						break;
					}
					case "cryptProviderType": {
						this.cryptProviderType = fromXml_ST_CryptProvType(reader.GetValueDecodeXml(), this.cryptProviderType);
						break;
					}
					case "cryptProvider": {
						this.cryptProvider = reader.GetValueDecodeXml();
						break;
					}
					case "cryptProviderTypeExt": {
						this.cryptProviderTypeExt = reader.GetValueDecodeXml();
						break;
					}
					case "cryptProviderTypeExtSource": {
						this.cryptProviderTypeExtSource = reader.GetValueDecodeXml();
						break;
					}
					case "cryptAlgorithmSid": {
						this.cryptAlgorithmSid = reader.GetValueInt(this.cryptAlgorithmSid);
						break;
					}
					case "cryptAlgorithmType": {
						this.cryptAlgorithmType = fromXml_ST_CryptAlgType(reader.GetValueDecodeXml(), this.cryptAlgorithmType);
						break;
					}
					case "cryptAlgorithmClass": {
						this.cryptAlgorithmClass = fromXml_ST_CryptAlgClass(reader.GetValueDecodeXml(), this.cryptAlgorithmClass);
						break;
					}
					case "edit": {
						this.edit = fromXml_ST_DocProtect(reader.GetValueDecodeXml(), this.edit);
						break;
					}
					case "enforcement": {
						this.enforcement = reader.GetValueBool();
						break;
					}
					case "formatting": {
						this.formatting = reader.GetValueBool();
						break;
					}
					case "hashValue":
					case "hash": {
						this.hashValue = reader.GetValueDecodeXml();
						break;
					}
					case "recommended": {
						this.recommended = reader.GetValueBool();
						break;
					}
					case "salt":
					case "saltValue": {
						this.saltValue = reader.GetValueDecodeXml();
						break;
					}
					case "cryptSpinCount":
					case "spinCount": {
						this.spinCount = reader.GetValueInt(this.SpinCount);
						break;
					}
				}
			}
		};
		CDocProtect.prototype.fromXml = function(reader) {
			this.readAttr(reader);
			reader.ReadTillEnd();
		};
		CDocProtect.prototype.toXml = function(writer, name) {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlNullableAttributeBool("w:recommended", this.formatting);
			writer.WriteXmlNullableAttributeStringEncode("w:edit", toXml_ST_DocProtect(this.edit));
			writer.WriteXmlNullableAttributeBool("w:enforcement", this.enforcement);
			writer.WriteXmlNullableAttributeBool("w:formatting", this.formatting);
			if(null !== this.cryptProviderType) {
				writer.WriteXmlNullableAttributeStringEncode("w:cryptProviderType", toXml_ST_CryptProvType(this.cryptProviderType));
				writer.WriteXmlNullableAttributeStringEncode("w:algIdExt", this.algIdExt);
				writer.WriteXmlNullableAttributeStringEncode("w:algIdExtSource", this.algIdExtSource);
				writer.WriteXmlNullableAttributeStringEncode("w:cryptAlgorithmClass", toXml_ST_CryptAlgClass(this.cryptAlgorithmClass));
				writer.WriteXmlNullableAttributeStringEncode("w:cryptAlgorithmType", toXml_ST_CryptAlgType(this.cryptAlgorithmType));
				writer.WriteXmlNullableAttributeInt("w:cryptAlgorithmSid", this.cryptAlgorithmSid);
				writer.WriteXmlNullableAttributeStringEncode("w:cryptProvider", this.cryptProvider);
				writer.WriteXmlNullableAttributeStringEncode("w:cryptProviderTypeExt", this.cryptProviderTypeExt);
				writer.WriteXmlNullableAttributeStringEncode("w:cryptProviderTypeExtSource", this.cryptProviderTypeExtSource);

				writer.WriteXmlNullableAttributeInt("w:cryptSpinCount", this.spinCount);
				writer.WriteXmlNullableAttributeStringEncode("w:hash", this.hashValue);
				writer.WriteXmlNullableAttributeStringEncode("w:salt", this.saltValue);
			} else {
				writer.WriteXmlNullableAttributeStringEncode("w:algorithmName", toXml_ST_CryptAlgoritmName(this.algorithmName));
				writer.WriteXmlNullableAttributeStringEncode("w:hashValue", this.hashValue);
				writer.WriteXmlNullableAttributeStringEncode("w:saltValue", this.saltValue);
				writer.WriteXmlNullableAttributeInt("w:spinCount", this.spinCount);
			}
			writer.WriteXmlAttributesEnd(true);
		};
	}
//sdt
	if(typeof CBlockLevelSdt !== "undefined") {
		CBlockLevelSdt.prototype.fromXml = function (reader) {
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "sdtPr" : {
						this.Pr.fromXml(reader);
						break;
					}
					case "sdtEndPr" : {
						break;
					}
					case "sdtContent" : {
						var Content = [];
						CDocument.prototype.fromXmlDocContent(reader, Content, this.LogicDocument.Get_DrawingDocument(), this.Content);
						if (Content.length > 0) {
							this.Content.ReplaceContent(Content);
						}
						break;
					}
				}
			}
		};
		CBlockLevelSdt.prototype.toXml = function (writer, name) {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(this.Pr, "w:sdtPr");
			// writer.WriteXmlArray(this.SdtEndPr, "w:sdtEndPr");
			writer.WriteXmlNodeStart("w:sdtContent");
			writer.WriteXmlAttributesEnd();
			this.Content.Content.forEach(function(item) {
				CDocument.prototype.toXmlDocContentElem(writer, item);
			});
			writer.WriteXmlNodeEnd("w:sdtContent");
			writer.WriteXmlNodeEnd(name);
		};
	}

	if(typeof CInlineLevelSdt !== "undefined") {
		CInlineLevelSdt.prototype.fromXml = function (reader) {
			let elem, depth = reader.GetDepth();
			while (reader.ReadNextSiblingNode(depth)) {
				switch (reader.GetNameNoNS()) {
					case "sdtPr" : {
						this.Pr.fromXml(reader);
						break;
					}
					case "sdtEndPr" : {
						break;
					}
					case "sdtContent" : {
						CParagraphContentWithParagraphLikeContent.prototype.fromXml.call(this, reader);
						break;
					}
				}
			}
		};
		CInlineLevelSdt.prototype.toXml = function (writer, name) {
			writer.WriteXmlNodeStart(name);
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(this.Pr, "w:sdtPr");
			// writer.WriteXmlArray(this.SdtEndPr, "w:sdtEndPr");
			writer.WriteXmlNodeStart("w:sdtContent");
			writer.WriteXmlAttributesEnd();
			CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
			writer.WriteXmlNodeEnd("w:sdtContent");
			writer.WriteXmlNodeEnd(name);
		};
	}
	function CT_SdtRow(table) {
		this.table = table;
		return this;
	}
	CT_SdtRow.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "sdtPr" : {
					break;
				}
				case "sdtEndPr" : {
					break;
				}
				case "sdtContent" : {
					this.table.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_SdtRow.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.SdtPr, "w:sdtPr");
		// writer.WriteXmlArray(this.SdtEndPr, "w:sdtEndPr");
		writer.WriteXmlNodeStart("w:sdtContent");
		writer.WriteXmlAttributesEnd();
		this.table.toXml(writer, name);
		writer.WriteXmlNodeEnd("w:sdtContent");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_SdtCell(row) {
		this.row = row;
		return this;
	}

	CT_SdtCell.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "sdtPr" : {
					break;
				}
				case "sdtEndPr" : {
					break;
				}
				case "sdtContent" : {
					this.row.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_SdtCell.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.SdtPr, "w:sdtPr");
		// writer.WriteXmlArray(this.SdtEndPr, "w:sdtEndPr");
		writer.WriteXmlNodeStart("w:sdtContent");
		writer.WriteXmlAttributesEnd();
		this.row.toXml(writer, name);
		writer.WriteXmlNodeEnd("w:sdtContent");
		writer.WriteXmlNodeEnd(name);
	};
	CSdtPr.prototype.fromXml = function(reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "rPr" : {
					this.TextPr = new CTextPr();
					this.TextPr.fromXml(reader);
					break;
				}
				case "alias" : {
					this.Alias = CT_StringW.prototype.toVal(reader, this.Alias);
					break;
				}
				case "tag" : {
					this.Tag = CT_StringW.prototype.toVal(reader, this.Tag);
					break;
				}
				case "id" : {
					this.Id = CT_IntW.prototype.toVal(reader, this.Id);
					break;
				}
				case "lock" : {
					this.Lock = fromXml_ST_Lock(CT_StringW.prototype.toVal(reader, this.Lock), this.Lock);
					break;
				}
				case "placeholder" : {
					var subDepth = reader.GetDepth();
					while (reader.ReadNextSiblingNode(subDepth)) {
						if ("docPart" === reader.GetNameNoNS()) {
							this.Placeholder = CT_StringW.prototype.toVal(reader, this.Placeholder);
						}
					}
					break;
				}
				case "temporary" : {
					this.Temporary = CT_BoolW.prototype.toVal(reader, this.Temporary);
					break;
				}
				case "showingPlcHdr" : {
					this.ShowingPlcHdr = CT_BoolW.prototype.toVal(reader, this.ShowingPlcHdr);
					break;
				}
				case "dataBinding" : {
					break;
				}
				case "label" : {
					this.Label = CT_IntW.prototype.toVal(reader, this.Label);
					break;
				}
				case "tabIndex" : {
					break;
				}
				case "equation" : {
					this.Equation = true;
					break;
				}
				case "comboBox" : {
					this.ComboBox = new AscWord.CSdtComboBoxPr();
					this.ComboBox.fromXml(reader);
					break;
				}
				case "date" : {
					this.Date = new AscWord.CSdtDatePickerPr();
					this.Date.fromXml(reader);
					break;
				}
				case "docPartObj" : {
					this.fromXmlDocPartObj(reader);
					break;
				}
				case "docPartList" : {
					break;
				}
				case "dropDownList" : {
					this.DropDown = new AscWord.CSdtComboBoxPr();
					this.DropDown.fromXml(reader);
					break;
				}
				case "picture" : {
					this.Picture = true;
					elem = new AscWord.CSdtPictureFormPr();
					elem.fromXml(reader);
					this.PictureFormPr = elem;
					break;
				}
				case "richText" : {
					break;
				}
				case "text" : {
					this.Text = true;
					//todo MultiLine
					break;
				}
				case "citation" : {
					break;
				}
				case "group" : {
					break;
				}
				case "bibliography" : {
					break;
				}
				case "checkbox" : {
					this.CheckBox = new AscWord.CSdtCheckBoxPr();
					this.CheckBox.fromXml(reader);
					break;
				}
				case "appearance" : {
					this.Appearance = fromXml_ST_SdtAppearance(CT_StringW15.prototype.toVal(reader, this.Appearance), this.Appearance);
					break;
				}
				case "color" : {
					//todo w15
					elem = new CT_Color("val", "themeColor", "themeTint", "themeShade");
					elem.fromXml(reader);
					this.Color = elem.getColor(0, 0, 0);
					break;
				}
				case "formPr" : {
					elem = new AscWord.CSdtFormPr();
					elem.fromXml(reader);
					this.FormPr = elem;
					break;
				}
				case "textFormPr" : {
					elem = new AscWord.CSdtTextFormPr();
					elem.fromXml(reader);
					this.TextForm = elem;
					break;
				}
				case "complexFormPr" : {
					elem = new AscWord.CSdtComplexFormPr();
					elem.fromXml(reader);
					this.ComplexFormPr = elem;
					break;
				}
			}
		}
	};
	CSdtPr.prototype.toXml = function(writer, name) {
		var Color = new CT_Color("val", "themeColor", "themeTint", "themeShade");
		Color.setColor(this.Color);

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Alias), "w:alias");
		writer.WriteXmlNullable(CT_StringW15.prototype.fromVal(toXml_ST_SdtAppearance(this.Appearance)), "w15:appearance");
		if (!Color.isEmpty()) {
			writer.WriteXmlNullable(Color, "w15:color");
		}
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.Id), "w:id");
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.Label), "w:label");
		// writer.WriteXmlNullable(this.tabIndex, "w:tabIndex");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_Lock(this.Lock)), "w:lock");
		if (this.Placeholder) {
			writer.WriteXmlNodeStart("w:placeholder");
			writer.WriteXmlAttributesEnd();
			writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Placeholder), "w:docPart");
			writer.WriteXmlNodeEnd("w:placeholder");
		}
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.ShowingPlcHdr), "w:showingPlcHdr");
		// writer.WriteXmlNullable(this.dataBinding, "w:dataBinding");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(this.Temporary), "w:temporary");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.Tag), "w:tag");
		if (this.ComboBox) {
			this.ComboBox.toXml(writer, "w:comboBox");
		} else if (this.Picture) {
			if (this.PictureFormPr) {
				this.PictureFormPr.toXml(writer, "w:picture");
			} else {
				writer.WriteXmlNodeStart("w:picture");
				writer.WriteXmlAttributesEnd(true);
			}
		} else if (this.Text) {
			writer.WriteXmlNodeStart("w:text");
			writer.WriteXmlAttributesEnd(true);
		} else if (this.Equation) {
			writer.WriteXmlNodeStart("w:equation");
			writer.WriteXmlAttributesEnd();
		} else if (this.IsBuiltInDocPart()) {
			this.toXmlDocPartObj(writer, "w:docPartObj");
		} else if (this.DropDown) {
			this.DropDown.toXml(writer, "w:dropDownList");
		} else if (this.Date) {
			this.Date.toXml(writer, "w:date");
		} else if (undefined !== this.CheckBox) {
			writer.WriteXmlNullable(this.CheckBox, "w14:checkbox");
		}
		writer.WriteXmlNullable(this.FormPr, "w:formPr");
		writer.WriteXmlNullable(this.TextForm, "w:textFormPr");
		writer.WriteXmlNullable(this.ComplexFormPr, "w:complexFormPr");
		writer.WriteXmlNullable(this.TextPr, "w:rPr");
		writer.WriteXmlNodeEnd(name);
	};
	CSdtPr.prototype.fromXmlDocPartObj = function(reader) {
		let docPartObj = this.DocPartObj;
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "docPartGallery" : {
					docPartObj.Gallery = CT_StringW.prototype.toVal(reader, docPartObj.Gallery);
					break;
				}
				case "docPartCategory" : {
					docPartObj.Category = CT_StringW.prototype.toVal(reader, docPartObj.Category);
					break;
				}
				case "docPartUnique" : {
					docPartObj.Unique = CT_BoolW.prototype.toVal(reader, docPartObj.Unique);
					break;
				}
			}
		}
	};
	CSdtPr.prototype.toXmlDocPartObj = function(writer, name) {
		let docPartObj = this.DocPartObj;

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(docPartObj.Gallery), "w:docPartGallery");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(docPartObj.Category), "w:docPartCategory");
		writer.WriteXmlNullable(CT_BoolW.prototype.fromVal(docPartObj.Unique), "w:docPartUnique");
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtComboBoxPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "lastValue": {
					break;
				}
			}
		}
	};
	AscWord.CSdtComboBoxPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "format" : {
					this.GetFormat().fromXml(reader);
					break;
				}
				case "listItem" : {
					elem = new AscWord.CSdtListItem();
					elem.fromXml(reader);
					this.ListItems.push(elem);
					break;
				}
			}
		}
	};
	AscWord.CSdtComboBoxPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeStringEncode("w:lastValue", this.LastValue);
		writer.WriteXmlAttributesEnd();
		let format = this.GetFormat();
		if (format && !format.IsEmpty()) {
			format.toXml(writer, "w:format");
		}
		writer.WriteXmlArray(this.ListItems, "w:listItem");
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtListItem.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "displayText": {
					this.DisplayText = reader.GetValueDecodeXml();
					break;
				}
				case "value": {
					this.Value = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	AscWord.CSdtListItem.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscWord.CSdtListItem.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		if (this.DisplayText) {
			writer.WriteXmlNullableAttributeStringEncode("w:displayText", this.DisplayText);
		}
		writer.WriteXmlNullableAttributeStringEncode("w:value", this.Value);
		writer.WriteXmlAttributesEnd(true);
	};
	AscWord.CSdtDatePickerPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "fullDate": {
					this.FullDate = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	AscWord.CSdtDatePickerPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "calendar" : {
					this.Calendar = fromXml_ST_CalendarType(CT_StringW.prototype.toVal(reader, this.Calendar), this.Calendar);
					break;
				}
				case "dateFormat" : {
					this.DateFormat = CT_StringW.prototype.toVal(reader, this.DateFormat);
					break;
				}
				case "lid" : {
					this.LangId = Asc.g_oLcidNameToIdMap[CT_StringW.prototype.toVal(reader, this.LangId)] || this.LangId;
					break;
				}
				case "storeMappedDataAs" : {
					break;
				}
			}
		}
	};
	AscWord.CSdtDatePickerPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:fullDate", this.FullDate);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(toXml_ST_CalendarType(this.Calendar)), "w:calendar");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(this.DateFormat), "w:dateFormat");
		writer.WriteXmlNullable(CT_StringW.prototype.fromVal(Asc.g_oLcidIdToNameMap[this.LangId]), "w:lid");
		// writer.WriteXmlNullable(this.StoreMappedDataAs, "w:storeMappedDataAs");
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtCheckBoxPr.prototype.fromXml = function (reader) {
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "checked" : {
					this.Checked = CT_BoolW14.prototype.toVal(reader, this.Checked);
					break;
				}
				case "checkedState" : {
					while (reader.MoveToNextAttribute()) {
						switch (reader.GetNameNoNS()) {
							case "val": {
								this.CheckedSymbol = reader.GetValueUInt(this.CheckedSymbol, 16);
								break;
							}
							case "font": {
								this.CheckedFont = reader.GetValueDecodeXml();
								break;
							}
						}
					}
					reader.ReadTillEnd();
					break;
				}
				case "uncheckedState" : {
					while (reader.MoveToNextAttribute()) {
						switch (reader.GetNameNoNS()) {
							case "val": {
								this.UncheckedSymbol = reader.GetValueUInt(this.CheckedSymbol, 16);
								break;
							}
							case "font": {
								this.UncheckedFont = reader.GetValueDecodeXml();
								break;
							}
						}
					}
					reader.ReadTillEnd();
					break;
				}
				case "groupKey" : {
					this.GroupKey = CT_StringW14.prototype.toVal(reader, this.GroupKey);
					break;
				}
			}
		}
	};
	AscWord.CSdtCheckBoxPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(CT_BoolW14.prototype.fromVal(this.Checked), "w14:checked");
		writer.WriteXmlNodeStart("w14:checkedState");
		writer.WriteXmlNullableAttributeString("w14:val", AscCommon.Int16ToHex(this.CheckedSymbol));
		writer.WriteXmlNullableAttributeString("w14:font", this.CheckedFont);
		writer.WriteXmlAttributesEnd(true);
		writer.WriteXmlNodeStart("w14:uncheckedState");
		writer.WriteXmlNullableAttributeString("w14:val", AscCommon.Int16ToHex(this.UncheckedSymbol));
		writer.WriteXmlNullableAttributeString("w14:font", this.UncheckedFont);
		writer.WriteXmlAttributesEnd(true);
		writer.WriteXmlNullable(CT_StringW14.prototype.fromVal(this.GroupKey), "w14:groupKey");
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtFormPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "key": {
					this.Key = reader.GetValueDecodeXml();
					break;
				}
				case "label": {
					this.Label = reader.GetValueDecodeXml();
					break;
				}
				case "helpText": {
					this.HelpText = reader.GetValueDecodeXml();
					break;
				}
				case "required": {
					this.Required = reader.GetValueBool();
					break;
				}
			}
		}
	};
	AscWord.CSdtFormPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "shd" : {
					this.Shd = new CDocumentShd();
					this.Shd.fromXml(reader);
					break;
				}
				case "border" : {
					this.Border = new CDocumentBorder();
					this.Border.fromXml(reader);
					break;
				}
			}
		}
	};
	AscWord.CSdtFormPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:key", this.Key);
		writer.WriteXmlNullableAttributeStringEncode("w:label", this.Label);
		writer.WriteXmlNullableAttributeStringEncode("w:helpText", this.HelpText);
		writer.WriteXmlNullableAttributeBool("w:required", this.Required);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Shd, "w:shd");
		writer.WriteXmlNullable(this.Border, "w:border");
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtTextFormPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "multiLine": {
					this.MultiLine = reader.GetValueBool();
					break;
				}
				case "autoFit": {
					this.AutoFit = reader.GetValueBool();
					break;
				}
			}
		}
	};
	AscWord.CSdtTextFormPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "comb" : {
					this.Comb = true;
					this.fromXmlComb(reader);
					break;
				}
				case "maxCharacters" : {
					this.MaxCharacters = CT_IntW.prototype.toVal(reader, this.MaxCharacters);
					break;
				}
				case "combBorder" : {
					this.CombBorder = new CDocumentBorder();
					this.CombBorder.fromXml(reader);
					break;
				}
				case "format" : {
					this.GetFormat().fromXml(reader);
					break;
				}
			}
		}
	};
	AscWord.CSdtTextFormPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("w:multiLine", this.MultiLine);
		writer.WriteXmlNullableAttributeBool("w:autoFit", this.AutoFit);
		writer.WriteXmlAttributesEnd();
		if (this.Comb) {
			this.toXmlComb(writer, "w:comb");
		}
		writer.WriteXmlNullable(CT_IntW.prototype.fromVal(this.MaxCharacters), "w:maxCharacters");
		writer.WriteXmlNullable(this.CombBorder, "w:combBorder");
		let format = this.GetFormat();
		if (format && !format.IsEmpty()) {
			format.toXml(writer, "w:format");
		}
		writer.WriteXmlNodeEnd(name);
	};
	AscWord.CSdtTextFormPr.prototype.readAttrComb = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "width": {
					this.Width = reader.GetValueInt();
					break;
				}
				case "sym": {
					this.CombPlaceholderSymbol = reader.GetValueDecodeXml();
					break;
				}
				case "font": {
					this.CombPlaceholderFont = reader.GetValueDecodeXml();
					break;
				}
				case "wRule": {
					this.WidthRule = fromXml_ST_CombFormWidthRule(reader.GetValue(), this.WidthRule);
					break;
				}
			}
		}
	};
	AscWord.CSdtTextFormPr.prototype.fromXmlComb = function (reader) {
		this.readAttrComb(reader);
		reader.ReadTillEnd();
	};
	AscWord.CSdtTextFormPr.prototype.toXmlComb = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:width", this.Width);
		writer.WriteXmlNonEmptyAttributeStringEncode("w:sym", this.CombPlaceholderSymbol);
		writer.WriteXmlNonEmptyAttributeStringEncode("w:font", this.CombPlaceholderFont);
		writer.WriteXmlNonEmptyAttributeStringEncode("w:wRule", toXml_ST_CombFormWidthRule(this.WidthRule));
		writer.WriteXmlAttributesEnd(true);
	};
	AscWord.CSdtPictureFormPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "scaleFlag": {
					this.ScaleFlag = reader.GetValueInt();
					break;
				}
				case "lockProportions": {
					this.Proportions = reader.GetValueBool();
					break;
				}
				case "respectBorders": {
					this.Borders = reader.GetValueBool();
					break;
				}
				case "shiftX": {
					this.ShiftX = reader.GetValueDouble();
					break;
				}
				case "shiftY": {
					this.ShiftY = reader.GetValueDouble();
					break;
				}
			}
		}
	};
	AscWord.CSdtPictureFormPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscWord.CSdtPictureFormPr.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:scaleFlag", this.ScaleFlag);
		writer.WriteXmlNullableAttributeBool("w:lockProportions", this.Proportions);
		writer.WriteXmlNullableAttributeBool("w:respectBorders", this.Borders);
		writer.WriteXmlNullableAttributeDouble("w:shiftX", this.ShiftX);
		writer.WriteXmlNullableAttributeDouble("w:shiftY", this.ShiftY);
		writer.WriteXmlAttributesEnd(true);
	};
	AscWord.CTextFormFormat.prototype.readAttr = function (reader) {
		let type = Asc.TextFormFormatType.None;
		let val = "";
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					type = fromXml_ST_TextFormFormatType(reader.GetValue(), type);
					break;
				}
				case "val": {
					val = reader.GetValueDecodeXml();
					break;
				}
				case "symbols": {
					this.SetSymbols(reader.GetValueDecodeXml());
					break;
				}
			}
		}
		this.SetType(type, val);
	};
	AscWord.CTextFormFormat.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscWord.CTextFormFormat.prototype.toXml = function (writer, name) {
		let type = this.GetType();
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:type", toXml_ST_TextFormFormatType(type));
		if (Asc.TextFormFormatType.Mask === type) {
			writer.WriteXmlNullableAttributeStringEncode("w:val", this.GetMask());
		} else if (Asc.TextFormFormatType.RegExp === type) {
			writer.WriteXmlNullableAttributeStringEncode("w:val", this.GetRegExp());
		}
		writer.WriteXmlNonEmptyAttributeStringEncode("w:symbols", this.GetSymbols(true));
		writer.WriteXmlAttributesEnd(true);
	};
	AscWord.CSdtComplexFormPr.prototype.readAttr = function (reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					this.Type = fromXml_ST_ComplexFormType(reader.GetValue(), this.Type);
					break;
				}
			}
		}
	};
	AscWord.CSdtComplexFormPr.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	AscWord.CSdtComplexFormPr.prototype.toXml = function (writer, name) {
		let type = this.GetType();
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:type", toXml_ST_ComplexFormType(this.GetType()));
		writer.WriteXmlAttributesEnd(true);
	};
//comments
	if(typeof CComments !== "undefined") {
		CComments.prototype.ReadFromXml = function(documentPart, context) {
			let reader, ct_comments, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible;
			let commentsPart = documentPart.getPartByRelationshipType(openXml.Types.wordComments.relationType);
			if (commentsPart) {
				let commentsContent = commentsPart.getDocumentContent();
				reader = new StaxParser(commentsContent, commentsPart, context);
				ct_comments = new CT_Comments();
				ct_comments.fromXml(reader);
			}
			let peoplePart = documentPart.getPartByRelationshipType(openXml.Types.wordPeople.relationType);
			if (peoplePart) {
				let peopleContent = peoplePart.getDocumentContent();
				reader = new StaxParser(peopleContent, peoplePart, context);
				ct_people = new CT_People();
				ct_people.fromXml(reader);
			}
			let commentsExtendedPart = documentPart.getPartByRelationshipType(openXml.Types.wordCommentsExtended.relationType);
			if (commentsExtendedPart) {
				let commentsExtendedContent = commentsExtendedPart.getDocumentContent();
				reader = new StaxParser(commentsExtendedContent, commentsExtendedPart, context);
				ct_commentsExt = new CT_CommentsEx();
				ct_commentsExt.fromXml(reader);
			}
			let commentsIdsPart = documentPart.getPartByRelationshipType(openXml.Types.wordCommentsIds.relationType);
			if (commentsIdsPart) {
				let commentsIdsContent = commentsIdsPart.getDocumentContent();
				reader = new StaxParser(commentsIdsContent, commentsIdsPart, context);
				ct_commentsIds = new CT_CommentsIds();
				ct_commentsIds.fromXml(reader);
			}
			let commentsExtensiblePart = documentPart.getPartByRelationshipType(openXml.Types.wordCommentsExtensible.relationType);
			if (commentsExtensiblePart) {
				let commentsExtensibleContent = commentsExtensiblePart.getDocumentContent();
				reader = new StaxParser(commentsExtensibleContent, commentsExtensiblePart, context);
				ct_commentsExtensible = new CT_CommentsExtensible();
				ct_commentsExtensible.fromXml(reader);
			}

			if (ct_comments) {
				if (ct_commentsExt) {
					ct_comments.initReplies(ct_commentsExt);
				}
				for (let i = 0; i < ct_comments.comment.length; ++i) {
					let ct_comment = ct_comments.comment[i];
					let commentData = new AscCommon.CCommentData();
					commentData.ReadFromXml(ct_comment, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible);
					let oNewComment = new AscCommon.CComment(this, commentData);
					this.Add(oNewComment);
					context.commentDataById[ct_comment.id] = oNewComment;
				}
			}
		};
		CComments.prototype.WriteToXml = function(memory, docPart) {
			let ct_comments = new CT_Comments();
			let ct_people = new CT_People();
			let ct_commentsExt = new CT_CommentsEx();
			let ct_commentsIds = new CT_CommentsIds();
			let ct_commentsExtensible = new CT_CommentsExtensible();
			//todo document commments
			for(let i in this.m_arrCommentsById) {
				if(this.m_arrCommentsById.hasOwnProperty(i)) {
					let oComment = this.m_arrCommentsById[i];
					if (oComment.IsGlobalComment()) {
					} else {
					}
					let ct_comment = new CT_Comment();
					ct_comments.comment.push(ct_comment);
					oComment.Data.WriteToXml(ct_comment, null, memory.context, ct_comments, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible);
					memory.context.commentDataById[oComment.Get_Id()] = ct_comment;
				}
			}
			if (0 === ct_comments.comment.length) {
				return;
			}
			let commentsPart = docPart.part.addPart(AscCommon.openXml.Types.wordComments);
			commentsPart.part.setDataXml(ct_comments, memory);
			memory.Seek(0);

			let peoplePart = docPart.part.addPart(AscCommon.openXml.Types.wordPeople);
			peoplePart.part.setDataXml(ct_people, memory);
			memory.Seek(0);

			let commentsExtendedPart = docPart.part.addPart(AscCommon.openXml.Types.wordCommentsExtended);
			commentsExtendedPart.part.setDataXml(ct_commentsExt, memory);
			memory.Seek(0);

			let commentsIdsPart = docPart.part.addPart(AscCommon.openXml.Types.wordCommentsIds);
			commentsIdsPart.part.setDataXml(ct_commentsIds, memory);
			memory.Seek(0);

			let commentsExtensiblePart = docPart.part.addPart(AscCommon.openXml.Types.wordCommentsExtensible);
			commentsExtensiblePart.part.setDataXml(ct_commentsExtensible, memory);
			memory.Seek(0);
		};
	}
	if(typeof CCommentData !== "undefined") {
		CCommentData.prototype.ReadFromXml = function(ct_comment, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible) {
			let paraId = ct_comment.paraId;
			let durableId = ct_commentsIds && ct_commentsIds.commentId[paraId];
			let commentsExt = ct_commentsExt && ct_commentsExt.commentEx[paraId];
			let commentExtensible = ct_commentsExtensible && ct_commentsExtensible.commentExtensible[durableId];
			let presenceInfo = ct_people && ct_people.person[ct_comment.author];

			ct_comment.content.SelectAll();
			let text = ct_comment.content.GetSelectedText(true);

			this.m_sText = text || this.m_sText;
			this.m_sTime = ct_comment.date || this.m_sTime;
			this.m_sOOTime  = (commentExtensible && commentExtensible.dateUtc) || this.m_sOOTime;
			this.m_sUserId = (presenceInfo && presenceInfo.userId) || this.m_sUserId;
			this.m_sProviderId = (presenceInfo && presenceInfo.providerId) || this.m_sProviderId;
			this.m_sUserName = ct_comment.author || this.m_sUserName;
			this.m_sInitials = ct_comment.initials || this.m_sInitials;
			this.m_sQuoteText = "QuoteText";
			this.m_bSolved = (commentsExt && commentsExt.done) || this.m_bSolved;
			this.m_nDurableId = durableId || this.m_nDurableId;
			this.m_sUserData = "";//todo
			for (let i = 0; i < ct_comment.replies.length; ++i) {
				let commentData = new AscCommon.CCommentData();
				commentData.ReadFromXml(ct_comment.replies[i], ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible);
				this.m_aReplies.push(commentData);
			}
		};
		CCommentData.prototype.WriteTextToParagraph = function(text, paraId, documentContent) {
			let Spacing = new CParaSpacing();
			Spacing.LineRule = Asc.linerule_Auto;
			Spacing.Line = 1;
			Spacing.After = 0;
			Spacing.Before = 0;
			let Ind = new CParaInd();
			Ind.firstLine = 0;
			Ind.Left = 0;
			Ind.Right = 0;
			let paragraph = new Paragraph(documentContent.DrawingDocument, documentContent);
			paragraph.ParaId = paraId;
			paragraph.Pr.Spacing = Spacing;
			paragraph.Pr.Ind = Ind;
			paragraph.Pr.Jc = AscCommon.align_Left;
			let run = new ParaRun(paragraph, false);
			run.Pr.RFonts.SetAll('Arial');
			run.Pr.FontSize = 11;
			run.AddText(text, -1);
			paragraph.AddToContent(paragraph.GetElementsCount(), run);
			return paragraph;
		};
		CCommentData.prototype.WriteToXml = function(ct_comment, paraIdParent, context, ct_comments, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible) {
			let id = context.commentIdIndex++;
			let paraId = context.paraIdIndex++;
			let durableId = this.m_nDurableId || 1;
			let presenceInfo = ct_people.person[this.m_sUserName];
			if(!presenceInfo) {
				presenceInfo = new CT_PresenceInfo();
				ct_people.person[this.m_sUserName] = presenceInfo;
			}
			let commentsExt = ct_commentsExt.commentEx[paraId];
			if(!commentsExt) {
				commentsExt = new CT_CommentEx();
				ct_commentsExt.commentEx[paraId] = commentsExt;
			}
			let commentExtensible = ct_commentsExtensible.commentExtensible[durableId];
			if(!commentExtensible) {
				commentExtensible = new CT_CommentExtensible();
				ct_commentsExtensible.commentExtensible[durableId] = commentExtensible;
			}

			ct_comment.author = this.m_sUserName;
			ct_comment.id = id;
			ct_comment.date = parseInt(this.m_sTime);
			ct_comment.initials = this.m_sInitials;
			//todo newline
			ct_comment.content.Content[0] = this.WriteTextToParagraph(this.m_sText, paraId, ct_comment.content);

			presenceInfo.userId = this.m_sUserId;
			presenceInfo.providerId = this.m_sProviderId;
			commentsExt.paraId = paraId;
			commentsExt.paraIdParent = paraIdParent;
			commentsExt.done = this.m_bSolved;
			ct_commentsIds.commentId[paraId] = durableId;
			commentExtensible.durableId = durableId;
			commentExtensible.dateUtc = parseInt(this.m_sOOTime);

			for (let i = 0; i < this.m_aReplies.length; ++i) {
				let ct_reply = new CT_Comment();
				ct_comment.replies.push(ct_reply);
				ct_comments.comment.push(ct_reply);
				this.m_aReplies[i].WriteToXml(ct_reply, paraId, context, ct_comments, ct_people, ct_commentsExt, ct_commentsIds, ct_commentsExtensible);
			}
		};
	}
	function CT_Comment() {
		this.id = null;
		this.author = null;
		this.date = null;
		this.initials = null;
		this.content = new CDocumentContent(this, undefined, 0, 0, 0, 0, false, true);

		this.paraId = null;//from first paragraph
		this.replies = [];//from commentsExt
		return this;
	}

	CT_Comment.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					this.id = reader.GetValueInt(this.id);
					break;
				}
				case "author": {
					this.author = reader.GetValueDecodeXml();
					break;
				}
				case "date": {
					let dateMs = AscCommon.getTimeISO8601(reader.GetValueDecodeXml());
					if (!isNaN(dateMs)) {
						this.date = dateMs;
					}
					break;
				}
				case "initials": {
					this.initials = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_Comment.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var Content = [];
		CDocument.prototype.fromXmlDocContent(reader, Content, null, this);
		if (Content.length > 0) {
			this.content.ReplaceContent(Content);
		}
		this.paraId = this.content.Content[0].ParaId;
	};
	CT_Comment.prototype.toXml = function(writer, name) {
		let date = this.date ? new Date(this.date).toISOString().slice(0, 19) + 'Z' : null;
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:initials", this.initials);
		writer.WriteXmlNullableAttributeStringEncode("w:author", this.author);
		writer.WriteXmlNullableAttributeString("w:date", date);
		writer.WriteXmlNullableAttributeInt("w:id", this.id);
		writer.WriteXmlAttributesEnd();
		this.content.Content.forEach(function(item) {
			CDocument.prototype.toXmlDocContentElem(writer, item);
		});
		writer.WriteXmlNodeEnd(name);
	};
	CT_Comment.prototype.toXmlRef = function(writer, Start) {
		let paraComment = new CT_MarkupRange();
		paraComment.id = this.id;
		if (Start) {
			paraComment.toXml(writer, "w:commentRangeStart");
		} else {
			paraComment.toXml(writer, "w:commentRangeEnd");
			//todo add type
			writer.WriteXmlString('<w:r><w:commentReference w:id="' + paraComment.id + '"/></w:r>');
		}
		this.replies.forEach(function(elem) {
			elem.toXmlRef(writer, Start);
		});
	};
	function CT_Comments() {
		this.comment = [];
		return this;
	}

	CT_Comments.prototype.fromXml = function(reader) {
		let name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("comments" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("comments" !== name) {
			return;
		}

		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "comment" : {
					elem = new CT_Comment();
					elem.fromXml(reader);
					this.comment.push(elem);
					break;
				}
			}
		}
	};
	CT_Comments.prototype.toXml = function(writer) {
		let name = 'w:comments';

		writer.WriteXmlString(AscCommonWord.g_sXmlHeader);
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlWordCommentsNamespaces);
		writer.WriteXmlAttributesEnd();
		for (let i = 0; i < this.comment.length; ++i) {
			this.comment[i].toXml(writer, "w:comment");
			for (let j = 0; j < this.comment[i].replies.length; ++j) {
				this.comment[i].replies[j].toXml(writer, "w:comment");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};
	CT_Comments.prototype.initReplies = function(ct_commentsExt) {
		let newComment = [];
		let replies = {};
		for (let i = 0; i < this.comment.length; ++i) {
			let commentsExt = ct_commentsExt.commentEx[this.comment[i].paraId];
			if (commentsExt && null !== commentsExt.paraIdParent) {
				if (!replies[commentsExt.paraIdParent]) {
					replies[commentsExt.paraIdParent] = [];
				}
				replies[commentsExt.paraIdParent].push(this.comment[i]);
			} else {
				newComment.push(this.comment[i]);
			}
		}
		for (let i = 0; i < newComment.length; ++i) {
			if (replies[newComment[i].paraId]) {
				newComment[i].replies = replies[newComment[i].paraId];
			}
		}
		this.comment = newComment;
	};
	function CT_CommentEx() {
		this.paraId = null;
		this.paraIdParent = null;
		this.done = null;
		return this;
	}
	CT_CommentEx.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "paraId": {
					this.paraId = reader.GetValueUInt(this.paraId, 16);
					break;
				}
				case "paraIdParent": {
					this.paraIdParent = reader.GetValueUInt(this.paraIdParent, 16);
					break;
				}
				case "done": {
					this.done = reader.GetValueBool();
					break;
				}
			}
		}
	};
	CT_CommentEx.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_CommentEx.prototype.toXml = function (writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w15:paraId", AscCommon.Int32ToHexOrNull(this.paraId));
		writer.WriteXmlNullableAttributeString("w15:paraIdParent", AscCommon.Int32ToHexOrNull(this.paraIdParent));
		writer.WriteXmlNullableAttributeBool("w15:done", this.done);
		writer.WriteXmlAttributesEnd(true);
	};

	function CT_CommentsEx() {
		this.commentEx = {};
		return this;
	}
	CT_CommentsEx.prototype.fromXml = function(reader) {
		let name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("commentsEx" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("commentsEx" !== name) {
			return;
		}

		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "commentEx" : {
					elem = new CT_CommentEx();
					elem.fromXml(reader);
					this.commentEx[elem.paraId] = elem;
					break;}
			}
		}
	};
	CT_CommentsEx.prototype.toXml = function (writer) {
		let name = 'w15:commentsEx';

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlWordCommentsExtendedNamespaces);
		writer.WriteXmlAttributesEnd();
		for (let paraId in this.commentEx) {
			if (this.commentEx.hasOwnProperty(paraId)) {
				this.commentEx[paraId].toXml(writer, "w15:commentEx");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	function CT_PresenceInfo() {
		this.providerId = null;
		this.userId = null;
		return this;
	}

	CT_PresenceInfo.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "providerId": {
					this.providerId = reader.GetValueDecodeXml();
					break;
				}
				case "userId": {
					this.userId = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_PresenceInfo.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_PresenceInfo.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w15:providerId", this.providerId);
		writer.WriteXmlNullableAttributeStringEncode("w15:userId", this.userId);
		writer.WriteXmlAttributesEnd(true);
	};
	function CT_Person() {
		this.author = null;
		this.presenceInfo = null;
		return this;
	}

	CT_Person.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "author": {
					this.author = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_Person.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "presenceInfo" : {
					this.presenceInfo = new CT_PresenceInfo();
					this.presenceInfo.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_Person.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w15:author", this.author);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.presenceInfo, "w15:presenceInfo");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_People() {
		this.person = {};
		return this;
	}
	CT_People.prototype.fromXml = function(reader) {
		let name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("people" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("people" !== name) {
			return;
		}

		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "person" : {
					elem = new CT_Person();
					elem.fromXml(reader);
					this.person[elem.author] = elem.presenceInfo;
					break;
				}
			}
		}
	};
	CT_People.prototype.toXml = function(writer) {
		let name = 'w15:people';

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlWordPeopleNamespaces);
		writer.WriteXmlAttributesEnd();
		for (let author in this.person) {
			if (this.person.hasOwnProperty(author)) {
				let elem = new CT_Person();
				elem.author = author;
				elem.presenceInfo = this.person[author];
				elem.toXml(writer, "w15:person");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	function CT_CommentId() {
		this.paraId = null;
		this.durableId = null;
		return this;
	}

	CT_CommentId.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "paraId": {
					this.paraId = reader.GetValueUInt(this.paraId, 16);
					break;
				}
				case "durableId": {
					this.durableId = AscCommon.FixDurableId(reader.GetValueUInt(this.durableId, 16));
					break;
				}
			}
		}
	};
	CT_CommentId.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_CommentId.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w16cid:paraId", AscCommon.Int32ToHexOrNull(this.paraId));
		writer.WriteXmlNullableAttributeString("w16cid:durableId", AscCommon.Int32ToHexOrNull(this.durableId));
		writer.WriteXmlAttributesEnd(true);
	};
	function CT_CommentsIds() {
		this.commentId = {};
		return this;
	}

	CT_CommentsIds.prototype.fromXml = function(reader) {
		let name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("commentsIds" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("commentsIds" !== name) {
			return;
		}

		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "commentId" : {
					elem = new CT_CommentId();
					elem.fromXml(reader);
					this.commentId[elem.paraId] = elem.durableId;
					break;
				}
			}
		}
	};
	CT_CommentsIds.prototype.toXml = function(writer) {
		let name = 'w16cid:commentsIds';

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlWordCommentsIdsNamespaces);
		writer.WriteXmlAttributesEnd();
		for (let paraId in this.commentId) {
			if (this.commentId.hasOwnProperty(paraId)) {
				let elem = new CT_CommentId();
				elem.paraId = parseInt(paraId);
				elem.durableId = this.commentId[paraId];
				elem.toXml(writer, "w16cid:commentId");
			}
		}
		writer.WriteXmlNodeEnd(name);
	};

	function CT_CommentExtensible() {
		this.durableId = null;
		this.dateUtc = null;
		this.intelligentPlaceholder = null;
		this.extLst = [];
		return this;
	}

	CT_CommentExtensible.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "durableId": {
					this.durableId = AscCommon.FixDurableId(reader.GetValueUInt(this.durableId, 16));
					break;
				}
				case "dateUtc": {
					let dateMs = AscCommon.getTimeISO8601(reader.GetValueDecodeXml());
					if (!isNaN(dateMs)) {
						this.dateUtc = dateMs;
					}
					break;
				}
				case "intelligentPlaceholder": {
					this.intelligentPlaceholder = reader.GetValueBool();
					break;
				}
			}
		}
	};
	CT_CommentExtensible.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		let depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "extLst" : {
					break;
				}
			}
		}
	};
	CT_CommentExtensible.prototype.toXml = function(writer, name) {
		let dateUtc = this.dateUtc ? new Date(this.dateUtc).toISOString().slice(0, 19) + 'Z' : null;
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w16cex:durableId", AscCommon.Int32ToHexOrNull(this.durableId));
		writer.WriteXmlNullableAttributeString("w16cex:dateUtc", dateUtc);
		writer.WriteXmlNullableAttributeBool("w16cex:intelligentPlaceholder", this.intelligentPlaceholder);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlArray(this.extLst, "w16cex:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_CommentsExtensible() {
		this.commentExtensible = {};
		this.extLst = [];
		return this;
	}

	CT_CommentsExtensible.prototype.fromXml = function(reader) {
		let name;
		if (!reader.ReadNextNode()) {
			return;
		}
		name = reader.GetNameNoNS();
		if ("commentsExtensible" !== name) {
			if (!reader.ReadNextNode()) {
				return;
			}
		}
		name = reader.GetNameNoNS();
		if ("commentsExtensible" !== name) {
			return;
		}
		let elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "commentExtensible" : {
					elem = new CT_CommentExtensible();
					elem.fromXml(reader);
					this.commentExtensible[elem.durableId] = elem;
					break;
				}
				// case "extLst" : {
				// 	break;
				// }
			}
		}
	};
	CT_CommentsExtensible.prototype.toXml = function(writer) {
		let name = 'w16cex:commentsExtensible';

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlString(AscCommonWord.g_sXmlWordCommentsExtensibleNamespaces);
		writer.WriteXmlAttributesEnd();
		for (let durableId in this.commentExtensible) {
			if (this.commentExtensible.hasOwnProperty(durableId)) {
				this.commentExtensible[durableId].toXml(writer, "w16cid:commentId");
			}
		}
		// writer.WriteXmlArray(this.extLst, "w16cex:extLst");
		writer.WriteXmlNodeEnd(name);
	};
//revision

//misc
	function CT_TblPPr() {
		this.LeftFromText = null;
		this.RightFromText = null;
		this.TopFromText = null;
		this.BottomFromText = null;
		this.VertAnchor = null;
		this.HorzAnchor = null;
		this.TblpXSpec = null;
		this.TblpX = null;
		this.TblpYSpec = null;
		this.TblpY = null;
		return this;
	}

	CT_TblPPr.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "leftFromText": {
					this.LeftFromText =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
				case "rightFromText": {
					this.RightFromText =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
				case "topFromText": {
					this.TopFromText =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
				case "bottomFromText": {
					this.BottomFromText =
						AscCommon.universalMeasureToUnsignedMm(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_mm, 0);
					break;
				}
				case "vertAnchor": {
					this.VertAnchor = fromXml_ST_VAnchor(reader.GetValue());
					break;
				}
				case "horzAnchor": {
					this.HorzAnchor = fromXml_ST_HAnchor(reader.GetValue());
					break;
				}
				case "tblpXSpec": {
					this.TblpXSpec = fromXml_ST_XAlign(reader.GetValue());
					break;
				}
				case "tblpX": {
					var TblpX = AscCommon.universalMeasureToPt(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_pt,
						null);
					if (null !== TblpX) {
						this.TblpX = AscCommonWord.g_dKoef_pt_to_twips * TblpX;
					}
					break;
				}
				case "tblpYSpec": {
					this.TblpYSpec = fromXml_ST_YAlign(reader.GetValue());
					break;
				}
				case "tblpY": {
					var TblpY = AscCommon.universalMeasureToPt(reader.GetValue(), AscCommonWord.g_dKoef_twips_to_pt,
						null);
					if (null !== TblpY) {
						this.TblpY = AscCommonWord.g_dKoef_pt_to_twips * TblpY;
					}
					break;
				}
			}
		}
	};
	CT_TblPPr.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
		this.initDefaults();
	};
	CT_TblPPr.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:leftFromText", this.LeftFromText,
			AscCommonWord.g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:rightFromText", this.RightFromText,
			AscCommonWord.g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:topFromText", this.TopFromText,
			AscCommonWord.g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeUIntWithKoef("w:bottomFromText", this.BottomFromText,
			AscCommonWord.g_dKoef_mm_to_twips);
		writer.WriteXmlNullableAttributeString("w:vertAnchor", toXml_ST_VAnchor(this.VertAnchor));
		writer.WriteXmlNullableAttributeString("w:horzAnchor", toXml_ST_HAnchor(this.HorzAnchor));
		writer.WriteXmlNullableAttributeString("w:tblpXSpec", toXml_ST_XAlign(this.TblpXSpec));
		writer.WriteXmlNullableAttributeInt("w:tblpX", this.TblpX);
		writer.WriteXmlNullableAttributeString("w:tblpYSpec", toXml_ST_YAlign(this.TblpYSpec));
		writer.WriteXmlNullableAttributeInt("w:tblpY", this.TblpY);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_TblPPr.prototype.fromTable = function(table) {
		var PositionH = table.PositionH;
		if (PositionH) {
			this.HorzAnchor = PositionH.RelativeFrom;
			if (PositionH.Align) {
				this.TblpXSpec = PositionH.Value;
			} else {
				this.TblpX = table.Get_PositionHValueInTwips();
			}
		}
		var PositionV = table.PositionV;
		if (PositionV) {
			this.VertAnchor = PositionV.RelativeFrom;
			if (PositionV.Align) {
				this.TblpYSpec = PositionV.Value;
			} else {
				this.TblpY = table.Get_PositionVValueInTwips();
			}
		}
		if (table.Distance.L > 0) {
			this.LeftFromText = table.Distance.L;
		}
		if (table.Distance.T > 0) {
			this.TopFromText = table.Distance.T;
		}
		if (table.Distance.R > 0) {
			this.RightFromText = table.Distance.R;
		}
		if (table.Distance.B > 0) {
			this.BottomFromText = table.Distance.B;
		}
	};
	CT_TblPPr.prototype.toTable = function(table) {
		table.Set_Inline(false);
		if (null !== this.HorzAnchor) {
			if (null !== this.TblpX) {
				table.Set_PositionH(this.HorzAnchor, false, this.TblpX * AscCommonWord.g_dKoef_twips_to_mm);
			} else if (null !== this.TblpXSpec) {
				table.Set_PositionH(this.HorzAnchor, true, this.TblpXSpec);
			}
		}
		if (null !== this.VertAnchor) {
			if (null !== this.TblpY) {
				table.Set_PositionV(this.VertAnchor, false, this.TblpY * AscCommonWord.g_dKoef_twips_to_mm);
			} else if (null !== this.TblpYSpec) {
				table.Set_PositionV(this.VertAnchor, true, this.TblpYSpec);
			}
		}
		if (null !== this.LeftFromText || null !== this.TopFromText || null !== this.RightFromText ||
			null !== this.BottomFromText) {
			table.Set_Distance(this.LeftFromText || 0, this.TopFromText || 0, this.RightFromText || 0,
				this.BottomFromText || 0);
		}
	};
	CT_TblPPr.prototype.initDefaults = function() {
		if (null === this.HorzAnchor) {
			this.HorzAnchor = Asc.c_oAscHAnchor.Text;
		}
		if (null === this.VertAnchor) {
			this.VertAnchor = Asc.c_oAscVAnchor.Margin;
		}
		if (null === this.TblpXSpec) {
			if (null === this.TblpX) {
				this.TblpX = 0;
			}
			//Several values of sprmTDxaAbs have special meanings as specified by
			//[ECMA-376] Part 4, Section 2.18.114. These values are specified as
			//follows.
			switch (this.TblpX) {
				case 0:
					this.TblpXSpec = Asc.c_oAscXAlign.Left;
					this.TblpX = null;
					break;
				case -4:
					this.TblpXSpec = Asc.c_oAscXAlign.Center;
					this.TblpX = null;
					break;
				case -8:
					this.TblpXSpec = Asc.c_oAscXAlign.Right;
					this.TblpX = null;
					break;
				case -12:
					this.TblpXSpec = Asc.c_oAscXAlign.Inside;
					this.TblpX = null;
					break;
				case -16:
					this.TblpXSpec = Asc.c_oAscXAlign.Outside;
					this.TblpX = null;
					break;
			}
		}
		if (null === this.TblpYSpec) {
			if (null === this.TblpY) {
				this.TblpY = 0;
			}
			//Several values of sprmTDxaAbs have special meanings as specified by
			//[ECMA-376] Part 4, Section 2.18.114. These values are specified as
			//follows.
			switch (this.TblpY) {
				case 0:
					this.VertAnchor = Asc.c_oAscVAnchor.Text;
					break;
				case -4:
					this.TblpYSpec = Asc.c_oAscYAlign.Top;
					this.TblpY = null;
					break;
				case -8:
					this.TblpYSpec = Asc.c_oAscYAlign.Center;
					this.TblpY = null;
					break;
				case -12:
					this.TblpYSpec = Asc.c_oAscYAlign.Bottom;
					this.TblpY = null;
					break;
				case -16:
					this.TblpYSpec = Asc.c_oAscYAlign.Inside;
					this.TblpY = null;
					break;
				case -16:
					this.TblpYSpec = Asc.c_oAscYAlign.Outside;
					this.TblpY = null;
					break;
			}
		}
	};
	function CT_TblCellMar() {
		this.Top = null;
		this.Start = null;
		this.Bottom = null;
		this.End = null;
		return this;
	}

	CT_TblCellMar.prototype.fromXml = function(reader) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "top" : {
					this.Top = new CTableMeasurement(tblwidth_Auto, 0);
					this.Top.fromXml(reader);
					break;
				}
				case "left" :
				case "start" : {
					this.Left = new CTableMeasurement(tblwidth_Auto, 0);
					this.Left.fromXml(reader);
					break;
				}
				case "bottom" : {
					this.Bottom = new CTableMeasurement(tblwidth_Auto, 0);
					this.Bottom.fromXml(reader);
					break;
				}
				case "right" :
				case "end" : {
					this.Right = new CTableMeasurement(tblwidth_Auto, 0);
					this.Right.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_TblCellMar.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Top, "w:top");
		writer.WriteXmlNullable(this.Left, "w:left");
		writer.WriteXmlNullable(this.Bottom, "w:bottom");
		writer.WriteXmlNullable(this.Right, "w:right");
		writer.WriteXmlNodeEnd(name);
	};
	CT_TblCellMar.prototype.fromObj = function(obj) {
		for (var elem in obj) {
			if (obj.hasOwnProperty(elem) && obj[elem]) {
				this[elem] = obj[elem];
			}
		}
	};
	CT_TblCellMar.prototype.toObj = function(obj) {
		for (var elem in this) {
			if (this.hasOwnProperty(elem) && this[elem]) {
				obj[elem] = this[elem];
			}
		}
	};
	CT_TblCellMar.prototype.isEmpty = function() {
		for (var elem in this) {
			if (this.hasOwnProperty(elem) && this[elem]) {
				return false;
			}
		}
		return true;
	};
	function CT_Bdr() {
		this.Top = null;
		this.Left = null;
		this.Bottom = null;
		this.Right = null;
		this.Between = null;
		this.Bar = null;
		this.InsideH = null;
		this.InsideV = null;
		return this;
	}

	CT_Bdr.prototype.fromXml = function(reader) {
		var depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "top" : {
					this.Top = new CDocumentBorder();
					this.Top.fromXml(reader);
					break;
				}
				case "left" : {
					this.Left = new CDocumentBorder();
					this.Left.fromXml(reader);
					break;
				}
				case "bottom" : {
					this.Bottom = new CDocumentBorder();
					this.Bottom.fromXml(reader);
					break;
				}
				case "right" : {
					this.Right = new CDocumentBorder();
					this.Right.fromXml(reader);
					break;
				}
				case "between" : {
					this.Between = new CDocumentBorder();
					this.Between.fromXml(reader);
					break;
				}
				case "bar" : {
					this.Bar = new CDocumentBorder();
					this.Bar.fromXml(reader);
					break;
				}
				case "insideH" : {
					this.InsideH = new CDocumentBorder();
					this.InsideH.fromXml(reader);
					break;
				}
				case "insideV" : {
					this.InsideV = new CDocumentBorder();
					this.InsideV.fromXml(reader);
					break;
				}
			}
		}
	};
	CT_Bdr.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.Top, "w:top");
		writer.WriteXmlNullable(this.Left, "w:left");
		writer.WriteXmlNullable(this.Bottom, "w:bottom");
		writer.WriteXmlNullable(this.Right, "w:right");
		writer.WriteXmlNullable(this.Between, "w:between");
		writer.WriteXmlNullable(this.Bar, "w:bar");
		writer.WriteXmlNullable(this.InsideH, "w:insideH");
		writer.WriteXmlNullable(this.InsideV, "w:insideV");
		writer.WriteXmlNodeEnd(name);
	};
	CT_Bdr.prototype.fromObj = function(obj) {
		for (var elem in obj) {
			if (obj.hasOwnProperty(elem) && obj[elem]) {
				this[elem] = obj[elem];
			}
		}
	};
	CT_Bdr.prototype.toObj = function(obj) {
		for (var elem in this) {
			if (this.hasOwnProperty(elem) && this[elem]) {
				obj[elem] = this[elem];
			}
		}
	};
	CT_Bdr.prototype.isEmpty = function() {
		for (var elem in this) {
			if (this.hasOwnProperty(elem) && this[elem]) {
				return false;
			}
		}
		return true;
	};
	function CT_Color(xmlVal, xmlThemeColor, xmlThemeTint, xmlThemeShade) {
		this.xmlVal = xmlVal;
		this.xmlThemeColor = xmlThemeColor;
		this.xmlThemeTint = xmlThemeTint;
		this.xmlThemeShade = xmlThemeShade;

		this.Val = null;
		this.ThemeColor = null;
		this.ThemeTint = null;
		this.ThemeShade = null;
		return this;
	}

	CT_Color.prototype.readAttrElem = function(reader, name) {
		if (this.xmlVal === name) {
			this.Val = reader.GetValueDecodeXml();
		} else if (this.xmlThemeColor === name) {
			this.ThemeColor = fromXml_ST_ThemeColor(reader.GetValue());
		} else if (this.xmlThemeTint === name) {
			this.ThemeTint = reader.GetValueByte(this.ThemeTint, 16);
		} else if (this.xmlThemeShade === name) {
			this.ThemeShade = reader.GetValueByte(this.ThemeTint, 16);
		} else {
			return false;
		}
	};
	CT_Color.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			this.readAttrElem(reader, name);
		}
	};
	CT_Color.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_Color.prototype.toXmlElems = function(writer) {
		var ThemeTint = null !== this.ThemeTint ? AscCommon.ByteToHex(this.ThemeTint) : null;
		var ThemeShade = null !== this.ThemeShade ? AscCommon.ByteToHex(this.ThemeShade) : null;
		writer.WriteXmlNullableAttributeStringEncode("w:" + this.xmlVal, this.Val);
		writer.WriteXmlNullableAttributeString("w:" + this.xmlThemeColor, toXml_ST_ThemeColor(this.ThemeColor));
		writer.WriteXmlNullableAttributeString("w:" + this.xmlThemeTint, ThemeTint);
		writer.WriteXmlNullableAttributeString("w:" + this.xmlThemeShade, ThemeShade);
	};
	CT_Color.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		this.toXmlElems(writer);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_Color.prototype.isEmpty = function() {
		return null === this.Val && null === this.ThemeColor && null === this.ThemeTint && null === this.ThemeShade;
	};
	CT_Color.prototype.getColor = function(defR, defG, defB) {
		if (this.Val) {
			var color = new CDocumentColor(defR, defG, defB);
			color.SetFromHexColor(this.Val);
			return color;
		}
		return undefined;
	};
	CT_Color.prototype.setColor = function(val) {
		this.Val = val ? val.ToHexColor() : null;
	};
	CT_Color.prototype.getUnifill = function() {
		var unifill = AscCommonWord.CreateThemeUnifill(this.ThemeColor, this.ThemeTint, this.ThemeShade);
		return null != unifill ? unifill : undefined;
	};
	CT_Color.prototype.setUniFill = function(val) {
		var obj = AscCommonWord.CreateFromThemeUnifill(val);
		this.ThemeColor = obj.Color;
		this.ThemeTint = obj.Tint;
		this.ThemeShade = obj.Shade;
	};
	function CT_Underline() {
		this.Val = undefined;
		this.Color = new CT_Color("color", "themeColor", "themeTint", "themeShade");
		return this;
	}

	CT_Underline.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			var name = reader.GetNameNoNS();
			switch (name) {
				case "val": {
					this.Val = fromXml_ST_Underline(reader.GetValue());
					break;
				}
				default:
					this.Color.readAttrElem(reader, name);
					break;
			}
		}
	};
	CT_Underline.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_Underline.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:val", toXml_ST_Underline(this.Val));
		this.Color.toXmlElems(writer);
		writer.WriteXmlAttributesEnd(true);
	};
	function CT_HdrFtrRef() {
		this.Id = null;
		this.Type = null;
		return this;
	}
	CT_HdrFtrRef.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					this.Id = reader.GetValueDecodeXml();
					break;
				}
				case "type": {
					this.Type = reader.GetValue();
					break;
				}
			}
		}
	};
	CT_HdrFtrRef.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_HdrFtrRef.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("r:id", this.Id);
		writer.WriteXmlNullableAttributeString("w:type", this.Type);
		writer.WriteXmlAttributesEnd(true);
	};

	function CT_Inline(drawing) {
		this.drawing = drawing;

		this.DistT = null;
		this.DistB = null;
		this.DistL = null;
		this.DistR = null;
		this.Extent = null;
		this.EffectExtent = null;
		this.DocPr = null;
		this.CNvGraphicFramePr = null;
//todo Graphic
		this.Graphic = null;
		return this;
	}
	CT_Inline.prototype.readAttr = function(reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "distT": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null, null);
					}
					break;
				}
				case "distB": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, null, null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val));
					}
					break;
				}
				case "distL": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null, null, null);
					}
					break;
				}
				case "distR": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null);
					}
					break;
				}
			}
		}
	};
	CT_Inline.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "extent" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var cx = parseInt(elem.attributes["cx"]);
					var cy = parseInt(elem.attributes["cy"]);
					cx = !isNaN(cx) ? AscCommonWord.g_dKoef_emu_to_mm * cx : null;
					cy = !isNaN(cy) ? AscCommonWord.g_dKoef_emu_to_mm * cy : null;
					this.drawing.setExtent(cx, cy);
					break;
				}
				case "effectExtent" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					let L = parseInt(elem.attributes["l"]);
					let T = parseInt(elem.attributes["t"]);
					let R = parseInt(elem.attributes["r"]);
					let B = parseInt(elem.attributes["b"]);
					L = !isNaN(L) ? AscCommonWord.g_dKoef_emu_to_mm * L : null;
					T = !isNaN(T) ? AscCommonWord.g_dKoef_emu_to_mm * T : null;
					R = !isNaN(R) ? AscCommonWord.g_dKoef_emu_to_mm * R : null;
					B = !isNaN(B) ? AscCommonWord.g_dKoef_emu_to_mm * B : null;
					this.drawing.setEffectExtent(L, T, R, B);
					break;
				}
				case "docPr" : {
					this.drawing.docPr.fromXml(reader);
					if(this.drawing.docPr.form !== null) {
						this.drawing.SetForm(this.drawing.docPr.form);
						this.drawing.docPr.form = null;
					}
					break;
				}
				case "cNvGraphicFramePr" : {
					elem = new CT_NonVisualGraphicFrameProperties();
					elem.fromXml(reader);
					break;
				}
				case "graphic" : {
					var graphic = new AscFormat.CT_GraphicalObject();
					graphic.fromXml(reader);
					let graphicObject = graphic.GraphicData && graphic.GraphicData.graphicObject;
					if (graphicObject) {
						//todo init in graphic.fromXml
						graphicObject.setBDeleted(false);
						graphicObject.setParent(this.drawing);
						this.drawing.Set_GraphicObject(graphicObject);
					}
					break;
				}
			}
		}
	};
	CT_Inline.prototype.toXml = function(writer, name) {
		var drawing = this.drawing;
		var Extent = new CT_XmlNode();
		Extent.attributes["cx"] = Math.round(drawing.Extent.W * g_dKoef_mm_to_emu);
		Extent.attributes["cy"] = Math.round(drawing.Extent.H * g_dKoef_mm_to_emu);

		var EffectExtent = new CT_XmlNode();
		EffectExtent.attributes["l"] = Math.round(drawing.EffectExtent.L * g_dKoef_mm_to_emu);
		EffectExtent.attributes["t"] = Math.round(drawing.EffectExtent.T * g_dKoef_mm_to_emu);
		EffectExtent.attributes["r"] = Math.round(drawing.EffectExtent.R * g_dKoef_mm_to_emu);
		EffectExtent.attributes["b"] = Math.round(drawing.EffectExtent.B * g_dKoef_mm_to_emu);
		var Graphic, nvGraphicFramePr;
		if (drawing.GraphicObj) {
			Graphic = new AscFormat.CT_GraphicalObject();
			Graphic.Namespace = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
			Graphic.GraphicData = new AscFormat.CT_GraphicalObjectData();
			Graphic.GraphicData.Uri = GetDrawingUri(drawing.GraphicObj);
			Graphic.GraphicData.graphicObject = drawing.GraphicObj;

			nvGraphicFramePr = new CT_NonVisualGraphicFrameProperties();
			nvGraphicFramePr.setLocks(drawing.GraphicObj.locks);
		}

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUIntWithKoef("distT", drawing.Distance.T, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distB", drawing.Distance.B, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distL", drawing.Distance.L, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distR", drawing.Distance.R, g_dKoef_mm_to_emu);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(Extent, "wp:extent");
		writer.WriteXmlNullable(EffectExtent, "wp:effectExtent");
		if(drawing.docPr && drawing.Form) {
			drawing.docPr.form = drawing.Form;
		}
		writer.WriteXmlNullable(drawing.docPr, "wp:docPr");
		writer.WriteXmlNullable(nvGraphicFramePr, "wp:cNvGraphicFramePr");
		writer.WriteXmlNullable(Graphic, "a:graphic");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_Anchor(drawing) {
		this.drawing = drawing;

		this.DistT = null;
		this.DistB = null;
		this.DistL = null;
		this.DistR = null;
		this.SimplePos = null;
		this.RelativeHeight = null;
		this.BehindDoc = null;
		this.Locked = null;
		this.LayoutInCell = null;
		this.Hidden = null;
		this.AllowOverlap = null;
		this.SimplePos = null;
		this.PositionH = null;
		this.PositionV = null;
		this.Extent = null;
		this.EffectExtent = null;
//todo Item
		this.Item = null;
		this.DocPr = null;
		this.CNvGraphicFramePr = null;
//todo Graphic
		this.Graphic = null;
		return this;
	}
	CT_Anchor.prototype.readAttr = function(reader) {
		var val;
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "distT": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null, null);
					}
					break;
				}
				case "distB": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, null, null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val));
					}
					break;
				}
				case "distL": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null, null, null);
					}
					break;
				}
				case "distR": {
					val = reader.GetValueUInt64();
					if (undefined !== val) {
						this.drawing.Set_Distance(null, null, Math.abs(AscCommonWord.g_dKoef_emu_to_mm * val), null);
					}
					break;
				}
				case "simplePos": {
					this.drawing.SimplePos.Use = reader.GetValueBool();
					break;
				}
				case "relativeHeight": {
					this.drawing.Set_RelativeHeight(reader.GetValueUInt(this.RelativeHeight));
					reader.context.checkZIndex(this.RelativeHeight);
					break;
				}
				case "behindDoc": {
					this.drawing.Set_BehindDoc(reader.GetValueBool());
					break;
				}
				case "locked": {
					this.drawing.Set_Locked(reader.GetValueBool());
					break;
				}
				case "layoutInCell": {
					this.drawing.Set_LayoutInCell(reader.GetValueBool());
					break;
				}
				// case "hidden": {
				// 	this.Hidden = reader.GetValueBool();
				// 	break;
				// }
				case "allowOverlap": {
					this.drawing.Set_AllowOverlap(reader.GetValueBool());
					break;
				}
			}
		}
	};
	CT_Anchor.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var drawing = this.drawing;
		var elem, align, posOffset, nvGraphicFramePr, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "simplePos" : {
					elem = new CPolygonPoint();
					elem.x = drawing.SimplePos.X;
					elem.y = drawing.SimplePos.Y;
					elem.fromXml(reader);
					drawing.SimplePos.X = elem.x;
					drawing.SimplePos.Y = elem.y;
					break;
				}
				case "positionH" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var PosH = drawing.PositionH;
					PosH.RelativeFrom = fromXml_ST_RelFromH(elem.attributes["relativeFrom"], PosH.RelativeFrom);
					align = fromXml_ST_AlignH(elem.members["align"] && elem.members["align"].text);
					posOffset = parseInt(elem.members["posOffset"] && elem.members["posOffset"].text);
					//todo percent
					var pctPosHOffset = parseFloat(elem.members["pctPosHOffset"] && elem.members["pctPosHOffset"].text);
					if (undefined !== align) {
						PosH.Align = true;
						PosH.Value = align;
					} else if (!isNaN(posOffset)) {
						PosH.Align = false;
						PosH.Value = AscCommonWord.g_dKoef_emu_to_mm * posOffset;
					} else if (!isNaN(pctPosHOffset)) {
						PosH.Percent = true;
						PosH.Value = pctPosHOffset;
					}
					drawing.Set_PositionH(PosH.RelativeFrom , PosH.Align , PosH.Value, PosH.Percent);
					break;
				}
				case "positionV" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var PosV = drawing.PositionV;
					PosV.RelativeFrom = fromXml_ST_RelFromV(elem.attributes["relativeFrom"], PosV.RelativeFrom);
					align = fromXml_ST_AlignV(elem.members["align"] && elem.members["align"].text);
					posOffset = parseInt(elem.members["posOffset"] && elem.members["posOffset"].text);
					var pctPosVOffset = parseFloat(elem.members["pctPosVOffset"] && elem.members["pctPosVOffset"].text);
					if (undefined !== align) {
						PosV.Align = true;
						PosV.Value = align;
					} else if (!isNaN(posOffset)) {
						PosV.Align = false;
						PosV.Value = AscCommonWord.g_dKoef_emu_to_mm * posOffset;
					} else if (!isNaN(pctPosVOffset)) {
						PosV.Percent = true;
						PosV.Value = pctPosVOffset;
					}
					drawing.Set_PositionV(PosV.RelativeFrom , PosV.Align , PosV.Value, PosV.Percent);
					break;
				}
				case "extent" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var cx = parseInt(elem.attributes["cx"]);
					var cy = parseInt(elem.attributes["cy"]);
					if(!isNaN(cx)) {
						drawing.Extent.W = AscCommonWord.g_dKoef_emu_to_mm * cx;
					}
					if(!isNaN(cy)) {
						drawing.Extent.H = AscCommonWord.g_dKoef_emu_to_mm * cy;
					}
					break;
				}
				case "effectExtent" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					let L = parseInt(elem.attributes["l"]);
					let T = parseInt(elem.attributes["t"]);
					let R = parseInt(elem.attributes["r"]);
					let B = parseInt(elem.attributes["b"]);
					L = !isNaN(L) ? AscCommonWord.g_dKoef_emu_to_mm * L : null;
					T = !isNaN(T) ? AscCommonWord.g_dKoef_emu_to_mm * T : null;
					R = !isNaN(R) ? AscCommonWord.g_dKoef_emu_to_mm * R : null;
					B = !isNaN(B) ? AscCommonWord.g_dKoef_emu_to_mm * B : null;
					this.drawing.setEffectExtent(L, T, R, B);
					break;
				}
				case "wrapNone" : {
					drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
					break;
				}
				case "wrapSquare" : {
					drawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
					break;
				}
				case "wrapTight" : {
					drawing.Set_WrappingType(WRAPPING_TYPE_TIGHT);
					elem = new CT_XmlNode(function(reader, name) {
						if ("wrapPolygon" === name) {
							drawing.wrappingPolygon.fromXml(reader);
							return drawing.wrappingPolygon;
						}
					});
					elem.fromXml(reader);
					break;
				}
				case "wrapThrough" : {
					drawing.Set_WrappingType(WRAPPING_TYPE_THROUGH);
					elem = new CT_XmlNode(function(reader, name) {
						if ("wrapPolygon" === name) {
							drawing.wrappingPolygon.fromXml(reader);
							return drawing.wrappingPolygon;
						}
					});
					elem.fromXml(reader);
					break;
				}
				case "wrapTopAndBottom" : {
					drawing.Set_WrappingType(WRAPPING_TYPE_TOP_AND_BOTTOM);
					break;
				}
				case "docPr" : {
					drawing.docPr.fromXml(reader);
					if(this.drawing.docPr.form !== null) {
						this.drawing.SetForm(this.drawing.docPr.form);
						this.drawing.docPr.form = null;
					}
					break;
				}
				case "cNvGraphicFramePr" : {
					nvGraphicFramePr = new CT_NonVisualGraphicFrameProperties();
					nvGraphicFramePr.fromXml(reader);
					break;
				}
				case "graphic" : {
					var graphic = new AscFormat.CT_GraphicalObject();
					graphic.fromXml(reader);
					let graphicObject = graphic.GraphicData && graphic.GraphicData.graphicObject;
					if (graphicObject) {
						//todo init in graphic.fromXml
						graphicObject.setBDeleted(false);
						graphicObject.setParent(drawing);
						drawing.Set_GraphicObject(graphicObject);
					}
					break;
				}
				case "sizeRelH" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var sizeRelH = {RelativeFrom: c_oAscRelativeFromV.Page, Percent: 0};//Percent 0-1
					sizeRelH.RelativeFrom = fromXml_ST_SizeRelFromH(elem.attributes["relativeFrom"], sizeRelH.RelativeFrom);
					sizeRelH.Percent = parseFloat(elem.members["pctWidth"] && elem.members["pctWidth"].text) / 100000;
					drawing.SetSizeRelH(sizeRelH);
					break;
				}
				case "sizeRelV" : {
					elem = new CT_XmlNode();
					elem.fromXml(reader);
					var sizeRelV = {RelativeFrom: c_oAscRelativeFromV.Page, Percent: 0};//Percent 0-1
					sizeRelV.RelativeFrom = fromXml_ST_SizeRelFromV(elem.attributes["relativeFrom"], sizeRelV.RelativeFrom);
					sizeRelV.Percent = parseFloat(elem.members["pctHeight"] && elem.members["pctHeight"].text) / 100000;
					drawing.SetSizeRelV(sizeRelV);
					break;
				}
			}
		}
		if(drawing.GraphicObj && nvGraphicFramePr) {
			drawing.GraphicObj.setLocks(nvGraphicFramePr.getLocks());
		}
	};
	CT_Anchor.prototype.toXml = function(writer, name) {
		var drawing = this.drawing;
		var SimplePos = new CPolygonPoint();
		SimplePos.x = drawing.SimplePos.X;
		SimplePos.y = drawing.SimplePos.Y;

		var PositionH = new CT_XmlNode();
		PositionH.attributes["relativeFrom"] = toXml_ST_RelFromH(drawing.PositionH.RelativeFrom);
		if (true == drawing.PositionH.Align) {
			PositionH.members["wp:align"] = new CT_XmlNode();
			PositionH.members["wp:align"].text = toXml_ST_AlignH(drawing.PositionH.Value);
		}
		else if (true == PositionH.Percent) {
			PositionH.members["wp14:pctPosHOffset"] = new CT_XmlNode();
			PositionH.members["wp14:pctPosHOffset"].text = drawing.PositionH.Value;
		} else {
			PositionH.members["wp:posOffset"] = new CT_XmlNode();
			PositionH.members["wp:posOffset"].text = Math.round(g_dKoef_mm_to_emu * drawing.PositionH.Value);
		}
		var PositionV = new CT_XmlNode();
		PositionV.attributes["relativeFrom"] = toXml_ST_RelFromV(drawing.PositionV.RelativeFrom);
		if (true == drawing.PositionV.Align) {
			PositionV.members["wp:align"] = new CT_XmlNode();
			PositionV.members["wp:align"].text = toXml_ST_AlignV(drawing.PositionV.Value);
		}
		else if (true == PositionV.Percent) {
			PositionV.members["wp14:pctPosVOffset"] = new CT_XmlNode();
			PositionV.members["wp14:pctPosVOffset"].text = drawing.PositionV.Value;
		} else {
			PositionV.members["wp:posOffset"] = new CT_XmlNode();
			PositionV.members["wp:posOffset"].text = Math.round(g_dKoef_mm_to_emu * drawing.PositionV.Value);
		}
		var Extent = new CT_XmlNode();
		Extent.attributes["cx"] = Math.round(drawing.Extent.W * g_dKoef_mm_to_emu);
		Extent.attributes["cy"] = Math.round(drawing.Extent.H * g_dKoef_mm_to_emu);

		var EffectExtent = new CT_XmlNode();
		EffectExtent.attributes["l"] = Math.round(drawing.EffectExtent.L * g_dKoef_mm_to_emu);
		EffectExtent.attributes["t"] = Math.round(drawing.EffectExtent.T * g_dKoef_mm_to_emu);
		EffectExtent.attributes["r"] = Math.round(drawing.EffectExtent.R * g_dKoef_mm_to_emu);
		EffectExtent.attributes["b"] = Math.round(drawing.EffectExtent.B * g_dKoef_mm_to_emu);
		var WrapNone, WrapSquare, WrapTight, WrapThrough, WrapTopAndBottom;
		switch(drawing.wrappingType) {
			case WRAPPING_TYPE_NONE:
				WrapNone = new CT_XmlNode();
				break;
			case WRAPPING_TYPE_SQUARE:
				WrapSquare = new CT_XmlNode();
				WrapSquare.attributes["wrapText"] = "bothSides";
				break;
			case WRAPPING_TYPE_TIGHT:
				WrapTight = new CT_XmlNode();
				WrapTight.attributes["wrapText"] = "bothSides";
				WrapTight.members["wp:wrapPolygon"] = drawing.wrappingPolygon;
				break;
			case WRAPPING_TYPE_THROUGH:
				WrapThrough = new CT_XmlNode();
				WrapThrough.attributes["wrapText"] = "bothSides";
				WrapThrough.members["wp:wrapPolygon"] = drawing.wrappingPolygon;
				break;
			case WRAPPING_TYPE_TOP_AND_BOTTOM:
				WrapTopAndBottom = new CT_XmlNode();
				break;
		}
		var Graphic, nvGraphicFramePr;
		if (drawing.GraphicObj) {
			Graphic = new AscFormat.CT_GraphicalObject();
			Graphic.Namespace = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
			Graphic.GraphicData = new AscFormat.CT_GraphicalObjectData();
			Graphic.GraphicData.Uri = GetDrawingUri(drawing.GraphicObj);
			Graphic.GraphicData.graphicObject = drawing.GraphicObj;

			nvGraphicFramePr = new CT_NonVisualGraphicFrameProperties();
			nvGraphicFramePr.setLocks(drawing.GraphicObj.locks);
		}
		var SizeRelH;
		if(drawing.SizeRelH) {
			SizeRelH = new CT_XmlNode();
			SizeRelH.attributes["relativeFrom"] = toXml_ST_SizeRelFromH(drawing.SizeRelH.RelativeFrom);
			SizeRelH.members["wp14:pctWidth"] = new CT_XmlNode();
			SizeRelH.members["wp14:pctWidth"].text = (drawing.SizeRelH.Percent * 100000 + 0.5) >> 0;
		}
		var SizeRelV;
		if(drawing.SizeRelV) {
			SizeRelV = new CT_XmlNode();
			SizeRelV.attributes["relativeFrom"] = toXml_ST_SizeRelFromV(drawing.SizeRelV.RelativeFrom);
			SizeRelV.members["wp14:pctHeight"] = new CT_XmlNode();
			SizeRelV.members["wp14:pctHeight"].text = (drawing.SizeRelV.Percent * 100000 + 0.5) >> 0;
		}

		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeUIntWithKoef("distT", drawing.Distance.T, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distB", drawing.Distance.B, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distL", drawing.Distance.L, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeUIntWithKoef("distR", drawing.Distance.R, g_dKoef_mm_to_emu);
		writer.WriteXmlNullableAttributeBool("simplePos", drawing.SimplePos.Use);
		writer.WriteXmlNullableAttributeUInt("relativeHeight", drawing.RelativeHeight);
		writer.WriteXmlNullableAttributeBool("behindDoc", drawing.behindDoc);
		writer.WriteXmlNullableAttributeBool("locked", drawing.Locked);
		writer.WriteXmlNullableAttributeBool("layoutInCell", drawing.LayoutInCell);
		// writer.WriteXmlNullableAttributeBool("hidden", drawing.Hidden);
		writer.WriteXmlNullableAttributeBool("allowOverlap", drawing.AllowOverlap);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(SimplePos, "wp:simplePos");
		writer.WriteXmlNullable(PositionH, "wp:positionH");
		writer.WriteXmlNullable(PositionV, "wp:positionV");
		writer.WriteXmlNullable(Extent, "wp:extent");
		writer.WriteXmlNullable(EffectExtent, "wp:effectExtent");
		writer.WriteXmlNullable(WrapNone, "wp:wrapNone");
		writer.WriteXmlNullable(WrapSquare, "wp:wrapSquare");
		writer.WriteXmlNullable(WrapTight, "wp:wrapTight");
		writer.WriteXmlNullable(WrapThrough, "wp:wrapThrough");
		writer.WriteXmlNullable(WrapTopAndBottom, "wp:wrapTopAndBottom");
		if(drawing.docPr && drawing.Form) {
			drawing.docPr.form = drawing.Form;
		}
		writer.WriteXmlNullable(drawing.docPr, "wp:docPr");
		writer.WriteXmlNullable(nvGraphicFramePr, "wp:cNvGraphicFramePr");
		writer.WriteXmlNullable(Graphic, "a:graphic");
		writer.WriteXmlNullable(SizeRelH, "wp14:sizeRelH");
		writer.WriteXmlNullable(SizeRelV, "wp14:sizeRelV");
		writer.WriteXmlNodeEnd(name);
	};

	function GetDrawingUri(oDrawing) {
		if (!oDrawing) {
			return "";
		}
		else if (oDrawing instanceof AscFormat.CLockedCanvas) {
			return "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
		}
		else if (oDrawing instanceof AscFormat.SmartArt) {
			return "http://schemas.openxmlformats.org/drawingml/2006/diagram";
		}
		else if (oDrawing instanceof AscFormat.CGroupShape) {
			return "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
		}
		else if (oDrawing instanceof AscFormat.CImageShape) {
			return "http://schemas.openxmlformats.org/drawingml/2006/picture";
		}
		else if (oDrawing instanceof AscFormat.CChartSpace) {
			return "http://schemas.openxmlformats.org/drawingml/2006/chart";
		}
		else {
			return "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
		}
	}
	function CT_PosH() {
		this.RelativeFrom = null;
//todo Item
		this.Item = null;
		return this;
	}
	CT_PosH.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "relativeFrom": {
					this.RelativeFrom = fromXml_ST_RelFromH(reader.GetValue());
					break;
				}
			}
		}
	};
	CT_PosH.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "Item" : {
//todo Item
					break;
				}
			}
		}
	};
	CT_PosH.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:relativeFrom", toXml_ST_RelFromH(this.RelativeFrom));
		writer.WriteXmlAttributesEnd();
//todo Item
		writer.WriteXmlNodeEnd(name);
	};
	function CT_NonVisualGraphicFrameProperties() {
		this.GraphicFrameLocks = null;
		this.ExtLst = null;
		return this;
	}
	CT_NonVisualGraphicFrameProperties.prototype.fromXml = function(reader) {
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			switch (reader.GetNameNoNS()) {
				case "graphicFrameLocks" : {
					this.GraphicFrameLocks = new CT_GraphicalObjectFrameLocking();
					this.GraphicFrameLocks.fromXml(reader);
					break;
				}
				// case "extLst" : {
				// 	this.ExtLst = new CT_OfficeArtExtensionList();
				// 	this.ExtLst.fromXml(reader);
				// 	break;
				// }
			}
		}
	};
	CT_NonVisualGraphicFrameProperties.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlAttributesEnd();
		writer.WriteXmlNullable(this.GraphicFrameLocks, "a:graphicFrameLocks");
		// writer.WriteXmlNullable(this.ExtLst, "w:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	CT_NonVisualGraphicFrameProperties.prototype.getLocks = function() {
		var locks = 0;
		if (this.GraphicFrameLocks) {
			var elem = this.GraphicFrameLocks;
			locks |= (AscFormat.LOCKS_MASKS.noGrp | (elem.NoGrp ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
			locks |= (AscFormat.LOCKS_MASKS.noDrilldown | (elem.NoDrilldown ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
			locks |= (AscFormat.LOCKS_MASKS.noSelect | (elem.NoSelect ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
			locks |= (AscFormat.LOCKS_MASKS.noChangeAspect | (elem.NoChangeAspect ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
			locks |= (AscFormat.LOCKS_MASKS.noMove | (elem.NoMove ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
			locks |= (AscFormat.LOCKS_MASKS.noResize | (elem.NoResize ? AscFormat.LOCKS_MASKS.noChangeAspect << 1 : 0));
		}
		return locks;
	};
	CT_NonVisualGraphicFrameProperties.prototype.setLocks = function(locks) {
		if (locks > 0) {
			var elem = new CT_GraphicalObjectFrameLocking();
			elem.noGrp = !!(locks & AscFormat.LOCKS_MASKS.NoGrp << 1);
			elem.noDrilldown = !!(locks & AscFormat.LOCKS_MASKS.NoDrilldown << 1);
			elem.noSelect = !!(locks & AscFormat.LOCKS_MASKS.NoSelect << 1);
			elem.noChangeAspect = !!(locks & AscFormat.LOCKS_MASKS.NoChangeAspect << 1);
			elem.noMove = !!(locks & AscFormat.LOCKS_MASKS.NoMove << 1);
			elem.noResize = !!(locks & AscFormat.LOCKS_MASKS.NoResize << 1);
			this.GraphicFrameLocks = elem;
		}
	};
	function CT_GraphicalObjectFrameLocking() {
		this.NoGrp = null;//False
		this.NoDrilldown = null;//False
		this.NoSelect = null;//False
		this.NoChangeAspect = null;//False
		this.NoMove = null;//False
		this.NoResize = null;//False
		this.ExtLst = null;
		return this;
	}
	CT_GraphicalObjectFrameLocking.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "noGrp": {
					this.NoGrp = reader.GetValueBool();
					break;
				}
				case "noDrilldown": {
					this.NoDrilldown = reader.GetValueBool();
					break;
				}
				case "noSelect": {
					this.NoSelect = reader.GetValueBool();
					break;
				}
				case "noChangeAspect": {
					this.NoChangeAspect = reader.GetValueBool();
					break;
				}
				case "noMove": {
					this.NoMove = reader.GetValueBool();
					break;
				}
				case "noResize": {
					this.NoResize = reader.GetValueBool();
					break;
				}
			}
		}
	};
	CT_GraphicalObjectFrameLocking.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		var elem, depth = reader.GetDepth();
		while (reader.ReadNextSiblingNode(depth)) {
			// switch (reader.GetNameNoNS()) {
			// 	case "extLst" : {
			// 		this.ExtLst = new CT_OfficeArtExtensionList();
			// 		this.ExtLst.fromXml(reader);
			// 		break;
			// 	}
			// }
		}
	};
	CT_GraphicalObjectFrameLocking.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("noGrp", this.NoGrp);
		writer.WriteXmlNullableAttributeBool("noDrilldown", this.NoDrilldown);
		writer.WriteXmlNullableAttributeBool("noSelect", this.NoSelect);
		writer.WriteXmlNullableAttributeBool("noChangeAspect", this.NoChangeAspect);
		writer.WriteXmlNullableAttributeBool("noMove", this.NoMove);
		writer.WriteXmlNullableAttributeBool("noResize", this.NoResize);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlNullable(this.ExtLst, "w:extLst");
		writer.WriteXmlNodeEnd(name);
	};
	function CT_FtnEdnRef() {
		this.customMarkFollows = undefined;
		this.id = undefined;
	}
	CT_FtnEdnRef.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "customMarkFollows": {
					this.customMarkFollows = reader.GetValueBool();
					break;
				}
				case "id": {
					this.id = reader.GetValueInt(this.id);
					break;
				}
			}
		}
	};
	CT_FtnEdnRef.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_FtnEdnRef.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeBool("w:customMarkFollows", this.customMarkFollows);
		writer.WriteXmlNullableAttributeInt("w:id", this.id);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_FtnEdnRef.prototype.fromVal = function(index, CustomMark) {
		this.customMarkFollows = CustomMark;
		this.id = index;
	};
	function CT_CompatSetting() {
		this.name = null;
		this.uri = null;
		this.val = null;
		return this;
	}
	CT_CompatSetting.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "name": {
					this.name = reader.GetValueDecodeXml();
					break;
				}
				case "uri": {
					this.uri = reader.GetValueDecodeXml();
					break;
				}
				case "val": {
					this.val = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_CompatSetting.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_CompatSetting.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:name", this.name);
		writer.WriteXmlNullableAttributeStringEncode("w:uri", this.uri);
		writer.WriteXmlNullableAttributeStringEncode("w:val", this.val);
		writer.WriteXmlAttributesEnd(true);
	};
	CT_CompatSetting.prototype.toXmlValues = function(writer, nodeName, name, uri, val) {
		if (null !== val && undefined !== val) {
			let elem = new CT_CompatSetting();
			elem.name = name;
			elem.uri = uri;
			elem.val = val;
			elem.toXml(writer, nodeName);
		}
	};
	function CT_MarkupRange() {
		this.id = null;
		this.displacedByCustomXml = null;
		return this;
	}
	CT_MarkupRange.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "id": {
					this.id = reader.GetValueInt(this.id);
					break;
				}
				// case "displacedByCustomXml": {
				// 	this.displacedByCustomXml = fromXml_ST_DisplacedByCustomXml(reader.GetValue(), this.displacedByCustomXml);
				// 	break;
				// }
			}
		}
	};
	CT_MarkupRange.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_MarkupRange.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeInt("w:id", this.id);
		// writer.WriteXmlNullableAttributeString("w:displacedByCustomXml", toXml_ST_DisplacedByCustomXml(this.displacedByCustomXml));
		writer.WriteXmlAttributesEnd(true);
	};
	function CT_Sym() {
		this.font = null;
		this.char = null;
		return this;
	}
	CT_Sym.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "font": {
					this.font = reader.GetValueDecodeXml();
					break;
				}
				case "char": {
					this.char = reader.GetValueByte(this.char, 16);
					if (null !== this.char) {
						this.char = 0x0FFF & this.char;
					}
					break;
				}
			}
		}
	};
	CT_Sym.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_Sym.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:font", this.font);
		writer.WriteXmlNullableAttributeByte("w:char", this.char);
		writer.WriteXmlAttributesEnd(true);
	};
	function CT_BdoDirContentRun() {
		CParagraphContentWithParagraphLikeContent.call(this);
		this.Val = null;
		return this;
	}
	CT_BdoDirContentRun.prototype = Object.create(CParagraphContentWithParagraphLikeContent.prototype);
	CT_BdoDirContentRun.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "val": {
					// this.val = fromXml_ST_Direction(reader.GetValue(), this.val);
					break;
				}
			}
		}
	};
	CT_BdoDirContentRun.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		CParagraphContentWithParagraphLikeContent.prototype.fromXml.call(this, reader);
	};
	CT_BdoDirContentRun.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		// writer.WriteXmlNullableAttributeString("w:val", toXml_ST_Direction(this.val));
		writer.WriteXmlAttributesEnd();
		CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
		writer.WriteXmlNodeEnd(name);
	};
	function CT_SmartTagRun() {
		CParagraphContentWithParagraphLikeContent.call(this);
		this.uri = null;
		this.element = null;
		this.smartTagPr = [];
		return this;
	}
	CT_SmartTagRun.prototype = Object.create(CParagraphContentWithParagraphLikeContent.prototype);
	CT_SmartTagRun.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "uri": {
					this.uri = reader.GetValueDecodeXml();
					break;
				}
				case "element": {
					this.element = reader.GetValueDecodeXml();
					break;
				}
			}
		}
	};
	CT_SmartTagRun.prototype.fromXml = function (reader) {
		this.readAttr(reader);
		//todo smartTagPr
		CParagraphContentWithParagraphLikeContent.prototype.fromXml.call(this, reader);
	};
	CT_SmartTagRun.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeStringEncode("w:uri", this.uri);
		writer.WriteXmlNullableAttributeStringEncode("w:element", this.element);
		writer.WriteXmlAttributesEnd();
		// writer.WriteXmlArray(this.smartTagPr, "w:smartTagPr");
		CParagraphContentWithParagraphLikeContent.prototype.toXml.call(this, writer);
		writer.WriteXmlNodeEnd(name);
	};
	function CT_TblLayoutType() {
		this.type = null;
		return this;
	}
	CT_TblLayoutType.prototype.readAttr = function(reader) {
		while (reader.MoveToNextAttribute()) {
			switch (reader.GetNameNoNS()) {
				case "type": {
					this.type = fromXml_ST_TblLayoutType(reader.GetValue(), this.type);
					break;
				}
			}
		}
	};
	CT_TblLayoutType.prototype.fromXml = function(reader) {
		this.readAttr(reader);
		reader.ReadTillEnd();
	};
	CT_TblLayoutType.prototype.toXml = function(writer, name) {
		writer.WriteXmlNodeStart(name);
		writer.WriteXmlNullableAttributeString("w:type", toXml_ST_TblLayoutType(this.type));
		writer.WriteXmlAttributesEnd(true);
	};
	//review
	if (typeof CRunRevisionMove !== "undefined" && typeof CParaRevisionMove !== "undefined") {
		CRunRevisionMove.prototype.readAttr = CParaRevisionMove.prototype.readAttr = function(reader, options) {
			while (reader.MoveToNextAttribute()) {
				switch (reader.GetNameNoNS()) {
					// case "colFirst": {
					// 	this.colFirst = reader.GetValueInt(this.colFirst);
					// 	break;
					// }
					// case "colLast": {
					// 	this.colLast = reader.GetValueInt(this.colLast);
					// 	break;
					// }
					case "id": {
						options.id = reader.GetValueInt(options.id);
						break;
					}
					case "name": {
						this.Name = reader.GetValueDecodeXml();
						break;
					}
					case "author": {
						this.ReviewInfo.UserName = reader.GetValueDecodeXml();
						break;
					}
					case "date": {
						let dateStr = reader.GetValueDecodeXml();
						let dateMs = AscCommon.getTimeISO8601(dateStr);
						if (isNaN(dateMs)) {
							dateMs = new Date().getTime();
						}
						this.ReviewInfo.DateTime = dateMs;
						break;
					}
					case "oouserid": {
						this.ReviewInfo.UserId = reader.GetValueDecodeXml();
						break;
					}
				}
			}
		};
		CRunRevisionMove.prototype.fromXml = CParaRevisionMove.prototype.fromXml = function(reader, options) {
			this.readAttr(reader, options);
			reader.ReadTillEnd();
		};
		CRunRevisionMove.prototype.toXml = CParaRevisionMove.prototype.toXml = function(writer, name, options) {
			writer.WriteXmlNodeStart(name);
			// writer.WriteXmlNullableAttributeInt("w:colFirst", this.colFirst);
			// writer.WriteXmlNullableAttributeInt("w:colLast", this.colLast);

			if (this.IsStart()) {
				if (this.ReviewInfo) {
					let dateUtc = this.ReviewInfo.DateTime ?
					new Date(this.ReviewInfo.DateTime).toISOString().slice(0, 19) + 'Z' : null;
					writer.WriteXmlNonEmptyAttributeStringEncode("w:author", this.ReviewInfo.UserName);
					writer.WriteXmlNullableAttributeString("w:date", dateUtc);
				}
				if (options) {
					writer.WriteXmlNullableAttributeInt("w:id", options.id);
				}
				writer.WriteXmlNullableAttributeStringEncode("w:name", this.Name);
				if (this.ReviewInfo) {
					writer.WriteXmlNonEmptyAttributeStringEncode("oouserid", this.ReviewInfo.UserId);
				}
			} else {
				if (options) {
					writer.WriteXmlNullableAttributeInt("w:id", options.id);
				}
			}

			writer.WriteXmlAttributesEnd(true);
		};
	}

	//enums
	function fromXml_ST_Border(val) {
		switch (val) {
			case "nil":
				return border_None;
			case "none":
				return border_None;
			case "single":
				return border_Single;
			case "thick":
				return border_Single;
			case "double":
				return border_Single;
			case "dotted":
				return border_Single;
			case "dashed":
				return border_Single;
			case "dotDash":
				return border_Single;
			case "dotDotDash":
				return border_Single;
			case "triple":
				return border_Single;
			case "thinThickSmallGap":
				return border_Single;
			case "thickThinSmallGap":
				return border_Single;
			case "thinThickThinSmallGap":
				return border_Single;
			case "thinThickMediumGap":
				return border_Single;
			case "thickThinMediumGap":
				return border_Single;
			case "thinThickThinMediumGap":
				return border_Single;
			case "thinThickLargeGap":
				return border_Single;
			case "thickThinLargeGap":
				return border_Single;
			case "thinThickThinLargeGap":
				return border_Single;
			case "wave":
				return border_Single;
			case "doubleWave":
				return border_Single;
			case "dashSmallGap":
				return border_Single;
			case "dashDotStroked":
				return border_Single;
			case "threeDEmboss":
				return border_Single;
			case "threeDEngrave":
				return border_Single;
			case "outset":
				return border_Single;
			case "inset":
				return border_Single;
			case "apples":
				return border_Single;
			case "archedScallops":
				return border_Single;
			case "babyPacifier":
				return border_Single;
			case "babyRattle":
				return border_Single;
			case "balloons3Colors":
				return border_Single;
			case "balloonsHotAir":
				return border_Single;
			case "basicBlackDashes":
				return border_Single;
			case "basicBlackDots":
				return border_Single;
			case "basicBlackSquares":
				return border_Single;
			case "basicThinLines":
				return border_Single;
			case "basicWhiteDashes":
				return border_Single;
			case "basicWhiteDots":
				return border_Single;
			case "basicWhiteSquares":
				return border_Single;
			case "basicWideInline":
				return border_Single;
			case "basicWideMidline":
				return border_Single;
			case "basicWideOutline":
				return border_Single;
			case "bats":
				return border_Single;
			case "birds":
				return border_Single;
			case "birdsFlight":
				return border_Single;
			case "cabins":
				return border_Single;
			case "cakeSlice":
				return border_Single;
			case "candyCorn":
				return border_Single;
			case "celticKnotwork":
				return border_Single;
			case "certificateBanner":
				return border_Single;
			case "chainLink":
				return border_Single;
			case "champagneBottle":
				return border_Single;
			case "checkedBarBlack":
				return border_Single;
			case "checkedBarColor":
				return border_Single;
			case "checkered":
				return border_Single;
			case "christmasTree":
				return border_Single;
			case "circlesLines":
				return border_Single;
			case "circlesRectangles":
				return border_Single;
			case "classicalWave":
				return border_Single;
			case "clocks":
				return border_Single;
			case "compass":
				return border_Single;
			case "confetti":
				return border_Single;
			case "confettiGrays":
				return border_Single;
			case "confettiOutline":
				return border_Single;
			case "confettiStreamers":
				return border_Single;
			case "confettiWhite":
				return border_Single;
			case "cornerTriangles":
				return border_Single;
			case "couponCutoutDashes":
				return border_Single;
			case "couponCutoutDots":
				return border_Single;
			case "crazyMaze":
				return border_Single;
			case "creaturesButterfly":
				return border_Single;
			case "creaturesFish":
				return border_Single;
			case "creaturesInsects":
				return border_Single;
			case "creaturesLadyBug":
				return border_Single;
			case "crossStitch":
				return border_Single;
			case "cup":
				return border_Single;
			case "decoArch":
				return border_Single;
			case "decoArchColor":
				return border_Single;
			case "decoBlocks":
				return border_Single;
			case "diamondsGray":
				return border_Single;
			case "doubleD":
				return border_Single;
			case "doubleDiamonds":
				return border_Single;
			case "earth1":
				return border_Single;
			case "earth2":
				return border_Single;
			case "earth3":
				return border_Single;
			case "eclipsingSquares1":
				return border_Single;
			case "eclipsingSquares2":
				return border_Single;
			case "eggsBlack":
				return border_Single;
			case "fans":
				return border_Single;
			case "film":
				return border_Single;
			case "firecrackers":
				return border_Single;
			case "flowersBlockPrint":
				return border_Single;
			case "flowersDaisies":
				return border_Single;
			case "flowersModern1":
				return border_Single;
			case "flowersModern2":
				return border_Single;
			case "flowersPansy":
				return border_Single;
			case "flowersRedRose":
				return border_Single;
			case "flowersRoses":
				return border_Single;
			case "flowersTeacup":
				return border_Single;
			case "flowersTiny":
				return border_Single;
			case "gems":
				return border_Single;
			case "gingerbreadMan":
				return border_Single;
			case "gradient":
				return border_Single;
			case "handmade1":
				return border_Single;
			case "handmade2":
				return border_Single;
			case "heartBalloon":
				return border_Single;
			case "heartGray":
				return border_Single;
			case "hearts":
				return border_Single;
			case "heebieJeebies":
				return border_Single;
			case "holly":
				return border_Single;
			case "houseFunky":
				return border_Single;
			case "hypnotic":
				return border_Single;
			case "iceCreamCones":
				return border_Single;
			case "lightBulb":
				return border_Single;
			case "lightning1":
				return border_Single;
			case "lightning2":
				return border_Single;
			case "mapPins":
				return border_Single;
			case "mapleLeaf":
				return border_Single;
			case "mapleMuffins":
				return border_Single;
			case "marquee":
				return border_Single;
			case "marqueeToothed":
				return border_Single;
			case "moons":
				return border_Single;
			case "mosaic":
				return border_Single;
			case "musicNotes":
				return border_Single;
			case "northwest":
				return border_Single;
			case "ovals":
				return border_Single;
			case "packages":
				return border_Single;
			case "palmsBlack":
				return border_Single;
			case "palmsColor":
				return border_Single;
			case "paperClips":
				return border_Single;
			case "papyrus":
				return border_Single;
			case "partyFavor":
				return border_Single;
			case "partyGlass":
				return border_Single;
			case "pencils":
				return border_Single;
			case "people":
				return border_Single;
			case "peopleWaving":
				return border_Single;
			case "peopleHats":
				return border_Single;
			case "poinsettias":
				return border_Single;
			case "postageStamp":
				return border_Single;
			case "pumpkin1":
				return border_Single;
			case "pushPinNote2":
				return border_Single;
			case "pushPinNote1":
				return border_Single;
			case "pyramids":
				return border_Single;
			case "pyramidsAbove":
				return border_Single;
			case "quadrants":
				return border_Single;
			case "rings":
				return border_Single;
			case "safari":
				return border_Single;
			case "sawtooth":
				return border_Single;
			case "sawtoothGray":
				return border_Single;
			case "scaredCat":
				return border_Single;
			case "seattle":
				return border_Single;
			case "shadowedSquares":
				return border_Single;
			case "sharksTeeth":
				return border_Single;
			case "shorebirdTracks":
				return border_Single;
			case "skyrocket":
				return border_Single;
			case "snowflakeFancy":
				return border_Single;
			case "snowflakes":
				return border_Single;
			case "sombrero":
				return border_Single;
			case "southwest":
				return border_Single;
			case "stars":
				return border_Single;
			case "starsTop":
				return border_Single;
			case "stars3d":
				return border_Single;
			case "starsBlack":
				return border_Single;
			case "starsShadowed":
				return border_Single;
			case "sun":
				return border_Single;
			case "swirligig":
				return border_Single;
			case "tornPaper":
				return border_Single;
			case "tornPaperBlack":
				return border_Single;
			case "trees":
				return border_Single;
			case "triangleParty":
				return border_Single;
			case "triangles":
				return border_Single;
			case "triangle1":
				return border_Single;
			case "triangle2":
				return border_Single;
			case "triangleCircle1":
				return border_Single;
			case "triangleCircle2":
				return border_Single;
			case "shapes1":
				return border_Single;
			case "shapes2":
				return border_Single;
			case "twistedLines1":
				return border_Single;
			case "twistedLines2":
				return border_Single;
			case "vine":
				return border_Single;
			case "waveline":
				return border_Single;
			case "weavingAngles":
				return border_Single;
			case "weavingBraid":
				return border_Single;
			case "weavingRibbon":
				return border_Single;
			case "weavingStrips":
				return border_Single;
			case "whiteFlowers":
				return border_Single;
			case "woodwork":
				return border_Single;
			case "xIllusions":
				return border_Single;
			case "zanyTriangles":
				return border_Single;
			case "zigZag":
				return border_Single;
			case "zigZagStitch":
				return border_Single;
			case "custom":
				return border_Single;
		}
		return undefined;
	}

	function toXml_ST_Border(val) {
		switch (val) {
			case border_None:
				return "none";
			case border_Single:
				return "single";
		}
		return null;
	}

	function fromXml_ST_DropCap(val) {
		switch (val) {
			case "none":
				return Asc.c_oAscDropCap.None;
			case "drop":
				return Asc.c_oAscDropCap.Drop;
			case "margin":
				return Asc.c_oAscDropCap.Margin;
		}
		return undefined;
	}

	function toXml_ST_DropCap(val) {
		switch (val) {
			case Asc.c_oAscDropCap.None:
				return "none";
			case Asc.c_oAscDropCap.Drop:
				return "drop";
			case Asc.c_oAscDropCap.Margin:
				return "margin";
		}
		return null;
	}

	function fromXml_ST_Wrap(val) {
		switch (val) {
			case "auto":
				return wrap_Auto;
			case "notBeside":
				return wrap_NotBeside;
			case "around":
				return wrap_Around;
			case "tight":
				return wrap_Tight;
			case "through":
				return wrap_Through;
			case "none":
				return wrap_None;
		}
		return undefined;
	}

	function toXml_ST_Wrap(val) {
		switch (val) {
			case wrap_Auto:
				return "auto";
			case wrap_NotBeside:
				return "notBeside";
			case wrap_Around:
				return "around";
			case wrap_Tight:
				return "tight";
			case wrap_Through:
				return "through";
			case wrap_None:
				return "none";
		}
		return null;
	}

	function fromXml_ST_HAnchor(val) {
		switch (val) {
			case "text":
				return Asc.c_oAscHAnchor.Text;
			case "margin":
				return Asc.c_oAscHAnchor.Margin;
			case "page":
				return Asc.c_oAscHAnchor.Page;
		}
		return undefined;
	}

	function toXml_ST_HAnchor(val) {
		switch (val) {
			case Asc.c_oAscHAnchor.Text:
				return "text";
			case Asc.c_oAscHAnchor.Margin:
				return "margin";
			case Asc.c_oAscHAnchor.Page:
				return "page";
		}
		return null;
	}

	function fromXml_ST_VAnchor(val) {
		switch (val) {
			case "text":
				return Asc.c_oAscVAnchor.Text;
			case "margin":
				return Asc.c_oAscVAnchor.Margin;
			case "page":
				return Asc.c_oAscVAnchor.Page;
		}
		return undefined;
	}

	function toXml_ST_VAnchor(val) {
		switch (val) {
			case Asc.c_oAscVAnchor.Text:
				return "text";
			case Asc.c_oAscVAnchor.Margin:
				return "margin";
			case Asc.c_oAscVAnchor.Page:
				return "page";
		}
		return null;
	}

	function fromXml_ST_XAlign(val) {
		switch (val) {
			case "left":
				return Asc.c_oAscXAlign.Left;
			case "center":
				return Asc.c_oAscXAlign.Center;
			case "right":
				return Asc.c_oAscXAlign.Right;
			case "inside":
				return Asc.c_oAscXAlign.Inside;
			case "outside":
				return Asc.c_oAscXAlign.Outside;
		}
		return undefined;
	}

	function toXml_ST_XAlign(val) {
		switch (val) {
			case Asc.c_oAscXAlign.Left:
				return "left";
			case Asc.c_oAscXAlign.Center:
				return "center";
			case Asc.c_oAscXAlign.Right:
				return "right";
			case Asc.c_oAscXAlign.Inside:
				return "inside";
			case Asc.c_oAscXAlign.Outside:
				return "outside";
		}
		return null;
	}

	function fromXml_ST_YAlign(val) {
		switch (val) {
			case "inline":
				return Asc.c_oAscYAlign.Inline;
			case "top":
				return Asc.c_oAscYAlign.Top;
			case "center":
				return Asc.c_oAscYAlign.Center;
			case "bottom":
				return Asc.c_oAscYAlign.Bottom;
			case "inside":
				return Asc.c_oAscYAlign.Inside;
			case "outside":
				return Asc.c_oAscYAlign.Outside;
		}
		return undefined;
	}

	function toXml_ST_YAlign(val) {
		switch (val) {
			case Asc.c_oAscYAlign.Inline:
				return "inline";
			case Asc.c_oAscYAlign.Top:
				return "top";
			case Asc.c_oAscYAlign.Center:
				return "center";
			case Asc.c_oAscYAlign.Bottom:
				return "bottom";
			case Asc.c_oAscYAlign.Inside:
				return "inside";
			case Asc.c_oAscYAlign.Outside:
				return "outside";
		}
		return null;
	}

	function fromXml_ST_HeightRule(val) {
		switch (val) {
			case "auto":
				return Asc.linerule_Auto;
			case "exact":
				return Asc.linerule_Exact;
			case "atLeast":
				return Asc.linerule_AtLeast;
		}
		return undefined;
	}

	function toXml_ST_HeightRule(val) {
		switch (val) {
			case Asc.linerule_Auto:
				return "auto";
			case Asc.linerule_Exact:
				return "exact";
			case Asc.linerule_AtLeast:
				return "atLeast";
		}
		return null;
	}

	function fromXml_ST_Shd(val) {
		switch (val) {
			case "nil":
				return Asc.c_oAscShd.Nil;
			case "clear":
				return Asc.c_oAscShd.Clear;
			case "solid":
				return Asc.c_oAscShd.Solid;
			case "horzStripe":
				return Asc.c_oAscShd.HorzStripe;
			case "vertStripe":
				return Asc.c_oAscShd.VertStripe;
			case "reverseDiagStripe":
				return Asc.c_oAscShd.ReverseDiagStripe;
			case "diagStripe":
				return Asc.c_oAscShd.DiagStripe;
			case "horzCross":
				return Asc.c_oAscShd.HorzCross;
			case "diagCross":
				return Asc.c_oAscShd.DiagCross;
			case "thinHorzStripe":
				return Asc.c_oAscShd.ThinHorzStripe;
			case "thinVertStripe":
				return Asc.c_oAscShd.ThinVertStripe;
			case "thinReverseDiagStripe":
				return Asc.c_oAscShd.ThinReverseDiagStripe;
			case "thinDiagStripe":
				return Asc.c_oAscShd.ThinDiagStripe;
			case "thinHorzCross":
				return Asc.c_oAscShd.ThinHorzCross;
			case "thinDiagCross":
				return Asc.c_oAscShd.ThinDiagCross;
			case "pct5":
				return Asc.c_oAscShd.Pct5;
			case "pct10":
				return Asc.c_oAscShd.Pct10;
			case "pct12":
				return Asc.c_oAscShd.Pct12;
			case "pct15":
				return Asc.c_oAscShd.Pct15;
			case "pct20":
				return Asc.c_oAscShd.Pct20;
			case "pct25":
				return Asc.c_oAscShd.Pct25;
			case "pct30":
				return Asc.c_oAscShd.Pct30;
			case "pct35":
				return Asc.c_oAscShd.Pct35;
			case "pct37":
				return Asc.c_oAscShd.Pct37;
			case "pct40":
				return Asc.c_oAscShd.Pct40;
			case "pct45":
				return Asc.c_oAscShd.Pct45;
			case "pct50":
				return Asc.c_oAscShd.Pct50;
			case "pct55":
				return Asc.c_oAscShd.Pct55;
			case "pct60":
				return Asc.c_oAscShd.Pct60;
			case "pct62":
				return Asc.c_oAscShd.Pct62;
			case "pct65":
				return Asc.c_oAscShd.Pct65;
			case "pct70":
				return Asc.c_oAscShd.Pct70;
			case "pct75":
				return Asc.c_oAscShd.Pct75;
			case "pct80":
				return Asc.c_oAscShd.Pct80;
			case "pct85":
				return Asc.c_oAscShd.Pct85;
			case "pct87":
				return Asc.c_oAscShd.Pct87;
			case "pct90":
				return Asc.c_oAscShd.Pct90;
			case "pct95":
				return Asc.c_oAscShd.Pct95;
		}
		return undefined;
	}

	function toXml_ST_Shd(val) {
		switch (val) {
			case Asc.c_oAscShd.Nil:
				return "nil";
			case Asc.c_oAscShd.Clear:
				return "clear";
			case Asc.c_oAscShd.Solid:
				return "solid";
			case Asc.c_oAscShd.HorzStripe:
				return "horzStripe";
			case Asc.c_oAscShd.VertStripe:
				return "vertStripe";
			case Asc.c_oAscShd.ReverseDiagStripe:
				return "reverseDiagStripe";
			case Asc.c_oAscShd.DiagStripe:
				return "diagStripe";
			case Asc.c_oAscShd.HorzCross:
				return "horzCross";
			case Asc.c_oAscShd.DiagCross:
				return "diagCross";
			case Asc.c_oAscShd.ThinHorzStripe:
				return "thinHorzStripe";
			case Asc.c_oAscShd.ThinVertStripe:
				return "thinVertStripe";
			case Asc.c_oAscShd.ThinReverseDiagStripe:
				return "thinReverseDiagStripe";
			case Asc.c_oAscShd.ThinDiagStripe:
				return "thinDiagStripe";
			case Asc.c_oAscShd.ThinHorzCross:
				return "thinHorzCross";
			case Asc.c_oAscShd.ThinDiagCross:
				return "thinDiagCross";
			case Asc.c_oAscShd.Pct5:
				return "pct5";
			case Asc.c_oAscShd.Pct10:
				return "pct10";
			case Asc.c_oAscShd.Pct12:
				return "pct12";
			case Asc.c_oAscShd.Pct15:
				return "pct15";
			case Asc.c_oAscShd.Pct20:
				return "pct20";
			case Asc.c_oAscShd.Pct25:
				return "pct25";
			case Asc.c_oAscShd.Pct30:
				return "pct30";
			case Asc.c_oAscShd.Pct35:
				return "pct35";
			case Asc.c_oAscShd.Pct37:
				return "pct37";
			case Asc.c_oAscShd.Pct40:
				return "pct40";
			case Asc.c_oAscShd.Pct45:
				return "pct45";
			case Asc.c_oAscShd.Pct50:
				return "pct50";
			case Asc.c_oAscShd.Pct55:
				return "pct55";
			case Asc.c_oAscShd.Pct60:
				return "pct60";
			case Asc.c_oAscShd.Pct62:
				return "pct62";
			case Asc.c_oAscShd.Pct65:
				return "pct65";
			case Asc.c_oAscShd.Pct70:
				return "pct70";
			case Asc.c_oAscShd.Pct75:
				return "pct75";
			case Asc.c_oAscShd.Pct80:
				return "pct80";
			case Asc.c_oAscShd.Pct85:
				return "pct85";
			case Asc.c_oAscShd.Pct87:
				return "pct87";
			case Asc.c_oAscShd.Pct90:
				return "pct90";
			case Asc.c_oAscShd.Pct95:
				return "pct95";
		}
		return null;
	}

	function fromXml_ST_ThemeColor(val) {
		switch (val) {
			case "dark1":
				return AscCommonWord.EThemeColor.themecolorDark1;
			case "light1":
				return AscCommonWord.EThemeColor.themecolorLight1;
			case "dark2":
				return AscCommonWord.EThemeColor.themecolorDark2;
			case "light2":
				return AscCommonWord.EThemeColor.themecolorLight2;
			case "accent1":
				return AscCommonWord.EThemeColor.themecolorAccent1;
			case "accent2":
				return AscCommonWord.EThemeColor.themecolorAccent2;
			case "accent3":
				return AscCommonWord.EThemeColor.themecolorAccent3;
			case "accent4":
				return AscCommonWord.EThemeColor.themecolorAccent4;
			case "accent5":
				return AscCommonWord.EThemeColor.themecolorAccent5;
			case "accent6":
				return AscCommonWord.EThemeColor.themecolorAccent6;
			case "hyperlink":
				return AscCommonWord.EThemeColor.themecolorHyperlink;
			case "followedHyperlink":
				return AscCommonWord.EThemeColor.themecolorFollowedHyperlink;
			case "none":
				return AscCommonWord.EThemeColor.themecolorNone;
			case "background1":
				return AscCommonWord.EThemeColor.themecolorBackground1;
			case "text1":
				return AscCommonWord.EThemeColor.themecolorText1;
			case "background2":
				return AscCommonWord.EThemeColor.themecolorBackground2;
			case "text2":
				return AscCommonWord.EThemeColor.themecolorText2;
		}
		return null;
	}

	function toXml_ST_ThemeColor(val) {
		switch (val) {
			case AscCommonWord.EThemeColor.themecolorDark1:
				return "dark1";
			case AscCommonWord.EThemeColor.themecolorLight1:
				return "light1";
			case AscCommonWord.EThemeColor.themecolorDark2:
				return "dark2";
			case AscCommonWord.EThemeColor.themecolorLight2:
				return "light2";
			case AscCommonWord.EThemeColor.themecolorAccent1:
				return "accent1";
			case AscCommonWord.EThemeColor.themecolorAccent2:
				return "accent2";
			case AscCommonWord.EThemeColor.themecolorAccent3:
				return "accent3";
			case AscCommonWord.EThemeColor.themecolorAccent4:
				return "accent4";
			case AscCommonWord.EThemeColor.themecolorAccent5:
				return "accent5";
			case AscCommonWord.EThemeColor.themecolorAccent6:
				return "accent6";
			case AscCommonWord.EThemeColor.themecolorHyperlink:
				return "hyperlink";
			case AscCommonWord.EThemeColor.themecolorFollowedHyperlink:
				return "followedHyperlink";
			case AscCommonWord.EThemeColor.themecolorNone:
				return "none";
			case AscCommonWord.EThemeColor.themecolorBackground1:
				return "background1";
			case AscCommonWord.EThemeColor.themecolorText1:
				return "text1";
			case AscCommonWord.EThemeColor.themecolorBackground2:
				return "background2";
			case AscCommonWord.EThemeColor.themecolorText2:
				return "text2";
		}
		return null;
	}

	function fromXml_ST_TabJc(val, def) {
		switch (val) {
			case "clear":
				return tab_Clear;
			case "start":
			case "left":
				return tab_Left;
			case "center":
				return tab_Center;
			case "end":
			case "right":
				return tab_Right;
			case "decimal":
				return tab_Decimal;
			case "bar":
				return tab_Bar;
			case "num":
				return tab_Num;
		}
		return def;
	}

	function toXml_ST_TabJc(val) {
		switch (val) {
			case tab_Clear:
				return "clear";
			case tab_Left:
				return "left";
			case tab_Center:
				return "center";
			case tab_Right:
				return "right";
			case tab_Decimal:
				return "decimal";
			case tab_Bar:
				return "bar";
			case tab_Num:
				return "num";
		}
		return null;
	}

	function fromXml_ST_TabTlc(val) {
		switch (val) {
			case "none":
				return Asc.c_oAscTabLeader.None;
			case "dot":
				return Asc.c_oAscTabLeader.Dot;
			case "hyphen":
				return Asc.c_oAscTabLeader.Hyphen;
			case "underscore":
				return Asc.c_oAscTabLeader.Underscore;
			case "heavy":
				return Asc.c_oAscTabLeader.Heavy;
			case "middleDot":
				return Asc.c_oAscTabLeader.MiddleDot;
		}
		return undefined;
	}

	function toXml_ST_TabTlc(val) {
		switch (val) {
			case Asc.c_oAscTabLeader.None:
				return "none";
			case Asc.c_oAscTabLeader.Dot:
				return "dot";
			case Asc.c_oAscTabLeader.Hyphen:
				return "hyphen";
			case Asc.c_oAscTabLeader.Underscore:
				return "underscore";
			case Asc.c_oAscTabLeader.Heavy:
				return "heavy";
			case Asc.c_oAscTabLeader.MiddleDot:
				return "middleDot";
		}
		return null;
	}

	function fromXml_ST_LineSpacingRule(val) {
		switch (val) {
			case "auto":
				return Asc.linerule_Auto;
			case "exact":
				return Asc.linerule_Exact;
			case "atLeast":
				return Asc.linerule_AtLeast;
		}
		return undefined;
	}

	function toXml_ST_LineSpacingRule(val) {
		switch (val) {
			case Asc.linerule_Auto:
				return "auto";
			case Asc.linerule_Exact:
				return "exact";
			case Asc.linerule_AtLeast:
				return "atLeast";
		}
		return null;
	}

	function fromXml_ST_Jc1(val, def) {
		switch (val) {
			case "start":
			case "left":
				return AscCommon.align_Left;
			case "center":
				return AscCommon.align_Center;
			case "end":
			case "right":
				return AscCommon.align_Right;
			case "both":
				return AscCommon.align_Justify;
			case "mediumKashida":
				return AscCommon.align_Left;
			case "distribute":
				return AscCommon.align_Distributed;
			case "numTab":
				return AscCommon.align_Left;
			case "highKashida":
				return AscCommon.align_Left;
			case "lowKashida":
				return AscCommon.align_Left;
			case "thaiDistribute":
				return AscCommon.align_Distributed;
		}
		return def;
	}

	function toXml_ST_Jc1(val) {
		switch (val) {
			case AscCommon.align_Left:
				return "left";
			case AscCommon.align_Center:
				return "center";
			case AscCommon.align_Right:
				return "right";
			case AscCommon.align_Justify:
				return "both";
			case AscCommon.align_Distributed:
				return "distribute";
			// case AscCommon.align_CenterContinuous:
			// 	return "distribute";
		}
		return null;
	}

	function fromXml_ST_FtnEdn(val, def) {
		switch (val) {
			case "normal":
				return footnote_Normal;
			case "separator":
				return footnote_Separator;
			case "continuationSeparator":
				return footnote_ContinuationSeparator;
			case "continuationNotice":
				return footnote_ContinuationNotice;
		}
		return def;
	}
	function toXml_ST_FtnEdn(val) {
		switch (val) {
			case footnote_Normal:
				return "normal";
			case footnote_Separator:
				return "separator";
			case footnote_ContinuationSeparator:
				return "continuationSeparator";
			case footnote_ContinuationNotice:
				return "continuationNotice";
		}
		return null;
	}

	function fromXml_ST_Hint(val) {
		switch (val) {
			case "default":
				return AscWord.fonthint_Default;
			case "cs":
				return AscWord.fonthint_CS;
			case "eastAsia":
				return AscWord.fonthint_EastAsia;
		}
		return undefined;
	}

	function toXml_ST_Hint(val) {
		switch (val) {
			case AscWord.fonthint_Default:
				return "default";
			case AscWord.fonthint_CS:
				return "cs";
			case AscWord.fonthint_EastAsia:
				return "eastAsia";
		}
		return null;
	}

	function fromXml_ST_Underline(val) {
		switch (val) {
			case "single":
				return true;
			case "words":
				return true;
			case "double":
				return true;
			case "thick":
				return true;
			case "dotted":
				return true;
			case "dottedHeavy":
				return true;
			case "dash":
				return true;
			case "dashedHeavy":
				return true;
			case "dashLong":
				return true;
			case "dashLongHeavy":
				return true;
			case "dotDash":
				return true;
			case "dashDotHeavy":
				return true;
			case "dotDotDash":
				return true;
			case "dashDotDotHeavy":
				return true;
			case "wave":
				return true;
			case "wavyHeavy":
				return true;
			case "wavyDouble":
				return true;
			case "none":
				return false;
		}
		return undefined;
	}

	function toXml_ST_Underline(val) {
		switch (val) {
			case true:
				return "single";
			case false:
				return "none";
		}
		return null;
	}

	function fromXml_ST_VerticalAlignRun(val) {
		switch (val) {
			case "baseline":
				return AscCommon.vertalign_Baseline;
			case "superscript":
				return AscCommon.vertalign_SuperScript;
			case "subscript":
				return AscCommon.vertalign_SubScript;
		}
		return undefined;
	}

	function toXml_ST_VerticalAlignRun(val) {
		switch (val) {
			case AscCommon.vertalign_Baseline:
				return "baseline";
			case AscCommon.vertalign_SuperScript:
				return "superscript";
			case AscCommon.vertalign_SubScript:
				return "subscript";
		}
		return null;
	}

	function fromXml_ST_TblWidth(val, def) {
		switch (val) {
			case "nil":
				return tblwidth_Nil;
			case "pct":
				return tblwidth_Pct;
			case "dxa":
				return tblwidth_Mm;
			case "auto":
				return tblwidth_Auto;
		}
		return def;
	}

	function toXml_ST_TblWidth(val) {
		switch (val) {
			case tblwidth_Nil:
				return "nil";
			case tblwidth_Pct:
				return "pct";
			case tblwidth_Mm:
				return "dxa";
			case tblwidth_Auto:
				return "auto";
		}
		return null;
	}

	function fromXml_ST_JcTable(val) {
		switch (val) {
			case "center":
				return AscCommon.align_Center;
			case "right":
			case "end":
				return AscCommon.align_Right;
			case "left":
			case "start":
				return AscCommon.align_Left;
		}
		return undefined;
	}

	function toXml_ST_JcTable(val) {
		switch (val) {
			case AscCommon.align_Center:
				return "center";
			case AscCommon.align_Right:
				return "right";
			case AscCommon.align_Left:
				return "left";
		}
		return null;
	}

	function fromXml_ST_TblLayoutType(val) {
		switch (val) {
			case "fixed":
				return tbllayout_Fixed;
			case "autofit":
				return tbllayout_AutoFit;
		}
		return undefined;
	}

	function toXml_ST_TblLayoutType(val) {
		switch (val) {
			case tbllayout_Fixed:
				return "fixed";
			case tbllayout_AutoFit:
				return "autofit";
		}
		return null;
	}

	function fromXml_ST_Merge(val, def) {
		switch (val) {
			case "continue":
				return vmerge_Continue;
			case "restart":
				return vmerge_Restart;
		}
		return def;
	}

	function toXml_ST_Merge(val) {
		switch (val) {
			case vmerge_Continue:
				return "continue";
			case vmerge_Restart:
				return "restart";
		}
		return null;
	}

	function fromXml_ST_TextDirection(val) {
		switch (val) {
			case "lrTb":
			case "tb":
				return textdirection_LRTB;
			case "tbRl":
			case "rl":
				return textdirection_TBRL;
			case "btLr":
			case "lr":
				return textdirection_BTLR;
			case "lrTbV":
			case "tbV":
				return textdirection_LRTBV;
			case "tbRlV":
			case "rlV":
				return textdirection_TBRLV;
			case "tbLrV":
			case "lrV":
				return textdirection_TBLRV;
		}
		return undefined;
	}
	function toXml_ST_TextDirection(val) {
		switch (val) {
			case textdirection_LRTB:
				return "lrTb";
			case textdirection_TBRL:
				return "tbRl";
			case textdirection_BTLR:
				return "btLr";
			case textdirection_LRTBV:
				return "lrTbV";
			case textdirection_TBRLV:
				return "tbRlV";
			case textdirection_TBLRV:
				return "tbLrV";
		}
		return null;
	}

	function fromXml_ST_VerticalJc(val) {
		switch (val) {
			case "top":
				return vertalignjc_Top;
			case "center":
				return vertalignjc_Center;
			case "both":
				return vertalignjc_Top;
			case "bottom":
				return vertalignjc_Bottom;
		}
		return undefined;
	}

	function toXml_ST_VerticalJc(val) {
		switch (val) {
			case vertalignjc_Top:
				return "top";
			case vertalignjc_Center:
				return "center";
			// case ST_VerticalJc.Both:
			// 	return "both";
			case vertalignjc_Bottom:
				return "bottom";
		}
		return null;
	}

	function fromXml_ST_StyleType(val) {
		switch (val) {
			case "paragraph":
				return styletype_Paragraph;
			case "character":
				return styletype_Character;
			case "table":
				return styletype_Table;
			case "numbering":
				return styletype_Numbering;
		}
		return undefined;
	}

	function toXml_ST_StyleType(val) {
		switch (val) {
			case styletype_Paragraph:
				return "paragraph";
			case styletype_Character:
				return "character";
			case styletype_Table:
				return "table";
			case styletype_Numbering:
				return "numbering";
		}
		return null;
	}
	function fromXml_ST_TblStyleOverrideType(style, elem, type) {
		switch (type) {
			case "wholeTable":
				style.TableWholeTable = elem;
				break;
			case "firstRow":
				style.TableFirstRow = elem;
				break;
			case "lastRow":
				style.TableLastRow = elem;
				break;
			case "firstCol":
				style.TableFirstCol = elem;
				break;
			case "lastCol":
				style.TableLastCol = elem;
				break;
			case "band1Vert":
				style.TableBand1Vert = elem;
				break;
			case "band2Vert":
				style.TableBand2Vert = elem;
				break;
			case "band1Horz":
				style.TableBand1Horz = elem;
				break;
			case "band2Horz":
				style.TableBand2Horz = elem;
				break;
			case "neCell":
				style.TableTRCell = elem;
				break;
			case "nwCell":
				style.TableTLCell = elem;
				break;
			case "seCell":
				style.TableBRCell = elem;
				break;
			case "swCell":
				style.TableBLCell = elem;
				break;
		}
	}
	function toXml_ST_TblStyleOverrideType(writer, style) {
		if (style.TableWholeTable) {
			style.TableWholeTable.toXml(writer, "w:tblStylePr", "wholeTable");
		}
		if (style.TableFirstRow) {
			style.TableFirstRow.toXml(writer, "w:tblStylePr", "firstRow");
		}
		if (style.TableLastRow) {
			style.TableLastRow.toXml(writer, "w:tblStylePr", "lastRow");
		}
		if (style.TableFirstCol) {
			style.TableFirstCol.toXml(writer, "w:tblStylePr", "firstCol");
		}
		if (style.TableLastCol) {
			style.TableLastCol.toXml(writer, "w:tblStylePr", "lastCol");
		}
		if (style.TableBand1Vert) {
			style.TableBand1Vert.toXml(writer, "w:tblStylePr", "band1Vert");
		}
		if (style.TableBand2Vert) {
			style.TableBand2Vert.toXml(writer, "w:tblStylePr", "band2Vert");
		}
		if (style.TableBand1Horz) {
			style.TableBand1Horz.toXml(writer, "w:tblStylePr", "band1Horz");
		}
		if (style.TableBand2Horz) {
			style.TableBand2Horz.toXml(writer, "w:tblStylePr", "band2Horz");
		}
		if (style.TableTRCell) {
			style.TableTRCell.toXml(writer, "w:tblStylePr", "neCell");
		}
		if (style.TableTLCell) {
			style.TableTLCell.toXml(writer, "w:tblStylePr", "nwCell");
		}
		if (style.TableBRCell) {
			style.TableBRCell.toXml(writer, "w:tblStylePr", "seCell");
		}
		if (style.TableBLCell) {
			style.TableBLCell.toXml(writer, "w:tblStylePr", "swCell");
		}
	}

	function fromXml_ST_FldCharType(val, def) {
		switch (val) {
			case "begin":
				return fldchartype_Begin;
			case "separate":
				return fldchartype_Separate;
			case "end":
				return fldchartype_End;
		}
		return def;
	}
	function toXml_ST_FldCharType(val) {
		switch (val) {
			case fldchartype_Begin:
				return "begin";
			case fldchartype_Separate:
				return "separate";
			case fldchartype_End:
				return "end";
		}
		return null;
	}
	function fromXml_ST_SectionMark(val) {
		switch (val) {
			case "nextPage":
				return c_oAscSectionBreakType.NextPage;
			case "nextColumn":
				return c_oAscSectionBreakType.Column;
			case "continuous":
				return c_oAscSectionBreakType.Continuous;
			case "evenPage":
				return c_oAscSectionBreakType.EvenPage;
			case "oddPage":
				return c_oAscSectionBreakType.OddPage;
		}
		return undefined;
	}
	function toXml_ST_SectionMark(val) {
		switch (val) {
			case c_oAscSectionBreakType.NextPage:
				return "nextPage";
			case c_oAscSectionBreakType.Column:
				return "nextColumn";
			case c_oAscSectionBreakType.Continuous:
				return "continuous";
			case c_oAscSectionBreakType.EvenPage:
				return "evenPage";
			case c_oAscSectionBreakType.OddPage:
				return "oddPage";
		}
		return null;
	}
	function fromXml_ST_PageOrientation(val) {
		switch (val) {
			case "portrait":
				return Asc.c_oAscPageOrientation.PagePortrait;
			case "landscape":
				return Asc.c_oAscPageOrientation.PageLandscape;
		}
		return undefined;
	}
	function toXml_ST_PageOrientation(val) {
		switch (val) {
			case Asc.c_oAscPageOrientation.PagePortrait:
				return "portrait";
			case Asc.c_oAscPageOrientation.PageLandscape:
				return "landscape";
		}
		return null;
	}
	function fromXml_ST_PageBorderZOrder(val) {
		switch (val) {
			case "front":
				return section_borders_ZOrderFront;
			case "back":
				return section_borders_ZOrderBack;
		}
		return undefined;
	}
	function toXml_ST_PageBorderZOrder(val) {
		switch (val) {
			case section_borders_ZOrderFront:
				return "front";
			case section_borders_ZOrderBack:
				return "back";
		}
		return null;
	}
	function fromXml_ST_PageBorderDisplay(val) {
		switch (val) {
			case "allPages":
				return section_borders_DisplayAllPages;
			case "firstPage":
				return section_borders_DisplayFirstPage;
			case "notFirstPage":
				return section_borders_DisplayNotFirstPage;
		}
		return undefined;
	}
	function toXml_ST_PageBorderDisplay(val) {
		switch (val) {
			case section_borders_DisplayAllPages:
				return "allPages";
			case section_borders_DisplayFirstPage:
				return "firstPage";
			case section_borders_DisplayNotFirstPage:
				return "notFirstPage";
		}
		return null;
	}
	function fromXml_ST_PageBorderOffset(val) {
		switch (val) {
			case "page":
				return section_borders_OffsetFromPage;
			case "text":
				return section_borders_OffsetFromText;
		}
		return undefined;
	}
	function toXml_ST_PageBorderOffset(val) {
		switch (val) {
			case section_borders_OffsetFromPage:
				return "page";
			case section_borders_OffsetFromText:
				return "text";
		}
		return null;
	}
	function fromXml_ST_LineNumberRestart(val, def) {
		switch (val) {
			case "newPage":
				return Asc.c_oAscLineNumberRestartType.NewPage;
			case "newSection":
				return Asc.c_oAscLineNumberRestartType.NewSection;
			case "continuous":
				return Asc.c_oAscLineNumberRestartType.Continuous;
		}
		return def;
	}
	function toXml_ST_LineNumberRestart(val) {
		switch (val) {
			case Asc.c_oAscLineNumberRestartType.NewPage:
				return "newPage";
			case Asc.c_oAscLineNumberRestartType.NewSection:
				return "newSection";
			case Asc.c_oAscLineNumberRestartType.Continuous:
				return "continuous";
		}
		return null;
	}
	function fromXml_ST_FtnPos(val, def) {
		switch (val) {
			case "pageBottom":
				return Asc.c_oAscFootnotePos.PageBottom;
			case "beneathText":
				return Asc.c_oAscFootnotePos.BeneathText;
			case "sectEnd":
				return Asc.c_oAscFootnotePos.SectEnd;
			case "docEnd":
				return Asc.c_oAscFootnotePos.DocEnd;
		}
		return def;
	}
	function toXml_ST_FtnPos(val) {
		switch (val) {
			case Asc.c_oAscFootnotePos.PageBottom:
				return "pageBottom";
			case Asc.c_oAscFootnotePos.BeneathText:
				return "beneathText";
			case Asc.c_oAscFootnotePos.SectEnd:
				return "sectEnd";
			case Asc.c_oAscFootnotePos.DocEnd:
				return "docEnd";
		}
		return null;
	}
	function fromXml_ST_EdnPos(val, def) {
		switch (val) {
			case "sectEnd":
				return Asc.c_oAscEndnotePos.SectEnd;
			case "docEnd":
				return Asc.c_oAscEndnotePos.DocEnd;
		}
		return def;
	}
	function toXml_ST_EdnPos(val) {
		switch (val) {
			case Asc.c_oAscEndnotePos.SectEnd:
				return "sectEnd";
			case Asc.c_oAscEndnotePos.DocEnd:
				return "docEnd";
		}
		return null;
	}
	function fromXml_ST_NumberFormat(val, def) {
		switch (val) {
			case "decimal":
				return Asc.c_oAscNumberingFormat.Decimal;
			case "upperRoman":
				return Asc.c_oAscNumberingFormat.UpperRoman;
			case "lowerRoman":
				return Asc.c_oAscNumberingFormat.LowerRoman;
			case "upperLetter":
				return Asc.c_oAscNumberingFormat.UpperLetter;
			case "lowerLetter":
				return Asc.c_oAscNumberingFormat.LowerLetter;
			// case "ordinal":
			// 	return Asc.c_oAscNumberingFormat.Ordinal;
			// case "cardinalText":
			// 	return Asc.c_oAscNumberingFormat.CardinalText;
			// case "ordinalText":
			// 	return Asc.c_oAscNumberingFormat.OrdinalText;
			// case "hex":
			// 	return Asc.c_oAscNumberingFormat.Hex;
			// case "chicago":
			// 	return Asc.c_oAscNumberingFormat.Chicago;
			// case "ideographDigital":
			// 	return Asc.c_oAscNumberingFormat.IdeographDigital;
			// case "japaneseCounting":
			// 	return Asc.c_oAscNumberingFormat.JapaneseCounting;
			// case "aiueo":
			// 	return Asc.c_oAscNumberingFormat.Aiueo;
			// case "iroha":
			// 	return Asc.c_oAscNumberingFormat.Iroha;
			// case "decimalFullWidth":
			// 	return Asc.c_oAscNumberingFormat.DecimalFullWidth;
			// case "decimalHalfWidth":
			// 	return Asc.c_oAscNumberingFormat.DecimalHalfWidth;
			// case "japaneseLegal":
			// 	return Asc.c_oAscNumberingFormat.JapaneseLegal;
			// case "japaneseDigitalTenThousand":
			// 	return Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand;
			case "decimalEnclosedCircle":
				return Asc.c_oAscNumberingFormat.DecimalEnclosedCircle;
			// case "decimalFullWidth2":
			// 	return Asc.c_oAscNumberingFormat.DecimalFullWidth2;
			// case "aiueoFullWidth":
			// 	return Asc.c_oAscNumberingFormat.AiueoFullWidth;
			// case "irohaFullWidth":
			// 	return Asc.c_oAscNumberingFormat.IrohaFullWidth;
			case "decimalZero":
				return Asc.c_oAscNumberingFormat.DecimalZero;
			case "bullet":
				return Asc.c_oAscNumberingFormat.Bullet;
			// case "ganada":
			// 	return Asc.c_oAscNumberingFormat.Ganada;
			// case "chosung":
			// 	return Asc.c_oAscNumberingFormat.Chosung;
			// case "decimalEnclosedFullstop":
			// 	return Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop;
			// case "decimalEnclosedParen":
			// 	return Asc.c_oAscNumberingFormat.DecimalEnclosedParen;
			// case "decimalEnclosedCircleChinese":
			// 	return Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese;
			// case "ideographEnclosedCircle":
			// 	return Asc.c_oAscNumberingFormat.IdeographEnclosedCircle;
			// case "ideographTraditional":
			// 	return Asc.c_oAscNumberingFormat.IdeographTraditional;
			// case "ideographZodiac":
			// 	return Asc.c_oAscNumberingFormat.IdeographZodiac;
			// case "ideographZodiacTraditional":
			// 	return Asc.c_oAscNumberingFormat.IdeographZodiacTraditional;
			// case "taiwaneseCounting":
			// 	return Asc.c_oAscNumberingFormat.TaiwaneseCounting;
			// case "ideographLegalTraditional":
			// 	return Asc.c_oAscNumberingFormat.IdeographLegalTraditional;
			// case "taiwaneseCountingThousand":
			// 	return Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand;
			// case "taiwaneseDigital":
			// 	return Asc.c_oAscNumberingFormat.TaiwaneseDigital;
			case "chineseCounting":
				return Asc.c_oAscNumberingFormat.ChineseCounting;
			case "chineseLegalSimplified":
				return Asc.c_oAscNumberingFormat.ChineseLegalSimplified;
			case "chineseCountingThousand":
				return Asc.c_oAscNumberingFormat.ChineseCountingThousand;
			// case "koreanDigital":
			// 	return Asc.c_oAscNumberingFormat.KoreanDigital;
			// case "koreanCounting":
			// 	return Asc.c_oAscNumberingFormat.KoreanCounting;
			// case "koreanLegal":
			// 	return Asc.c_oAscNumberingFormat.KoreanLegal;
			// case "koreanDigital2":
			// 	return Asc.c_oAscNumberingFormat.KoreanDigital2;
			// case "vietnameseCounting":
			// 	return Asc.c_oAscNumberingFormat.VietnameseCounting;
			case "russianLower":
				return Asc.c_oAscNumberingFormat.RussianLower;
			case "russianUpper":
				return Asc.c_oAscNumberingFormat.RussianUpper;
			case "none":
				return Asc.c_oAscNumberingFormat.None;
			// case "numberInDash":
			// 	return Asc.c_oAscNumberingFormat.NumberInDash;
			// case "hebrew1":
			// 	return Asc.c_oAscNumberingFormat.Hebrew1;
			// case "hebrew2":
			// 	return Asc.c_oAscNumberingFormat.Hebrew2;
			// case "arabicAlpha":
			// 	return Asc.c_oAscNumberingFormat.ArabicAlpha;
			// case "arabicAbjad":
			// 	return Asc.c_oAscNumberingFormat.ArabicAbjad;
			// case "hindiVowels":
			// 	return Asc.c_oAscNumberingFormat.HindiVowels;
			// case "hindiConsonants":
			// 	return Asc.c_oAscNumberingFormat.HindiConsonants;
			// case "hindiNumbers":
			// 	return Asc.c_oAscNumberingFormat.HindiNumbers;
			// case "hindiCounting":
			// 	return Asc.c_oAscNumberingFormat.HindiCounting;
			// case "thaiLetters":
			// 	return Asc.c_oAscNumberingFormat.ThaiLetters;
			// case "thaiNumbers":
			// 	return Asc.c_oAscNumberingFormat.ThaiNumbers;
			// case "thaiCounting":
			// 	return Asc.c_oAscNumberingFormat.ThaiCounting;
			// case "bahtText":
			// 	return Asc.c_oAscNumberingFormat.BahtText;
			// case "dollarText":
			// 	return Asc.c_oAscNumberingFormat.DollarText;
			// case "custom":
			// 	return Asc.c_oAscNumberingFormat.Custom;
		}
		return def;
	}
	function toXml_ST_NumberFormat(val) {
		switch (val) {
			case Asc.c_oAscNumberingFormat.Decimal:
				return "decimal";
			case Asc.c_oAscNumberingFormat.UpperRoman:
				return "upperRoman";
			case Asc.c_oAscNumberingFormat.LowerRoman:
				return "lowerRoman";
			case Asc.c_oAscNumberingFormat.UpperLetter:
				return "upperLetter";
			case Asc.c_oAscNumberingFormat.LowerLetter:
				return "lowerLetter";
			// case Asc.c_oAscNumberingFormat.Ordinal:
			// 	return "ordinal";
			// case Asc.c_oAscNumberingFormat.CardinalText:
			// 	return "cardinalText";
			// case Asc.c_oAscNumberingFormat.OrdinalText:
			// 	return "ordinalText";
			// case Asc.c_oAscNumberingFormat.Hex:
			// 	return "hex";
			// case Asc.c_oAscNumberingFormat.Chicago:
			// 	return "chicago";
			// case Asc.c_oAscNumberingFormat.IdeographDigital:
			// 	return "ideographDigital";
			// case Asc.c_oAscNumberingFormat.JapaneseCounting:
			// 	return "japaneseCounting";
			// case Asc.c_oAscNumberingFormat.Aiueo:
			// 	return "aiueo";
			// case Asc.c_oAscNumberingFormat.Iroha:
			// 	return "iroha";
			// case Asc.c_oAscNumberingFormat.DecimalFullWidth:
			// 	return "decimalFullWidth";
			// case Asc.c_oAscNumberingFormat.DecimalHalfWidth:
			// 	return "decimalHalfWidth";
			// case Asc.c_oAscNumberingFormat.JapaneseLegal:
			// 	return "japaneseLegal";
			// case Asc.c_oAscNumberingFormat.JapaneseDigitalTenThousand:
			// 	return "japaneseDigitalTenThousand";
			case Asc.c_oAscNumberingFormat.DecimalEnclosedCircle:
				return "decimalEnclosedCircle";
			// case Asc.c_oAscNumberingFormat.DecimalFullWidth2:
			// 	return "decimalFullWidth2";
			// case Asc.c_oAscNumberingFormat.AiueoFullWidth:
			// 	return "aiueoFullWidth";
			// case Asc.c_oAscNumberingFormat.IrohaFullWidth:
			// 	return "irohaFullWidth";
			case Asc.c_oAscNumberingFormat.DecimalZero:
				return "decimalZero";
			case Asc.c_oAscNumberingFormat.Bullet:
				return "bullet";
			// case Asc.c_oAscNumberingFormat.Ganada:
			// 	return "ganada";
			// case Asc.c_oAscNumberingFormat.Chosung:
			// 	return "chosung";
			// case Asc.c_oAscNumberingFormat.DecimalEnclosedFullstop:
			// 	return "decimalEnclosedFullstop";
			// case Asc.c_oAscNumberingFormat.DecimalEnclosedParen:
			// 	return "decimalEnclosedParen";
			// case Asc.c_oAscNumberingFormat.DecimalEnclosedCircleChinese:
			// 	return "decimalEnclosedCircleChinese";
			// case Asc.c_oAscNumberingFormat.IdeographEnclosedCircle:
			// 	return "ideographEnclosedCircle";
			// case Asc.c_oAscNumberingFormat.IdeographTraditional:
			// 	return "ideographTraditional";
			// case Asc.c_oAscNumberingFormat.IdeographZodiac:
			// 	return "ideographZodiac";
			// case Asc.c_oAscNumberingFormat.IdeographZodiacTraditional:
			// 	return "ideographZodiacTraditional";
			// case Asc.c_oAscNumberingFormat.TaiwaneseCounting:
			// 	return "taiwaneseCounting";
			// case Asc.c_oAscNumberingFormat.IdeographLegalTraditional:
			// 	return "ideographLegalTraditional";
			// case Asc.c_oAscNumberingFormat.TaiwaneseCountingThousand:
			// 	return "taiwaneseCountingThousand";
			// case Asc.c_oAscNumberingFormat.TaiwaneseDigital:
			// 	return "taiwaneseDigital";
			case Asc.c_oAscNumberingFormat.ChineseCounting:
				return "chineseCounting";
			case Asc.c_oAscNumberingFormat.ChineseLegalSimplified:
				return "chineseLegalSimplified";
			case Asc.c_oAscNumberingFormat.ChineseCountingThousand:
				return "chineseCountingThousand";
			// case Asc.c_oAscNumberingFormat.KoreanDigital:
			// 	return "koreanDigital";
			// case Asc.c_oAscNumberingFormat.KoreanCounting:
			// 	return "koreanCounting";
			// case Asc.c_oAscNumberingFormat.KoreanLegal:
			// 	return "koreanLegal";
			// case Asc.c_oAscNumberingFormat.KoreanDigital2:
			// 	return "koreanDigital2";
			// case Asc.c_oAscNumberingFormat.VietnameseCounting:
			// 	return "vietnameseCounting";
			case Asc.c_oAscNumberingFormat.RussianLower:
				return "russianLower";
			case Asc.c_oAscNumberingFormat.RussianUpper:
				return "russianUpper";
			case Asc.c_oAscNumberingFormat.None:
				return "none";
			// case Asc.c_oAscNumberingFormat.NumberInDash:
			// 	return "numberInDash";
			// case Asc.c_oAscNumberingFormat.Hebrew1:
			// 	return "hebrew1";
			// case Asc.c_oAscNumberingFormat.Hebrew2:
			// 	return "hebrew2";
			// case Asc.c_oAscNumberingFormat.ArabicAlpha:
			// 	return "arabicAlpha";
			// case Asc.c_oAscNumberingFormat.ArabicAbjad:
			// 	return "arabicAbjad";
			// case Asc.c_oAscNumberingFormat.HindiVowels:
			// 	return "hindiVowels";
			// case Asc.c_oAscNumberingFormat.HindiConsonants:
			// 	return "hindiConsonants";
			// case Asc.c_oAscNumberingFormat.HindiNumbers:
			// 	return "hindiNumbers";
			// case Asc.c_oAscNumberingFormat.HindiCounting:
			// 	return "hindiCounting";
			// case Asc.c_oAscNumberingFormat.ThaiLetters:
			// 	return "thaiLetters";
			// case Asc.c_oAscNumberingFormat.ThaiNumbers:
			// 	return "thaiNumbers";
			// case Asc.c_oAscNumberingFormat.ThaiCounting:
			// 	return "thaiCounting";
			// case Asc.c_oAscNumberingFormat.BahtText:
			// 	return "bahtText";
			// case Asc.c_oAscNumberingFormat.DollarText:
			// 	return "dollarText";
			// case Asc.c_oAscNumberingFormat.Custom:
			// 	return "custom";
		}
		return null;
	}
	function fromXml_ST_RestartNumber(val, def) {
		switch (val) {
			case "continuous":
				return section_footnote_RestartContinuous;
			case "eachSect":
				return section_footnote_RestartEachSect;
			case "eachPage":
				return section_footnote_RestartEachPage;
		}
		return def;
	}
	function toXml_ST_RestartNumber(val) {
		switch (val) {
			case section_footnote_RestartContinuous:
				return "continuous";
			case section_footnote_RestartEachSect:
				return "eachSect";
			case section_footnote_RestartEachPage:
				return "eachPage";
		}
		return null;
	}
	function fromXml_ST_RelFromH(val, def) {
		switch (val) {
			case "margin":
				return c_oAscRelativeFromH.Margin;
			case "page":
				return c_oAscRelativeFromH.Page;
			case "column":
				return c_oAscRelativeFromH.Column;
			case "character":
				return c_oAscRelativeFromH.Character;
			case "leftMargin":
				return c_oAscRelativeFromH.LeftMargin;
			case "rightMargin":
				return c_oAscRelativeFromH.RightMargin;
			case "insideMargin":
				return c_oAscRelativeFromH.InsideMargin;
			case "outsideMargin":
				return c_oAscRelativeFromH.OutsideMargin;
		}
		return def;
	}
	function toXml_ST_RelFromH(val) {
		switch (val) {
			case c_oAscRelativeFromH.Margin:
				return "margin";
			case c_oAscRelativeFromH.Page:
				return "page";
			case c_oAscRelativeFromH.Column:
				return "column";
			case c_oAscRelativeFromH.Character:
				return "character";
			case c_oAscRelativeFromH.LeftMargin:
				return "leftMargin";
			case c_oAscRelativeFromH.RightMargin:
				return "rightMargin";
			case c_oAscRelativeFromH.InsideMargin:
				return "insideMargin";
			case c_oAscRelativeFromH.OutsideMargin:
				return "outsideMargin";
		}
		return null;
	}
	function fromXml_ST_SizeRelFromH(val, def) {
		switch (val) {
			case "margin":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhMargin;
			case "page":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhPage;
			case "leftMargin":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhLeftMargin;
			case "rightMargin":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhRightMargin;
			case "insideMargin":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhInsideMargin;
			case "outsideMargin":
				return AscCommon.c_oAscSizeRelFromH.sizerelfromhOutsideMargin;
		}
		return def;
	}
	function toXml_ST_SizeRelFromH(val) {
		switch (val) {
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhMargin:
				return "margin";
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhPage:
				return "page";
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhLeftMargin:
				return "leftMargin";
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhRightMargin:
				return "rightMargin";
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhInsideMargin:
				return "insideMargin";
			case AscCommon.c_oAscSizeRelFromH.sizerelfromhOutsideMargin:
				return "outsideMargin";
		}
		return null;
	}
	function fromXml_ST_AlignH(val, def) {
		switch (val) {
			case "left":
				return Asc.c_oAscAlignH.Left;
			case "right":
				return Asc.c_oAscAlignH.Right;
			case "center":
				return Asc.c_oAscAlignH.Center;
			case "inside":
				return Asc.c_oAscAlignH.Inside;
			case "outside":
				return Asc.c_oAscAlignH.Outside;
		}
		return def;
	}
	function toXml_ST_AlignH(val) {
		switch (val) {
			case Asc.c_oAscAlignH.Left:
				return "left";
			case Asc.c_oAscAlignH.Right:
				return "right";
			case Asc.c_oAscAlignH.Center:
				return "center";
			case Asc.c_oAscAlignH.Inside:
				return "inside";
			case Asc.c_oAscAlignH.Outside:
				return "outside";
		}
		return null;
	}
	function fromXml_ST_RelFromV(val, def) {
		switch (val) {
			case "margin":
				return c_oAscRelativeFromV.Margin;
			case "page":
				return c_oAscRelativeFromV.Page;
			case "paragraph":
				return c_oAscRelativeFromV.Paragraph;
			case "line":
				return c_oAscRelativeFromV.Line;
			case "topMargin":
				return c_oAscRelativeFromV.TopMargin;
			case "bottomMargin":
				return c_oAscRelativeFromV.BottomMargin;
			case "insideMargin":
				return c_oAscRelativeFromV.InsideMargin;
			case "outsideMargin":
				return c_oAscRelativeFromV.OutsideMargin;
		}
		return def;
	}
	function toXml_ST_RelFromV(val) {
		switch (val) {
			case c_oAscRelativeFromV.Margin:
				return "margin";
			case c_oAscRelativeFromV.Page:
				return "page";
			case c_oAscRelativeFromV.Paragraph:
				return "paragraph";
			case c_oAscRelativeFromV.Line:
				return "line";
			case c_oAscRelativeFromV.TopMargin:
				return "topMargin";
			case c_oAscRelativeFromV.BottomMargin:
				return "bottomMargin";
			case c_oAscRelativeFromV.InsideMargin:
				return "insideMargin";
			case c_oAscRelativeFromV.OutsideMargin:
				return "outsideMargin";
		}
		return null;
	}
	function fromXml_ST_SizeRelFromV(val, def) {
		switch (val) {
			case "margin":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvMargin;
			case "page":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvPage;
			case "topMargin":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvTopMargin;
			case "bottomMargin":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvBottomMargin;
			case "insideMargin":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvInsideMargin;
			case "outsideMargin":
				return AscCommon.c_oAscSizeRelFromV.sizerelfromvOutsideMargin;
		}
		return def;
	}
	function toXml_ST_SizeRelFromV(val) {
		switch (val) {
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvMargin:
				return "margin";
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvPage:
				return "page";
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvTopMargin:
				return "topMargin";
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvBottomMargin:
				return "bottomMargin";
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvInsideMargin:
				return "insideMargin";
			case AscCommon.c_oAscSizeRelFromV.sizerelfromvOutsideMargin:
				return "outsideMargin";
		}
		return null;
	}
	function fromXml_ST_AlignV(val, def) {
		switch (val) {
			case "top":
				return Asc.c_oAscAlignV.Top;
			case "bottom":
				return Asc.c_oAscAlignV.Bottom;
			case "center":
				return Asc.c_oAscAlignV.Center;
			case "inside":
				return Asc.c_oAscAlignV.Inside;
			case "outside":
				return Asc.c_oAscAlignV.Outside;
		}
		return def;
	}
	function toXml_ST_AlignV(val) {
		switch (val) {
			case Asc.c_oAscAlignV.Top:
				return "top";
			case Asc.c_oAscAlignV.Bottom:
				return "bottom";
			case Asc.c_oAscAlignV.Center:
				return "center";
			case Asc.c_oAscAlignV.Inside:
				return "inside";
			case Asc.c_oAscAlignV.Outside:
				return "outside";
		}
		return null;
	}
	function fromXml_ST_LevelSuffix(val, def) {
		switch (val) {
			case "tab":
				return Asc.c_oAscNumberingSuff.Tab;
			case "space":
				return Asc.c_oAscNumberingSuff.Space;
			case "nothing":
				return Asc.c_oAscNumberingSuff.None;
		}
		return def;
	}
	function toXml_ST_LevelSuffix(val) {
		switch (val) {
			case Asc.c_oAscNumberingSuff.Tab:
				return "tab";
			case Asc.c_oAscNumberingSuff.Space:
				return "space";
			case Asc.c_oAscNumberingSuff.None:
				return "nothing";
		}
		return null;
	}
	function fromXml_ST_CryptAlgoritmName(val, def) {
		switch (val) {
			case "1" :
			case "MD2" :
				return ECryptAlgoritmName.MD2;
			case "2" :
			case "MD4" :
				return ECryptAlgoritmName.MD4;
			case "3" :
			case "MD5" :
				return ECryptAlgoritmName.MD5;
			case "6" :
			case "RIPEMD-128" :
				return ECryptAlgoritmName.RIPEMD_128;
			case "7" :
			case "RIPEMD-160" :
				return ECryptAlgoritmName.RIPEMD_160;
			case "4" :
			case "SHA-1" :
				return ECryptAlgoritmName.SHA1;
			case "12" :
			case "SHA-256" :
				return ECryptAlgoritmName.SHA_256;
			case "13" :
			case "SHA-384" :
				return ECryptAlgoritmName.SHA_384;
			case "14" :
			case "SHA-512" :
				return ECryptAlgoritmName.SHA_512;
			case "WHIRLPOOL" :
				return ECryptAlgoritmName.WHIRLPOOL;
		}
		return def;
	}
	function toXml_ST_CryptAlgoritmName(val) {
		switch (val) {
			case ECryptAlgoritmName.MD2 :
				return "MD2";
			case ECryptAlgoritmName.MD4 :
				return "MD4";
			case ECryptAlgoritmName.MD5 :
				return "MD5";
			case ECryptAlgoritmName.RIPEMD_128 :
				return "RIPEMD-128";
			case ECryptAlgoritmName.RIPEMD_160 :
				return "RIPEMD-160";
			case ECryptAlgoritmName.SHA1 :
				return "SHA-1";
			case ECryptAlgoritmName.SHA_256 :
				return "SHA-256";
			case ECryptAlgoritmName.SHA_384 :
				return "SHA-384";
			case ECryptAlgoritmName.SHA_512 :
				return "SHA-512";
			case ECryptAlgoritmName.WHIRLPOOL :
				return "WHIRLPOOL";
		}
		return null;
	}
	function fromXml_ST_DocProtect(val, def) {
		switch (val) {
			case "comments":
				return EDocProtect.Comments;
			case "forms":
				return EDocProtect.Forms;
			case "none":
				return EDocProtect.None;
			case "readOnly":
				return EDocProtect.ReadOnly;
			case "trackedChanges":
				return EDocProtect.TrackedChanges;
		}
		return def;
	}
	function toXml_ST_DocProtect(val) {
		switch (val) {
			case EDocProtect.Comments:
				return "comments";
			case EDocProtect.Forms:
				return "forms";
			case EDocProtect.None:
				return "none";
			case EDocProtect.ReadOnly:
				return "readOnly";
			case EDocProtect.TrackedChanges:
				return "trackedChanges";
		}
		return null;
	}
	function fromXml_ST_CryptAlgClass(val, def) {
		switch (val) {
			case "custom":
				return ECryptAlgClass.Custom;
			case "hash":
				return ECryptAlgClass.Hash;
		}
		return def;
	}
	function toXml_ST_CryptAlgClass(val) {
		switch (val) {
			case ECryptAlgClass.Custom:
				return "custom";
			case ECryptAlgClass.Hash:
				return "hash";
		}
		return null;
	}
	function fromXml_ST_CryptAlgType(val, def) {
		switch (val) {
			case "custom":
				return ECryptAlgType.Custom;
			case "typeAny":
				return ECryptAlgType.TypeAny;
		}
		return def;
	}
	function toXml_ST_CryptAlgType(val) {
		switch (val) {
			case ECryptAlgType.Custom:
				return "custom";
			case ECryptAlgType.TypeAny:
				return "typeAny";
		}
		return null;
	}
	function fromXml_ST_CryptProvType(val, def) {
		switch (val) {
			case "custom":
				return ECryptProv.Custom;
			case "rsaAES":
				return ECryptProv.RsaAES;
			case "rsaFull":
				return ECryptProv.RsaFull;
		}
		return def;
	}
	function toXml_ST_CryptProvType(val) {
		switch (val) {
			case ECryptProv.Custom:
				return "custom";
			case ECryptProv.RsaAES:
				return "rsaAES";
			case ECryptProv.RsaFull:
				return "rsaFull";
		}
		return null;
	}
	function fromXml_ST_DocPartGallery(val, def) {
		switch (val) {
			case "placeholder":
				return c_oAscDocPartGallery.Placeholder;
			case "any":
				return c_oAscDocPartGallery.Any;
			case "default":
				return c_oAscDocPartGallery.Default;
			case "docParts":
				return c_oAscDocPartGallery.DocParts;
			case "coverPg":
				return c_oAscDocPartGallery.CoverPg;
			case "eq":
				return c_oAscDocPartGallery.Eq;
			case "ftrs":
				return c_oAscDocPartGallery.Ftrs;
			case "hdrs":
				return c_oAscDocPartGallery.Hdrs;
			case "pgNum":
				return c_oAscDocPartGallery.PgNum;
			case "tbls":
				return c_oAscDocPartGallery.Tbls;
			case "watermarks":
				return c_oAscDocPartGallery.Watermarks;
			case "autoTxt":
				return c_oAscDocPartGallery.AutoTxt;
			case "txtBox":
				return c_oAscDocPartGallery.TxtBox;
			case "pgNumT":
				return c_oAscDocPartGallery.PgNumT;
			case "pgNumB":
				return c_oAscDocPartGallery.PgNumB;
			case "pgNumMargins":
				return c_oAscDocPartGallery.PgNumMargins;
			case "tblOfContents":
				return c_oAscDocPartGallery.TblOfContents;
			case "bib":
				return c_oAscDocPartGallery.Bib;
			case "custQuickParts":
				return c_oAscDocPartGallery.CustQuickParts;
			case "custCoverPg":
				return c_oAscDocPartGallery.CustCoverPg;
			case "custEq":
				return c_oAscDocPartGallery.CustEq;
			case "custFtrs":
				return c_oAscDocPartGallery.CustFtrs;
			case "custHdrs":
				return c_oAscDocPartGallery.CustHdrs;
			case "custPgNum":
				return c_oAscDocPartGallery.CustPgNum;
			case "custTbls":
				return c_oAscDocPartGallery.CustTbls;
			case "custWatermarks":
				return c_oAscDocPartGallery.CustWatermarks;
			case "custAutoTxt":
				return c_oAscDocPartGallery.CustAutoTxt;
			case "custTxtBox":
				return c_oAscDocPartGallery.CustTxtBox;
			case "custPgNumT":
				return c_oAscDocPartGallery.CustPgNumT;
			case "custPgNumB":
				return c_oAscDocPartGallery.CustPgNumB;
			case "custPgNumMargins":
				return c_oAscDocPartGallery.CustPgNumMargins;
			case "custTblOfContents":
				return c_oAscDocPartGallery.CustTblOfContents;
			case "custBib":
				return c_oAscDocPartGallery.CustBib;
			case "custom1":
				return c_oAscDocPartGallery.Custom1;
			case "custom2":
				return c_oAscDocPartGallery.Custom2;
			case "custom3":
				return c_oAscDocPartGallery.Custom3;
			case "custom4":
				return c_oAscDocPartGallery.Custom4;
			case "custom5":
				return c_oAscDocPartGallery.Custom5;
		}
		return def;
	}
	function toXml_ST_DocPartGallery(val) {
		switch (val) {
			case c_oAscDocPartGallery.Placeholder:
				return "placeholder";
			case c_oAscDocPartGallery.Any:
				return "any";
			case c_oAscDocPartGallery.Default:
				return "default";
			case c_oAscDocPartGallery.DocParts:
				return "docParts";
			case c_oAscDocPartGallery.CoverPg:
				return "coverPg";
			case c_oAscDocPartGallery.Eq:
				return "eq";
			case c_oAscDocPartGallery.Ftrs:
				return "ftrs";
			case c_oAscDocPartGallery.Hdrs:
				return "hdrs";
			case c_oAscDocPartGallery.PgNum:
				return "pgNum";
			case c_oAscDocPartGallery.Tbls:
				return "tbls";
			case c_oAscDocPartGallery.Watermarks:
				return "watermarks";
			case c_oAscDocPartGallery.AutoTxt:
				return "autoTxt";
			case c_oAscDocPartGallery.TxtBox:
				return "txtBox";
			case c_oAscDocPartGallery.PgNumT:
				return "pgNumT";
			case c_oAscDocPartGallery.PgNumB:
				return "pgNumB";
			case c_oAscDocPartGallery.PgNumMargins:
				return "pgNumMargins";
			case c_oAscDocPartGallery.TblOfContents:
				return "tblOfContents";
			case c_oAscDocPartGallery.Bib:
				return "bib";
			case c_oAscDocPartGallery.CustQuickParts:
				return "custQuickParts";
			case c_oAscDocPartGallery.CustCoverPg:
				return "custCoverPg";
			case c_oAscDocPartGallery.CustEq:
				return "custEq";
			case c_oAscDocPartGallery.CustFtrs:
				return "custFtrs";
			case c_oAscDocPartGallery.CustHdrs:
				return "custHdrs";
			case c_oAscDocPartGallery.CustPgNum:
				return "custPgNum";
			case c_oAscDocPartGallery.CustTbls:
				return "custTbls";
			case c_oAscDocPartGallery.CustWatermarks:
				return "custWatermarks";
			case c_oAscDocPartGallery.CustAutoTxt:
				return "custAutoTxt";
			case c_oAscDocPartGallery.CustTxtBox:
				return "custTxtBox";
			case c_oAscDocPartGallery.CustPgNumT:
				return "custPgNumT";
			case c_oAscDocPartGallery.CustPgNumB:
				return "custPgNumB";
			case c_oAscDocPartGallery.CustPgNumMargins:
				return "custPgNumMargins";
			case c_oAscDocPartGallery.CustTblOfContents:
				return "custTblOfContents";
			case c_oAscDocPartGallery.CustBib:
				return "custBib";
			case c_oAscDocPartGallery.Custom1:
				return "custom1";
			case c_oAscDocPartGallery.Custom2:
				return "custom2";
			case c_oAscDocPartGallery.Custom3:
				return "custom3";
			case c_oAscDocPartGallery.Custom4:
				return "custom4";
			case c_oAscDocPartGallery.Custom5:
				return "custom5";
		}
		return null;
	}

	function fromXml_ST_DocPartType(val, def) {
		switch (val) {
			case "none":
				return c_oAscDocPartType.None;
			case "normal":
				return c_oAscDocPartType.Normal;
			case "autoExp":
				return c_oAscDocPartType.AutoExp;
			case "toolbar":
				return c_oAscDocPartType.Toolbar;
			case "speller":
				return c_oAscDocPartType.Speller;
			case "formFld":
				return c_oAscDocPartType.FormFld;
			case "bbPlcHdr":
				return c_oAscDocPartType.BBPlcHolder;
		}
		return def;
	}
	function toXml_ST_DocPartType(val) {
		switch (val) {
			case c_oAscDocPartType.None:
				return "none";
			case c_oAscDocPartType.Normal:
				return "normal";
			case c_oAscDocPartType.AutoExp:
				return "autoExp";
			case c_oAscDocPartType.Toolbar:
				return "toolbar";
			case c_oAscDocPartType.Speller:
				return "speller";
			case c_oAscDocPartType.FormFld:
				return "formFld";
			case c_oAscDocPartType.BBPlcHolder:
				return "bbPlcHdr";
		}
		return null;
	}
	function fromXml_ST_DocPartBehavior(val, def) {
		switch (val) {
			case "content":
				return c_oAscDocPartBehavior.Content;
			case "p":
				return c_oAscDocPartBehavior.P;
			case "pg":
				return c_oAscDocPartBehavior.Pg;
		}
		return def;
	}
	function toXml_ST_DocPartBehavior(val) {
		switch (val) {
			case c_oAscDocPartBehavior.Content:
				return "content";
			case c_oAscDocPartBehavior.P:
				return "p";
			case c_oAscDocPartBehavior.Pg:
				return "pg";
		}
		return null;
	}
	function fromXml_ST_BrType(val, def) {
		switch (val) {
			case "page":
				return AscWord.break_Page;
			case "column":
				return AscWord.break_Column;
			case "textWrapping":
				return AscWord.break_Line;
		}
		return def;
	}
	function toXml_ST_BrType(val) {
		switch (val) {
			case AscWord.break_Page:
				return "page";
			case AscWord.break_Column:
				return "column";
			case AscWord.break_Line:
				return "textWrapping";
		}
		return null;
	}
	function fromXml_ST_Lock(val, def) {
		switch (val) {
			case "sdtLocked":
				return c_oAscSdtLockType.SdtLocked;
			case "contentLocked":
				return c_oAscSdtLockType.ContentLocked;
			case "unlocked":
				return c_oAscSdtLockType.Unlocked;
			case "sdtContentLocked":
				return c_oAscSdtLockType.SdtContentLocked;
		}
		return def;
	}
	function toXml_ST_Lock(val) {
		switch (val) {
			case c_oAscSdtLockType.SdtLocked:
				return "sdtLocked";
			case c_oAscSdtLockType.ContentLocked:
				return "contentLocked";
			case c_oAscSdtLockType.Unlocked:
				return "unlocked";
			case c_oAscSdtLockType.SdtContentLocked:
				return "sdtContentLocked";
		}
		return null;
	}
	function fromXml_ST_SdtAppearance(val, def) {
		switch (val) {
			case "tags":
				return 0;
			case "boundingBox":
				return Asc.c_oAscSdtAppearance.Frame;
			case "hidden":
				return Asc.c_oAscSdtAppearance.Hidden;
		}
		return def;
	}
	function toXml_ST_SdtAppearance(val) {
		switch (val) {
			case 0:
				return "tags";
			case Asc.c_oAscSdtAppearance.Frame:
				return "boundingBox";
			case Asc.c_oAscSdtAppearance.Hidden:
				return "hidden";
		}
		return null;
	}
	function fromXml_ST_CalendarType(val, def) {
		switch (val) {
			case "gregorian":
				return Asc.c_oAscCalendarType.Gregorian;
			case "gregorianUs":
				return Asc.c_oAscCalendarType.GregorianUs;
			case "gregorianMeFrench":
				return Asc.c_oAscCalendarType.GregorianMeFrench;
			case "gregorianArabic":
				return Asc.c_oAscCalendarType.GregorianArabic;
			case "hijri":
				return Asc.c_oAscCalendarType.Hijri;
			case "hebrew":
				return Asc.c_oAscCalendarType.Hebrew;
			case "taiwan":
				return Asc.c_oAscCalendarType.Taiwan;
			case "japan":
				return Asc.c_oAscCalendarType.Japan;
			case "thai":
				return Asc.c_oAscCalendarType.Thai;
			case "korea":
				return Asc.c_oAscCalendarType.Korea;
			case "saka":
				return Asc.c_oAscCalendarType.Saka;
			case "gregorianXlitEnglish":
				return Asc.c_oAscCalendarType.GregorianXlitEnglish;
			case "gregorianXlitFrench":
				return Asc.c_oAscCalendarType.GregorianXlitFrench;
			case "none":
				return Asc.c_oAscCalendarType.None;
		}
		return def;
	}
	function toXml_ST_CalendarType(val) {
		switch (val) {
			case Asc.c_oAscCalendarType.Gregorian:
				return "gregorian";
			case Asc.c_oAscCalendarType.GregorianUs:
				return "gregorianUs";
			case Asc.c_oAscCalendarType.GregorianMeFrench:
				return "gregorianMeFrench";
			case Asc.c_oAscCalendarType.GregorianArabic:
				return "gregorianArabic";
			case Asc.c_oAscCalendarType.Hijri:
				return "hijri";
			case Asc.c_oAscCalendarType.Hebrew:
				return "hebrew";
			case Asc.c_oAscCalendarType.Taiwan:
				return "taiwan";
			case Asc.c_oAscCalendarType.Japan:
				return "japan";
			case Asc.c_oAscCalendarType.Thai:
				return "thai";
			case Asc.c_oAscCalendarType.Korea:
				return "korea";
			case Asc.c_oAscCalendarType.Saka:
				return "saka";
			case Asc.c_oAscCalendarType.GregorianXlitEnglish:
				return "gregorianXlitEnglish";
			case Asc.c_oAscCalendarType.GregorianXlitFrench:
				return "gregorianXlitFrench";
			case Asc.c_oAscCalendarType.None:
				return "none";
		}
		return null;
	}
	function fromXml_ST_CombFormWidthRule(val, def) {
		switch (val) {
			case "atLeast":
				return Asc.CombFormWidthRule.AtLeast;
			case "auto":
				return Asc.CombFormWidthRule.Auto;
			case "exact":
				return Asc.CombFormWidthRule.Exact;
		}
		return def;
	}
	function toXml_ST_CombFormWidthRule(val) {
		switch (val) {
			case Asc.CombFormWidthRule.AtLeast:
				return "atLeast";
			case Asc.CombFormWidthRule.Auto:
				return "auto";
			case Asc.CombFormWidthRule.Exact:
				return "exact";
		}
		return null;
	}
	function fromXml_ST_TextFormFormatType(val, def) {
		switch (val) {
			case "none":
				return Asc.TextFormFormatType.None;
			case "digit":
				return Asc.TextFormFormatType.Digit;
			case "letter":
				return Asc.TextFormFormatType.Letter;
			case "mask":
				return Asc.TextFormFormatType.Mask;
			case "regExp":
				return Asc.TextFormFormatType.RegExp;
		}
		return def;
	}
	function toXml_ST_TextFormFormatType(val) {
		switch (val) {
			case Asc.TextFormFormatType.None:
				return "none";
			case Asc.TextFormFormatType.Digit:
				return "digit";
			case Asc.TextFormFormatType.Letter:
				return "letter";
			case Asc.TextFormFormatType.Mask:
				return "mask";
			case Asc.TextFormFormatType.RegExp:
				return "regExp";
		}
		return null;
	}
	function fromXml_ST_ComplexFormType(val, def) {
		switch (val) {
			case "custom":
				return Asc.ComplexFormType.Custom;
			case "telephone":
				return Asc.ComplexFormType.Telephone;
			case "email":
				return Asc.ComplexFormType.Email;
		}
		return def;
	}
	function toXml_ST_ComplexFormType(val) {
		switch (val) {
			case Asc.ComplexFormType.Custom:
				return null;//default
			case Asc.ComplexFormType.Telephone:
				return "telephone";
			case Asc.ComplexFormType.Email:
				return "email";
		}
		return null;
	}

	function writeRunText(writer, ns, name, text) {
		if(text) {
			writer.WriteXmlNullableValueStringEncode(ns + name, text);
		}
		return "";
	}
	function WiteMoveRangeXml(writer, revisionMove) {
		let docSaveParams = writer.context.docSaveParams;
		let moveRangeNameToId, name;
		if (revisionMove.IsFrom()) {
			moveRangeNameToId = docSaveParams.moveRangeFromNameToId;
			name = revisionMove.IsStart() ? "w:moveFromRangeStart" : "w:moveFromRangeEnd";
		} else {
			moveRangeNameToId = docSaveParams.moveRangeToNameToId;
			name = revisionMove.IsStart() ? "w:moveToRangeStart" : "w:moveToRangeEnd";
		}
		var revisionId = moveRangeNameToId[revisionMove.GetMarkId()];
		if (undefined === revisionId) {
			revisionId = docSaveParams.trackRevisionId++;
			moveRangeNameToId[revisionMove.GetMarkId()] = revisionId;
		}
		let options = {id: revisionId};
		revisionMove.toXml(writer, name, options);
	}

	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CT_TblGrid = CT_TblGrid;
})(window);
