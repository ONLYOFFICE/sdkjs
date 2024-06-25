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

(function(window, document)
{
	/**
	 * 	Docs:
	 * 	VisioDocument_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/visiodocument_type-complextypevisio-xml
	 * @constructor
	 */
	function CVisioDocument(Api, DrawingDocument, isMainLogicDocument) {
		this.start = null;
		this.key = null;
		this.metric = null;
		this.buildnum = null;
		this.version = null;
		this.docLangID = null;
		this.documentProperties = null;
		this.documentSettings = null;
		this.colors = [];
		this.faceNames = [];
		/**
		 * see loadFonts function
		 * @type {CFont[]}
		 */
		this.loadedFonts = [];
		this.styleSheets = [];
		this.documentSheet = null;
		this.eventList = [];
		this.headerFooter = null;
		this.dataTransferInfo = null;
		this.publishSettings = null;
		this.comments = null;
		this.any = null;
		this.anyAttr = null;

		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;

		// TODO mb consider 'this'(CVisioDocument) contains parts(.xml files) like document.xml and windows.xml
		// only but not XMLmethods and call class representing document.xml VisioDocument_Type
		// this.visioDocument_Type = null;
		this.windows = null;
		this.masters = null;
		this.masterContents = [];
		this.pages = null;
		this.pageContents = [];
		this.themes = [];
		this.app = null;
		this.core = null;
		this.customProperties = null;
		this.thumbnail = null;
		this.commentsPart = null;
		this.extensions = null;
		this.dataConnections = null;
		this.dataRecordSets = null;
		this.validation = null;

		// Not realized, file defines schema and data of that schema
		this.solutions = null;
		// solution contents
		this.solutionXMLs = [];

		/**
		 * not parsed data. inits in OpenDocumentFromZipNoInit, xmlParserContext.loadDataLinks()
		 * {index: number, path: string}
		 * @type {{}}
		 */
		this.imageMap = null;

		// unfinished
		// this.EmbeddedData = null;

		//------------------------------------------------------------------------------------------------------------------
		//  Сохраняем ссылки на глобальные объекты
		//------------------------------------------------------------------------------------------------------------------
		this.history              = History;
		this.idCounter            = AscCommon.g_oIdCounter;
		this.tableId              = AscCommon.g_oTableId;
		// this.collaborativeEditing = (("undefined" !== typeof(AscCommon.CWordCollaborativeEditing) && AscCommon.CollaborativeEditing instanceof AscCommon.CWordCollaborativeEditing) ? AscCommon.CollaborativeEditing : null);
		this.api                  = Api;
		this.DrawingDocument = DrawingDocument
		//------------------------------------------------------------------------------------------------------------------
		//  Выставляем ссылки на главный класс
		//------------------------------------------------------------------------------------------------------------------
		if (false !== isMainLogicDocument)
		{
			// if (this.history)
			// 	this.history.Set_LogicDocument(this);
		}
		this.mainDocument = false !== isMainLogicDocument;
		this.zoom = 100;

		//stubs for compatibility with DocumentContent
		AscCommon.mockLogicDoc(CVisioDocument.prototype);
	}

	/**
	 * TODO Check thumbnail parse in fromZip and setData in toZip
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.fromZip = function(zip, context, oReadResult) {
		// Maybe it should be moved to 	sdkjs-ooxml/visio/Editor/SerializeXml.js like in 'word' case?
		// 'word' case: 								sdkjs-ooxml/word/Editor/SerializeXml.js
		context.zip = zip;

		let reader;
		let doc = new AscCommon.openXml.OpenXmlPackage(zip, null);

		parseApp.call(this, doc, reader, context);
		parseCore.call(this, doc, reader, context);
		parseCustomProperties.call(this, doc, reader, context);
		parseThumbnail.call(this, doc, reader, context);

		let documentPart = doc.getPartByRelationshipType(AscCommon.openXml.Types.visioDocument.relationType);
		if (documentPart) {
			let contentDocument = documentPart.getDocumentContent();
			reader = new StaxParser(contentDocument, documentPart, context);
			this.fromXml(reader);
			// TODO mb consider 'this' contains parts(.xml files) only but not XML like document.xml and windows.xml
			// this.visioDocument_Type = new AscCommonDraw.VisioDocument_Type();
			// this.visioDocument_Type.fromXml(reader);

			parseWindows.call(this, documentPart, reader, context);
			parseMasters.call(this, documentPart, reader, context);
			parsePages.call(this, documentPart, reader, context);
			parseThemes.call(this, documentPart, reader, context);
			parseComments.call(this, documentPart, reader, context);
			parseExtensions.call(this, documentPart, reader, context);
			parseDataConnections.call(this, documentPart, reader, context);
			parseDataRecordSets.call(this, documentPart, reader, context);
			parseValidation.call(this, documentPart, reader, context);
			// Not realized, file defines schema and data of that schema
			parseSolutions.call(this, documentPart, reader, context);
		}
		// unfinished
		// saveEmbeddedData.call(this, doc);
		// handleEmbeddedDataRels.call(this, zip);
	};

	/**
	 * 	TODO mb rewrite consider 'CVisioDocument' contains parts(.xml files) only but not XML
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.toZip = function(zip, context) {
		let memory = new AscCommon.CMemory();
		memory.SetXmlAttributeQuote(0x27);
		memory.context = context;
		context.document = this;

		let filePart = new AscCommon.openXml.OpenXmlPackage(zip, memory);

		let docPart = filePart.addPart(AscCommon.openXml.Types.visioDocument);
		let appPart = filePart.addPart(AscCommon.openXml.Types.extendedFileProperties);
		let corePart = filePart.addPart(AscCommon.openXml.Types.coreFileProperties);
		let customPrPart = filePart.addPart(AscCommon.openXml.Types.customFileProperties);
		let thumbNailPart = filePart.addPart(AscCommon.openXml.Types.thumbnail);
		let windowsPart = docPart.part.addPart(AscCommon.openXml.Types.visioDocumentWindows);
		let mastersPart = docPart.part.addPart(AscCommon.openXml.Types.masters);
		let themesPart = docPart.part.addPart(AscCommon.openXml.Types.theme);
		let commentsPart = docPart.part.addPart(AscCommon.openXml.Types.visioComments);
		let extensionsPart = docPart.part.addPart(AscCommon.openXml.Types.visioExtensions);
		let dataConnectionsPart = docPart.part.addPart(AscCommon.openXml.Types.visioDataConnections);
		let dataRecordSetsPart = docPart.part.addPart(AscCommon.openXml.Types.visioDataRecordSets);
		let validationPart = docPart.part.addPart(AscCommon.openXml.Types.validation);
		// Not realized, file defines schema and data of that schema
		let solutionsPart = docPart.part.addPart(AscCommon.openXml.Types.solutions);

		for (let i = 0; i < this.masterContents.length; i++) {
			let masterContent = mastersPart.part.addPart(AscCommon.openXml.Types.master);
			masterContent.part.setDataXml(this.masterContents[i], memory);
		}

		let pagesPart = docPart.part.addPart(AscCommon.openXml.Types.pages);

		for (let i = 0; i < this.pageContents.length; i++) {
			let pageContent = pagesPart.part.addPart(AscCommon.openXml.Types.page);
			pageContent.part.setDataXml(this.pageContents[i], memory);

			// I add page[N].xml.rels below
			// It has links to all masters but
			// in examplevsdx file in page[N].xml.rels rId[N] states to random master[N]
			// e.g. rId3 to ../masters/master1.xml
			// here rId1 will state to master1, rId2 to master2, etc.
			// TODO check if this is important
			// in page[N].xml there is no rId used only <Shape ... Master="[ID]">
			// e. g. <Shape ... Master="1">
			for (let i = 0; i < this.masterContents.length; i++) {
				pageContent.part.addRelationship(AscCommon.openXml.Types.masterFromPage.relationType,
					"../masters/master" + (i + 1) + ".xml");
			}
		}

		// Not realized, file defines schema and data of that schema
		for (let i = 0; i < this.solutionXMLs.length; i++) {
			let solutionContent = solutionsPart.part.addPart(AscCommon.openXml.Types.solution);
			solutionContent.part.setDataXml(this.solutionXMLs[i], memory);
		}

		docPart.part.setDataXml(this, memory);
		appPart.part.setDataXml(this.app, memory);
		corePart.part.setDataXml(this.core, memory);
		if (this.customProperties) {
			// if Custom part exists
			customPrPart.part.setDataXml(this.customProperties, memory);
		}
		if (this.thumbnail) {
			thumbNailPart.part.setData(this.thumbnail, memory);
		}
		if (this.windows) {
			// if Windows part exists
			windowsPart.part.setDataXml(this.windows, memory);
		}
		if (this.masters) {
			mastersPart.part.setDataXml(this.masters, memory);
		}
		pagesPart.part.setDataXml(this.pages, memory);
		for (let i = 0; i < this.themes.length; i++) {
			let themeContent = themesPart.part.addPart(AscCommon.openXml.Types.theme);
			themeContent.part.setDataXml(this.themes[i], memory);
		}
		if (this.commentsPart) {
			commentsPart.part.setDataXml(this.commentsPart, memory);
		}
		if (this.extensions) {
			extensionsPart.part.setDataXml(this.extensions, memory);
		}
		if (this.dataConnections) {
			dataConnectionsPart.part.setDataXml(this.dataConnections, memory);
		}
		if (this.dataRecordSets) {
			dataRecordSetsPart.part.setDataXml(this.dataRecordSets, memory);
		}
		if (this.validation) {
			validationPart.part.setDataXml(this.validation, memory);
		}
		// Not realized, file defines schema and data of that schema
		if (this.solutions) {
			solutionsPart.part.setDataXml(this.solutions, memory);
		}
		memory.Seek(0);
	};

	/**
	 * @memberOf CVisioDocument
	 * @return {number}
	 */
	CVisioDocument.prototype.getObjectType = function() {
		//to be parent of shape
		return 0;
	};

	/**
	 * get zoom from 0 to 100
	 * @memberOf CVisioDocument
	 * @param pageIndex
	 * @param displayedWidthPx
	 * @param displayedHeightPX
	 * @return {number}
	 */
	CVisioDocument.prototype.getFitZoomValue = function(pageIndex, displayedWidthPx, displayedHeightPX) {
		let logic_w_mm = this.GetWidthMM(pageIndex);
		let logic_h_mm = this.GetHeightMM(pageIndex);

		var _value = 100;

		var _pageWidth  = logic_w_mm * g_dKoef_mm_to_pix;
		var _pageHeight = logic_h_mm * g_dKoef_mm_to_pix;

		var _hor_Zoom = 100;
		if (0 != _pageWidth)
			_hor_Zoom = (100 * displayedWidthPx) / _pageWidth;
		var _ver_Zoom = 100;
		if (0 != _pageHeight)
			_ver_Zoom = (100 * displayedHeightPX) / _pageHeight;

		_value = Math.min(_hor_Zoom, _ver_Zoom);

		return _value;
	};

	/**
	 * @param pageIndex
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.GetWidthMM = function(pageIndex) {
		//todo units, indexes
		let logic_w_inch = this.pages.page[pageIndex].pageSheet.getCellNumberValue("PageWidth");
		return logic_w_inch * g_dKoef_in_to_mm;
	}


	/**
	 * @param pageIndex
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.GetHeightMM = function(pageIndex) {
		let logic_h_inch = this.pages.page[pageIndex].pageSheet.getCellNumberValue("PageHeight");
		return logic_h_inch * g_dKoef_in_to_mm;
	}

	/**
	 * Load fonts which are used in document and do callback (can be used to call CVisioDocument.prototype.draw).
	 * Inits this.loadedFonts.
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.loadFonts = function() {
		let api = this.api;

		let aFonts = this.loadedFonts;
		// load Arial by default
		aFonts.push(new AscFonts.CFont("Arial", 0, "", 0));
		let newFontIndex = 1;
		this.faceNames.forEach(function (faceName_Type) {
			let nameU = faceName_Type.nameU;
			let fontInfo = api.FontLoader.fontInfos.find(function(cFontInfo) {
				return cFontInfo.Name === nameU;
			});
			if (fontInfo === undefined || fontInfo === null) {
				console.log("Unknown font used in visio file");
			} else {
				console.log("Font", nameU, "will be loaded");
				aFonts.push(new AscFonts.CFont(nameU, newFontIndex, "", 0));
				newFontIndex += 1;
			}
		});

		// may immediately call callback
		api.FontLoader.LoadDocumentFonts(aFonts, false);
	}


	/**
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.draw = function(Zoom, pGraphics) {
		//HOTFIX
		this.theme = this.themes[0];

		let pageIndex = 0;
		let pageInfo = this.pages.page[pageIndex];
		let pageContent = this.pageContents[pageIndex];

		let topLevelShapesAndGroups = this.convertToCShapesAndGroups(pageInfo, pageContent);

		let api = this.api;
		let logic_w_mm = this.GetWidthMM(pageIndex);
		let logic_h_mm = this.GetHeightMM(pageIndex);

		/**
		 * sometimes text suddenly gets huge sizes that take all of a picture space.
		 * It usually happens on low zoom (when we scroll down a lot). This coefficient fixes this issue but
		 * makes extra space and scroll lines get wrong sizes. fixScale number was randomly selected.
		 * High fixScale value can make high load and break page, especially on high zooming. Low values
		 * may not fix text on low zoom. 1 is default.
		 * @type {number}
		 */
		let fixScale = 1;

		let graphics;

		let useFitToScreenZoom = !pGraphics;
		let pageScale;
		if (useFitToScreenZoom) {
			Zoom = Zoom * this.getFitZoomValue(pageIndex, api.HtmlElement.offsetWidth,api.HtmlElement.offsetHeight) / 100 * fixScale;
		}
		pageScale = Zoom / 100;

		if (pGraphics) {
			graphics = pGraphics;
		} else {
			let dKoef = pageScale * g_dKoef_mm_to_pix;
			dKoef *= AscCommon.AscBrowser.retinaPixelRatio;

			let w_mm = logic_w_mm;
			let h_mm = logic_h_mm;

			var w_px = (w_mm * dKoef + 0.5) >> 0;
			var h_px = (h_mm * dKoef + 0.5) >> 0;

			let _canvas = api.canvas;

			// Version 1
			// let parentElement = _canvas.parentElement;
			// _canvas.style.width  = parentElement.offsetWidth + "px";
			// _canvas.style.height = parentElement.offsetHeight + "px";

			// Version 2 with correct scroll lines
			// setup scroll lines
			_canvas.style.width  = w_px + "px";
			_canvas.style.height = h_px + "px";
			// set pixels count for width and height
			_canvas.width = AscCommon.AscBrowser.convertToRetinaValue(_canvas.clientWidth, true);
			_canvas.height = AscCommon.AscBrowser.convertToRetinaValue(_canvas.clientHeight, true);

			// canvas#id_viewer_overlay and div#id_target_cursor creates empty gray space below any drawing

			AscCommon.calculateCanvasSize(_canvas);

			let ctx = _canvas.getContext('2d');

			graphics = new AscCommon.CGraphics();
			graphics.init(ctx, w_px, h_px, w_mm, h_mm);
			graphics.m_oFontManager = AscCommon.g_fontManager;
		}

		if (graphics.m_oContext) {
			graphics.m_oContext.clearRect(0, 0, api.canvas.width, api.canvas.height);
		}

		//visio y coordinate goes up while
		//ECMA-376-11_5th_edition and Geometry.js y coordinate goes down
		let baseMatrix = new AscCommon.CMatrix();
		// baseMatrix.SetValues(1, 0, 0, 1, 0, 0);
		baseMatrix.SetValues(1 / fixScale, 0, 0, -1 / fixScale, 0, logic_h_mm / fixScale);
		graphics.SetBaseTransform(baseMatrix);

		let baseTextMatrix = new AscCommon.CMatrix();
		baseTextMatrix.SetValues(1 / fixScale, 0, 0, 1 / fixScale, 0, 0);
		// baseTextMatrix.SetValues(1, 0, 0, -1, 0, logic_h_mm);

		/**
		 * @type {boolean}
		 */
		let changeTextDirection = true;


		graphics.SaveGrState();
		graphics.SetIntegerGrid(false);
		graphics.transform3(new AscCommon.CMatrix());
		graphics.b_color1( 255, 255, 255, 255 );
		graphics.rect( 0, 0, logic_w_mm, logic_h_mm );
		graphics.df();
		graphics.RestoreGrState();

		// see sdkjs/common/Shapes/Serialize.js this.ReadGroupShape = function(type) to
		// learn how to work with shape groups
		function drawShapeOrGroupRecursively(shapeOrGroup, currentGroupHandling) {
			if (shapeOrGroup.spTree) {
				// group came to argument
				/** @type CGroupShape */
				let group = shapeOrGroup;
				// if we use CGroupShape.draw it doesn't draw group geometry, only its children

				// draw group geometry
				graphics.SaveGrState();
				graphics.SetIntegerGrid(false);

				graphics.transform3(group.transform);

				// create shape to draw group geometry
				let сGroupShape = new AscFormat.CShape();
				// cShape.setParent();
				сGroupShape.setLocks(0);
				сGroupShape.setBDeleted(false);

				сGroupShape.setSpPr(group.spPr);
				сGroupShape.spPr.setParent(сGroupShape);
				сGroupShape.rot = group.rot;
				сGroupShape.Id = group.Id;
				сGroupShape.brush = group.brush;
				сGroupShape.bounds = group.bounds;
				сGroupShape.flipH = group.flipH;
				сGroupShape.flipV = group.flipV;
				сGroupShape.localTransform = group.localTransform;
				сGroupShape.pen = group.pen;


				let shape_drawer = new AscCommon.CShapeDrawer();
				shape_drawer.fromShape2(сGroupShape, graphics, сGroupShape.getGeometry());
				let groupGeometry = сGroupShape.getGeometry();
				shape_drawer.draw(groupGeometry);

				shape_drawer.Clear();
				graphics.RestoreGrState();

				// handle group children
				group.spTree.forEach(function(shapeOrGroup) {
					drawShapeOrGroupRecursively(shapeOrGroup, group);
				});
			} else {
				// shape came to argument

				// flip images
				if (shapeOrGroup.constructor.name === "CImageShape") {
					shapeOrGroup.transform.sy = -1;
					shapeOrGroup.transform.ty += shapeOrGroup.spPr.xfrm.extY;
				}

				if (changeTextDirection && shapeOrGroup.Id.substring(shapeOrGroup.Id.length - 4) === "Text") {
					graphics.SetBaseTransform(baseTextMatrix);

					// let reflectedTextHeight;
					// if (shapeOrGroup.txBody.content.Content[0].Lines.length === 2) {
					// 	reflectedTextHeight = shapeOrGroup.txBody.content.Content[0].Lines[0].Metrics.TextAscent;
					// } else if (shapeOrGroup.txBody.content.Content[0].Lines.length > 2) {
					// 	let contentHeight = shapeOrGroup.txBody.content.GetSummaryHeight();
					// 	let emptyLineHeight = shapeOrGroup.txBody.content.Content[0].Lines.slice(-1)[0].Bottom - shapeOrGroup.txBody.content.Content[0].Lines.slice(-2)[0].Bottom;
					// 	let lastNonEmptyLineDescent = shapeOrGroup.txBody.content.Content[0].Lines.slice(-2)[0].Metrics.Descent;
					// 	reflectedTextHeight = contentHeight - lastNonEmptyLineDescent - emptyLineHeight;
					// } else {
					// 	reflectedTextHeight = 0;
					// }
					// shapeOrGroup.transformText.sy = -1;
					// shapeOrGroup.spPr.xfrm.rot = Math.PI;
					// shapeOrGroup.recalculateTransform();
					// shapeOrGroup.recalculateTransformText();
					// shapeOrGroup.recalculateLocalTransform(shapeOrGroup.transform);
					// shapeOrGroup.recalculate();
					// shapeOrGroup.transformText.ty = shapeOrGroup.transformText.ty + reflectedTextHeight * 2;

					// shapeOrGroup.transformText.ty = 0;

					// shapeOrGroup.transformText.sy = -1;
					// let offset = shapeOrGroup.transform.ty - shapeOrGroup.transformText.ty; // > 0
					// let correctTextTy = shapeOrGroup.transform.ty + offset;
					shapeOrGroup.transform.ty = logic_h_mm - shapeOrGroup.transform.ty - shapeOrGroup.spPr.xfrm.extY;
					// let oContent = shapeOrGroup.getDocContent();
					// let height = oContent.GetSummaryHeight();
					// shapeOrGroup.transform.ty += height/2;
					// shapeOrGroup.recalculate();
					// shapeOrGroup.recalculateContent2();
					// shapeOrGroup.transformText.ty = logic_h_mm - shapeOrGroup.transformText.ty;
					// shapeOrGroup.transformText.ty = 10; // mm
					shapeOrGroup.recalculateTransformText();
					// transformText before recalculateTransformText doesn't change anything
					// let scale = 50
					// shapeOrGroup.transformText.sx = scale;
					// shapeOrGroup.transformText.sy = scale;
					// shapeOrGroup.transformText.ty -= shapeOrGroup.spPr.xfrm.extY * scale / 2;
					// shapeOrGroup.transformText.tx -= shapeOrGroup.spPr.xfrm.extX * scale / 2;


					// graphics.SaveGrState();
					// graphics.SetIntegerGrid(false);
					// graphics.transform3(new AscCommon.CMatrix());
					// graphics.b_color1( 255, 0, 0, 255 );
					// graphics.rect( shapeOrGroup.transform.tx, shapeOrGroup.transform.ty, shapeOrGroup.spPr.xfrm.extX, shapeOrGroup.spPr.xfrm.extY );
					// // graphics.rect( shapeOrGroup.transform.tx, shapeOrGroup.transform.ty, 10, 1 );
					// graphics.df();
					// graphics.RestoreGrState();
					//
					// graphics.SaveGrState();
					// graphics.SetIntegerGrid(false);
					// graphics.transform3(new AscCommon.CMatrix());
					// graphics.b_color1( 0, 255, 0, 100 );
					// graphics.rect( shapeOrGroup.transformText.tx, shapeOrGroup.transformText.ty, 10, 1);
					// graphics.df();
					// graphics.RestoreGrState();
				}

				shapeOrGroup.draw(graphics, shapeOrGroup.transform, shapeOrGroup.transformText);

				if (changeTextDirection && shapeOrGroup.Id.substring(shapeOrGroup.Id.length - 4) === "Text") {
					graphics.SetBaseTransform(baseMatrix);
				}
			}
		}

		topLevelShapesAndGroups.forEach(function(shapeOrGroup) {
			drawShapeOrGroupRecursively(shapeOrGroup);
		});
	};
	function getRandomPrst() {
		let types = AscCommon.g_oAutoShapesTypes[Math.floor(Math.random()*AscCommon.g_oAutoShapesTypes.length)];
		return types[Math.floor(Math.random()*types.length)].Type;
	}

	/**
	 * @memberOf CVisioDocument
	 * @param pageInfo
	 * @param pageContent
	 * @return {(CShape | CGroupShape | CImageShape)[]} topLevelShapesAndGroups
	 */
	CVisioDocument.prototype.convertToCShapesAndGroups = function(pageInfo, pageContent) {
		/** @type {Shape_Type[]} */
		let shapeClasses = [];

		/** @type {(CShape | CGroupShape | CImageShape)[]} */
		let topLevelShapesAndGroups = [];

		let masters = this.joinMastersInfoAndContents();

		for(let i = 0; i < pageContent.shapes.length; i++) {
			let shape = pageContent.shapes[i];

			shape.realizeMasterInheritanceRecursively(masters);
			shape.realizeStyleInheritanceRecursively(this.styleSheets);

			if (shape.type === "Group") {
				let cGroupShapeAndText = shape.toCGroupShapeRecursively(this,pageInfo);
				topLevelShapesAndGroups.push(cGroupShapeAndText.cGroupShape);
				if (cGroupShapeAndText.textCShape) {
					topLevelShapesAndGroups.push(cGroupShapeAndText.textCShape);
				}
			} else {
				let cShapes = shape.toGeometryAndTextCShapes(this, pageInfo);
				topLevelShapesAndGroups.push(cShapes.geometryCShape);
				if (cShapes.textCShape !== null) {
					topLevelShapesAndGroups.push(cShapes.textCShape);
				}
			}
		}

		return topLevelShapesAndGroups;
	};

	/**
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.joinMastersInfoAndContents = function() {
		// join Master_Type and MasterContents_Type
		if (this.masters === null || this.masters === undefined) {
			return [];
		}
		let masterFromMastersInfoArray = this.masters.master;
		let master = null;
		let mastersJoined = [];

		let thisContext = this;
		for (let i = 0; i < masterFromMastersInfoArray.length; i++) {
			const masterFromMasters = masterFromMastersInfoArray[i];
			let masterFromMastersArrayRelId = masterFromMasters.rel.id;
			// TODO find file by relationships
			let masterContentNum = +masterFromMastersArrayRelId.match(/\d+/)[0];
			let masterContent = thisContext.masterContents[masterContentNum - 1];
			master = masterFromMasters;
			master.content = masterContent;
			mastersJoined.push(master);
		}
		return mastersJoined;
	}

	// CVisioDocument.prototype.getMasterByID = function(ID) {
	// 	// join Master_Type and MasterContents_Type
	// 	let masterFromMastersArray = this.masters.master;
	// 	let master = null;
	//
	// 	let thisContext = this;
	// 	for (let i = 0; i < masterFromMastersArray.length; i++) {
	// 		const masterFromMasters = masterFromMastersArray[i];
	// 		if (masterFromMasters.iD === ID) {
	// 			let masterFromMastersArrayRelId = masterFromMasters.rel.id;
	// 			// TODO find file by relationships
	// 			let masterContentNum = +masterFromMastersArrayRelId.match(/\d+/)[0];
	// 			let masterContent = thisContext.masterContents[masterContentNum - 1];
	// 			master = masterFromMasters;
	// 			master.content = masterContent;
	// 			break;
	// 		}
	// 	}
	// 	return master;
	// }
	//
	//

	// Main classes for reading

	// Docs:
	// Windows_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/windows_type-complextypevisio-xml
	function CWindows() {
		this.clientWidth = null;
		this.clientHeight = null;
		this.window = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
	// Masters_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/masters_type-complextypevisio-xml
	function CMasters() {
		this.master = [];
		this.masterShortcut = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Another name in docs PageContents_Type
	function CMasterContents() {
		this.shapes = [];
		this.connects = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
	// Pages_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pages_type-complextypevisio-xml
	function CPages() {
		this.page = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
	// Элемент Shapes (PageContents_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/shapes-element-pagecontents_type-complextypevisio-xml
	// PageContents_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pagecontents_type-complextypevisio-xml
	function CPageContents() {
		/**
		 *
		 * @type {Shape_Type[]}
		 */
		this.shapes = [];
		this.connects = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
// Comments_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/comments_type-complextypevisio-xml
	function CComments() {
		this.showCommentTags = null;
		this.authorList = [];
		this.commentList = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
// Extensions_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/extensions_type-complextypevisio-xml
	function CExtensions() {
		this.cellDef = [];
		this.functionDef = [];
		this.sectionDef = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
// DataConnections_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/dataconnections_type-complextypevisio-xml
	function CDataConnections() {
		this.nextID = null;
		this.dataConnection = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs:
	// DataRecordSets_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datarecordsets_type-complextypevisio-xml
	function CDataRecordSets() {
		this.nextID = null;
		this.activeRecordsetID = null;
		this.dataWindowOrder = null;
		this.dataRecordSet = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Docs old:
	// Validation_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/validation_type-complextypevisio-xml
	function CValidation() {
		this.validationProperties = null;
		this.ruleSets = [];
		this.issues = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	// Not realized, file defines schema and data of that schema
	// Docs:
	// Solutions_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/solutions_type-complextypevisio-xml
	function CSolutions() {
		this.solution = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
		return this;
	}

	function CSolutionXML() {
		this.name = null;
		// string containing overall xml
		this.fileContents = null;
		return this;
	}

	function parseApp(doc, reader, context) {
		let appPart = doc.getPartByRelationshipType(AscCommon.openXml.Types.extendedFileProperties.relationType);
		if (appPart) {
			let appContent = appPart.getDocumentContent();
			reader = new StaxParser(appContent, appPart, context);
			this.app = new AscCommon.CApp();
			this.app.fromXml(reader, true);
		}
	}

	function parseCore(doc, reader, context) {
		let corePart = doc.getPartByRelationshipType(AscCommon.openXml.Types.coreFileProperties.relationType);
		if (corePart) {
			let coreContent = corePart.getDocumentContent();
			reader = new StaxParser(coreContent, corePart, context);
			this.core = new AscCommon.CCore();
			this.core.fromXml(reader, true);
		}
	}

	function parseCustomProperties(doc, reader, context) {
		let customPrPart = doc.getPartByRelationshipType(AscCommon.openXml.Types.customFileProperties.relationType);
		if (customPrPart) {
			let customPrPartContent = customPrPart.getDocumentContent();
			reader = new StaxParser(customPrPartContent, customPrPart, context);
			this.customProperties = new AscCommon.CCustomProperties();
			this.customProperties.fromXml(reader, true);
		}
	}

	function parseThumbnail(doc, reader, context) {
		let thumbnailPart = doc.getPartByRelationshipType(AscCommon.openXml.Types.thumbnail.relationType);
		if (thumbnailPart) {
			let thumbnailPartContent = thumbnailPart.getDocumentContent();
			this.thumbnail = thumbnailPartContent;
		}
	}

	function parseWindows(documentPart, reader, context) {
		let windowsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.visioDocumentWindows.relationType);
		if (windowsPart) {
			let contentWindows = windowsPart.getDocumentContent();
			reader = new StaxParser(contentWindows, windowsPart, context);
			this.windows = new CWindows();
			this.windows.fromXml(reader);
		}
	}

	function parseMasters(documentPart, reader, context) {
		let mastersPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.masters.relationType);
		if (mastersPart) {
			let contentMasters = mastersPart.getDocumentContent();
			reader = new StaxParser(contentMasters, mastersPart, context);
			this.masters = new CMasters();
			this.masters.fromXml(reader);

			let masters = mastersPart.getPartsByRelationshipType(AscCommon.openXml.Types.master.relationType);
			if (masters.length > 0) {
				// order is important so sort masters using uri
				let mastersSort = [];
				for (let i = 0; i < masters.length; i++) {
					let masterNumber = +masters[i].uri.match(/\d+/)[0]; // for master3.xml we get 3
					if (!isNaN(parseFloat(masterNumber)) && !isNaN(masterNumber - 0)) {
						// if masterNumber is number
						mastersSort[masterNumber - 1] = masters[i];
					} else {
						console.log('check sdkjs/draw/model/VisioDocument.js : parseMasters');
						mastersSort = masters;
						break;
					}
				}

				masters = mastersSort;
				for (let i = 0; i < masters.length; i++) {
					let masterPart = masters[i];
					let contentMaster = masterPart.getDocumentContent();
					reader = new StaxParser(contentMaster, masterPart, context);
					let MasterContent = new CMasterContents();
					MasterContent.fromXml(reader);
					this.masterContents.push(MasterContent);
				}
			}
		}
	}

	function parsePages(documentPart, reader, context) {
		let pagesPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.pages.relationType);
		if (pagesPart) {
			let contentPages = pagesPart.getDocumentContent();
			reader = new StaxParser(contentPages, pagesPart, context);
			this.pages = new CPages();
			this.pages.fromXml(reader);

			let pages = pagesPart.getPartsByRelationshipType(AscCommon.openXml.Types.page.relationType);
			if (pages.length  > 0) {
				// order is important so sort masters using uri
				let pagesSort = [];
				for (let i = 0; i < pages.length; i++) {
					let pageNumber = pages[i].uri.slice(-5)[0]; // for page3.xml we get 3
					if (!isNaN(parseFloat(pageNumber)) && !isNaN(pageNumber - 0)) {
						// if masterNumber is number
						pagesSort[pageNumber - 1] = pages[i];
					} else {
						console.log('check sdkjs/draw/model/VisioDocument.js : parsePages');
						pagesSort = pages;
						break;
					}
				}
				pages = pagesSort;
				for (let i = 0; i < pages.length; i++) {
					let pagePart = pages[i];
					let contentPage = pagePart.getDocumentContent();
					reader = new StaxParser(contentPage, pagePart, context);
					let PageContent = new CPageContents();
					PageContent.fromXml(reader);
					this.pageContents.push(PageContent);
				}
			}
		}
	}

	function parseThemes(documentPart, reader, context) {
		let themeParts = documentPart.getPartsByRelationshipType(AscCommon.openXml.Types.theme.relationType);
		if (themeParts.length > 0) {
			// order is important so sort themes using uri
			let themesSort = [];
			for (let i = 0; i < themeParts.length; i++) {
				let themeNumber = +themeParts[i].uri.match(/\d+/)[0];
				if (!isNaN(parseFloat(themeNumber)) && !isNaN(themeNumber)) {
					// if themeNumber is number
					themesSort[themeNumber - 1] = themeParts[i];
				} else {
					console.log('check sdkjs/draw/model/VisioDocument.js : parseThemes');
					themesSort = themeParts;
					break;
				}
			}
			themeParts = themesSort;
			for (let i = 0; i < themeParts.length; i++) {
				let themePart = themeParts[i];
				let themePartContent = themePart.getDocumentContent();
				reader = new StaxParser(themePartContent, themePart, context);
				let theme = new AscFormat.CTheme();
				theme.fromXml(reader, true);
				this.themes.push(theme);
			}
		} else {
			this.themes.push(AscFormat.GenerateDefaultTheme(null, null));
		}
	}

	function parseComments(documentPart, reader, context) {
		let commentsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.visioComments.relationType);
		if (commentsPart) {
			let commentsPartContent = commentsPart.getDocumentContent();
			reader = new StaxParser(commentsPartContent, commentsPart, context);
			this.commentsPart = new CComments();
			this.commentsPart.fromXml(reader, true);
		}
	}

	function parseExtensions(documentPart, reader, context) {
		let extensionsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.visioExtensions.relationType);
		if (extensionsPart) {
			let extensionsPartContent = extensionsPart.getDocumentContent();
			reader = new StaxParser(extensionsPartContent, extensionsPart, context);
			this.extensions = new CExtensions();
			this.extensions.fromXml(reader, true);
		}
	}

	function parseDataConnections(documentPart, reader, context) {
		let dataConnectionsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.visioDataConnections.relationType);
		if (dataConnectionsPart) {
			let dataConnectionsPartContent = dataConnectionsPart.getDocumentContent();
			reader = new StaxParser(dataConnectionsPartContent, dataConnectionsPart, context);
			this.dataConnections = new CDataConnections();
			this.dataConnections.fromXml(reader, true);
		}
	}

	function parseDataRecordSets(documentPart, reader, context) {
		let dataRecordSetsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.visioDataRecordSets.relationType);
		if (dataRecordSetsPart) {
			let dataRecordSetsPartContent = dataRecordSetsPart.getDocumentContent();
			reader = new StaxParser(dataRecordSetsPartContent, dataRecordSetsPart, context);
			this.dataRecordSets = new CDataRecordSets();
			this.dataRecordSets.fromXml(reader, true);
		}
	}

	function parseValidation(documentPart, reader, context) {
		let validationPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.validation.relationType);
		if (validationPart) {
			let validationPartContent = validationPart.getDocumentContent();
			reader = new StaxParser(validationPartContent, validationPart, context);
			this.validation = new CValidation();
			this.validation.fromXml(reader, true);
		}
	}

	// Not realized, file defines schema and data of that schema
	function parseSolutions(documentPart, reader, context) {
		let solutionsPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.solutions.relationType);
		if (solutionsPart) {
			let solutionsPartContent = solutionsPart.getDocumentContent();
			reader = new StaxParser(solutionsPartContent, solutionsPart, context);
			this.solutions = new CSolutions();
			this.solutions.fromXml(reader, true);

			let solutions = solutionsPart.getPartsByRelationshipType(AscCommon.openXml.Types.solution.relationType);
			if (solutions.length > 0) {
				// order is important so sort masters using uri
				let solutionsSort = [];
				for (let i = 0; i < solutions.length; i++) {
					let solutionNumber = +solutions[i].uri.match(/\d+/)[0];
					if (!isNaN(parseFloat(solutionNumber)) && !isNaN(solutionNumber - 0)) {
						// if masterNumber is number
						solutionsSort[solutionNumber - 1] = solutions[i];
					} else {
						console.log('check sdkjs/draw/model/VisioDocument.js : parseSolutions');
						solutionsSort = solutions;
						break;
					}
				}
				solutions = solutionsSort;
				for (let i = 0; i < solutions.length; i++) {
					let solutionPart = solutions[i];
					let contentSolution = solutionPart.getDocumentContent();
					reader = new StaxParser(contentSolution, solutionPart, context);
					let solutionContent = new CSolutionXML();
					solutionContent.fromXml(reader);
					this.solutionXMLs.push(solutionContent);
				}
			}
		}
	}

	// function handleEmbeddedDataRels(fullDocPart) {
	// 	// unfinished
	// 	// Proposal: find embedded data files related to parts
	//	// and add links to this.EmbeddedData(see below)[n] consider embedded data file path from it
	// 	// save this links so varibles like StyleSheet.embeddedData['rId1'] or
	//	// Document.embeddedData['rId2']
	// 	let relationTypes = ["http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"];
	//
	// 	let parts = fullDocPart.getParts();
	// 	let partsWithForeignDataLinks = parts.filter(function findPartsWhichLinksToForeignData(part) {
	// 		return -1 !== part.getRelationships().findIndex(function (relationship) {
	// 			return relationTypes.includes(relationship.relationshipType);
	// 		})
	// 	});
	//
	// 	let a = 1;
	// }

	// function saveEmbeddedData(zip) {
	// 	// unfinished
	// 	// Proposal: save embedded files (e.g. .xmls .emf, ...) data like base64 string
	// 	// Create obj with that string, filepath, and filename
	// 	// to this(VisioDocument) like this.EmbeddedData = [];
	// }


	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonDraw']  = window['AscCommonDraw'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

	window['AscCommonDraw'].CVisioDocument = CVisioDocument;
	window['AscCommonDraw'].CWindows = CWindows;
	window['AscCommonDraw'].CMasters = CMasters;
	window['AscCommonDraw'].CMasterContents = CMasterContents;
	window['AscCommonDraw'].CPages = CPages;
	window['AscCommonDraw'].CPageContents = CPageContents;
	window['AscCommonDraw'].CComments = CComments;
	window['AscCommonDraw'].CExtensions = CExtensions;
	window['AscCommonDraw'].CDataConnections = CDataConnections;
	window['AscCommonDraw'].CDataRecordSets = CDataRecordSets;
	window['AscCommonDraw'].CValidation = CValidation;
	// Not realized, file defines schema and data of that schema
	window['AscCommonDraw'].CSolutions = CSolutions;
	window['AscCommonDraw'].CSolutionXML = CSolutionXML;
})(window, window.document);
