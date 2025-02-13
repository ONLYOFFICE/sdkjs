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
		this.isVisioDocument = true;
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
		this.ImageMap = null;

		// unfinished
		// this.EmbeddedData = null;

		//------------------------------------------------------------------------------------------------------------------
		//  Сохраняем ссылки на глобальные объекты
		//------------------------------------------------------------------------------------------------------------------
		this.History              = History;
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
			if (this.History)
				this.History.Set_LogicDocument(this);
		}
		this.mainDocument = false !== isMainLogicDocument;
		this.pageIndex = 0;

		/**
		 * topLevelShapesAndGroups taken from visio shapes conversion
		 * @type {(CShape | CGroupShape | CImageShape)[][]} topLevelShapesAndGroups
		 */
		this.pageShapesCache = [];
		/**
		 * Stores pages for which background shapes were add already. Stores indexes in pageInfo array
		 * @type {number[]}
		 */
		this.backgroundAppliedFor = [];

		this.isPagesArranged = false;

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
			// this.visioDocument_Type = new AscVisio.VisioDocument_Type();
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

	CVisioDocument.prototype.Get_Api = function() {
		return this.api;
	};
	CVisioDocument.prototype.Get_CollaborativeEditing = function() {
		return this.collaborativeEditing;
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
	 * @memberOf CVisioDocument
	 * @return {number}
	 */
	CVisioDocument.prototype.GetFirstSelectedType = function() {
		return -1;
	};
	/**
	 * @memberOf CVisioDocument
	 * @param nIdx
	 * @return {number}
	 */
	CVisioDocument.prototype.GetSlideType = function(nIdx) {
		//чтобы не работал select thumbnail c ctrl или shift
		return nIdx;
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
		let logic_w_mm = this.GetWidthScaledMM(pageIndex);
		let logic_h_mm = this.GetHeightScaledMM(pageIndex);

		var _value = 100;

		var _pageWidth  = logic_w_mm * g_dKoef_mm_to_pix;
		var _pageHeight = logic_h_mm * g_dKoef_mm_to_pix;

		var _hor_Zoom = 100;
		if (0 != _pageWidth)
			_hor_Zoom = 100 * displayedWidthPx / _pageWidth;
		var _ver_Zoom = 100;
		if (0 != _pageHeight)
			_ver_Zoom = 100 * displayedHeightPX / _pageHeight;

		_value = Math.min(_hor_Zoom, _ver_Zoom);

		return _value;
	};

	/**
	 * @param pageIndex
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.GetWidthScaledMM = function(pageIndex) {
		let pageInfo = this.pages.page[pageIndex];
		let drawingScale = pageInfo.pageSheet.getCellNumberValue("DrawingScale");
		let pageScale = pageInfo.pageSheet.getCellNumberValue("PageScale");
		let logic_w_inch = pageInfo.pageSheet.getCellNumberValueWithScale("PageWidth", drawingScale / pageScale);
		return logic_w_inch * g_dKoef_in_to_mm;
	}
	CVisioDocument.prototype.GetWidthMM = function(pageIndex) {
		if (undefined !== pageIndex) {
			return this.GetWidthScaledMM(pageIndex);
		} else {
			return this.GetWidthScaledMM(this.pageIndex);
		}
	}


	/**
	 * @param pageIndex
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.GetHeightScaledMM = function(pageIndex) {
		let pageInfo = this.pages.page[pageIndex];
		let drawingScale = pageInfo.pageSheet.getCellNumberValue("DrawingScale");
		let pageScale = pageInfo.pageSheet.getCellNumberValue("PageScale");
		let logic_h_inch = pageInfo.pageSheet.getCellNumberValueWithScale("PageHeight", drawingScale / pageScale);
		return logic_h_inch * g_dKoef_in_to_mm;
	}
	CVisioDocument.prototype.GetHeightMM = function(pageIndex) {
		if (undefined !== pageIndex) {
			return this.GetHeightScaledMM(pageIndex);
		} else {
			return this.GetHeightScaledMM(this.pageIndex);
		}
	}

	/**
	 * Load fonts which are used in document and do callback (can be used to call CVisioDocument.prototype.draw).
	 * Inits this.loadedFonts.
	 * @memberOf CVisioDocument
	 */
	CVisioDocument.prototype.loadFonts = function() {
		/**
		 * load font to aFonts
		 * @param fontName
		 * @param {AscFonts.CFont[]} aFonts
		 * @param api
		 */
		function loadFontByName(fontName, aFonts, api) {
			// if font is not loaded already
			if (aFonts.findIndex(function (cFont) {
				return cFont.name === fontName;
			}) !== -1) {
				return;
			}
			aFonts.push(new AscFonts.CFont(fontName, newFontIndex, "", 0));
			newFontIndex++;

			let fontInfo = api.FontLoader.fontInfos.find(function(cFontInfo) {
				return cFontInfo.Name === fontName;
			});
			if (fontInfo === undefined || fontInfo === null) {
				AscCommon.consoleLog("Unknown font used in visio file: " + fontName);
			} else {
				AscCommon.consoleLog("Font: " + fontName + " will be loaded");
			}
		}

		let api = this.api;
		let aFonts = this.loadedFonts;

		let newFontIndex = 0;

		// load Arial and Calibri by default
		loadFontByName("Arial", aFonts, api);
		loadFontByName("Calibri", aFonts, api);

		// read theme.xml fonts
		var oFontMap = {};
		this.themes.forEach(function (theme) {
			theme.Document_Get_AllFontNames(oFontMap);
		});
		for (const fontName in oFontMap) {
			if (oFontMap.hasOwnProperty(fontName)) {
				loadFontByName(fontName, aFonts, api);
			}
		}

		// read document.xml FaceNames tag
		this.faceNames.forEach(function (faceName_Type) {
			let nameU = faceName_Type.nameU;
			loadFontByName(nameU, aFonts, api);
		});

		// may immediately call callback
		api.FontLoader.LoadDocumentFonts(aFonts, false);
	}


	/**
	 *
	 * @memberOf CVisioDocument
	 * @param {number} Zoom
	 * @param pGraphics
	 * @param {number} pageIndex
	 */
	CVisioDocument.prototype.draw = function(Zoom, pGraphics, pageIndex) {
		function drawShapeOrGroupRecursively(graphics, shapeOrGroup, baseMatrix, baseTextMatrix,
											 isRecalculateTextY, isFlipImages, isAdditionalRecalculate, logic_h_mm, currentGroupHandling) {
			// see sdkjs/common/Shapes/Serialize.js this.ReadGroupShape = function(type) to
			// learn how to work with shape groups

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
				let cGroupShape = new AscFormat.CShape();
				// cShape.setParent();
				cGroupShape.setLocks(0);
				cGroupShape.setBDeleted(false);

				cGroupShape.setSpPr(group.spPr);
				cGroupShape.spPr.setParent(cGroupShape);
				cGroupShape.rot = group.rot;
				cGroupShape.Id = group.Id;
				cGroupShape.brush = group.brush;
				cGroupShape.bounds = group.bounds;
				cGroupShape.flipH = group.flipH;
				cGroupShape.flipV = group.flipV;
				cGroupShape.localTransform = group.localTransform;
				cGroupShape.pen = group.pen;


				let shape_drawer = new AscCommon.CShapeDrawer();
				shape_drawer.fromShape2(cGroupShape, graphics, cGroupShape.getGeometry());
				let groupGeometry = cGroupShape.getGeometry();
				shape_drawer.draw(groupGeometry);

				shape_drawer.Clear();
				graphics.RestoreGrState();


				// handle group children
				group.spTree.forEach(function(shapeOrGroup) {
					drawShapeOrGroupRecursively(graphics, shapeOrGroup, baseMatrix, baseTextMatrix, isRecalculateTextY,
						isFlipImages, isAdditionalRecalculate, logic_h_mm, group);
				});
			} else {
				// shape came to argument

				// flip images
				if (isFlipImages && shapeOrGroup.getObjectType() === AscDFH.historyitem_type_ImageShape) {
					shapeOrGroup.transform.sy = -1;
					shapeOrGroup.transform.ty += shapeOrGroup.spPr.xfrm.extY;
				}

				if (isRecalculateTextY && shapeOrGroup.Id.substring(shapeOrGroup.Id.length - 4) === "Text") {
					if (graphics.SetBaseTransform) {
						//todo CSlideBoundsChecker
						graphics.SetBaseTransform(baseTextMatrix);
					}
					shapeOrGroup.transform.ty = logic_h_mm - shapeOrGroup.transform.ty - shapeOrGroup.spPr.xfrm.extY;
					shapeOrGroup.recalculateTransformText();
				}

				if (!shapeOrGroup.bDeleted) {
					shapeOrGroup.draw(graphics, shapeOrGroup.transform, shapeOrGroup.transformText);
				}

				// set shape transform that was before fix for future drawShapeOrGroupRecursively() calls
				if (isRecalculateTextY && shapeOrGroup.Id.substring(shapeOrGroup.Id.length - 4) === "Text") {
					if (graphics.SetBaseTransform) {
						//todo CSlideBoundsChecker
						graphics.SetBaseTransform(baseMatrix);
					}
					shapeOrGroup.transform.ty = logic_h_mm - shapeOrGroup.transform.ty - shapeOrGroup.spPr.xfrm.extY;
					shapeOrGroup.recalculateTransformText();
				}
				if (isFlipImages && shapeOrGroup.getObjectType() === AscDFH.historyitem_type_ImageShape) {
					shapeOrGroup.transform.sy = 1;
					shapeOrGroup.transform.ty -= shapeOrGroup.spPr.xfrm.extY;
				}

				if (isAdditionalRecalculate) {
					shapeOrGroup.recalculate();
					shapeOrGroup.recalculateTransformText && shapeOrGroup.recalculateTransformText();
					shapeOrGroup.recalculateLocalTransform(shapeOrGroup.transform);
					shapeOrGroup.recalculateContent && shapeOrGroup.recalculateContent();

				}
			}
		}

		function drawOnCanvas(pageIndex, visioDocument, canvas, isThumbnail) {
			let isRecalculateTextY = false;
			let isFlipYMatrix = false;
			let isFlipImages = false;

			let isAdditionalRecalculate = false;

			// let pageInfo = visioDocument.pages.page[pageIndex];
			// let pageContent = visioDocument.pageContents[pageIndex];
			// let topLevelShapesAndGroups = visioDocument.convertToCShapesAndGroups(pageInfo, pageContent);
			// calculated in CVisioDocument.prototype.draw
			let topLevelShapesAndGroups = visioDocument.pageShapesCache[pageIndex];

			let logic_w_mm = visioDocument.GetWidthScaledMM(pageIndex);
			let logic_h_mm = visioDocument.GetHeightScaledMM(pageIndex);

			let graphics;

			let useFitToScreenZoom = !pGraphics;
			let fitZoom;
			if (useFitToScreenZoom) {
				if (isThumbnail) {
					fitZoom = 100 *
						visioDocument.getFitZoomValue(pageIndex, canvas.offsetWidth, canvas.offsetHeight) / 100;
				} else {
					let api = visioDocument.api;
					let apiHtmlElement = api.HtmlElement.querySelector("#id_main");
					fitZoom = Zoom *
						visioDocument.getFitZoomValue(pageIndex, apiHtmlElement.offsetWidth, apiHtmlElement.offsetHeight) / 100;

				}
			} else {
				if (isThumbnail) {
					fitZoom = 10;
				} else {
					fitZoom = Zoom;
				}
			}

			let pageScale = fitZoom / 100;

			if (pGraphics) {
				graphics = pGraphics;
			} else {
				/**
				 * mm to px coef
				 * @type {number}
				 */
				let dKoef = pageScale * g_dKoef_mm_to_pix * AscCommon.AscBrowser.retinaPixelRatio;

				let w_mm = logic_w_mm;
				let h_mm = logic_h_mm;

				let w_px = (w_mm * dKoef + 0.5) >> 0;
				let h_px = (h_mm * dKoef + 0.5) >> 0;

				// Version 1
				// let parentElement = documentCanvas.parentElement;
				// documentCanvas.style.width  = parentElement.offsetWidth + "px";
				// documentCanvas.style.height = parentElement.offsetHeight + "px";

				// Version 2 with correct scroll lines
				// setup scroll lines
				//todo remove
				canvas.style.width  = w_px + "px";
				canvas.style.height = h_px + "px";
				// set pixels count for width and height
				canvas.width = AscCommon.AscBrowser.convertToRetinaValue(canvas.clientWidth, true);
				canvas.height = AscCommon.AscBrowser.convertToRetinaValue(canvas.clientHeight, true);

				// canvas#id_viewer_overlay and div#id_target_cursor creates empty gray space below any drawing

				AscCommon.calculateCanvasSize(canvas);

				let ctx = canvas.getContext('2d');

				graphics = new AscCommon.CGraphics();
				graphics.init(ctx, w_px, h_px, w_mm, h_mm);
				graphics.m_oFontManager = AscCommon.g_fontManager;
			}

			if (graphics.m_oContext) {
				graphics.m_oContext.clearRect(0, 0, canvas.width, canvas.height);
			}

			//visio y coordinate goes up while
			//ECMA-376-11_5th_edition and Geometry.js y coordinate goes down
			let baseMatrix = new AscCommon.CMatrix();
			if (isFlipYMatrix) {
				baseMatrix.SetValues(1, 0, 0, -1, 0, logic_h_mm);
			} else {
				baseMatrix.SetValues(1, 0, 0, 1, 0, 0);
			}
			if (graphics.SetBaseTransform) {
				//todo CSlideBoundsChecker
				graphics.SetBaseTransform(baseMatrix);
			}


			let baseTextMatrix = new AscCommon.CMatrix();
			baseTextMatrix.SetValues(1, 0, 0, 1, 0, 0);

			graphics.SaveGrState();
			graphics.SetIntegerGrid(false);
			graphics.transform3(new AscCommon.CMatrix());
			graphics.b_color1( 255, 255, 255, 255 );
			graphics.rect( 0, 0, logic_w_mm, logic_h_mm );
			graphics.df();
			graphics.RestoreGrState();

			topLevelShapesAndGroups.forEach(function(shapeOrGroup) {
				if (isFlipImages || isRecalculateTextY || isAdditionalRecalculate) {
					drawShapeOrGroupRecursively(graphics, shapeOrGroup, baseMatrix, baseTextMatrix, isRecalculateTextY,
						isFlipImages, isAdditionalRecalculate, logic_h_mm);
				} else {
					shapeOrGroup.draw(graphics);
				}
			});
		}

		if (this.pages === null) {
			// doesn't work after correct file read
			alert("No pages or wrong file");
			return;
		}

		// arrange pages
		if (!this.isPagesArranged) {

			// count backgrounds
			let backgroundsCount = 0;
			for (let pageIndex = 0; pageIndex < this.pages.page.length; pageIndex++) {
				let pageInfo = this.pages.page[pageIndex];
				if (pageInfo.background === true) {
					backgroundsCount++;
				}
			}

			// move background pages to back
			for (let i = 0; i < backgroundsCount; i++) {
				let backgroundInfo = this.pages.page.shift();
				let backgroundContent = this.pageContents.shift();

				this.pages.page.push(backgroundInfo);
				this.pageContents.push(backgroundContent);
			}
			this.isPagesArranged = true;
		}

		// convert shapes
		for (let pageIndex = 0; pageIndex < this.pages.page.length; pageIndex++) {
			if (this.pageShapesCache[pageIndex] === undefined) {
				let pageInfo = this.pages.page[pageIndex];
				let pageContent = this.pageContents[pageIndex];

				// Scale should be applied. Drawing scale should not be considered for text font size and stoke size
				// https://support.microsoft.com/en-us/office/change-the-drawing-scale-on-a-page-in-visio-05c24456-67bf-47f7-b5dc-d5caa9974f19
				// https://stackoverflow.com/questions/63295483/how-properly-set-line-scaling-in-ms-visio
				// also arrow size
				let drawingScale = pageInfo.pageSheet.getCellNumberValue("DrawingScale");
				let pageScale = pageInfo.pageSheet.getCellNumberValue("PageScale");
				let drawingPageScale = drawingScale / pageScale;

				let topLevelShapesAndGroups = this.convertToCShapesAndGroups(pageInfo, pageContent, drawingPageScale);
				this.pageShapesCache[pageIndex] = topLevelShapesAndGroups;
			}
		}

		// handle backgrounds
		for (let pageIndex = 0; pageIndex < this.pages.page.length; pageIndex++) {
			let pageInfo = this.pages.page[pageIndex];
			let pageContent = this.pageContents[pageIndex];

			if (!this.backgroundAppliedFor.includes(pageIndex)) {
				this.backgroundAppliedFor.push(pageIndex);
				let backgroundPageId = pageInfo.backPage;
				if (backgroundPageId !== null && backgroundPageId !== undefined) {
					// find background page
					let backgroundPageIndex = this.pages.page.findIndex(function (pageInfo) {
						return pageInfo.iD === backgroundPageId;
					});
					if (backgroundPageIndex !== -1) {
						let backgroundPageContent = this.pageShapesCache[backgroundPageIndex];
						this.pageShapesCache[pageIndex] = backgroundPageContent.concat(this.pageShapesCache[pageIndex]);
					}
				}
			}
		}

		//HOTFIX
		this.theme = this.themes[0];

		let api = this.api;
		let documentCanvas = api.canvas;
		let drawThumbnails = false;
		if (drawThumbnails) {
			let panelThumbnails = api.HtmlElement.querySelector("#id_panel_thumbnails");
			panelThumbnails.innerHTML = "";
			for (let thumbPageIndex = 0; thumbPageIndex < this.pages.page.length; thumbPageIndex++) {
				// var thumbnailCanvas = document.createElement("<canvas id=\"id_thumbnails\" class=\"block_elem\" style=\"user-select: none; z-index: 2; left: 0px; top: 0px; width: 100%; height:100%; cursor: default;\"></canvas>");
				let thumbnailCanvas = document.createElement("canvas");
				thumbnailCanvas.style.width = "70%";
				thumbnailCanvas.style.height = "30%";
				let thisContext = this;
				thumbnailCanvas.onclick = function () {
					// Zoom = 100;
					thisContext.api.goToPage(thumbPageIndex);
				}
				panelThumbnails.appendChild(thumbnailCanvas);

				drawOnCanvas(thumbPageIndex, this, thumbnailCanvas, true);
			}
		}

		drawOnCanvas(pageIndex, this, documentCanvas, false);
	};
	function getRandomPrst() {
		let types = AscCommon.g_oAutoShapesTypes[Math.floor(Math.random()*AscCommon.g_oAutoShapesTypes.length)];
		return types[Math.floor(Math.random()*types.length)].Type;
	}

	/**
	 * @memberOf CVisioDocument
	 * @param pageInfo
	 * @param pageContent
	 * @param {Number} drawingPageScale
	 * @return {(CShape | CGroupShape | CImageShape)[]} topLevelShapesAndGroups
	 */
	CVisioDocument.prototype.convertToCShapesAndGroups = function(pageInfo, pageContent, drawingPageScale) {
		/** @type {(CShape | CGroupShape | CImageShape)[]} */
		let topLevelShapesAndGroups = [];

		let masters = this.joinMastersInfoAndContents();

		for(let i = 0; i < pageContent.shapes.length; i++) {
			let shape = pageContent.shapes[i];

			// inherit styles
			let stylesWithRealizedInheritance = new Set();
			shape.realizeStyleInheritanceRecursively(this.styleSheets, stylesWithRealizedInheritance);
			// inherit master and links to master styles
			shape.realizeMasterInheritanceRecursively(masters);
			// inherit master styles
			// TODO performance: realize style inheritance only if style is inherited from master
			shape.realizeStyleInheritanceRecursively(this.styleSheets, stylesWithRealizedInheritance);

			if (shape.type === "Group") {
				let cGroupShape = shape.convertGroup(this, pageInfo, drawingPageScale);
				if (cGroupShape) {
					topLevelShapesAndGroups.push(cGroupShape);
				}
			} else {
				let cShapeOrCGroupShape = shape.convertShape(this, pageInfo, drawingPageScale);
				if (cShapeOrCGroupShape !== null) {
					topLevelShapesAndGroups.push(cShapeOrCGroupShape);
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
	CVisioDocument.prototype.getCountPages = function() {
		return this.pages && this.pages.page.length || 0;
	}
	CVisioDocument.prototype.getCurrentPage = function() {
		return this.pageIndex;
	}
	CVisioDocument.prototype.setCurrentPage = function(pageIndex) {
		return this.pageIndex = pageIndex;
	}
	CVisioDocument.prototype.getFirstSlideNumber = function() {
		//todo remove
		return 0;
	}
	CVisioDocument.prototype.IsVisibleSlide = function (nIndex) {
		//todo remove
		return true;
	};
	CVisioDocument.prototype.GetSlide = function() {
		//todo remove
		return {getObjectType: function(){}, isVisible: function(){return true}};
	}
	CVisioDocument.prototype.ContinueSpellCheck = function () {
		//this.Spelling.ContinueSpellCheck();
	};
	CVisioDocument.prototype.GetSlidesCount = function () {
		return this.getCountPages();
	};
	CVisioDocument.prototype.Recalculate = function (RecalcData) {
		//todo
		this.DrawingDocument.OnStartRecalculate(this.GetSlidesCount());
		let _RecalcData = RecalcData ? RecalcData : History.Get_RecalcData();
		if (_RecalcData.Drawings.All) {
			for(let pageIndex in this.pageShapesCache) {
				this.pageShapesCache[pageIndex].forEach(function(shapeOrGroup) {
					shapeOrGroup.recalcText();
					shapeOrGroup.recalculate();
				});
			}
		}
		this.DrawingDocument.OnEndRecalculate();
		History.Reset_RecalcIndex();
		this.RecalculateCurPos();
		this.Document_UpdateSelectionState();
	};

	CVisioDocument.prototype.RecalculateCurPos = function () {
	};
	CVisioDocument.prototype.Document_UpdateSelectionState = function () {
	};
	CVisioDocument.prototype.Document_UpdateInterfaceState = function () {
	};
	CVisioDocument.prototype.Document_UpdateRulersState = function () {
	};
	CVisioDocument.prototype.private_UpdateCursorXY = function (bUpdateX, bUpdateY) {
	};

	CVisioDocument.prototype.OnMouseUp = function (e, X, Y, PageIndex) {
	};
	CVisioDocument.prototype.OnMouseDown = function (e, X, Y, PageIndex) {
	};
	CVisioDocument.prototype.OnMouseMove = function (e, X, Y, PageIndex) {
	};
	CVisioDocument.prototype.DrawPage = function (pageIndex, pGraphics) {
		this.draw(100, pGraphics, pageIndex);
	};
	CVisioDocument.prototype.GetSlideNumber = function (nIdx) {
		return nIdx + 1;
	};
	CVisioDocument.prototype.isSlideAnimated = function () {
		return false;
	};
	CVisioDocument.prototype.CheckTargetUpdate = function () {
	};
	CVisioDocument.prototype.Document_CreateFontMap = function () {
		return {};
	};
	CVisioDocument.prototype.resetStateCurSlide = function () {
	};
	CVisioDocument.prototype.IsStartedPreview = function () {
		return false;
	};
	CVisioDocument.prototype.Set_FastCollaborativeEditing = function (isOn) {
		//todo
	};
	CVisioDocument.prototype.shiftSlides = function (pos, array, bCopy) {
		//todo
	};
	CVisioDocument.prototype.executeShortcut = function(type) {
		let result = false;

		switch (type) {
			case Asc.c_oAscDiagramShortcutType.Print: {
				this.api.onPrint();
				result = true;
				break;
			}
			default: {
				var oCustom = this.api.getCustomShortcutAction(type);
				if (oCustom) {
					if (AscCommon.c_oAscCustomShortcutType.Symbol === oCustom.Type) {
						this.api["asc_insertSymbol"](oCustom.Font, oCustom.CharCode);
					}
				}
				break;
			}
		}
		return result;
	}
	CVisioDocument.prototype.OnKeyDown = function (e) {
		this.api.sendEvent("asc_onBeforeKeyDown", e);

		var bUpdateSelection = true;
		var bRetValue = keydownresult_PreventNothing;
		let nStartHistoryIndex = this.History.Index;


		// // Сбрасываем текущий элемент в поиске
		// if (this.SearchEngine.Count > 0)
		// 	this.SearchEngine.ResetCurrent();

		let shortcutType = this.api.getShortcut(e);
		if (this.executeShortcut(shortcutType))
		{
			bRetValue = keydownresult_PreventAll;
			bUpdateSelection = false;
		}
		else {
			const bIsMacOs = AscCommon.AscBrowser.isMacOs;
			let WordControl = this.api.WordControl;
			if (e.KeyCode === 33) // PgUp
			{
				//
			}
			else if (e.KeyCode === 34) // PgDn
			{
				//
			}
			else if (e.KeyCode === 35) // клавиша End
			{
				if (true === e.CtrlKey) // Ctrl + End - переход в конец документа
				{
					WordControl.m_oScrollVerApi.scrollTo(WordControl.m_dScrollX_max, WordControl.m_dScrollY_max);
				}
				else
				{
					WordControl.m_oScrollVerApi.scrollTo(0, WordControl.m_dScrollY_max);
				}
			}
			else if (e.KeyCode === 36) // клавиша Home
			{
				if (true === e.CtrlKey) // Ctrl + Home - переход в начало документа
				{
					WordControl.m_oScrollVerApi.scrollTo(WordControl.m_dScrollX_max, 0);
				}
				else
				{
					WordControl.m_oScrollVerApi.scrollTo(0, 0);
				}
			}
			else if (e.KeyCode === 37) // Left Arrow
			{
				if (true || WordControl.m_bIsHorScrollVisible)
				{
					WordControl.m_oScrollHorApi.scrollBy(-30, 0, false);
				}
			}
			else if (e.KeyCode === 38) // Top Arrow
			{
				WordControl.m_oScrollVerApi.scrollBy(0, -30, false);
			}
			else if (e.KeyCode === 39) // Right Arrow
			{
				if (true || WordControl.m_bIsHorScrollVisible)
				{
					WordControl.m_oScrollHorApi.scrollBy(30, 0, false);
				}
			}
			else if (e.KeyCode === 40) // Bottom Arrow
			{
				WordControl.m_oScrollVerApi.scrollBy(0, 30, false);
			}
		}

		if (bRetValue & keydownflags_PreventKeyPress && true === bUpdateSelection)
			this.Document_UpdateSelectionState();

		if(nStartHistoryIndex === this.History.Index) {
			this.private_UpdateCursorXY(true, true);
		}

		this.api.sendEvent("asc_onKeyDown", e);
		return bRetValue;
	};
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
						AscCommon.consoleLog('check sdkjs/draw/model/VisioDocument.js : parseMasters');
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
			let pagesXml = pagesPart.getDocumentContent();
			reader = new StaxParser(pagesXml, pagesPart, context);
			this.pages = new CPages();
			this.pages.fromXml(reader);

			// page content parts
			let pages = pagesPart.getPartsByRelationshipType(AscCommon.openXml.Types.page.relationType);
			if (pages.length  > 0) {
				// this.pageContents order is important it must correspond to this.pages but pages is messed up by default
				// so now let's get page contents by page relationship to get pageContents in correct order
				// 1) find page rId number
				// 2) find pageContent by rId

				for (let i = 0; i < this.pages.page.length; i++) {
					let pageContentRid = this.pages.page[i] && this.pages.page[i].rel && this.pages.page[i].rel.id;
					if (pageContentRid) {
						// let pageContentRel = pagesPart.getRelationship(pageContentRid);
						let pageContentPart = pagesPart.getPartById(pageContentRid);
						let contentPage = pageContentPart.getDocumentContent();
						reader = new StaxParser(contentPage, pageContentPart, context);
						let PageContent = new CPageContents();
						PageContent.fromXml(reader);
						this.pageContents.push(PageContent);
					} else {
						AscCommon.consoleLog("Page content rId not found");
					}
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
					AscCommon.consoleLog('check sdkjs/draw/model/VisioDocument.js : parseThemes');
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
			AscCommon.consoleLog("Themes to parse not found. Mb no rels. Trying to get themes by filenames");
			let themeNum = 1;
			while (true) {
				let uInt8ArrayTheme = documentPart.pkg.zip.getFile("visio/theme/theme" + themeNum + ".xml");
				if (uInt8ArrayTheme === null) {
					break;
				}
				if (!uInt8ArrayTheme) {
					uInt8ArrayTheme = new Uint8Array(0);
				}
				let themeXml = AscCommon.UTF8ArrayToString(uInt8ArrayTheme, 0, uInt8ArrayTheme.length);

				reader = new StaxParser(themeXml, undefined, context);
				let theme = new AscFormat.CTheme();
				theme.fromXml(reader, true);
				this.themes.push(theme);

				themeNum++;
			}
			if (themeNum === 1) {
				AscCommon.consoleLog("No themes found by filenames. Creating default theme");
				this.themes.push(AscFormat.GenerateDefaultTheme(null, null));
			}
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
						AscCommon.consoleLog('check sdkjs/draw/model/VisioDocument.js : parseSolutions');
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
	window['AscVisio']  = window['AscVisio'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

	window['AscVisio'].CVisioDocument = CVisioDocument;
	window['AscVisio'].CWindows = CWindows;
	window['AscVisio'].CMasters = CMasters;
	window['AscVisio'].CMasterContents = CMasterContents;
	window['AscVisio'].CPages = CPages;
	window['AscVisio'].CPageContents = CPageContents;
	window['AscVisio'].CComments = CComments;
	window['AscVisio'].CExtensions = CExtensions;
	window['AscVisio'].CDataConnections = CDataConnections;
	window['AscVisio'].CDataRecordSets = CDataRecordSets;
	window['AscVisio'].CValidation = CValidation;
	// Not realized, file defines schema and data of that schema
	window['AscVisio'].CSolutions = CSolutions;
	window['AscVisio'].CSolutionXML = CSolutionXML;
})(window, window.document);
