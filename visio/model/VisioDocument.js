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

AscDFH.historyitem_type_VisioDocument = 301;
AscDFH.historyitem_type_VisioWindows = 302;
AscDFH.historyitem_type_VisioMasters = 303;
AscDFH.historyitem_type_VisioMasterContents = 304;
AscDFH.historyitem_type_VisioPages = 305;
AscDFH.historyitem_type_VisioPageContents = 306;
AscDFH.historyitem_type_VisioComments = 307;
AscDFH.historyitem_type_VisioExtensions = 308;
AscDFH.historyitem_type_VisioDataConnections = 309;
AscDFH.historyitem_type_VisioDataRecordSets = 310;
AscDFH.historyitem_type_VisioValidation = 311;
AscDFH.historyitem_type_VisioSolutions = 312;
AscDFH.historyitem_type_VisioSolutionXML = 313;
AscDFH.historyitem_type_VisioDocumentSettings = 314;
AscDFH.historyitem_type_VisioDocumentProperties = 315;
AscDFH.historyitem_type_VisioHeaderFooter = 316;
AscDFH.historyitem_type_VisioDataTransferInfo = 317;
AscDFH.historyitem_type_VisioPublishSettings = 318;
AscDFH.historyitem_type_VisioColorEntry = 319;
AscDFH.historyitem_type_VisioFaceName = 320;
AscDFH.historyitem_type_VisioRefBy = 321;
AscDFH.historyitem_type_VisioMaster = 322;
AscDFH.historyitem_type_VisioMasterShortcut = 323;
AscDFH.historyitem_type_VisioConnect = 324;
AscDFH.historyitem_type_VisioDataConnection = 325;
AscDFH.historyitem_type_VisioSolution = 326;
AscDFH.historyitem_type_VisioEventItem = 327;
AscDFH.historyitem_type_VisioWindow = 328;

(function(window, document)
{

	var c_oVsdxTextKind = {
		CP : 0,
		PP : 1,
		TP : 2,
		FLD : 3
	};

	/**
	 * 	Docs:
	 * 	VisioDocument_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/visiodocument_type-complextypevisio-xml
	 * @constructor
	 * @extends AscFormat.CBaseFormatNoIdObject
	 */
	function CVisioDocument(Api, DrawingDocument, isMainLogicDocument) {
		AscFormat.CBaseFormatNoIdObject.call(this);
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
		/**
		 * @type {StyleSheet_Type[]} styles
		 */
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
		/**
		 * with rId resolved as content
		 * @type {CMasters}
		 */
		this.masters = null;
		/**
		 * with rId resolved as content
		 * @type {CPages}
		 */
		this.pages = null;
		this.themes = [];
		this.app = null;
		this.core = null;
		this.customProperties = new AscCommon.CCustomProperties();
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
	AscFormat.InitClass(CVisioDocument, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);
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

		if (this.masters) {
			for (let i = 0; i < this.masters.master.length; i++) {
				let masterContent = mastersPart.part.addPart(AscCommon.openXml.Types.master);
				masterContent.part.setDataXml(this.masters.master[i].content, memory);
			}
		}
		let pagesPart = docPart.part.addPart(AscCommon.openXml.Types.pages);
		if (this.pages) {
			for (let i = 0; i < this.pages.page.length; i++) {
				let pageContent = pagesPart.part.addPart(AscCommon.openXml.Types.page);
				pageContent.part.setDataXml(this.pages.page[i].content, memory);

				// I add page[N].xml.rels below
				// It has links to all masters but
				// in examplevsdx file in page[N].xml.rels rId[N] states to random master[N]
				// e.g. rId3 to ../masters/master1.xml
				// here rId1 will state to master1, rId2 to master2, etc.
				// TODO check if this is important
				// in page[N].xml there is no rId used only <Shape ... Master="[ID]">
				// e. g. <Shape ... Master="1">
				if (this.masters) {
					for (let i = 0; i < this.masters.master.length; i++) {
						pageContent.part.addRelationship(AscCommon.openXml.Types.masterFromPage.relationType,
							"../masters/master" + (i + 1) + ".xml");
					}
				}
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
	CVisioDocument.prototype.AfterOpenDocument = function(zip, context) {
		if (!this.themes.length) {
			AscCommon.consoleLog("No themes found by filenames. Creating default theme");
			this.themes.push(AscFormat.GenerateDefaultTheme(null, null));
		}
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
	 * using visio shapes data from this inits this.pageContents which is array of CShapes
	 * @memberof CVisioDocument
	 */
	CVisioDocument.prototype.toCShapes = function() {
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
				this.pages.page.push(backgroundInfo);
			}
			this.isPagesArranged = true;
		}

		// convert shapes
		for (let pageIndex = 0; pageIndex < this.pages.page.length; pageIndex++) {
			if (this.pageShapesCache[pageIndex] === undefined) {
				let page = this.pages.page[pageIndex];
				let pageContent = page.content;

				// Scale should be applied. Drawing scale should not be considered for text font size and stoke size
				// https://support.microsoft.com/en-us/office/change-the-drawing-scale-on-a-page-in-visio-05c24456-67bf-47f7-b5dc-d5caa9974f19
				// https://stackoverflow.com/questions/63295483/how-properly-set-line-scaling-in-ms-visio
				// also arrow size
				let drawingScale = page.pageSheet.getCellNumberValue("DrawingScale");
				let pageScale = page.pageSheet.getCellNumberValue("PageScale");
				let drawingPageScale = drawingScale / pageScale;

				let topLevelShapesAndGroups = this.convertToCShapesAndGroups(page, pageContent, drawingPageScale);
				this.pageShapesCache[pageIndex] = topLevelShapesAndGroups;

				topLevelShapesAndGroups.forEach(function (shapeOrGroup) {
					shapeOrGroup.recalculate();
				});
			}
		}

		// handle backgrounds
		for (let pageIndex = 0; pageIndex < this.pages.page.length; pageIndex++) {
			let pageInfo = this.pages.page[pageIndex];

			if (!this.backgroundAppliedFor.includes(pageIndex)) {
				this.backgroundAppliedFor.push(pageIndex);
				let backgroundPageId = pageInfo.backPage;
				if (backgroundPageId !== null && backgroundPageId !== undefined) {
					// find background page
					let backgroundPageIndex = this.pages.page.findIndex(function (pageInfo) {
						return pageInfo.id === backgroundPageId;
					});
					if (backgroundPageIndex !== -1) {
						let backgroundPageContent = this.pageShapesCache[backgroundPageIndex];
						this.pageShapesCache[pageIndex] = backgroundPageContent.concat(this.pageShapesCache[pageIndex]);
					}
				}
			}
		}
	}

	/**
	 *
	 * @memberOf CVisioDocument
	 * @param {number} Zoom
	 * @param pGraphics
	 * @param {number} pageIndex
	 */
	CVisioDocument.prototype.draw = function(Zoom, pGraphics, pageIndex) {
		/**
		 *
		 * @param graphics
		 * @param {(CShape | CGroupShape)} shapeOrGroup
		 * @param baseMatrix
		 * @param baseTextMatrix
		 * @param isRecalculateTextY
		 * @param isFlipImages
		 * @param isAdditionalRecalculate
		 * @param logic_h_mm
		 * @param {CGroupShape?} currentGroupHandling
		 */
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

		for(let i = 0; i < pageContent.shapes.length; i++) {
			let shape = pageContent.shapes[i];

			// inherit styles
			let stylesWithRealizedInheritance = new Set();

			shape.realizeStyleInheritanceRecursively(this.styleSheets, stylesWithRealizedInheritance);

			// inherit master and links to master styles
			if (this.masters) {
				shape.realizeMasterInheritanceRecursively(this.masters.master);
			}

			// inherit master styles
			// TODO performance: realize style inheritance only if style is inherited from master
			shape.realizeStyleInheritanceRecursively(this.styleSheets, stylesWithRealizedInheritance);

			if (shape.type === AscVisio.SHAPE_TYPES_GROUP) {
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
	CVisioDocument.prototype.getCountPages = function() {
		return this.pages && this.pages.page.length || 0;
	}
	CVisioDocument.prototype.getCurrentPage = function() {
		return this.pageIndex;
	}
	CVisioDocument.prototype.Set_CurPage = CVisioDocument.prototype.setCurrentPage = function(pageIndex) {
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
	CVisioDocument.prototype.GetAllSlides = function () {
		return this.pages.page;
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
	CVisioDocument.prototype.isSlidePreserved = function () {
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
	CVisioDocument.prototype.Viewer_OnChangePosition = function () {
	};
	// CVisioDocument.prototype.getMasterByID = function(ID) {
	// 	// join Master_Type and MasterContents_Type
	// 	let masterFromMastersArray = this.masters.master;
	// 	let master = null;
	//
	// 	let thisContext = this;
	// 	for (let i = 0; i < masterFromMastersArray.length; i++) {
	// 		const masterFromMasters = masterFromMastersArray[i];
	// 		if (masterFromMasters.id === ID) {
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
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CWindows() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.clientWidth = null;
		this.clientHeight = null;
		this.window = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CWindows, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioWindows);

	// Docs:
	// Masters_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/masters_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CMasters() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.master = [];
		this.masterShortcut = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CMasters, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioMasters);

	// Another name in docs PageContents_Type
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CMasterContents() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.shapes = [];
		this.connects = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CMasterContents, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioMasterContents);

	// Docs:
	// Pages_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pages_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CPages() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.page = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CPages, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioPages);

	// Docs:
	// Элемент Shapes (PageContents_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/shapes-element-pagecontents_type-complextypevisio-xml
	// PageContents_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pagecontents_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CPageContents() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		/**
		 *
		 * @type {Shape_Type[]}
		 */
		this.shapes = [];
		this.connects = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CPageContents, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioPageContents);

	// Docs:
	// Comments_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/comments_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CComments() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.showCommentTags = null;
		this.authorList = [];
		this.commentList = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CComments, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioComments);

	// Docs:
	// Extensions_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/extensions_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CExtensions() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.cellDef = [];
		this.functionDef = [];
		this.sectionDef = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CExtensions, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioExtensions);

	// Docs:
	// DataConnections_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/dataconnections_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CDataConnections() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.nextID = null;
		this.dataConnection = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CDataConnections, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDataConnections);

	// Docs:
	// DataRecordSets_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datarecordsets_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CDataRecordSets() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.nextID = null;
		this.activeRecordsetID = null;
		this.dataWindowOrder = null;
		this.dataRecordSet = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CDataRecordSets, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDataRecordSets);

	// Docs old:
	// Validation_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/validation_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CValidation() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.validationProperties = null;
		this.ruleSets = [];
		this.issues = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CValidation, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioValidation);

	// Not realized, file defines schema and data of that schema
	// Docs:
	// Solutions_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/solutions_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CSolutions() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.solution = [];
		this.xmlSpace = null;
		this.xmlns = null;
		this.r = null;
	}
	AscFormat.InitClass(CSolutions, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioSolutions);

	// Not realized, file defines schema and data of that schema
	// Docs:
	// https://learn.microsoft.com/ru-ru/office/vba/api/visio.document.solutionxmlelement
	// 									https://learn.microsoft.com/ru-ru/office/client-developer/visio/solution_type-complextypevisio-xml
	// SolutionXML_Type complexType: 	https://learn.microsoft.com/ru-ru/office/client-developer/visio/solutionxml_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CSolutionXML() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.name = null;
		// string containing overall xml
		this.fileContents = null;
	}
	AscFormat.InitClass(CSolutionXML, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioSolutionXML);

	// Docs:
	// PublishSettings_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/publishsettings_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function PublishSettings_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.publishedPage = [];
		this.refreshableData = [];
	}
	AscFormat.InitClass(PublishSettings_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioPublishSettings);

	// Docs:
	// ColorEntry_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/colorentry_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function ColorEntry_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		this.rgb = null;
	}
	AscFormat.InitClass(ColorEntry_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioColorEntry);

	// Docs:
	// FaceName_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/facename_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FaceName_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.nameU = null;
		this.unicodeRanges = null;
		this.charSets = null;
		this.panose = null;
		this.panos = null;
		this.flags = null;
		this.id = null;
		this.name = null;
	}
	AscFormat.InitClass(FaceName_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioFaceName);

	// Docs:
	// RefBy_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/refby_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RefBy_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.t = null;
		this.id = null;
	}
	AscFormat.InitClass(RefBy_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioRefBy);

	// Docs:
	// Элемент PageSheet (Master_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/pagesheet-element-master_type-complextypevisio-xml
	// Элемент Rel (Master_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/rel-element-master_type-complextypevisio-xml
	// Master_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/master_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Master_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		/**
		 * Not rel id
		 * @type {Number}
		 */
		this.id = null;
		this.baseID = null;
		this.uniqueID = null;
		this.matchByName = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.iconSize = null;
		this.patternFlags = null;
		this.prompt = null;
		this.hidden = null;
		this.iconUpdate = null;
		this.alignName = null;
		this.masterType = null;
		this.pageSheet = null;
		this.rel = null;
		this.icon = null;

		/**
		 *
		 * @type {CMasterContents}
		 */
		this.content = null;
	}
	AscFormat.InitClass(Master_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioMaster);

	// Docs:
	// Элемент Icon (MasterShortcut_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/icon-element-mastershortcut_type-complextypevisio-xml
	// MasterShortcut_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/mastershortcut_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function MasterShortcut_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.name = null;
		this.nameU = null;
		this.isCustomName = null;
		this.isCustomNameU = null;
		this.iconSize = null;
		this.patternFlags = null;
		this.prompt = null;
		this.shortcutURL = null;
		this.shortcutHelp = null;
		this.alignName = null;
		this.masterType = null;
		this.icon = null;

		/**
		 *
		 * @type {CMasterContents}
		 */
		this.content = null;
	}
	AscFormat.InitClass(MasterShortcut_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioMasterShortcut);

	// Docs:
	// Connect_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/connect_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Connect_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.fromSheet = null;
		this.fromCell = null;
		this.fromPart = null;
		this.toSheet = null;
		this.toCell = null;
		this.toPart = null;
	}
	AscFormat.InitClass(Connect_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioConnect);

	// Docs:
	// DataConnection_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/dataconnection_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DataConnection_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.fileName = null;
		this.connectionString = null;
		this.command = null;
		this.friendlyName = null;
		this.timeout = null;
		this.alwaysUseConnectionFile = null;
	}
	AscFormat.InitClass(DataConnection_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDataConnection);

	// Docs:
	// Элемент DynamicGridEnabled (DocumentSettings_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/dynamicgridenabled-element-documentsettings_type-complextypevisio-xml
	// Элемент GlueSettings (DocumentSettings_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/gluesettings-element-documentsettings_type-complextypevisio-xml
	// Элемент SnapAngles (DocumentSettings_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapangles-element-documentsettings_type-complextypevisio-xml
	// Элемент SnapExtensions (DocumentSettings_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapextensions-element-documentsettings_type-complextypevisio-xml
	// Элемент SnapSettings (DocumentSettings_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapsettings-element-documentsettings_type-complextypevisio-xml
	// DocumentSettings_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/documentsettings_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DocumentSettings_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.topPage = null;
		this.defaultTextStyle = null;
		this.defaultLineStyle = null;
		this.defaultFillStyle = null;
		this.defaultGuideStyle = null;
		this.glueSettings = null;
		this.snapSettings = null;
		this.snapExtensions = null;
		this.snapAngles = [];
		this.dynamicGridEnabled = null;
		this.protectStyles = null;
		this.protectShapes = null;
		this.protectMasters = null;
		this.protectBkgnds = null;
		this.customMenusFile = null;
		this.customToolbarsFile = null;
		this.attachedToolbars = null;
	}
	AscFormat.InitClass(DocumentSettings_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDocumentSettings);

	// Docs:
	// DocumentProperties_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/documentproperties_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DocumentProperties_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.title = null;
		this.subject = null;
		this.creator = null;
		this.manager = null;
		this.company = null;
		this.category = null;
		this.keywords = null;
		this.desc = null;
		this.hyperlinkBase = null;
		this.alternateNames = null;
		this.template = null;
		this.buildNumberCreated = null;
		this.buildNumberEdited = null;
		this.previewPicture = null;
		this.customProps = [];
		this.timeCreated = null;
		this.timeSaved = null;
		this.timeEdited = null;
		this.timePrinted = null;
	}
	AscFormat.InitClass(DocumentProperties_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDocumentProperties);

	// Docs:
	// HeaderFooter_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headerfooter_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderFooter_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.headerFooterColor = null;
		this.headerMargin = null;
		this.footerMargin = null;
		this.headerLeft = null;
		this.headerCenter = null;
		this.headerRight = null;
		this.footerLeft = null;
		this.footerCenter = null;
		this.footerRight = null;
		this.headerFooterFont = null;
	}
	AscFormat.InitClass(HeaderFooter_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioHeaderFooter);

	// Docs:
	// DataTransferInfo_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datatransferinfo_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DataTransferInfo_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.context = null;
		this.containerType = null;
		this.container = null;
		this.view = null;
		this.sheet = null;
		this.transferType = null;
		this.transferTime = null;
	}
	AscFormat.InitClass(DataTransferInfo_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioDataTransferInfo);

	// Docs:
	// Solution_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/solution_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Solution_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.name = null;
		this.rel = null;
		this.content = null;
	}
	AscFormat.InitClass(Solution_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioSolution);

	// Docs old:
	// EventItem_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/eventitem_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function EventItem_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.action = null;
		this.eventCode = null;
		this.enabled = null;
		this.target = null;
		this.targetArgs = null;
		return this;
	}
	AscFormat.InitClass(EventItem_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_VisioEventItem);

	// Docs old:
// Элемент DynamicGridEnabled (Window_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/dynamicgridenabled-element-window_type-complextypevisio-xml
// Элемент GlueSettings (Window_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/gluesettings-element-window_type-complextypevisio-xml
// Элемент SnapAngles (Window_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapangles-element-window_type-complextypevisio-xml
// Элемент SnapExtensions (Window_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapextensions-element-window_type-complextypevisio-xml
// Элемент SnapSettings (Window_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapsettings-element-window_type-complextypevisio-xml
// Window_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/window_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Window_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.windowType = null;
		this.windowState = null;
		this.document = null;
		this.windowLeft = null;
		this.windowTop = null;
		this.windowWidth = null;
		this.windowHeight = null;
		this.master = null;
		this.containerType = null;
		this.container = null;
		this.sheet = null;
		this.readOnly = null;
		this.parentWindow = null;
		this.page = null;
		this.viewScale = null;
		this.viewCenterX = null;
		this.viewCenterY = null;
		this.stencilGroup = null;
		this.stencilGroupPos = null;
		this.showRulers = null;
		this.showGrid = null;
		this.showPageBreaks = null;
		this.showGuides = null;
		this.showConnectionPoints = null;
		this.glueSettings = null;
		this.snapSettings = null;
		this.snapExtensions = null;
		this.snapAngles = [];
		this.dynamicGridEnabled = null;
		this.tabSplitterPos = null;
		return this;
	}
	AscFormat.InitClass(Window_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// tp_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/tp_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function tp_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		return this;
	}
	AscFormat.InitClass(tp_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// pp_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/pp_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function pp_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		return this;
	}
	AscFormat.InitClass(pp_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// fld_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/fld_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function fld_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(fld_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// cp_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/cp_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function cp_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		return this;
	}
	AscFormat.InitClass(cp_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

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
			// don't read MasterContents separately. Master contents are read in masters using rId.
		}
	}

	function parsePages(documentPart, reader, context) {
		let pagesPart = documentPart.getPartByRelationshipType(AscCommon.openXml.Types.pages.relationType);
		if (pagesPart) {
			let pagesXml = pagesPart.getDocumentContent();
			reader = new StaxParser(pagesXml, pagesPart, context);
			this.pages = new CPages();
			this.pages.fromXml(reader);
			// don't read PageContents separately. Page contents are read in pages using rId.
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


	// OTHER CLASSES

	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Comments_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.showCommentTags = null;
		this.authorList = [];
		this.commentList = [];
		return this;
	}
	AscFormat.InitClass(Comments_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	// RuleTest_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/ruletest_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RuleTest_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(RuleTest_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	// RuleFilter_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rulefilter_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RuleFilter_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(RuleFilter_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RowKeyValue_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rowkeyvalue_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RowKeyValue_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.rowID = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(RowKeyValue_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// DataColumn_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datacolumn_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DataColumn_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.columnNameID = null;
		this.name = null;
		this.label = null;
		this.origLabel = null;
		this.langID = null;
		this.calendar = null;
		this.dataType = null;
		this.unitType = null;
		this.currency = null;
		this.degree = null;
		this.displayWidth = null;
		this.displayOrder = null;
		this.mapped = null;
		this.hyperlink = null;
		return this;
	}
	AscFormat.InitClass(DataColumn_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RuleInfo_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/ruleinfo_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RuleInfo_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ruleSetID = null;
		this.ruleID = null;
		return this;
	}
	AscFormat.InitClass(RuleInfo_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// IssueTarget_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/issuetarget_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function IssueTarget_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.pageID = null;
		this.shapeID = null;
		return this;
	}
	AscFormat.InitClass(IssueTarget_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// Rule_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rule_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Rule_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.nameU = null;
		this.category = null;
		this.description = null;
		this.ruleTarget = null;
		this.ignored = null;
		this.ruleFilter = null;
		this.ruleTest = null;
		return this;
	}
	AscFormat.InitClass(Rule_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RuleSetFlags_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rulesetflags_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RuleSetFlags_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.hidden = null;
		return this;
	}
	AscFormat.InitClass(RuleSetFlags_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// AutoLinkComparison_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/autolinkcomparison_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function AutoLinkComparison_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.columnName = null;
		this.contextType = null;
		this.contextTypeLabel = null;
		return this;
	}
	AscFormat.InitClass(AutoLinkComparison_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RefreshConflict_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/refreshconflict_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RefreshConflict_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.rowID = null;
		this.shapeID = null;
		this.pageID = null;
		return this;
	}
	AscFormat.InitClass(RefreshConflict_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RowMap_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rowmap_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RowMap_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.rowID = null;
		this.pageID = null;
		this.shapeID = null;
		return this;
	}
	AscFormat.InitClass(RowMap_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// PrimaryKey_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/primarykey_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function PrimaryKey_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.columnNameID = null;
		this.rowKeyValue = [];
		return this;
	}
	AscFormat.InitClass(PrimaryKey_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// DataColumns_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datacolumns_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DataColumns_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.sortColumn = null;
		this.sortAsc = null;
		this.dataColumn = [];
		return this;
	}
	AscFormat.InitClass(DataColumns_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function ADOData_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		return this;
	}
	AscFormat.InitClass(ADOData_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// Rel_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/rel_type-complextypevisio-xml
	// In fact rel may have r:id attribute
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Rel_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		return this;
	}
	AscFormat.InitClass(Rel_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// CommentEntry_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/commententry_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CommentEntry_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.authorID = null;
		this.pageID = null;
		this.shapeID = null;
		this.date = null;
		this.editDate = null;
		this.done = null;
		this.commentID = null;
		this.autoCommentType = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(CommentEntry_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// AuthorEntry_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/authorentry_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function AuthorEntry_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.name = null;
		this.initials = null;
		this.sIP = null;
		this.sMTP = null;
		this.id = null;
		this.resolutionID = null;
		return this;
	}
	AscFormat.InitClass(AuthorEntry_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RefreshableData_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/refreshabledata_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RefreshableData_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		return this;
	}
	AscFormat.InitClass(RefreshableData_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// PublishedPage_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/publishedpage_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function PublishedPage_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		return this;
	}
	AscFormat.InitClass(PublishedPage_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// HeaderFooterFont_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headerfooterfont_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderFooterFont_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.height = null;
		this.width = null;
		this.escapement = null;
		this.orientation = null;
		this.weight = null;
		this.italic = null;
		this.underline = null;
		this.strikeOut = null;
		this.charSet = null;
		this.outPrecision = null;
		this.clipPrecision = null;
		this.quality = null;
		this.pitchAndFamily = null;
		this.faceName = null;
		return this;
	}
	AscFormat.InitClass(HeaderFooterFont_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// FooterRight_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/footerright_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FooterRight_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(FooterRight_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// FooterCenter_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/footercenter_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FooterCenter_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(FooterCenter_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// FooterLeft_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/footerleft_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FooterLeft_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(FooterLeft_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// HeaderRight_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headerright_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderRight_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(HeaderRight_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// HeaderCenter_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headercenter_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderCenter_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(HeaderCenter_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// HeaderLeft_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headerleft_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderLeft_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(HeaderLeft_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// FooterMargin_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/footermargin_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FooterMargin_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.unit = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(FooterMargin_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// HeaderMargin_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/headermargin_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HeaderMargin_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.unit = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(HeaderMargin_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// CustomMenusFile_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/custommenusfile_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CustomMenusFile_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(CustomMenusFile_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// SnapAngle_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapangle_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function SnapAngle_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(SnapAngle_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// SnapExtensions_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/snapextensions_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function SnapExtensions_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(SnapExtensions_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	// https://learn.microsoft.com/ru-ru/search/?terms=TimePrinted
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function TimePrinted_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(TimePrinted_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function TimeEdited_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(TimeEdited_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	// https://learn.microsoft.com/ru-ru/office/vba/api/visio.document.timesaved
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function TimeSaved_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(TimeSaved_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function TimeCreated_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(TimeCreated_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CustomProp_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.name = null;
		this.propType = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(CustomProp_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function PreviewPicture_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.size = null;
		this.value = null;
		return this;
	}
	AscFormat.InitClass(PreviewPicture_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function BuildNumberEdited_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(BuildNumberEdited_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function BuildNumberCreated_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(BuildNumberCreated_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Template_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Template_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function AlternateNames_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(AlternateNames_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function HyperlinkBase_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.href = null;
		return this;
	}
	AscFormat.InitClass(HyperlinkBase_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Desc_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Desc_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Keywords_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Keywords_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Category_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Category_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Company_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Company_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Manager_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Manager_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Creator_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Creator_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Subject_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Subject_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Title_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.value = null;
		return this;
	}
	AscFormat.InitClass(Title_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// SectionDef_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/sectiondef_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function SectionDef_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.n = null;
		this.t = null;
		this.s = null;
		this.cellDef = [];
		this.rowDef = [];
		return this;
	}
	AscFormat.InitClass(SectionDef_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// FunctionDef_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/functiondef_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function FunctionDef_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.n = null;
		return this;
	}
	AscFormat.InitClass(FunctionDef_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// CellDef_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/celldef_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CellDef_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		return this;
	}
	AscFormat.InitClass(CellDef_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// Issue_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/issue_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function Issue_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.ignored = null;
		this.issueTarget = null;
		this.ruleInfo = null;
		return this;
	}
	AscFormat.InitClass(Issue_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// RuleSet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/ruleset_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function RuleSet_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.name = null;
		this.nameU = null;
		this.description = null;
		this.enabled = null;
		this.ruleSetFlags = null;
		this.rule = [];
		return this;
	}
	AscFormat.InitClass(RuleSet_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// ValidationProperties_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/validationproperties_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function ValidationProperties_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.lastValidated = null;
		this.showIgnored = null;
		return this;
	}
	AscFormat.InitClass(ValidationProperties_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
// Элемент Rel (DataRecordSet_Type complexType): https://learn.microsoft.com/ru-ru/office/client-developer/visio/rel-element-datarecordset_type-complextypevisio-xml
// DataRecordSet_Type complexType: https://learn.microsoft.com/ru-ru/office/client-developer/visio/datarecordset_type-complextypevisio-xml
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function DataRecordSet_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.id = null;
		this.connectionID = null;
		this.command = null;
		this.options = null;
		this.timeRefreshed = null;
		this.nextRowID = null;
		this.name = null;
		this.rowOrder = null;
		this.refreshOverwriteAll = null;
		this.refreshNoReconciliationUI = null;
		this.refreshInterval = null;
		this.replaceLinks = null;
		this.checksum = null;
		this.aDOData = null;
		this.rel = null;
		this.dataColumns = null;
		this.primaryKey = [];
		this.rowMap = [];
		this.refreshConflict = [];
		this.autoLinkComparison = [];
		this.content = null;
		return this;
	}
	AscFormat.InitClass(DataRecordSet_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function CellDefBase_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.n = null;
		this.t = null;
		this.f = null;
		this.ix = null;
		this.s = null;
		return this;
	}
	AscFormat.InitClass(CellDefBase_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function GeometryRow_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.t = null;
		return this;
	}
	AscFormat.InitClass(GeometryRow_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function IndexedRow_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.ix = null;
		return this;
	}
	AscFormat.InitClass(IndexedRow_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function NamedIndexedRow_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.n = null;
		this.localName = null;
		this.ix = null;
		return this;
	}
	AscFormat.InitClass(NamedIndexedRow_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function SolutionXML_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.name = null;
		this.any = null;
		this.anyAttr = null;
		return this;
	}
	AscFormat.InitClass(SolutionXML_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);

	// Docs old:
	/**
	 * @extends AscFormat.CBaseFormatNoIdObject
	 * @constructor
	 */
	function ExtendableCell_Type() {
		AscFormat.CBaseFormatNoIdObject.call(this);
		this.solutionXML = null;
		return this;
	}
	AscFormat.InitClass(ExtendableCell_Type, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_Unknown);







	//-------------------------------------------------------------export---------------------------------------------------
	window['Asc']            = window['Asc'] || {};
	window['AscCommon']      = window['AscCommon'] || {};
	window['AscCommonWord']  = window['AscCommonWord'] || {};
	window['AscCommonSlide'] = window['AscCommonSlide'] || {};
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscVisio']  = window['AscVisio'] || {};
	window['AscFormat']  = window['AscFormat'] || {};
	window['AscWord'] = window['AscWord'] || {};

	window['AscVisio'].c_oVsdxTextKind = c_oVsdxTextKind;
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
	window['AscVisio'].PublishSettings_Type = PublishSettings_Type;
	window['AscVisio'].ColorEntry_Type = ColorEntry_Type;
	window['AscVisio'].FaceName_Type = FaceName_Type;
	window['AscVisio'].RefBy_Type = RefBy_Type;
	window['AscVisio'].Master_Type = Master_Type;
	window['AscVisio'].MasterShortcut_Type = MasterShortcut_Type;
	window['AscVisio'].Connect_Type = Connect_Type;
	window['AscVisio'].DataConnection_Type = DataConnection_Type;
	window['AscVisio'].DocumentSettings_Type = DocumentSettings_Type;
	window['AscVisio'].DocumentProperties_Type = DocumentProperties_Type;
	window['AscVisio'].HeaderFooter_Type = HeaderFooter_Type;
	window['AscVisio'].DataTransferInfo_Type = DataTransferInfo_Type;
	window['AscVisio'].Solution_Type = Solution_Type;
	window['AscVisio'].EventItem_Type = EventItem_Type;
	window['AscVisio'].Window_Type = Window_Type;
	window['AscVisio'].tp_Type = tp_Type;
	window['AscVisio'].pp_Type = pp_Type;
	window['AscVisio'].cp_Type = cp_Type;
	window['AscVisio'].fld_Type = fld_Type;
	window['AscVisio'].Comments_Type = Comments_Type;
	window['AscVisio'].RuleTest_Type = RuleTest_Type;
	window['AscVisio'].RuleFilter_Type = RuleFilter_Type;
	window['AscVisio'].RowKeyValue_Type = RowKeyValue_Type;
	window['AscVisio'].DataColumn_Type = DataColumn_Type;
	window['AscVisio'].RuleInfo_Type = RuleInfo_Type;
	window['AscVisio'].IssueTarget_Type = IssueTarget_Type;
	window['AscVisio'].Rule_Type = Rule_Type;
	window['AscVisio'].RuleSetFlags_Type = RuleSetFlags_Type;
	window['AscVisio'].AutoLinkComparison_Type = AutoLinkComparison_Type;
	window['AscVisio'].RefreshConflict_Type = RefreshConflict_Type;
	window['AscVisio'].RowMap_Type = RowMap_Type;
	window['AscVisio'].PrimaryKey_Type = PrimaryKey_Type;
	window['AscVisio'].DataColumns_Type = DataColumns_Type;
	window['AscVisio'].ADOData_Type = ADOData_Type;
	window['AscVisio'].Rel_Type = Rel_Type;
	window['AscVisio'].CommentEntry_Type = CommentEntry_Type;
	window['AscVisio'].AuthorEntry_Type = AuthorEntry_Type;
	window['AscVisio'].RefreshableData_Type = RefreshableData_Type;
	window['AscVisio'].PublishedPage_Type = PublishedPage_Type;
	window['AscVisio'].HeaderFooterFont_Type = HeaderFooterFont_Type;
	window['AscVisio'].FooterRight_Type = FooterRight_Type;
	window['AscVisio'].FooterCenter_Type = FooterCenter_Type;
	window['AscVisio'].FooterLeft_Type = FooterLeft_Type;
	window['AscVisio'].HeaderRight_Type = HeaderRight_Type;
	window['AscVisio'].HeaderCenter_Type = HeaderCenter_Type;
	window['AscVisio'].HeaderLeft_Type = HeaderLeft_Type;
	window['AscVisio'].FooterMargin_Type = FooterMargin_Type;
	window['AscVisio'].HeaderMargin_Type = HeaderMargin_Type;
	window['AscVisio'].CustomMenusFile_Type = CustomMenusFile_Type;
	window['AscVisio'].SnapAngle_Type = SnapAngle_Type;
	window['AscVisio'].SnapExtensions_Type = SnapExtensions_Type;
	window['AscVisio'].TimePrinted_Type = TimePrinted_Type;
	window['AscVisio'].TimeEdited_Type = TimeEdited_Type;
	window['AscVisio'].TimeSaved_Type = TimeSaved_Type;
	window['AscVisio'].TimeCreated_Type = TimeCreated_Type;
	window['AscVisio'].CustomProp_Type = CustomProp_Type;
	window['AscVisio'].PreviewPicture_Type = PreviewPicture_Type;
	window['AscVisio'].BuildNumberEdited_Type = BuildNumberEdited_Type;
	window['AscVisio'].BuildNumberCreated_Type = BuildNumberCreated_Type;
	window['AscVisio'].Template_Type = Template_Type;
	window['AscVisio'].AlternateNames_Type = AlternateNames_Type;
	window['AscVisio'].HyperlinkBase_Type = HyperlinkBase_Type;
	window['AscVisio'].Desc_Type = Desc_Type;
	window['AscVisio'].Keywords_Type = Keywords_Type;
	window['AscVisio'].Category_Type = Category_Type;
	window['AscVisio'].Company_Type = Company_Type;
	window['AscVisio'].Manager_Type = Manager_Type;
	window['AscVisio'].Creator_Type = Creator_Type;
	window['AscVisio'].Subject_Type = Subject_Type;
	window['AscVisio'].Title_Type = Title_Type;
	window['AscVisio'].SectionDef_Type = SectionDef_Type;
	window['AscVisio'].FunctionDef_Type = FunctionDef_Type;
	window['AscVisio'].CellDef_Type = CellDef_Type;
	window['AscVisio'].Issue_Type = Issue_Type;
	window['AscVisio'].RuleSet_Type = RuleSet_Type;
	window['AscVisio'].ValidationProperties_Type = ValidationProperties_Type;
	window['AscVisio'].DataRecordSet_Type = DataRecordSet_Type;
	window['AscVisio'].CellDefBase_Type = CellDefBase_Type;
	window['AscVisio'].GeometryRow_Type = GeometryRow_Type;
	window['AscVisio'].IndexedRow_Type = IndexedRow_Type;
	window['AscVisio'].NamedIndexedRow_Type = NamedIndexedRow_Type;
	window['AscVisio'].SolutionXML_Type = SolutionXML_Type;
	window['AscVisio'].ExtendableCell_Type = SolutionXML_Type;

	window['AscVisio'].FOREIGN_TYPES_BITMAP = 0;
	window['AscVisio'].FOREIGN_TYPES_ENHMETAFILE = 1;
	window['AscVisio'].FOREIGN_TYPES_INK = 2;
	window['AscVisio'].FOREIGN_TYPES_OBJECT = 3;

	window['AscVisio'].COMPRESSION_TYPES_JPEG = 0;
	window['AscVisio'].COMPRESSION_TYPES_DIB = 1;
	window['AscVisio'].COMPRESSION_TYPES_PNG = 2;
	window['AscVisio'].COMPRESSION_TYPES_TIFF = 3;
	window['AscVisio'].COMPRESSION_TYPES_GIF = 4;
	
	window['AscVisio'].SHAPE_TYPES_GROUP = 0;
	window['AscVisio'].SHAPE_TYPES_GUIDE = 1;
	window['AscVisio'].SHAPE_TYPES_FOREIGN = 2;
	window['AscVisio'].SHAPE_TYPES_SHAPE = 3;

	window['AscVisio'].WINDOW_TYPES_DRAWING = 0;
	window['AscVisio'].WINDOW_TYPES_SHEET = 1;
	window['AscVisio'].WINDOW_TYPES_STENCIL = 2;
	
	window['AscVisio'].CONTAINER_TYPES_PAGE = 0;
	window['AscVisio'].CONTAINER_TYPES_SHEET = 1;
	window['AscVisio'].CONTAINER_TYPES_MASTER = 2;
})(window, window.document);
