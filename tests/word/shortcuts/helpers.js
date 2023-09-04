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

(function (window)
{
	window.setTimeout = function (callback)
	{
		callback();
	}

	window.AscFonts = window.AscFonts || {};
	AscFonts.g_fontApplication = {
		GetFontInfo    : function (sFontName)
		{
			if (sFontName === 'Cambria Math')
			{
				return new AscFonts.CFontInfo('Cambria Math', 40, 1, 433, 1, -1, -1, -1, -1, -1, -1);
			}
		},
		Init           : function ()
		{

		},
		LoadFont       : function ()
		{

		},
		GetFontInfoName: function () {}
	}

	window.g_fontApplication = AscFonts.g_fontApplication;
	AscCommon.CTableId = Object;
	const AscTestShortcut = window.AscTestShortcut = {};

	AscCommon.CGraphics.prototype.SetFontSlot = function () {};
	AscCommon.CGraphics.prototype.SetFont = function () {};
	AscCommon.CGraphics.prototype.SetFontInternal = function () {};
	AscCommon.loadSmartArtBinary = function ()
	{

	};

	Asc.asc_docs_api.prototype._loadModules = function () {};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
		this.ImageLoader = AscCommon.g_image_loader;
		this.chartPreviewManager   = new AscCommon.ChartPreviewManager();
		this.textArtPreviewManager = new AscCommon.TextArtPreviewManager();

		AscFormat.initStyleManager();
	};
	let editor = new Asc.asc_docs_api({'id-view': 'editor_sdk'});
	window.editor = editor;
	editor.WordControl.m_oDrawingDocument.GetVisibleMMHeight = () => 100;
	editor.WordControl.OnUpdateOverlay = () => {};
	AscCommon.loadSdk = function ()
	{
		editor._onEndLoadSdk();
	}

	Asc.createPluginsManager = function ()
	{

	};

	AscFonts.FontPickerByCharacter = {
		checkText      : function (text, _this, callback)
		{
			callback.call(_this);
		},
		getFontBySymbol: function ()
		{

		}
	};

	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		callback({"saveLock": false});
	};

	let oGlobalLogicDocument;

	function createLogicDocument()
	{
		if (oGlobalLogicDocument)
			return oGlobalLogicDocument;

		editor.InitEditor();
		editor.bInit_word_control = true;
		editor.WordControl.StartMainTimer = function ()
		{

		};
		editor.WordControl.InitControl();

		oGlobalLogicDocument = editor.WordControl.m_oLogicDocument;
		editor.WordControl.m_oDrawingDocument.m_oLogicDocument = oGlobalLogicDocument;
		oGlobalLogicDocument.UpdateAllSectionsInfo();
		oGlobalLogicDocument.Set_DocumentPageSize(100, 100);
		var props = new Asc.CDocumentSectionProps();
		props.put_TopMargin(0);
		props.put_LeftMargin(0);
		props.put_BottomMargin(0);
		props.put_RightMargin(0);
		oGlobalLogicDocument.Set_SectionProps(props);
		oGlobalLogicDocument.private_IsStartTimeoutOnRecalc = function ()
		{
			return false;
		}
		return oGlobalLogicDocument;
	}

	createLogicDocument();
	oGlobalLogicDocument.UpdateAllSectionsInfo();

	AscCommon.g_font_loader.LoadFont = function ()
	{
		return false;
	}
	editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function ()
	{
	};
	AscTest.CreateParagraph = function ()
	{
		return new AscWord.CParagraph(editor.WordControl.m_oDrawingDocument);
	}

	function addPropertyToDocument(oPr)
	{
		oGlobalLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(oPr), true);
	}

	function checkTextAfterKeyDownHelperEmpty(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, '');
	}

	function moveCursorDown(AddToSelect, CtrlKey)
	{
		oGlobalLogicDocument.MoveCursorDown(AddToSelect, CtrlKey);
	}

	function moveCursorUp(AddToSelect, CtrlKey)
	{
		oGlobalLogicDocument.MoveCursorUp(AddToSelect, CtrlKey);
	}

	function moveToParagraph(oParagraph, bIsStart, bSkipRemoveSelection)
	{
		if (!bSkipRemoveSelection)
		{
			oGlobalLogicDocument.RemoveSelection();
		}
		if (oParagraph.Parent.SetThisElementCurrent)
		{
			oParagraph.Parent.SetThisElementCurrent();
		}

		oParagraph.SetThisElementCurrent();
		if (bIsStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}
		oGlobalLogicDocument.private_UpdateCursorXY(true, true);
	}

	function insertManualBreak()
	{
		const oProps = new CMathMenuBase();
		oProps.insert_ManualBreak()
		editor.asc_SetMathProps(oProps);
	}

	function resetLogicDocument(oLogicDocument)
	{
		oLogicDocument.SetDocPosType(AscCommonWord.docpostype_Content);
	}

	function clean()
	{
		oGlobalLogicDocument.RemoveFromContent(0, oGlobalLogicDocument.GetElementsCount(), false);
	}

	function getLogicDocumentWithParagraphs(arrText, bRecalculate)
	{
		resetLogicDocument(oGlobalLogicDocument);
		if (!oGlobalLogicDocument.TurnOffRecalc)
		{
			oGlobalLogicDocument.Start_SilentMode();
		}
		clean();
		if (Array.isArray(arrText))
		{
			for (let i = 0; i < arrText.length; i += 1)
			{
				addParagraphToDocumentWithText(arrText[i]);
			}
		}
		if (oGlobalLogicDocument.TurnOffRecalc && bRecalculate)
		{
			oGlobalLogicDocument.End_SilentMode(true);
			oGlobalLogicDocument.private_UpdateCursorXY(true, true);
			recalculate();
		}

		const oFirstParagraph = oGlobalLogicDocument.Content[0];
		return {oLogicDocument: oGlobalLogicDocument, oParagraph: oFirstParagraph};
	}

	function recalculate()
	{
		oGlobalLogicDocument.RecalculateFromStart(false);
	}

	function addParagraphToDocumentWithText(sText)
	{
		const oParagraph = AscTest.CreateParagraph();
		oParagraph.Set_Ind({FirstLine: 0, Left: 0, Right: 0});
		oGlobalLogicDocument.Internal_Content_Add(oGlobalLogicDocument.Content.length, oParagraph);
		oParagraph.MoveCursorToEndPos();
		const oRun = new AscWord.CRun();
		oParagraph.AddToContent(0, oRun);
		oRun.AddText(sText);
		return oParagraph;
	}

	function remove()
	{
		oGlobalLogicDocument.Remove();
	}

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs([sInitText]);
		oParagraph.SetThisElementCurrent();
		oLogicDocument.MoveCursorToEndPos();
		onKeyDown(oEvent);
		const sTextAfterKeyDown = AscTest.GetParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function onKeyDown(oEvent)
	{
		editor.WordControl.onKeyDown(oEvent);
	}

	const arrCheckCodes = [48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 189, 187, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83,
		84, 85, 86, 87, 88, 89, 90, 219, 221, 186, 222, 220, 188, 190, 191, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 111, 106,
		109, 110, 107];

	function createNativeEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsMetaKey)
	{
		const bIsMacOs = AscCommon.AscBrowser.isMacOs;
		const oEvent = {};
		oEvent.isDefaultPrevented = false;
		oEvent.isPropagationStopped = false;
		oEvent.preventDefault = function ()
		{
			if (bIsMacOs && oEvent.altKey && !(oEvent.ctrlKey || oEvent.metaKey) && (arrCheckCodes.indexOf(nKeyCode) !== -1))
			{
				throw new Error('Alt key must not be disabled on macOS');
			}
			oEvent.isDefaultPrevented = true;
		};
		oEvent.stopPropagation = function ()
		{
			oEvent.isPropagationStopped = true;
		};

		oEvent.keyCode = nKeyCode;
		oEvent.ctrlKey = bIsCtrl;
		oEvent.shiftKey = bIsShift;
		oEvent.altKey = bIsAlt;
		oEvent.metaKey = bIsMetaKey;
		return oEvent;
	}

	function checkDirectTextPrAfterKeyDown(fCallback, nExpectedValue, sPrompt, oEvent, oAssert)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
		let oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
		return function recursive(fCallback2, nExpectedValue2, sPrompt2, oEvent2, oAssert2)
		{
			oTextPr = getDirectTextPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oTextPr), nExpectedValue2, sPrompt2);
			return recursive;
		}
	}

	function checkDirectParaPrAfterKeyDown(fCallback, nExpectedValue, sPrompt, oEvent, oAssert)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
		let oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);
		return function recursive(fCallback2, nExpectedValue2, sPrompt2, oEvent2, oAssert2)
		{
			oParaPr = getDirectParaPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oParaPr), nExpectedValue2, sPrompt2);
			return recursive;
		}
	}

	function getDirectTextPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		onKeyDown(oEvent);
		return oGlobalLogicDocument.GetDirectTextPr();
	}

	function getDirectParaPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		oGlobalLogicDocument.SelectAll();
		onKeyDown(oEvent);
		return oGlobalLogicDocument.GetDirectParaPr();
	}

	function getParagraphText(oParagraph)
	{
		return AscTest.GetParagraphText(oParagraph);
	}

	function executeTestWithCatchEvent(sSendEvent, fCustomCheck, customExpectedValue, oEvent, oAssert, fBeforeCallback)
	{
		fBeforeCallback && fBeforeCallback();

		let bCheck = false;

		const fCheck = function (...args)
		{
			if (fCustomCheck)
			{
				bCheck = fCustomCheck(...args);
			} else
			{
				bCheck = true;
			}
		}
		editor.asc_registerCallback(sSendEvent, fCheck);

		onKeyDown(oEvent);
		oAssert.strictEqual(bCheck, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
		editor.asc_unregisterCallback(sSendEvent, fCheck);
	}

	function moveCursorRight(AddToSelect, Word)
	{
		oGlobalLogicDocument.MoveCursorRight(AddToSelect, Word);
	}

	function moveCursorLeft(AddToSelect, Word)
	{
		oGlobalLogicDocument.MoveCursorLeft(AddToSelect, Word);
	}

	function selectAll()
	{
		oGlobalLogicDocument.SelectAll();
	}

	function getSelectedText()
	{
		return oGlobalLogicDocument.GetSelectedText(false, {TabSymbol: '\t'});
	}

	function createTest(oAssert)
	{
		return {deep: oAssert.deepEqual.bind(oAssert), equal: oAssert.strictEqual.bind(oAssert)};
	}

	function createChart()
	{
		var oDrawingDocument = editor.WordControl.m_oDrawingDocument;

		var oDrawing = new ParaDrawing(100, 100, null, oDrawingDocument, null, null);
		const oChartSpace = editor.asc_getChartObject(Asc.c_oAscChartTypeSettings.lineNormal);
		oChartSpace.spPr.setXfrm(new AscFormat.CXfrm());
		oChartSpace.spPr.xfrm.setOffX(0);
		oChartSpace.spPr.xfrm.setOffY(0);
		oChartSpace.spPr.xfrm.setExtX(100);
		oChartSpace.spPr.xfrm.setExtY(100);

		oChartSpace.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oChartSpace);
		oDrawing.setExtent(oChartSpace.spPr.xfrm.extX, oChartSpace.spPr.xfrm.extY);

		oDrawing.Set_DrawingType(drawing_Anchor);
		oDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		oDrawing.Set_Distance(0, 0, 0, 0);
		const oNearestPos = oGlobalLogicDocument.Get_NearestPos(0, oChartSpace.x, oChartSpace.y, true, oDrawing);
		oDrawing.Set_XYForAdd(oChartSpace.x, oChartSpace.y, oNearestPos, 0);
		oDrawing.AddToDocument(oNearestPos);
		oDrawing.CheckWH();
		recalculate();

		return oDrawing;
	}

	function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey)
	{
		const oKeyBoardEvent = new AscCommon.CKeyboardEvent();
		oKeyBoardEvent.KeyCode = nKeyCode;
		oKeyBoardEvent.ShiftKey = bIsShift;
		oKeyBoardEvent.AltKey = bIsAlt;
		oKeyBoardEvent.CtrlKey = bIsCtrl;
		oKeyBoardEvent.MacCmdKey = bIsMacCmdKey;
		oKeyBoardEvent.AltGr = bIsAltGr;
		return oKeyBoardEvent;
	}

	function checkInsertElementByType(nType, sPrompt, oAssert, oEvent)
	{
		const {oParagraph} = getLogicDocumentWithParagraphs(['']);
		onKeyDown(oEvent);
		let bCheck = false;
		for (let i = 0; i < oParagraph.Content.length; i += 1)
		{
			const oRun = oParagraph.Content[i];
			for (let j = 0; j < oRun.Content.length; j += 1)
			{
				if (oRun.Content[j].Type === nType)
				{
					bCheck = true;
				}
			}
		}
		oAssert.true(bCheck, sPrompt);
	}

	function createParagraphWithText(sText)
	{
		const oParagraph = AscTest.CreateParagraph();
		const oRun = new AscWord.CRun();
		oParagraph.AddToContent(0, oRun);
		oRun.AddText(sText);
		return oParagraph;
	}

	function checkApplyParagraphStyle(sStyleName, sPrompt, oEvent, oAssert)
	{
		const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
		oLogicDocument.SelectAll();
		onKeyDown(oEvent);
		const oParagraphPr = oLogicDocument.GetDirectParaPr();
		const sPStyleName = oLogicDocument.Styles.Get_Name(oParagraphPr.Get_PStyle());
		oAssert.strictEqual(sPStyleName, sStyleName, sPrompt);
	}

	function createHyperlink()
	{
		const oProps = new Asc.CHyperlinkProperty({Anchor: '_top', Text: "Beginning of document"});
		editor.add_Hyperlink(oProps);
	}

	function createTable(nRows, nCols)
	{
		const {oLogicDocument} = getLogicDocumentWithParagraphs(['']);
		return oLogicDocument.AddInlineTable(nCols, nRows);
	}

	function moveToTable(oTable, bToStart)
	{
		oTable.Document_SetThisElementCurrent();
		if (bToStart)
		{
			oTable.MoveCursorToStartPos();
		} else
		{
			oTable.MoveCursorToEndPos();
		}
	}

	function createShape()
	{
		AscCommon.History.Create_NewPoint();
		const oDrawing = new ParaDrawing(200, 100, null, oGlobalLogicDocument.GetDrawingDocument(), oGlobalLogicDocument, null);
		const oShapeTrack = new AscFormat.NewShapeTrack('rect', 0, 0, oGlobalLogicDocument.theme, null, null, null, 0);
		oShapeTrack.track({}, 0, 0);
		const oShape = oShapeTrack.getShape(true, oGlobalLogicDocument.GetDrawingDocument(), null);
		oShape.setBDeleted(false);
		oShape.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oShape);
		oDrawing.Set_DrawingType(drawing_Anchor);
		oDrawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		oDrawing.Set_Distance(0, 0, 0, 0);
		const oNearestPos = oGlobalLogicDocument.Get_NearestPos(0, oShape.x, oShape.y, true, oDrawing);
		oDrawing.Set_XYForAdd(oShape.x, oShape.y, oNearestPos, 0);
		oDrawing.AddToDocument(oNearestPos);
		oDrawing.CheckWH();
		recalculate();
		return oDrawing;
	}

	function getShapeWithText(sText, bStartRecalculate)
	{
		const oParaDrawing = createShape();
		selectParaDrawing(oParaDrawing);
		const oShape = oParaDrawing.GraphicObj;
		oShape.createTextBoxContent();
		const oParagraph = oShape.getDocContent().Content[0];
		moveToParagraph(oParagraph);
		addText(sText)
		if (bStartRecalculate)
		{
			startRecalculate();
		}
		return {oParaDrawing, oParagraph};
	}

	function createGroup(arrDrawings)
	{
		selectOnlyObjects(arrDrawings);
		return drawingObjects().groupSelectedObjects();
	}

	function selectOnlyObjects(arrDrawings)
	{
		const oController = drawingObjects();
		oController.resetSelection();
		for (let i = 0; i < arrDrawings.length; i += 1)
		{
			const oObject = arrDrawings[i].GraphicObj;
			if (oObject.group)
			{
				const oMainGroup = oObject.group.getMainGroup();
				oMainGroup.select(oController, 0);
				oController.selection.groupSelection = oMainGroup;
			}
			oObject.select(oController, 0);
		}
	}

	function selectParaDrawing(oParaDrawing)
	{
		oGlobalLogicDocument.SelectDrawings([oParaDrawing], oGlobalLogicDocument);
	}

	function drawingObjects()
	{
		return oGlobalLogicDocument.DrawingObjects;
	}

	function logicContent()
	{
		return oGlobalLogicDocument.Content;
	}

	function directParaPr()
	{
		return oGlobalLogicDocument.GetDirectParaPr();
	}

	function mouseClick(x, y, page, isRight, count)
	{
		mouseDown(x, y, page, isRight, count);
		mouseUp(x, y, page, isRight, count);
	}

	function mouseDown(x, y, page, isRight, count)
	{
		if (!oGlobalLogicDocument)
			return;

		let e = new AscCommon.CMouseEventHandler();

		e.Button = isRight ? AscCommon.g_mouse_button_right : AscCommon.g_mouse_button_left;
		e.ClickCount = count ? count : 1;

		e.Type = AscCommon.g_mouse_event_type_down;
		oGlobalLogicDocument.OnMouseDown(e, x, y, page);
	}

	function mouseUp(x, y, page, isRight, count)
	{
		if (!oGlobalLogicDocument)
			return;

		let e = new AscCommon.CMouseEventHandler();

		e.Button = isRight ? AscCommon.g_mouse_button_right : AscCommon.g_mouse_button_left;
		e.ClickCount = count ? count : 1;

		e.Type = AscCommon.g_mouse_event_type_up;
		oGlobalLogicDocument.OnMouseUp(e, x, y, page);
	}

	function mouseMove(x, y, page, isRight, count)
	{
		if (!oGlobalLogicDocument)
			return;

		let e = new AscCommon.CMouseEventHandler();

		e.Button = isRight ? AscCommon.g_mouse_button_right : AscCommon.g_mouse_button_left;
		e.ClickCount = count ? count : 1;

		e.Type = AscCommon.g_mouse_event_type_move;
		oGlobalLogicDocument.OnMouseMove(e, x, y, page);
	}

	function directTextPr()
	{
		return oGlobalLogicDocument.GetDirectTextPr();
	}

	function addToParagraph(oElement)
	{
		oGlobalLogicDocument.AddToParagraph(oElement);
	}

	function addBreakPage()
	{
		addToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
	}

	function startRecalculate()
	{
		if (oGlobalLogicDocument.TurnOffRecalc)
		{
			oGlobalLogicDocument.End_SilentMode(true);
		}
		recalculate();
		oGlobalLogicDocument.private_UpdateCursorXY(true, true);
	}

	function createMath(nType)
	{
		return editor.asc_AddMath(nType);
	}

	function addText(sText)
	{
		oGlobalLogicDocument.AddTextWithPr(sText);
	}

	function moveShapeHelper(nX, nY, oAssert, oEvent)
	{
		getLogicDocumentWithParagraphs([''], true);
		const oParaDrawing = createShape();
		const oShape = oParaDrawing.GraphicObj;
		selectOnlyObjects([oParaDrawing]);
		onKeyDown(oEvent);
		oAssert.deepEqual({x: round(oShape.x, 13), y: round(oShape.y, 13)}, {
			x: round(nX * AscCommon.g_dKoef_pix_to_mm, 13),
			y: round(nY * AscCommon.g_dKoef_pix_to_mm, 13)
		});
		drawingObjects().resetSelection()
	}

	function drawingContentPosition()
	{
		const arrPos = drawingObjects().getTargetDocContent().GetContentPosition();
		return arrPos[arrPos.length - 1].Position;
	}

	let fOldCheckOFormUserMaster;
	function setFillingFormsMode(bState)
	{
		if (bState)
		{
			fOldCheckOFormUserMaster = oGlobalLogicDocument.CheckOFormUserMaster;
			oGlobalLogicDocument.CheckOFormUserMaster = function ()
			{
				return true;
			}
		}
		else
		{
			oGlobalLogicDocument.CheckOFormUserMaster = fOldCheckOFormUserMaster;
		}
		var oRole = new AscCommon.CRestrictionSettings();
		oRole.put_OFormRole("Anyone");
		editor.asc_setRestriction(bState ? Asc.c_oAscRestrictionType.OnlyForms : Asc.c_oAscRestrictionType.None, oRole);
		editor.asc_SetPerformContentControlActionByClick(bState);
		editor.asc_SetHighlightRequiredFields(bState);
	}

	let nKeyId = 0;

	function createCheckBox()
	{
		const oCheckBox = oGlobalLogicDocument.AddContentControlCheckBox();
		var props = new AscCommon.CContentControlPr();
		var specProps = new AscCommon.CSdtCheckBoxPr();
		var oFormProps = new AscCommon.CSdtFormPr('key' + nKeyId++, '', '', false);
		props.SetFormPr(oFormProps);
		props.put_CheckBoxPr(specProps);
		editor.asc_SetContentControlProperties(props, oCheckBox.GetId());
		return oCheckBox;
	}

	function createComboBox()
	{
		const oComboBox = oGlobalLogicDocument.AddContentControlComboBox();
		var props = new AscCommon.CContentControlPr();
		var specProps = new AscCommon.CSdtComboBoxPr();
		var oFormProps = new AscCommon.CSdtFormPr('key' + nKeyId++, '', '', false);
		props.SetFormPr(oFormProps);
		specProps.clear();
		specProps.add_Item('Hello', 'Hello');
		specProps.add_Item('World', 'World');
		specProps.add_Item('yo', 'yo');
		props.put_ComboBoxPr(specProps);
		editor.asc_SetContentControlProperties(props, oComboBox.GetId());
		return oComboBox;
	}
	function round(nNumber, nAmount)
	{
		const nPower = Math.pow(10, nAmount);
		return Math.round(nNumber * nPower) / nPower;
	}
	function createComplexForm()
	{
		const oComplexForm = oGlobalLogicDocument.AddContentControlTextForm();
		oComplexForm.SetFormPr(new AscWord.CSdtFormPr());
		var props = new AscCommon.CContentControlPr();
		var formTextPr = new AscCommon.CSdtTextFormPr();
		formTextPr.put_MultiLine(true);
		props.put_TextFormPr(formTextPr);
		editor.asc_SetContentControlProperties(props, oComplexForm.GetId());
		return oComplexForm;
	}

	function contentPosition()
	{
		const oPos = oGlobalLogicDocument.GetContentPosition();
		return oPos[oPos.length - 1].Position;
	}

	AscTestShortcut.addPropertyToDocument = addPropertyToDocument;
	AscTestShortcut.checkTextAfterKeyDownHelperEmpty = checkTextAfterKeyDownHelperEmpty;
	AscTestShortcut.getLogicDocumentWithParagraphs = getLogicDocumentWithParagraphs;
	AscTestShortcut.resetLogicDocument = resetLogicDocument;
	AscTestShortcut.oGlobalLogicDocument = oGlobalLogicDocument;
	AscTestShortcut.onKeyDown = onKeyDown;
	AscTestShortcut.moveToParagraph = moveToParagraph;
	AscTestShortcut.createNativeEvent = createNativeEvent;
	AscTestShortcut.addParagraphToDocumentWithText = addParagraphToDocumentWithText;

	AscTestShortcut.checkDirectTextPrAfterKeyDown = checkDirectTextPrAfterKeyDown;
	AscTestShortcut.checkDirectParaPrAfterKeyDown = checkDirectParaPrAfterKeyDown;
	AscTestShortcut.executeTestWithCatchEvent = executeTestWithCatchEvent;
	AscTestShortcut.remove = remove;
	AscTestShortcut.recalculate = recalculate;
	AscTestShortcut.clean = clean;
	AscTestShortcut.moveCursorDown = moveCursorDown;
	AscTestShortcut.moveCursorUp = moveCursorUp;
	AscTestShortcut.moveCursorLeft = moveCursorLeft;
	AscTestShortcut.moveCursorRight = moveCursorRight;
	AscTestShortcut.selectAll = selectAll;
	AscTestShortcut.getSelectedText = getSelectedText;
	AscTestShortcut.createTest = createTest;
	AscTestShortcut.createChart = createChart;
	AscTestShortcut.createEvent = createEvent;
	AscTestShortcut.checkInsertElementByType = checkInsertElementByType;
	AscTestShortcut.createParagraphWithText = createParagraphWithText;
	AscTestShortcut.checkApplyParagraphStyle = checkApplyParagraphStyle;
	AscTestShortcut.createHyperlink = createHyperlink;
	AscTestShortcut.createTable = createTable;
	AscTestShortcut.moveToTable = moveToTable;
	AscTestShortcut.createShape = createShape;
	AscTestShortcut.getShapeWithText = getShapeWithText;
	AscTestShortcut.createGroup = createGroup;
	AscTestShortcut.selectOnlyObjects = selectOnlyObjects;
	AscTestShortcut.selectParaDrawing = selectParaDrawing;
	AscTestShortcut.drawingObjects = drawingObjects;
	AscTestShortcut.logicContent = logicContent;
	AscTestShortcut.directParaPr = directParaPr;
	AscTestShortcut.directTextPr = directTextPr;
	AscTestShortcut.addToParagraph = addToParagraph;
	AscTestShortcut.addBreakPage = addBreakPage;
	AscTestShortcut.startRecalculate = startRecalculate;
	AscTestShortcut.createMath = createMath;
	AscTestShortcut.addText = addText;
	AscTestShortcut.moveShapeHelper = moveShapeHelper;
	AscTestShortcut.drawingContentPosition = drawingContentPosition;
	AscTestShortcut.setFillingFormsMode = setFillingFormsMode;
	AscTestShortcut.createCheckBox = createCheckBox;
	AscTestShortcut.createComboBox = createComboBox;
	AscTestShortcut.createComplexForm = createComplexForm;
	AscTestShortcut.contentPosition = contentPosition;

	AscTestShortcut.mouseDown = mouseDown;
	AscTestShortcut.mouseUp = mouseUp;
	AscTestShortcut.mouseClick = mouseClick;
	AscTestShortcut.mouseMove = mouseMove;
	AscTestShortcut.resetLogicDocument = resetLogicDocument;
	AscTestShortcut.insertManualBreak = insertManualBreak;
})(window);
