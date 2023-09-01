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

'use strict';

(function (window)
{

	AscCommon.CGraphics.prototype.SetFontSlot = function () {};
	AscCommon.CGraphics.prototype.SetFont = function () {};
	AscCommon.CGraphics.prototype.SetFontInternal = function () {};
	window.AscFonts = window.AscFonts || {};
	AscFonts.g_fontApplication = {
		GetFontInfo    : function (sFontName)
		{
			if (sFontName === 'Cambria Math')
			{
				return new AscFonts.CFontInfo('Cambria Math', 40, 1, 433, 1, -1, -1, -1, -1, -1, -1);
			}
		},
		Init           : function () {},
		LoadFont       : function () {},
		GetFontInfoName: function () {}
	}

	window.g_fontApplication = AscFonts.g_fontApplication;

	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		callback({'saveLock': false});
	};
	AscCommon.loadSmartArtBinary = function ()
	{

	};
	let oGlobalShape
	const oGlobalLogicDocument = AscTest.CreateLogicDocument()
	editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState = function ()
	{
	};

	function getController()
	{
		return oGlobalLogicDocument.GetCurrentController();
	}

	function createGroup(arrObjects)
	{
		const oController = getController();
		oController.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			arrObjects[i].select(oController, 0);
		}

		return oController.createGroup();
	}

	function addToSelection(oObject)
	{
		const oController = getController();
		if (oObject.group)
		{
			const oMainGroup = oObject.group.getMainGroup();
			oMainGroup.select(oController, 0);
			oController.selection.groupSelection = oMainGroup;
		}
		oObject.select(oController, 0);
	}

	function remove()
	{
		oGlobalLogicDocument.Remove();
	}

	function selectOnlyObjects(arrObjects)
	{
		const oController = getController();
		oController.resetSelection();
		for (let i = 0; i < arrObjects.length; i += 1)
		{
			const oObject = arrObjects[i];
			if (oObject.group)
			{
				const oMainGroup = oObject.group.getMainGroup();
				oMainGroup.select(oController, 0);
				oController.selection.groupSelection = oMainGroup;
			}
			oObject.select(oController, 0);
		}
	}

	function getFirstSlide()
	{
		return oGlobalLogicDocument.Slides[0];
	}

	function moveToParagraph(oParagraph, bIsStart)
	{
		oParagraph.SetThisElementCurrent();
		if (bIsStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}
	}

	function getShapeWithParagraphHelper(sTextIntoShape, bResetSelection)
	{
		const oController = oGlobalLogicDocument.GetCurrentController();
		if (bResetSelection)
		{
			oController.resetSelection();
		}
		oGlobalShape = AscTest.createShape(oGlobalLogicDocument.Slides[0]);


		oGlobalShape.setTxBody(AscFormat.CreateTextBodyFromString(sTextIntoShape, editor.WordControl.m_oDrawingDocument, oGlobalShape));
		const oContent = oGlobalShape.getDocContent();
		const oParagraph = oContent.Content[0];

		return {oLogicDocument: oGlobalLogicDocument, oShape: oGlobalShape, oParagraph, oController, oContent};
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

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oLogicDocument, oParagraph} = getShapeWithParagraphHelper(sInitText);
		oParagraph.SetThisElementCurrent();
		oLogicDocument.MoveCursorToEndPos();
		onKeyDown(oEvent);
		const sTextAfterKeyDown = AscTest.GetParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function checkTextAfterKeyDownHelperEmpty(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, '');
	}

	function checkTextAfterKeyDownHelperHelloWorld(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, 'Hello World');
	}

	function checkRemoveObject(oObject, arrSpTree)
	{
		let bCheckRemoveFromSpTree = true;
		for (let nDrawingIndex = 0; nDrawingIndex < arrSpTree.length; nDrawingIndex += 1)
		{
			if (arrSpTree[nDrawingIndex] === oObject)
			{
				bCheckRemoveFromSpTree = false;
			}
		}
		return bCheckRemoveFromSpTree && oObject.bDeleted;
	}

	function createTable(nRows, nColumns)
	{
		const oGraphicFrame = oGlobalLogicDocument.Add_FlowTable(nColumns, nRows);
		const oTable = oGraphicFrame.graphicObject;
		//oTable.Resize(nColumns * 300, nRows * 200);
		for (let nRow = 0; nRow < nRows; nRow += 1)
		{
			for (let nColumn = 0; nColumn < nColumns; nColumn += 1)
			{
				const oCell = oTable.Content[nRow].Get_Cell(nColumn);
				const oContent = oCell.GetContent();
				AscFormat.AddToContentFromString(oContent, 'Cell' + nRow + 'x' + nColumn);
			}
		}
		oGlobalLogicDocument.Recalculate();
		return oGraphicFrame;
	}

	function createChart()
	{
		const oChart = editor.asc_getChartObject(Asc.c_oAscChartTypeSettings.lineNormal);
		oChart.setParent(oGlobalLogicDocument.Slides[0]);

		oChart.addToDrawingObjects();
		oChart.spPr.setXfrm(new AscFormat.CXfrm());
		oChart.spPr.xfrm.setOffX(0);
		oChart.spPr.xfrm.setOffY(0);
		oChart.spPr.xfrm.setExtX(100);
		oChart.spPr.xfrm.setExtY(100);
		oGlobalLogicDocument.Recalculate();
		oGlobalLogicDocument.Document_UpdateInterfaceState();
		oGlobalLogicDocument.CheckEmptyPlaceholderNotes();

		oGlobalLogicDocument.DrawingDocument.m_oWordControl.OnUpdateOverlay();
		return oChart;
	}

	function createShapeWithTitlePlaceholder()
	{
		const oShape = AscTest.createShape(oGlobalLogicDocument.Slides[0]);
		oShape.setNvSpPr(new AscFormat.UniNvPr());
		let oPh = new AscFormat.Ph();
		oPh.setType(AscFormat.phType_title);
		oShape.nvSpPr.nvPr.setPh(oPh);
		oShape.txBody = AscFormat.CreateTextBodyFromString('', oShape.getDrawingDocument(), oShape);

		oShape.recalculateContentWitCompiledPr();
		return oShape;
	}

	function testMoveHelper(oEvent, bMoveToEndPosition, bGetPos, bGetSelectedText)
	{
		const {
			oShape,
			oParagraph,
			oLogicDocument
		} = getShapeWithParagraphHelper('HelloworldHelloworldHelloworldHelloworldHelloworldHelloworldHello', true);
		oShape.setPaddings({Left: 0, Top: 0, Right: 0, Bottom: 0});
		oParagraph.SetThisElementCurrent();
		oParagraph.Pr.SetInd(0, 0, 0);
		oParagraph.Set_Align(AscCommon.align_Left);
		if (bMoveToEndPosition)
		{
			oLogicDocument.MoveCursorToEndPos();
		} else
		{
			oLogicDocument.MoveCursorToStartPos();
		}
		oShape.recalculateContentWitCompiledPr();
		oLogicDocument.RecalculateCurPos(true, true);

		onKeyDown(oEvent);

		let oPos;
		oLogicDocument.RecalculateCurPos(true, true);
		if (bGetPos)
		{

			oPos = oParagraph.GetCurPosXY(true, true);
		}
		let sSelectedText;
		if (bGetSelectedText)
		{
			sSelectedText = oParagraph.GetSelectedText();
		}
		return {oPos, sSelectedText};
	}

	function addPropertyToDocument(oPr)
	{
		oGlobalLogicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(oPr), true);
	}

	function executeCheckMoveShape(oEvent)
	{
		const oController = oGlobalLogicDocument.GetCurrentController();
		oController.resetSelection();
		oGlobalShape.spPr.xfrm.setOffX(0);
		oGlobalShape.spPr.xfrm.setOffY(0);
		oGlobalShape.select(oController, 0);
		oGlobalShape.recalculateTransform();
		editor.zoom100();
		onKeyDown(oEvent);
		return oGlobalShape;
	}

	function cleanPresentation()
	{
		goToPageWithFocus(0, FOCUS_OBJECT_THUMBNAILS);
		editor.WordControl.Thumbnails.SelectAll();
		const arrSelectedArray = oGlobalLogicDocument.GetSelectedSlides();
		oGlobalLogicDocument.deleteSlides(arrSelectedArray);
	}

	function checkSelectedSlides(arrSelectedSlides)
	{
		const arrPresentationSelectedSlides = oGlobalLogicDocument.GetSelectedSlides();
		return arrSelectedSlides.length === arrPresentationSelectedSlides.length && arrSelectedSlides.every((el, ind) => el === arrPresentationSelectedSlides[ind]);
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

	function onKeyDown(oEvent)
	{
		editor.WordControl.onKeyDown(oEvent);
	}

	function addSlide()
	{
		oGlobalLogicDocument.addNextSlide(0);
	}

	function createShape()
	{
		return AscTest.createShape(oGlobalLogicDocument.Slides[0]);
	}

	function goToPage(nPage)
	{
		editor.WordControl.GoToPage(nPage);
	}

	function goToPageWithFocus(nPage, eFocus)
	{
		goToPage(nPage);
		oGlobalLogicDocument.SetThumbnailsFocusElement(eFocus);
	}

	function onInput(arrCodes)
	{
		for (let i = 0; i < arrCodes.length; i += 1)
		{
			oGlobalLogicDocument.OnKeyPress(createEvent(arrCodes[i], false, false, false, false, false));
		}
	}

	function moveCursorRight(bWord, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorRight(!!bAddToSelect, !!bWord);
	}

	function moveCursorLeft(bWord, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorLeft(!!bAddToSelect, !!bWord);
	}

	function moveCursorDown(bCtrlKey, bAddToSelect)
	{
		oGlobalLogicDocument.MoveCursorDown(!!bAddToSelect, !!bCtrlKey);
	}

	function createMathInShape()
	{
		const {oShape, oParagraph} = getShapeWithParagraphHelper('', true);
		selectOnlyObjects([oShape]);
		moveToParagraph(oParagraph, true);
		editor.asc_AddMath2(c_oAscMathType.FractionVertical);
		return {oShape, oParagraph};
	}

	function checkDirectTextPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		const oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
	}

	function checkDirectParaPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = getShapeWithParagraphHelper('Hello World');
		const oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);
	}

	const AscTestShortcut = window.AscTestShortcut = {};
	AscTestShortcut.createMathInShape = createMathInShape;
	AscTestShortcut.moveCursorDown = moveCursorDown;
	AscTestShortcut.moveCursorLeft = moveCursorLeft;
	AscTestShortcut.moveCursorRight = moveCursorRight;
	AscTestShortcut.onInput = onInput;
	AscTestShortcut.goToPageWithFocus = goToPageWithFocus;
	AscTestShortcut.goToPage = goToPage;
	AscTestShortcut.createShape = createShape;
	AscTestShortcut.addSlide = addSlide;
	AscTestShortcut.onKeyDown = onKeyDown;
	AscTestShortcut.executeTestWithCatchEvent = executeTestWithCatchEvent;
	AscTestShortcut.createNativeEvent = createNativeEvent;
	AscTestShortcut.checkSelectedSlides = checkSelectedSlides;
	AscTestShortcut.cleanPresentation = cleanPresentation;
	AscTestShortcut.executeCheckMoveShape = executeCheckMoveShape;
	AscTestShortcut.addPropertyToDocument = addPropertyToDocument;
	AscTestShortcut.testMoveHelper = testMoveHelper;
	AscTestShortcut.createShapeWithTitlePlaceholder = createShapeWithTitlePlaceholder;
	AscTestShortcut.createChart = createChart;
	AscTestShortcut.createTable = createTable;
	AscTestShortcut.checkRemoveObject = checkRemoveObject;
	AscTestShortcut.checkTextAfterKeyDownHelperHelloWorld = checkTextAfterKeyDownHelperHelloWorld;
	AscTestShortcut.checkTextAfterKeyDownHelperEmpty = checkTextAfterKeyDownHelperEmpty;
	AscTestShortcut.getDirectParaPrHelper = getDirectParaPrHelper;
	AscTestShortcut.getDirectTextPrHelper = getDirectTextPrHelper;
	AscTestShortcut.createEvent = createEvent;
	AscTestShortcut.getShapeWithParagraphHelper = getShapeWithParagraphHelper;
	AscTestShortcut.moveToParagraph = moveToParagraph;
	AscTestShortcut.getFirstSlide = getFirstSlide;
	AscTestShortcut.selectOnlyObjects = selectOnlyObjects;
	AscTestShortcut.remove = remove;
	AscTestShortcut.addToSelection = addToSelection;
	AscTestShortcut.createGroup = createGroup;
	AscTestShortcut.getController = getController;
	AscTestShortcut.oGlobalLogicDocument = oGlobalLogicDocument;
	AscTestShortcut.oGlobalShape = oGlobalShape;
	AscTestShortcut.checkDirectTextPrAfterKeyDown = checkDirectTextPrAfterKeyDown;
	AscTestShortcut.checkDirectParaPrAfterKeyDown = checkDirectParaPrAfterKeyDown;
})(window)
