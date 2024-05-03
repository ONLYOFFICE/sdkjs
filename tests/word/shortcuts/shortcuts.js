/*
 * (c) Copyright Ascensio System SIA 2010-2024
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

'use strict';


(function (window)
{
	const testHotkeyEvents = AscTestShortcut.testHotkeyEvents;
	const testHotkeyActions = AscTestShortcut.testHotkeyActions;
	const ExecuteShortcut = AscTestShortcut.ExecuteShortcut;
	const ExecuteHotkey = AscTestShortcut.ExecuteHotkey;
	let logicDocument = AscTest.CreateLogicDocument();

	const pageWidth = 100;
	const pageHeight = 100;
	logicDocument.Set_DocumentPageSize(pageWidth, pageHeight);
	const props = new Asc.CDocumentSectionProps();
	props.put_TopMargin(0);
	props.put_LeftMargin(0);
	props.put_BottomMargin(0);
	props.put_RightMargin(0);
	logicDocument.Set_SectionProps(props);

	AscFormat.CHART_STYLE_MANAGER.init();

	const drawingDocument = editor.WordControl.m_oDrawingDocument;
	drawingDocument.GetVisibleMMHeight = function ()
	{
		return 100;
	};

	let isTrack = false;
	drawingDocument.StartTrackText = function ()
	{
		isTrack = true;
	};
	drawingDocument.EndTrackText = function ()
	{
		isTrack = false;
	};
	drawingDocument.CancelTrackText = function ()
	{
		isTrack = false;
	};

	drawingDocument.IsTrackText = function ()
	{
		return isTrack;
	};

	drawingDocument.SetCursorType = function () {};
	drawingDocument.OnRePaintAttack = function() {};
	drawingDocument.ConvertCoordsToCursorWR  = function(X, Y) {return {X: X, Y: Y}};
	drawingDocument.OnUpdateOverlay = function() {};
	drawingDocument.Set_RulerState_HdrFtr = function() {};

	editor.retrieveFormatPainterData = Asc.asc_docs_api.prototype.retrieveFormatPainterData.bind(editor);
	editor.get_ShowParaMarks = Asc.asc_docs_api.prototype.get_ShowParaMarks.bind(editor);
	editor.put_ShowParaMarks = Asc.asc_docs_api.prototype.put_ShowParaMarks.bind(editor);
	editor.sync_ShowParaMarks = Asc.asc_docs_api.prototype.sync_ShowParaMarks.bind(editor);
	editor.private_GetLogicDocument = Asc.asc_docs_api.prototype.private_GetLogicDocument.bind(editor);
	editor.asc_AddTableOfContents = Asc.asc_docs_api.prototype.asc_AddTableOfContents.bind(editor);
	editor.asc_registerCallback = Asc.asc_docs_api.prototype.asc_registerCallback.bind(editor);
	editor.asc_unregisterCallback = Asc.asc_docs_api.prototype.asc_unregisterCallback.bind(editor);
	editor.sendEvent = Asc.asc_docs_api.prototype.sendEvent.bind(editor);
	editor.sync_DialogAddHyperlink = Asc.asc_docs_api.prototype.sync_DialogAddHyperlink.bind(editor);
	editor.sync_ParaStyleName = Asc.asc_docs_api.prototype.sync_ParaStyleName.bind(editor);
	editor.sync_MouseMoveStartCallback = Asc.asc_docs_api.prototype.sync_MouseMoveStartCallback.bind(editor);
	editor.sync_MouseMoveCallback = Asc.asc_docs_api.prototype.sync_MouseMoveCallback.bind(editor);
	editor.sync_MouseMoveEndCallback = Asc.asc_docs_api.prototype.sync_MouseMoveEndCallback.bind(editor);
	editor.sync_HideComment = Asc.asc_docs_api.prototype.sync_HideComment.bind(editor);
	editor.sync_ContextMenuCallback = Asc.asc_docs_api.prototype.sync_ContextMenuCallback.bind(editor);
	editor.asc_AddMath = Asc.asc_docs_api.prototype.asc_AddMath2.bind(editor);
	editor.sync_StartAddShapeCallback = Asc.asc_docs_api.prototype.sync_StartAddShapeCallback.bind(editor);
	editor.SetPaintFormat = Asc.asc_docs_api.prototype.SetPaintFormat.bind(editor);
	editor.SetMarkerFormat = Asc.asc_docs_api.prototype.SetMarkerFormat.bind(editor);
	editor.sync_MarkerFormatCallback = Asc.asc_docs_api.prototype.sync_MarkerFormatCallback.bind(editor);
	editor.sync_PaintFormatCallback = Asc.asc_docs_api.prototype.sync_PaintFormatCallback.bind(editor);
	editor.sync_EndAddShape = function () {};
	editor.isDocumentEditor = true;

	editor.getShortcut = function (e)
	{
		if (typeof e === 'number')
		{
			return e;
		}
	};
	editor.FontSizeIn = function ()
	{
		logicDocument.IncreaseDecreaseFontSize(true);
	};
	editor.FontSizeOut = function ()
	{
		logicDocument.IncreaseDecreaseFontSize(false);
	};
	editor.StartAddShape = function ()
	{
		this.isStartAddShape = true;
	};

	AscCommon.CDocsCoApi.prototype.askSaveChanges = function (callback)
	{
		window.setTimeout(function ()
		{
			callback({"saveLock": false});
		}, 0);
	};

	function GoToHeaderFooter(page, isHeader)
	{
		logicDocument.SetDocPosType(AscCommonWord.docpostype_HdrFtr);
		const mouseEvent = new AscCommon.CMouseEventHandler();
		mouseEvent.Button     = 0;
		mouseEvent.ClickCount = 1;

		logicDocument.OnMouseDown(mouseEvent, 0, isHeader ? 0 : pageHeight, page);
		logicDocument.OnMouseUp(mouseEvent, 0, isHeader ? 0 : pageHeight, page);
		logicDocument.OnMouseMove(mouseEvent, 0, 0, page);
		logicDocument.MoveCursorLeft();
	}

	function RemoveHeader(page)
	{
		logicDocument.RemoveHdrFtr(page, true);
	}

	function RemoveFooter(page)
	{
		logicDocument.RemoveHdrFtr(page, false);
	}

	function ClearDocumentAndAddParagraph(text)
	{
		logicDocument.RemoveSelection();
		AscTest.ClearDocument();
		const paragraph = CreateParagraphWithText(text);
		logicDocument.AddToContent(0, paragraph);
		return paragraph;
	}

	function CreateParagraphWithText(text)
	{
		let paragraph = AscTest.CreateParagraph();

		if (text)
		{
			let run = AscTest.CreateRun();
			run.AddText(text);
			paragraph.AddToContentToEnd(run);
		}

		return paragraph;
	}

	logicDocument.Start_SilentMode();

	function TurnOnRecalculate()
	{
		logicDocument.TurnOn_Recalculate();
	}

	function TurnOffRecalculate()
	{
		logicDocument.TurnOff_Recalculate();
	}

	function TurnOnRecalculateCurPos()
	{
		logicDocument.TurnOn_RecalculateCurPos();
	}

	function TurnOffRecalculateCurPos()
	{
		logicDocument.TurnOff_RecalculateCurPos();
	}

	function ApplyTextPrToDocument(textPr)
	{
		logicDocument.AddToParagraph(new AscCommonWord.ParaTextPr(textPr));
	}

	function GetDirectTextPr()
	{
		return logicDocument.GetDirectTextPr();
	}

	function GetDirectParaPr()
	{
		return logicDocument.GetDirectParaPr();
	}

	function AddShape(x, y, h, w)
	{
		const drawing = new ParaDrawing(w, h, null, logicDocument.GetDrawingDocument(), logicDocument, null);
		const shapeTrack = new AscFormat.NewShapeTrack('rect', x, y, logicDocument.theme, null, null, null, 0);
		shapeTrack.track({}, x+ w, y + h);
		const shape = shapeTrack.getShape(true, logicDocument.GetDrawingDocument(), null);
		shape.setBDeleted(false);

		shape.setParent(drawing);
		drawing.Set_GraphicObject(shape);
		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_Distance(0, 0, 0, 0);
		logicDocument.AddToParagraph(drawing);
		return drawing;
	}

	function SelectDrawings(drawings)
	{
		logicDocument.SelectDrawings(drawings, logicDocument);
	}

	function GetDrawingObjects()
	{
		return logicDocument.DrawingObjects;
	}

	function AddCheckBox()
	{
		const checkBox = logicDocument.AddContentControlCheckBox();
		const specProps = new AscCommon.CSdtCheckBoxPr();
		checkBox.ApplyCheckBoxPr(specProps);
		checkBox.SetFormPr(new AscCommon.CSdtFormPr());
		return checkBox;
	}

	function AddComboBox(items)
	{
		const comboBox = logicDocument.AddContentControlComboBox();
		const specProps = new AscCommon.CSdtComboBoxPr();
		specProps.clear();
		for (let i = 0; i < items.length; i++)
		{
			specProps.add_Item(items[i], items[i]);
		}

		comboBox.ApplyComboBoxPr(specProps);
		comboBox.SetFormPr(new AscCommon.CSdtFormPr());

		return comboBox;
	}

	function AddTable(row, column)
	{
		let table = AscTest.CreateTable(row, column);
		logicDocument.PushToContent(table);
		return table;
	}

	function AddChart()
	{
		const drawing = new ParaDrawing(100, 100, null, drawingDocument, null, null);
		const chartSpace = logicDocument.GetChartObject(Asc.c_oAscChartTypeSettings.lineNormal);
		chartSpace.spPr.setXfrm(new AscFormat.CXfrm());
		chartSpace.spPr.xfrm.setOffX(0);
		chartSpace.spPr.xfrm.setOffY(0);
		chartSpace.spPr.xfrm.setExtX(100);
		chartSpace.spPr.xfrm.setExtY(100);

		chartSpace.setParent(drawing);
		drawing.Set_GraphicObject(chartSpace);
		drawing.setExtent(chartSpace.spPr.xfrm.extX, chartSpace.spPr.xfrm.extY);

		drawing.Set_DrawingType(drawing_Anchor);
		drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		drawing.Set_Distance(0, 0, 0, 0);
		logicDocument.RemoveSelection();
		logicDocument.AddToParagraph(drawing);
		return drawing;
	}

	function AddComplexForm()
	{
		let complexForm = logicDocument.AddComplexForm();
		const formPr = new AscWord.CSdtFormPr();
		const formTextPr = new AscCommon.CSdtTextFormPr();
		formTextPr.put_MultiLine(true);
		complexForm.SetFormPr(formPr);
		complexForm.SetTextFormPr(formTextPr);
		return complexForm;
	}

	function round(number, amount)
	{
		const power = Math.pow(10, amount);
		return Math.round(number * power) / power;
	}

	$(function ()
	{
		QUnit.module("Test shortcut actions");
		QUnit.test('Check page break shortcut', (assert) =>
		{
			TurnOnRecalculate();
			ClearDocumentAndAddParagraph();
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertPageBreak);
			assert.strictEqual(logicDocument.GetPagesCount(), 2, 'Check page break shortcut');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertPageBreak);
			assert.strictEqual(logicDocument.GetPagesCount(), 3, 'Check page break shortcut');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertPageBreak);
			assert.strictEqual(logicDocument.GetPagesCount(), 4, 'Check page break shortcut');
			TurnOffRecalculate();
		});

		QUnit.test('Check line break shortcut', (assert) =>
		{
			TurnOnRecalculate();
			let paragraph = ClearDocumentAndAddParagraph();
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertLineBreak);
			assert.strictEqual(paragraph.GetLinesCount(), 2, 'Check line break shortcut');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertLineBreak);
			assert.strictEqual(paragraph.GetLinesCount(), 3, 'Check line break shortcut');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertLineBreak);
			assert.strictEqual(paragraph.GetLinesCount(), 4, 'Check line break shortcut');
			TurnOffRecalculate();
		});

		QUnit.test('Check column break shortcut', (assert) =>
		{
			TurnOnRecalculate();
			let paragraph = ClearDocumentAndAddParagraph();
			let sectionPr = AscTest.GetFinalSection();
			sectionPr.SetColumnsNum(3);
			AscTest.Recalculate();

			function CheckColumns(colCount)
			{
				assert.strictEqual(logicDocument.GetPagesCount(), 1, 'Check logic document page count');
				assert.strictEqual(paragraph.GetPagesCount(), colCount, 'Check paragraph page count');
				for (let i = 0; i < colCount; ++i)
				{
					assert.strictEqual(paragraph.GetAbsoluteColumn(i), i, 'Check paragraph column index');
					assert.strictEqual(paragraph.GetAbsolutePage(i), 0, 'Check paragraph page index');
				}
			}

			CheckColumns(1);
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertColumnBreak);
			CheckColumns(2);
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertColumnBreak);
			CheckColumns(3);

			sectionPr.SetColumnsNum(1);
			TurnOffRecalculate();
		});

		QUnit.test('Check reset char shortcut', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello world');
			logicDocument.SelectAll();
			ApplyTextPrToDocument({Bold: true, Italic: true, Underline: true});

			let textPr = GetDirectTextPr();
			assert.true(true === textPr.GetBold() && true === textPr.GetItalic() && true === textPr.GetUnderline(), 'Check before reset');
			ExecuteShortcut(c_oAscDocumentShortcutType.ResetChar);
			textPr = GetDirectTextPr();
			assert.true(undefined === textPr.GetBold() && undefined === textPr.GetItalic() && undefined === textPr.GetUnderline(), 'Check after reset');
		});

		QUnit.test('Check adding various characters', (assert) =>
		{
			let paragraph = ClearDocumentAndAddParagraph();

			ExecuteShortcut(c_oAscDocumentShortcutType.NonBreakingSpace);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0), 'Check add non breaking space');
			ExecuteShortcut(c_oAscDocumentShortcutType.CopyrightSign);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9), 'Check add CopyrightSign');
			ExecuteShortcut(c_oAscDocumentShortcutType.EuroSign);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC), 'Check add EuroSign');
			ExecuteShortcut(c_oAscDocumentShortcutType.RegisteredSign);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE), 'Check add RegisteredSign');
			ExecuteShortcut(c_oAscDocumentShortcutType.TrademarkSign);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122), 'Check add TrademarkSign');
			ExecuteShortcut(c_oAscDocumentShortcutType.EnDash);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122, 0x2013), 'Check add EnDash');
			ExecuteShortcut(c_oAscDocumentShortcutType.EmDash);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122, 0x2013, 0x2014), 'Check add EmDash');
			ExecuteShortcut(c_oAscDocumentShortcutType.NonBreakingHyphen);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122, 0x2013, 0x2014, 0x002D), 'Check add NonBreakingHyphen');
			ExecuteShortcut(c_oAscDocumentShortcutType.HorizontalEllipsis);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122, 0x2013, 0x2014, 0x002D, 0x2026), 'Check add HorizontalEllipsis');
			ExecuteHotkey(testHotkeyActions.addSJKSpace);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), String.fromCharCode(0x00A0, 0x00A9, 0x20AC, 0x00AE, 0x2122, 0x2013, 0x2014, 0x002D, 0x2026, 0x0020), 'Check add SJK space');
		});

		QUnit.test('Check text property change', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello world');
			logicDocument.SelectAll();

			ExecuteShortcut(c_oAscDocumentShortcutType.Bold);
			assert.strictEqual(GetDirectTextPr().GetBold(), true, 'Check turn on bold');
			ExecuteShortcut(c_oAscDocumentShortcutType.Bold);
			assert.strictEqual(GetDirectTextPr().GetBold(), false, 'Check turn off bold');

			ExecuteShortcut(c_oAscDocumentShortcutType.Italic);
			assert.strictEqual(GetDirectTextPr().GetItalic(), true, 'Check turn on italic');
			ExecuteShortcut(c_oAscDocumentShortcutType.Italic);
			assert.strictEqual(GetDirectTextPr().GetItalic(), false, 'Check turn off italic');

			ExecuteShortcut(c_oAscDocumentShortcutType.Strikeout);
			assert.strictEqual(GetDirectTextPr().GetStrikeout(), true, 'Check turn on strikeout');
			ExecuteShortcut(c_oAscDocumentShortcutType.Strikeout);
			assert.strictEqual(GetDirectTextPr().GetStrikeout(), false, 'Check turn off strikeout');

			ExecuteShortcut(c_oAscDocumentShortcutType.Underline);
			assert.strictEqual(GetDirectTextPr().GetUnderline(), true, 'Check turn on underline');
			ExecuteShortcut(c_oAscDocumentShortcutType.Underline);
			assert.strictEqual(GetDirectTextPr().GetUnderline(), false, 'Check turn off underline');

			ExecuteShortcut(c_oAscDocumentShortcutType.Superscript);
			assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_SuperScript, 'Check turn on superscript');
			ExecuteShortcut(c_oAscDocumentShortcutType.Superscript);
			assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off superscript');

			ExecuteShortcut(c_oAscDocumentShortcutType.Subscript);
			assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_SubScript, 'Check turn on subscript');
			ExecuteShortcut(c_oAscDocumentShortcutType.Subscript);
			assert.strictEqual(GetDirectTextPr().GetVertAlign(), AscCommon.vertalign_Baseline, 'Check turn off subscript');

			// defaultSize = 10
			// 10 -> 11 -> 12 -> 14 -> 16 -> 14 -> 12 -> 11 -> 10
			ExecuteShortcut(c_oAscDocumentShortcutType.IncreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 11, 'Check increase font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.IncreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 12, 'Check increase font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.IncreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 14, 'Check increase font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.IncreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 16, 'Check increase font size');

			ExecuteShortcut(c_oAscDocumentShortcutType.DecreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 14, 'Check decrease font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.DecreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 12, 'Check decrease font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.DecreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 11, 'Check decrease font size');
			ExecuteShortcut(c_oAscDocumentShortcutType.DecreaseFontSize);
			assert.strictEqual(GetDirectTextPr().GetFontSize(), 10, 'Check decrease font size');
		});

		QUnit.test('Check select all shortcut', (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('Hello world');
			const table = AscTest.CreateTable(2, 2);
			logicDocument.AddToContent(1, table);
			assert.strictEqual(logicDocument.IsSelectionUse(), false, 'Check document selection');
			ExecuteShortcut(c_oAscDocumentShortcutType.EditSelectAll);
			assert.strictEqual(logicDocument.IsSelectionUse(), true, 'Check document selection');
			assert.strictEqual(paragraph.IsSelectedAll(), true, 'Check paragraph selection');
			assert.strictEqual(table.IsSelectedAll(), true, 'Check table selection');
		});

		QUnit.test('Check paragraph property change', (assert) =>
		{
			let paragraph = ClearDocumentAndAddParagraph('Hello world');

			function GetStyleName()
			{
				return logicDocument.GetStyleManager().GetName(paragraph.GetParagraphStyle());
			}

			assert.strictEqual(GetStyleName(), "", "Check style");
			ExecuteShortcut(c_oAscDocumentShortcutType.ApplyHeading1);
			assert.strictEqual(GetStyleName(), "Heading 1", "Check apply heading 1");
			ExecuteShortcut(c_oAscDocumentShortcutType.ApplyHeading2);
			assert.strictEqual(GetStyleName(), "Heading 2", "Check apply heading 2");
			ExecuteShortcut(c_oAscDocumentShortcutType.ApplyHeading3);
			assert.strictEqual(GetStyleName(), "Heading 3", "Check apply heading 3");

			assert.strictEqual(GetDirectParaPr().GetJc(), undefined, "Check justification");
			ExecuteShortcut(c_oAscDocumentShortcutType.CenterPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Center, "Check turn on center para");
			ExecuteShortcut(c_oAscDocumentShortcutType.CenterPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check turn off center para");

			ExecuteShortcut(c_oAscDocumentShortcutType.JustifyPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");
			ExecuteShortcut(c_oAscDocumentShortcutType.JustifyPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check turn off justify para");

			ExecuteShortcut(c_oAscDocumentShortcutType.JustifyPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Justify, "Check turn on justify para");
			ExecuteShortcut(c_oAscDocumentShortcutType.LeftPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check turn on left para");
			ExecuteShortcut(c_oAscDocumentShortcutType.LeftPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Justify, "Check turn off left para");

			ExecuteShortcut(c_oAscDocumentShortcutType.RightPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Right, "Check turn on right para");
			ExecuteShortcut(c_oAscDocumentShortcutType.RightPara);
			assert.strictEqual(GetDirectParaPr().GetJc(), AscCommon.align_Left, "Check turn off right para");

			ExecuteShortcut(c_oAscDocumentShortcutType.Indent);
			assert.strictEqual(GetDirectParaPr().GetIndLeft(), 12.5, "Check indent");

			ExecuteShortcut(c_oAscDocumentShortcutType.UnIndent);
			assert.strictEqual(GetDirectParaPr().GetIndLeft(), 0, "Check indent");

			CreateParagraphWithText('Hello');

			logicDocument.SelectAll();

			ExecuteHotkey(testHotkeyActions.testIndent);
			assert.strictEqual(GetDirectParaPr().GetIndLeft(), 12.5, "Check multi indent");

			ExecuteHotkey(testHotkeyActions.testUnIndent);
			assert.strictEqual(GetDirectParaPr().GetIndLeft(), 0, "Check multi unindent");
		});

		QUnit.test('Check insert note elements', (assert) =>
		{
			let paragraph = ClearDocumentAndAddParagraph('');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertFootnoteNow);
			const footnotes = logicDocument.GetFootnotesList();
			assert.equal(footnotes.length, 1, 'Check insert footnote');

			paragraph.SetThisElementCurrent();
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertEndnoteNow);
			const endNotes = logicDocument.GetEndnotesList();
			assert.equal(endNotes.length, 1, 'Check insert endnote');
			logicDocument.MoveCursorToStartPos();
		});


		QUnit.test('Check sending event to interface', (assert) =>
		{
			function checkSendingEvent(sSendEvent, oEvent)
			{
				let isCheck = false;
				const fCheck = function ()
				{
						isCheck = true;
				}
				editor.asc_registerCallback(sSendEvent, fCheck);

				ExecuteShortcut(oEvent);
				assert.strictEqual(isCheck, true, 'Check catch ' + sSendEvent + ' event');
				editor.asc_unregisterCallback(sSendEvent, fCheck);
			}

			checkSendingEvent("asc_onDialogAddHyperlink", c_oAscDocumentShortcutType.InsertHyperlink);
			checkSendingEvent("asc_onPrint", c_oAscDocumentShortcutType.PrintPreviewAndPrint);

			checkSendingEvent('asc_onMouseMoveStart', testHotkeyEvents[testHotkeyActions.closeAllWindowsPopups][0]);
			checkSendingEvent('asc_onMouseMove', testHotkeyEvents[testHotkeyActions.closeAllWindowsPopups][0]);
			checkSendingEvent('asc_onMouseMoveEnd', testHotkeyEvents[testHotkeyActions.closeAllWindowsPopups][0]);

			checkSendingEvent('asc_onContextMenu', testHotkeyEvents[testHotkeyActions.showContextMenu][0]);
			AscCommon.AscBrowser.isOpera = true;
			checkSendingEvent('asc_onContextMenu', testHotkeyEvents[testHotkeyActions.showContextMenu][1]);
			AscCommon.AscBrowser.isOpera = false;
			checkSendingEvent('asc_onContextMenu', testHotkeyEvents[testHotkeyActions.showContextMenu][2]);
		});

		QUnit.test('Check insert equation', (assert) =>
		{
			ClearDocumentAndAddParagraph('');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertEquation);
			const math = logicDocument.GetCurrentMath();
			assert.true(!!math, 'Check insert equation');
		});

		QUnit.test('Check insert page number', (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			ExecuteShortcut(c_oAscDocumentShortcutType.InsertPageNumber);

			const firstRun = paragraph.Content[0];
			assert.strictEqual(firstRun.Content[0].Type, para_PageNum, "Check insert page number");
		});

		QUnit.test('Check toggle bullet list', (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			assert.false(paragraph.IsBulletedNumbering(), 'check disable bullet list');
			ExecuteShortcut(c_oAscDocumentShortcutType.ApplyListBullet);
			assert.true(paragraph.IsBulletedNumbering(), 'check apply bullet list');
		});

		QUnit.test('Check copy/paste format', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello');
			ApplyTextPrToDocument({Bold: true, Italic: true, Underline: true});
			GetDirectTextPr();
			ExecuteShortcut(c_oAscDocumentShortcutType.CopyFormat);
			let textPr = editor.getFormatPainterData().TextPr;
			assert.true(textPr.Get_Bold(), 'Check copy bold format from paragraph');
			assert.true(textPr.Get_Italic(), 'Check copy italic format from paragraph');
			assert.true(textPr.Get_Underline(), 'Check copy underline format from paragraph');

			ClearDocumentAndAddParagraph('');
			ExecuteShortcut(c_oAscDocumentShortcutType.PasteFormat);
			textPr = GetDirectTextPr();
			assert.true(textPr.Get_Bold(), 'Check paste bold format from paragraph');
			assert.true(textPr.Get_Italic(), 'Check paste italic format from paragraph');
			assert.true(textPr.Get_Underline(), 'Check paste underline format from paragraph');
		});

		QUnit.test('Check undo/redo history', (assert) =>
		{
			let paragraph = ClearDocumentAndAddParagraph('Hello');
			paragraph.MoveCursorToEndPos();
			logicDocument.AddTextWithPr(' World');
			ExecuteShortcut(c_oAscDocumentShortcutType.EditUndo);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello', 'Check undo');

			ExecuteShortcut(c_oAscDocumentShortcutType.EditRedo);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello World', 'Check redo');
		});

		QUnit.test('Check show/hide non printing symbols', (assert) =>
		{
			ExecuteShortcut(c_oAscDocumentShortcutType.ShowAll);
			assert.true(editor.ShowParaMarks, 'Check show non printing characters shortcut');

			ExecuteShortcut(c_oAscDocumentShortcutType.ShowAll);
			assert.false(editor.ShowParaMarks, 'Check hide non printing characters shortcut');
		});

		QUnit.test('Check save', (assert) =>
		{
			assert.timeout(100);
			const done = assert.async();

			const fOldSave = editor._onSaveCallbackInner;
			editor._onSaveCallbackInner = function ()
			{
				assert.true(true, 'Check save shortcut');
				done();
				editor._onSaveCallbackInner = fOldSave;
			};
			editor._saveCheck = () => true;
			editor.asc_isDocumentCanSave = () => true;
			ExecuteShortcut(c_oAscDocumentShortcutType.Save, 'Check save');
		});

		QUnit.test('Check update fields', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello');
			const paragraph2 = CreateParagraphWithText('Hello');
			const paragraph3 = CreateParagraphWithText('Hello');

			logicDocument.PushToContent(paragraph2);
			logicDocument.PushToContent(paragraph3);
			for (let i = 0; i < logicDocument.Content.length; i += 1)
			{
				logicDocument.Set_CurrentElement(i, true);
				logicDocument.SetParagraphStyle("Heading 1");
			}

			logicDocument.MoveCursorToStartPos();
			logicDocument.AddTableOfContents(null, new Asc.CTableOfContentsPr());

			logicDocument.MoveCursorToEndPos();
			const paragraph4 = CreateParagraphWithText('Hello');
			logicDocument.PushToContent(paragraph4);
			paragraph4.SetThisElementCurrent(true);
			logicDocument.SetParagraphStyle("Heading 1");

			logicDocument.MoveCursorToStartPos();
			AscTest.MoveCursorRight()

			ExecuteShortcut(c_oAscDocumentShortcutType.UpdateFields);
			assert.strictEqual(logicDocument.Content[0].Content.Content.length, 5, 'Check update fields from 3 to 4 items');
		});

		QUnit.test('Check remove symbols', (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('Hello Hello Hello Hello Hello Hello Hello');

			ExecuteHotkey(testHotkeyActions.removeBackSymbol);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello Hell', 'Check removing back symbol');

			ExecuteHotkey(testHotkeyActions.removeBackWord);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello Hello ', 'Check removing back word');

			ExecuteHotkey(testHotkeyActions.removeBackWord, 1);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello Hello ', 'Check removing back word');

			logicDocument.MoveCursorToStartPos();
			ExecuteHotkey(testHotkeyActions.removeFrontSymbol);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'ello Hello Hello Hello Hello ', 'Check removing front symbol');
			ExecuteHotkey(testHotkeyActions.removeFrontWord);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello Hello ', 'Check removing front word');

			ExecuteHotkey(testHotkeyActions.removeFrontWord, 1);
			assert.strictEqual(AscTest.GetParagraphText(paragraph), 'Hello Hello Hello ', 'Check removing front word');
		});
		
		QUnit.test('Check move/select in no calculated text', (assert) =>
		{
			function CheckCursorPosition(expected, description)
			{
				const position = logicDocument.GetContentPosition();
				assert.strictEqual(position[position.length - 1].Position, expected, description);
			}
			
			ClearDocumentAndAddParagraph('The quick brown fox jumps over the lazy dog');
			
			logicDocument.MoveCursorToStartPos();
			
			ExecuteHotkey(testHotkeyActions.moveToRightChar);
			ExecuteHotkey(testHotkeyActions.moveToRightChar);
			ExecuteHotkey(testHotkeyActions.moveToRightChar);
			CheckCursorPosition(3, 'Check move to right char');
			
			ExecuteHotkey(testHotkeyActions.moveToEndLine);
			CheckCursorPosition(43, 'Check move to end line');
			
			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			CheckCursorPosition(38, 'Check move to left char');
			
			ExecuteHotkey(testHotkeyActions.moveToStartLine);
			CheckCursorPosition(0, 'Check move to start line');
			
			ExecuteHotkey(testHotkeyActions.moveToEndDocument);
			CheckCursorPosition(43, 'Check move to the end of the document');
			
			ExecuteHotkey(testHotkeyActions.moveToStartDocument);
			CheckCursorPosition(0, 'Check move to the start of the document');
			
			ExecuteHotkey(testHotkeyActions.moveToRightWord);
			ExecuteHotkey(testHotkeyActions.moveToRightWord);
			ExecuteHotkey(testHotkeyActions.moveToRightWord);
			CheckCursorPosition(16, 'Check move to the right by words');
			
			ExecuteHotkey(testHotkeyActions.moveToLeftWord);
			CheckCursorPosition(10, 'Check move to the left by words');
			
			ExecuteHotkey(testHotkeyActions.moveUp);
			CheckCursorPosition(10, 'Check move up');

			ExecuteHotkey(testHotkeyActions.moveDown);
			CheckCursorPosition(10, 'Check move down');
			
			ExecuteHotkey(testHotkeyActions.moveToNextPage);
			CheckCursorPosition(10, 'Check move to the next page');

			ExecuteHotkey(testHotkeyActions.moveToPreviousPage);
			CheckCursorPosition(10, 'Check move to the previous page');
		});
		
		QUnit.test('Check move/select in text', (assert) =>
		{
			function CheckCursorPosition(expected, description)
			{
				const position = logicDocument.GetContentPosition();
				assert.strictEqual(position[position.length - 1].Position, expected, description);
			}

			ClearDocumentAndAddParagraph(
				'Hello World Hello ' +
				'World Hello World ' +
				'Hello World Hello ' +
				'World Hello World ' +
				'Hello World Hello ' +
				'Hello World Hello ' +
				'Hello World Hello ' +
				'Hello World Hello ' +
				'World Hello World');

			logicDocument.MoveCursorToStartPos();
			TurnOnRecalculate();
			TurnOnRecalculateCurPos();
			AscTest.Recalculate();
			TurnOffRecalculate();
			TurnOffRecalculateCurPos();

			ExecuteHotkey(testHotkeyActions.moveToEndLine);
			CheckCursorPosition(18, 'Check move to end line');

			ExecuteHotkey(testHotkeyActions.moveToRightChar);
			CheckCursorPosition(19, 'Check move to right char');

			ExecuteHotkey(testHotkeyActions.moveToEndLine, 1);
			CheckCursorPosition(36, 'Check move to end line');

			ExecuteHotkey(testHotkeyActions.moveToEndLine, 1);
			CheckCursorPosition(36, 'Check move to end line');

			ExecuteHotkey(testHotkeyActions.moveToStartLine, 1);
			CheckCursorPosition(18, 'Check move to start line');

			ExecuteHotkey(testHotkeyActions.moveToStartLine, 1);
			CheckCursorPosition(18, 'Check move to start line');

			ExecuteHotkey(testHotkeyActions.moveToRightChar);
			CheckCursorPosition(19, 'Check move to right char');

			ExecuteHotkey(testHotkeyActions.moveToLeftChar);
			CheckCursorPosition(18, 'Check move to left char');

			ExecuteHotkey(testHotkeyActions.moveToLeftWord);
			CheckCursorPosition(12, 'Check move to left word');

			ExecuteHotkey(testHotkeyActions.moveToLeftWord, 1);
			CheckCursorPosition(6, 'Check move to left word');

			ExecuteHotkey(testHotkeyActions.moveToRightWord, 1);
			CheckCursorPosition(12, 'Check move to right word');

			ExecuteHotkey(testHotkeyActions.moveToRightWord);
			CheckCursorPosition(18, 'Check move to right word');

			ExecuteHotkey(testHotkeyActions.moveToRightWord);
			CheckCursorPosition(24, 'Check move to right word');


			ExecuteHotkey(testHotkeyActions.moveToStartLine);
			CheckCursorPosition(18, 'Check move to start line');

			ExecuteHotkey(testHotkeyActions.moveDown);
			CheckCursorPosition(36, 'Check move down');

			ExecuteHotkey(testHotkeyActions.moveUp);
			CheckCursorPosition(18, 'Check move up');

			ExecuteHotkey(testHotkeyActions.moveToEndDocument);
			CheckCursorPosition(161, 'Check move to end document');

			ExecuteHotkey(testHotkeyActions.moveToStartDocument);
			CheckCursorPosition(0, 'Check move to start document');

			AscTest.MoveCursorRight();

			ExecuteHotkey(testHotkeyActions.moveToNextPage);
			CheckCursorPosition(91, 'Check move to next page');

			ExecuteHotkey(testHotkeyActions.moveToPreviousPage);
			CheckCursorPosition(1, 'Check move to previous page');

			ExecuteHotkey(testHotkeyActions.moveToStartNextPage);
			CheckCursorPosition(90, 'Check move to start next page');

			ExecuteHotkey(testHotkeyActions.moveToStartPreviousPage);
			CheckCursorPosition(0, 'Check move to start previous page');

			function CheckSelectedText(expectedText, description)
			{
				const selectedText = logicDocument.GetSelectedText();
				assert.strictEqual(selectedText, expectedText, description);
			}

			ExecuteHotkey(testHotkeyActions.selectToEndLine);
			CheckSelectedText('Hello World Hello ', 'Select to end line');

			ExecuteHotkey(testHotkeyActions.selectRightChar);
			CheckSelectedText('Hello World Hello W', 'Select to right char');

			ExecuteHotkey(testHotkeyActions.selectToEndLine, 1);
			CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

			ExecuteHotkey(testHotkeyActions.selectToEndLine, 1);
			CheckSelectedText('Hello World Hello World Hello World ', 'Select to end line content');

			ExecuteHotkey(testHotkeyActions.selectToStartLine, 1);
			CheckSelectedText('Hello World Hello ', 'Select to start line content');

			ExecuteHotkey(testHotkeyActions.selectToStartLine, 1);
			CheckSelectedText('Hello World Hello ', 'Select to start line content');

			ExecuteHotkey(testHotkeyActions.selectRightChar);
			CheckSelectedText('Hello World Hello W', 'Select to right char');

			ExecuteHotkey(testHotkeyActions.selectLeftChar);
			CheckSelectedText('Hello World Hello ', 'Select to left char');

			ExecuteHotkey(testHotkeyActions.selectLeftWord);
			CheckSelectedText('Hello World ', 'Select to left word');

			ExecuteHotkey(testHotkeyActions.selectRightWord);
			CheckSelectedText('Hello World Hello ', 'Select to right word');

			ExecuteHotkey(testHotkeyActions.selectLeftWord, 1);
			CheckSelectedText('Hello World ', 'Select to left word');

			ExecuteHotkey(testHotkeyActions.selectRightWord, 1);
			CheckSelectedText('Hello World Hello ', 'Select to right word');

			ExecuteHotkey(testHotkeyActions.selectRightWord);
			CheckSelectedText('Hello World Hello World ', 'Select to right word');

			ExecuteHotkey(testHotkeyActions.selectRightWord);
			CheckSelectedText('Hello World Hello World Hello ', 'Select to right word');

			ExecuteHotkey(testHotkeyActions.selectToStartLine);
			CheckSelectedText('Hello World Hello ', 'Select to start line');

			ExecuteHotkey(testHotkeyActions.selectDown);
			CheckSelectedText('Hello World Hello World Hello World ', 'Select down');

			ExecuteHotkey(testHotkeyActions.selectUp);
			CheckSelectedText('Hello World Hello ', 'Select up');

			ExecuteHotkey(testHotkeyActions.selectToEndDocument);
			CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello Hello World Hello Hello World Hello Hello World Hello World Hello World', 'Select to end document');

			ExecuteHotkey(testHotkeyActions.selectToStartDocument);
			CheckSelectedText('', 'Select to start document');

			logicDocument.MoveCursorToEndPos();
			ExecuteHotkey(testHotkeyActions.selectLeftChar);
			CheckSelectedText('d', 'Select to left char');

			ExecuteHotkey(testHotkeyActions.selectLeftWord);
			CheckSelectedText('World', 'Select to left word');

			logicDocument.MoveCursorToStartPos();
			AscTest.MoveCursorRight();
			ExecuteHotkey(testHotkeyActions.selectToNextPage);
			CheckSelectedText('ello World Hello World Hello World Hello World Hello World Hello World Hello World Hello H', 'Select to next page');
			AscTest.MoveCursorRight();

			ExecuteHotkey(testHotkeyActions.selectToPreviousPage);
			CheckSelectedText('ello World Hello World Hello World Hello World Hello World Hello World Hello World Hello H', 'Select to previous word');
			AscTest.MoveCursorLeft();
			ExecuteHotkey(testHotkeyActions.selectToStartNextPage);
			CheckSelectedText('ello World Hello World Hello World Hello World Hello World Hello World Hello World Hello ', 'Select to start next page');
			AscTest.MoveCursorRight();
			ExecuteHotkey(testHotkeyActions.selectToStartPreviousPage);
			CheckSelectedText('Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello ', 'Select to start previous page');
		});

		QUnit.test('Check actions with shapes', (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			TurnOnRecalculate();
			AscTest.Recalculate();

			const drawing1 = AddShape(0, 0, 100, 200);

			const dotsPerMM = logicDocument.DrawingDocument.GetDotsPerMM();

			function CheckShapePosition(X, Y)
			{
				assert.deepEqual([round(drawing1.X * dotsPerMM, 10), round(drawing1.Y * dotsPerMM, 10), drawing1.Extent.W, drawing1.Extent.H], [X, Y, 200, 100], 'Check shape position after movement');
			}

			SelectDrawings([drawing1]);

			ExecuteHotkey(testHotkeyActions.bigMoveGraphicObjectLeft);
			CheckShapePosition(-5, 0);

			ExecuteHotkey(testHotkeyActions.littleMoveGraphicObjectLeft);
			CheckShapePosition(-6, 0);

			ExecuteHotkey(testHotkeyActions.bigMoveGraphicObjectRight);
			CheckShapePosition(-1, 0);

			ExecuteHotkey(testHotkeyActions.littleMoveGraphicObjectRight);
			CheckShapePosition(0, 0);

			ExecuteHotkey(testHotkeyActions.bigMoveGraphicObjectDown);
			CheckShapePosition(0, 5);

			ExecuteHotkey(testHotkeyActions.littleMoveGraphicObjectDown);
			CheckShapePosition(0, 6);

			ExecuteHotkey(testHotkeyActions.bigMoveGraphicObjectUp);
			CheckShapePosition(0, 1);

			ExecuteHotkey(testHotkeyActions.littleMoveGraphicObjectUp);
			CheckShapePosition(0, 0);


			function CheckSelectedObjects(arrOfDrawings)
			{
				const length = Math.max(arrOfDrawings.length, GetDrawingObjects().selectedObjects.length);
				for (let i = 0; i < length; i++)
				{
					assert.true(GetDrawingObjects().selectedObjects[i] === arrOfDrawings[i].GraphicObj, 'Check selection movement between objects');
				}
			}

			logicDocument.RemoveSelection();
			const drawing2 = AddShape(0, 0, 10, 10);
			const drawing3 = AddShape(0, 0, 10, 10);
			SelectDrawings([drawing3]);

			ExecuteHotkey(testHotkeyActions.selectNextObject);
			CheckSelectedObjects([drawing1]);

			ExecuteHotkey(testHotkeyActions.selectNextObject);
			CheckSelectedObjects([drawing2]);

			ExecuteHotkey(testHotkeyActions.selectNextObject);
			CheckSelectedObjects([drawing3]);

			ExecuteHotkey(testHotkeyActions.selectPreviousObject);
			CheckSelectedObjects([drawing2]);

			ExecuteHotkey(testHotkeyActions.selectPreviousObject);
			CheckSelectedObjects([drawing1]);

			ExecuteHotkey(testHotkeyActions.selectPreviousObject);
			CheckSelectedObjects([drawing3]);
			TurnOffRecalculate();

			ExecuteHotkey(testHotkeyActions.createTextBoxContent);
			assert.true(!!drawing3.GraphicObj.textBoxContent, 'Check create textBoxContent after enter');

			drawing2.GraphicObj.setWordShape(false);
			SelectDrawings([drawing2]);

			ExecuteHotkey(testHotkeyActions.createTextBody);
			assert.true(!!drawing2.GraphicObj.txBody, 'Check create textBody after enter');

			SelectDrawings([drawing2]);
			ExecuteHotkey(testHotkeyActions.moveCursorToStartPositionShapeEnter);
			const para = drawing2.GraphicObj.getDocContent().Content[0];
			assert.true(para.IsThisElementCurrent() && para.IsCursorAtBegin(), 'Check movement to start position in empty content');

			AscTest.EnterText('Hello');
			SelectDrawings([drawing2]);

			ExecuteHotkey(testHotkeyActions.selectAllShapeEnter);
			assert.strictEqual(logicDocument.GetSelectedText(), 'Hello', 'Check select non empty content');

			SelectDrawings([drawing1, drawing2]);

			const group = GetDrawingObjects().groupSelectedObjects();
			group.GraphicObj.selectObject(drawing1.GraphicObj, 0);
			GetDrawingObjects().selection.groupSelection = group.GraphicObj;

			ExecuteHotkey(testHotkeyActions.resetShapeSelection);
			assert.strictEqual(GetDrawingObjects().selectedObjects.length, 0, 'Check reset shape selection');

			SelectDrawings([group]);
			ExecuteHotkey(testHotkeyActions.removeShape, 0);
			assert.strictEqual(paragraph.GetRunByElement(group), null, 'Check remove shape');

			SelectDrawings([drawing3]);

			ExecuteHotkey(testHotkeyActions.removeShape, 1);
			assert.strictEqual(paragraph.GetRunByElement(drawing3), null, 'Check remove shape');

		});


		QUnit.test('Check actions with headers/footers', (assert) =>
		{
			TurnOnRecalculate();
			TurnOnRecalculateCurPos();
			ClearDocumentAndAddParagraph("Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World");
			AscTest.Recalculate();

			GoToHeaderFooter(2);
			GoToHeaderFooter(2, true);

			ExecuteHotkey(testHotkeyActions.moveToPreviousHeaderFooter);
			assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check current page');
			assert.true(!!logicDocument.Controller.HdrFtr.Pages[1].Footer, 'Footer exists');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[1].Footer, 'Check move to previous footer');
			ExecuteHotkey(testHotkeyActions.moveToPreviousHeaderFooter);
			assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check current page');
			assert.true(!!logicDocument.Controller.HdrFtr.Pages[1].Header, 'Header exists');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[1].Header, 'Check move to previous header');

			ExecuteHotkey(testHotkeyActions.moveToNextHeaderFooter);
			assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check current page');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[1].Footer, 'Check move to next footer');
			ExecuteHotkey(testHotkeyActions.moveToNextHeaderFooter);
			assert.strictEqual(logicDocument.Get_CurPage(), 2, 'Check current page');
			assert.true(!!logicDocument.Controller.HdrFtr.Pages[2].Header, 'Header exists');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[2].Header, 'Check move to next header');

			ExecuteHotkey(testHotkeyActions.moveToPreviousHeader, 0);
			assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check current page');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[1].Header, 'Check move to previous header');
			ExecuteHotkey(testHotkeyActions.moveToPreviousHeader, 1);
			assert.strictEqual(logicDocument.Get_CurPage(), 0, 'Check current page');
			assert.true(!!logicDocument.Controller.HdrFtr.Pages[0].Header, 'Header exists');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[0].Header, 'Check move to previous header');

			ExecuteHotkey(testHotkeyActions.moveToNextHeader, 0);
			assert.strictEqual(logicDocument.Get_CurPage(), 1, 'Check current page');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[1].Header, 'Check move to next header');
			ExecuteHotkey(testHotkeyActions.moveToNextHeader, 1);
			assert.strictEqual(logicDocument.Get_CurPage(), 2, 'Check current page');
			assert.true(logicDocument.Controller.HdrFtr.CurHdrFtr === logicDocument.Controller.HdrFtr.Pages[2].Header, 'Check move to next header');

			ExecuteHotkey(testHotkeyActions.endEditing);
			assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Check end editing footer");

			GoToHeaderFooter(0);
			ExecuteHotkey(testHotkeyActions.endEditing);
			assert.strictEqual(logicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Check end editing footer");

			RemoveHeader(0);
			RemoveFooter(0);
			TurnOffRecalculateCurPos();
			TurnOffRecalculate();
		});

		QUnit.test('Check reset actions shortcut', (assert) =>
		{
			TurnOnRecalculate();
			ClearDocumentAndAddParagraph("");
			AscTest.Recalculate()
			editor.StartAddShape('rect');
			ExecuteHotkey(testHotkeyActions.resetStartAddShape);
			assert.strictEqual(editor.isStartAddShape, false, "Check reset add shape");
			TurnOffRecalculate();
			editor.SetPaintFormat(AscCommon.c_oAscFormatPainterState.kOn);
			ExecuteHotkey(testHotkeyActions.resetFormattingByExample);
			assert.strictEqual(editor.isFormatPainterOn(), false, "Check reset formatting by example");

			editor.SetMarkerFormat(true, true, 0, 0, 0);
			ExecuteHotkey(testHotkeyActions.resetMarkerFormat);
			assert.strictEqual(editor.isMarkerFormat, false, "Check reset marker");
		});

		QUnit.test('Check disable shortcuts', (assert) =>
		{
			assert.strictEqual(ExecuteHotkey(testHotkeyActions.disableNumLock) & keydownresult_PreventAll, keydownresult_PreventAll);
			assert.strictEqual(ExecuteHotkey(testHotkeyActions.disableScrollLock) & keydownresult_PreventAll, keydownresult_PreventAll);
			assert.strictEqual(ExecuteHotkey(testHotkeyActions.disableBrowserZoomIn) & keydownresult_PreventAll, keydownresult_PreventAll);
		});

		QUnit.test('Check filling forms', (assert) =>
		{
			ClearDocumentAndAddParagraph('');

			const checkBox = AddCheckBox();
			AscTest.SetFillingFormMode(true);
			ExecuteHotkey(testHotkeyActions.toggleCheckBox);
			assert.true(checkBox.IsCheckBoxChecked(), 'Check turn on checkbox');

			ExecuteHotkey(testHotkeyActions.toggleCheckBox);
			assert.false(checkBox.IsCheckBoxChecked(), 'Check turn off checkbox');
			AscTest.SetEditingMode();

			ClearDocumentAndAddParagraph('');
			AddComboBox(['Hello', 'World', 'yo']);
			AscTest.SetFillingFormMode(true);
			ExecuteHotkey(testHotkeyActions.nextOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'Hello', 'Check select next option in combobox');

			ExecuteHotkey(testHotkeyActions.nextOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'World', 'Check select next option in combobox');

			ExecuteHotkey(testHotkeyActions.nextOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'yo', 'Check select next option in combobox');

			ExecuteHotkey(testHotkeyActions.previousOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'World', 'Check select previous option in combobox');

			ExecuteHotkey(testHotkeyActions.previousOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'Hello', 'Check select previous option in combobox');

			ExecuteHotkey(testHotkeyActions.previousOptionComboBox);
			assert.strictEqual(logicDocument.GetSelectedText(), 'yo', 'Check select previous option in combobox');
			AscTest.SetEditingMode();
		});



		QUnit.test('Check movement selecting forms', (assert) =>
		{
			ClearDocumentAndAddParagraph('');
			let checkBox1 = AddCheckBox();
			AscTest.MoveCursorRight();
			let checkBox2 = AddCheckBox();
			AscTest.MoveCursorRight();
			let checkBox3 = AddCheckBox();
			AscTest.SetFillingFormMode(true);


			ExecuteHotkey(testHotkeyActions.moveToNextForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox1, 'Check move to next form');

			ExecuteHotkey(testHotkeyActions.moveToNextForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox2, 'Check move to next form');

			ExecuteHotkey(testHotkeyActions.moveToNextForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox3, 'Check move to next form');

			ExecuteHotkey(testHotkeyActions.moveToPreviousForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox2, 'Check move to previous form');
			ExecuteHotkey(testHotkeyActions.moveToPreviousForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox1, 'Check move to previous form');
			ExecuteHotkey(testHotkeyActions.moveToPreviousForm);
			assert.true(logicDocument.GetSelectedElementsInfo().GetInlineLevelSdt() === checkBox3, 'Check move to previous form');

			AscTest.SetEditingMode();
		});

		QUnit.test('Check movement in table', (assert) =>
		{
			ClearDocumentAndAddParagraph();
			const table = AddTable(3, 4);
			table.Document_SetThisElementCurrent();
			table.MoveCursorToStartPos();
			ExecuteHotkey(testHotkeyActions.moveToNextCell);
			assert.strictEqual(table.CurCell.Index, 1, 'Check move to next cell');
			ExecuteHotkey(testHotkeyActions.moveToNextCell);
			assert.strictEqual(table.CurCell.Index, 2, 'Check move to next cell');
			ExecuteHotkey(testHotkeyActions.moveToNextCell);
			assert.strictEqual(table.CurCell.Index, 3, 'Check move to next cell');

			ExecuteHotkey(testHotkeyActions.moveToPreviousCell);
			assert.strictEqual(table.CurCell.Index, 2, 'Check move to previous cell');
			ExecuteHotkey(testHotkeyActions.moveToPreviousCell);
			assert.strictEqual(table.CurCell.Index, 1, 'Check move to previous cell');
			ExecuteHotkey(testHotkeyActions.moveToPreviousCell);
			assert.strictEqual(table.CurCell.Index, 0, 'Check move to previous cell');
		});

		QUnit.test('Check select all in chart title', (assert) =>
		{
			TurnOnRecalculate();
			ClearDocumentAndAddParagraph('');
			AscTest.Recalculate();
			TurnOffRecalculate();
			const paraDrawing = AddChart();

			const chart = paraDrawing.GraphicObj;
			SelectDrawings([paraDrawing]);
			const titles = chart.getAllTitles();
			const controller = GetDrawingObjects();
			controller.selection.chartSelection = chart;
			chart.selectTitle(titles[0], 0);

			ExecuteHotkey(testHotkeyActions.selectAllInChartTitle);
			assert.strictEqual(logicDocument.GetSelectedText(), 'Diagram Title', 'Check select all in title');
		});

		QUnit.test('Check add new paragraph in content', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello text');
			ExecuteHotkey(testHotkeyActions.addNewParagraphContent);
			assert.strictEqual(logicDocument.Content.length, 2, 'Check add new paragraph');
		});

		QUnit.test('Check add new paragraph math', (assert) =>
		{
			ClearDocumentAndAddParagraph('Hello text');
			logicDocument.AddParaMath();
			AscTest.EnterText('abcd');
			AscTest.MoveCursorLeft();
			ExecuteHotkey(testHotkeyActions.addNewParagraphMath)
			assert.strictEqual(logicDocument.Content.length, 2, 'Check add new paragraph');
		});

		QUnit.test("Test add new line to math", (oAssert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			logicDocument.AddParaMath(c_oAscMathType.FractionVertical);
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			AscTest.EnterText('Hello');
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			ExecuteHotkey(testHotkeyActions.addNewLineToMath);
			const paraMath = paragraph.GetAllParaMaths()[0];
			const fraction = paraMath.Root.GetFirstElement();
			const numerator = fraction.getNumerator();
			const eqArray = numerator.GetFirstElement();
			oAssert.strictEqual(eqArray.getRowsCount(), 2, 'Check add new line');
		});

		QUnit.test("Check remove form", (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			let form = AddComboBox(['hdfh']);
			ExecuteHotkey(testHotkeyActions.removeForm, 0);
			assert.strictEqual(paragraph.GetPosByElement(form), null, 'Check remove form');

			AddComboBox(['hdfh']);
			ExecuteHotkey(testHotkeyActions.removeForm, 1);
			assert.strictEqual(paragraph.GetPosByElement(form), null, 'Check remove form');
			AscTest.SetEditingMode();
		});

		QUnit.test("Check add tab to paragraph", (assert) =>
		{
			const paragraph = ClearDocumentAndAddParagraph('');
			ExecuteHotkey(testHotkeyActions.addTabToParagraph);
			assert.true(paragraph.GetPrevRunElement().IsTab(), "Check add tab to paragraph");
		});


		QUnit.test("Check add break line to inlinelvlsdt", (assert) =>
		{
			TurnOnRecalculate();
			ClearDocumentAndAddParagraph('');
			const complexForm = AddComplexForm();
			ExecuteHotkey(testHotkeyActions.addBreakLineInlineLvlSdt);
			assert.strictEqual(complexForm.Lines[0], 2, "Check add break line");
			TurnOffRecalculate();
		});

		QUnit.test("Check visit hyperlink", (assert) =>
		{
			TurnOnRecalculate()
			ClearDocumentAndAddParagraph('');
			AscTest.Recalculate()
			TurnOffRecalculate();
			logicDocument.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page))
			logicDocument.AddHyperlink(new Asc.CHyperlinkProperty({Anchor: '_top', Text: "Beginning of document"}));
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			ExecuteHotkey(testHotkeyActions.visitHyperlink);
			assert.true(logicDocument.GetCurrentParagraph() === logicDocument.Content[0], "Check visit hyperlink");
			assert.strictEqual(logicDocument.Get_CurPage(), 0, "Check visit hyperlink");
		});

		QUnit.test("Check handle tab in math", (assert) =>
		{
			ClearDocumentAndAddParagraph('');
			logicDocument.AddParaMath();
			AscTest.EnterText('abcd+abcd+abcd');
			logicDocument.MoveCursorToEndPos();
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();
			AscTest.MoveCursorLeft();

			const props = new CMathMenuBase();
			props.insert_ManualBreak();
			logicDocument.Set_MathProps(props);
			ExecuteHotkey(testHotkeyActions.handleTab);
			AscTest.MoveCursorRight();
			const contentPosition = logicDocument.GetContentPosition();
			const currentRun = contentPosition[contentPosition.length - 1].Class;
			assert.strictEqual(currentRun.MathPrp.Get_AlnAt(), 1, "Check handle tab in math");
		});


		QUnit.test("Check end editing form", (assert) =>
		{
			logicDocument.End_SilentMode();
			ClearDocumentAndAddParagraph('');
			const checkBox = AddCheckBox();
			AscTest.SetFillingFormMode(true);
			checkBox.MoveCursorToContentControl(true);
			ExecuteHotkey(testHotkeyActions.endEditing);
			const selectedInfo = logicDocument.GetSelectedElementsInfo();
			assert.strictEqual(!!selectedInfo.GetInlineLevelSdt(), false, "Check end editing form");
			AscTest.SetEditingMode();
			logicDocument.Start_SilentMode();
		});

		QUnit.test("Check replace unicode to char hotkeys", (assert) =>
		{
			ClearDocumentAndAddParagraph('2601');
			AscTest.MoveCursorLeft(true, true);
			ExecuteHotkey(testHotkeyActions.unicodeToChar, 0);
			assert.strictEqual(logicDocument.GetSelectedText(), '☁', 'Check replace unicode code to symbol');
			AscTest.MoveCursorRight();
			AscTest.EnterText(' 261d');

			AscTest.MoveCursorLeft(true, true);

			ExecuteHotkey(testHotkeyActions.unicodeToChar, 1);
			assert.strictEqual(logicDocument.GetSelectedText(), '☝', 'Check replace unicode code to symbol');
		});

		QUnit.test("Check reset drag'n'drop", (oAssert) =>
		{
			TurnOnRecalculate();
			ClearDocumentAndAddParagraph('Hello Hello');
			AscTest.Recalculate();
			logicDocument.MoveCursorToStartPos();
			AscTest.MoveCursorRight(true, true);

			let e = new AscCommon.CMouseEventHandler();
			e.Button = AscCommon.g_mouse_button_left;
			e.ClickCount = 1;

			e.Type = AscCommon.g_mouse_event_type_down;
			logicDocument.OnMouseDown(e, 5, 10, 0);

			e.Type = AscCommon.g_mouse_event_type_move;
			logicDocument.OnMouseMove(e, 45, 10, 0);

			ExecuteHotkey(testHotkeyActions.resetDragNDrop);
			oAssert.true(!drawingDocument.IsTrackText(), "Check reset drag'n'drop");

			e.Type = AscCommon.g_mouse_event_type_up;
			logicDocument.OnMouseUp(e, 45, 10, 0);

			TurnOffRecalculate();
		});
	});
})(window);
