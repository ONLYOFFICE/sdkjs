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
	const {
		addPropertyToDocument,
		getLogicDocumentWithParagraphs,
		checkTextAfterKeyDownHelperEmpty,
		checkDirectTextPrAfterKeyDown,
		checkDirectParaPrAfterKeyDown,
		oGlobalLogicDocument,
		addParagraphToDocumentWithText,
		remove,
		clean,
		recalculate,
		onKeyDown,
		moveToParagraph,
		createNativeEvent,
		moveCursorDown,
		moveCursorLeft,
		moveCursorRight,
		selectAll,
		getSelectedText,
		executeTestWithCatchEvent,
		startTest,
		createChart,
		createEvent,
		checkInsertElementByType,
		createParagraphWithText,
		checkApplyParagraphStyle,
		createHyperlink,
		createTable,
		moveToTable,
		createShape,
		createGroup,
		selectOnlyObjects,
		selectParaDrawing,
		drawingObjects,
		logicContent,
		directParaPr,
		directTextPr,
		addBreakPage,
		createMath,
		resetLogicDocument,
		addText,
		moveShapeHelper,
		setFillingFormsMode,
		createCheckBox,
		createComboBox,
		createComplexForm,
		contentPosition,
		oTestTypes,
		mouseMove,
		mouseDown,
		mouseUp,
		insertManualBreak
	} = AscTestShortcut;

	$(function ()
	{
		let fOldGetShortcut;
		QUnit.module("Test shortcut actions", {
			before    : function ()
			{
				editor.initDefaultShortcuts();
			},
			beforeEach: function ()
			{
				fOldGetShortcut = editor.getShortcut;
			},
			afterEach : function ()
			{
				editor.getShortcut = fOldGetShortcut;
			},
			after     : function ()
			{
				editor.Shortcuts = new AscCommon.CShortcuts();
			}
		});

		QUnit.test('Check page break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertPageBreak};
			let oEvent = createNativeEvent();
			const {oLogicDocument} = getLogicDocumentWithParagraphs([''], true);
			onKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 2, 'Check page break shortcut');
			onKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 3, 'Check page break shortcut');
			onKeyDown(oEvent);
			oAssert.strictEqual(oLogicDocument.GetPagesCount(), 4, 'Check page break shortcut');
		});

		QUnit.test('Check line break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertLineBreak};
			const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs([''], true);
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 2, 'Check line break shortcut');
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 3, 'Check line break shortcut');
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(oParagraph.GetLinesCount(), 4, 'Check line break shortcut');
		});

		QUnit.test('Check column break shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertColumnBreak};
			let oColumnProps = new Asc.CDocumentColumnsProps();
			oColumnProps.put_Num(2);
			editor.asc_SetColumnsProps(oColumnProps);
			const {oLogicDocument} = getLogicDocumentWithParagraphs([''], true);

			onKeyDown(createNativeEvent());
			const oParagraph = oLogicDocument.GetCurrentParagraph();

			oAssert.strictEqual(oParagraph.Get_CurrentColumn(), 1, 'Check column break shortcut');
			oColumnProps = new Asc.CDocumentColumnsProps();
			oColumnProps.put_Num(1);
			editor.asc_SetColumnsProps(oColumnProps);
		});

		QUnit.test('Check reset char shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ResetChar};

			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello world']);
			oLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});
			onKeyDown(createNativeEvent());
			const oDirectTextPr = directTextPr();
			oAssert.true(!(oDirectTextPr.Get_Bold() || oDirectTextPr.Get_Italic() || oDirectTextPr.Get_Underline()), 'Check reset char shortcut');
		});

		QUnit.test('Check add non breaking space shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.NonBreakingSpace};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A0), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check add strikeout shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Strikeout};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Strikeout(), true, 'Check add strikeout shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Strikeout(), false, 'Check add strikeout shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check show non printing characters shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ShowAll};
			editor.put_ShowParaMarks(false);
			onKeyDown(createNativeEvent());
			oAssert.true(editor.get_ShowParaMarks(), 'Check show non printing characters shortcut');
		});

		QUnit.test('Check select all shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditSelectAll};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(oLogicDocument.GetSelectedText(), 'Hello World', 'Check select all shortcut');
		});

		QUnit.test('Check bold shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Bold};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Bold(), true, 'Check bold shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Bold(), false, 'Check bold shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check copy format shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CopyFormat};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true, Underline: true});

			onKeyDown(createNativeEvent());
			const oCopyParagraphTextPr = new AscCommonWord.CTextPr();
			oCopyParagraphTextPr.SetUnderline(true);
			oCopyParagraphTextPr.SetBold(true);
			oCopyParagraphTextPr.BoldCS = true;
			oCopyParagraphTextPr.SetItalic(true);
			oCopyParagraphTextPr.ItalicCS = true;
			oAssert.deepEqual(editor.getFormatPainterData().TextPr, oCopyParagraphTextPr, 'Check copy format shortcut');
		});

		QUnit.test('Check insert copyright shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CopyrightSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00A9), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check insert endnote shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertEndnoteNow};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());
			const arrEndnotes = oLogicDocument.GetEndnotesList();
			oAssert.deepEqual(arrEndnotes.length, 1, 'Check insert endnote shortcut');
		});

		QUnit.test('Check center para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.CenterPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Center, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check insert euro sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EuroSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x20AC), createNativeEvent(), oAssert, 'Check add non breaking space shortcut');
		});

		QUnit.test('Check italic shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Italic};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Italic(), true, 'Check add italic shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Italic(), false, 'Check add italic shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check justify para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.JustifyPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Justify, 'Check justify para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check justify para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check bullet list shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyListBullet};
			const {oLogicDocument, oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());

			oAssert.true(oParagraph.IsBulletedNumbering(), 'check apply bullet list');
		});

		QUnit.test('Check left para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.LeftPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Justify, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check indent shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Indent};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oParagraph.Pr.SetInd(0, 0, 0);
			moveToParagraph(oParagraph, true);

			onKeyDown(createNativeEvent());
			oAssert.strictEqual(directParaPr().GetIndLeft(), 12.5, 'Check increase indent');
		});

		QUnit.test('Check unindent shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.UnIndent};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello']);
			oParagraph.Pr.SetInd(0, 12.5, 0);
			moveToParagraph(oParagraph, true);

			onKeyDown(createNativeEvent());
			oAssert.true(AscFormat.fApproxEqual(directParaPr().GetIndLeft(), 0), 'Check decrease indent');
		});

		QUnit.test('Check insert page number shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertPageNumber};
			checkInsertElementByType(para_PageNum, 'Check insert page number shortcut', oAssert, createNativeEvent());
		});

		QUnit.test('Check right para shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.RightPara};
			const fAnotherCheck = checkDirectParaPrAfterKeyDown((oParaPr) => oParaPr.Get_Jc(), align_Right, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oParaPr) => oParaPr.Get_Jc(), align_Left, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check registered sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.RegisteredSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x00AE), createNativeEvent(), oAssert, 'Check registered sign shortcut');
		});

		QUnit.test('Check trademark sign shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.TrademarkSign};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2122), createNativeEvent(), oAssert, 'Check registered sign shortcut');
		});

		QUnit.test('Check underline shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Underline};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_Underline(), true, 'Check underline shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_Underline(), false, 'Check underline shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check paste format shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.PasteFormat};
			const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
			oParagraph.SetThisElementCurrent();
			oGlobalLogicDocument.SelectAll();
			addPropertyToDocument({Bold: true, Italic: true});
			oGlobalLogicDocument.Document_Format_Copy();
			remove();
			addParagraphToDocumentWithText('Hello');
			oGlobalLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());
			const oDirectTextPr = directTextPr();
			oAssert.true(oDirectTextPr.Get_Bold() && oDirectTextPr.Get_Italic(), 'Check paste format shortcut');
		});

		QUnit.test('Check redo shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditRedo};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello World']);
			oLogicDocument.SelectAll();
			oLogicDocument.Remove(undefined, undefined, true);
			oLogicDocument.Document_Undo();
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(AscTest.GetParagraphText(logicContent()[0]), '', 'Check redo shortcut');
		});

		QUnit.test('Check undo shortcut', (oAssert) =>
		{
			getLogicDocumentWithParagraphs(['Hello World']);
			selectAll();
			editor.asc_Remove();

			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EditUndo};
			onKeyDown(createNativeEvent());
			selectAll();
			oAssert.strictEqual(getSelectedText(), 'Hello World', 'Check redo shortcut');
		});

		QUnit.test('Check en dash shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EnDash};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2013), createNativeEvent(), oAssert, 'Check en dash shortcut');
		});

		QUnit.test('Check em dash shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.EmDash};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2014), createNativeEvent(), oAssert, 'Check em dash shortcut');
		});

		QUnit.test('Check update fields shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.UpdateFields};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello', 'Hello', 'Hello'], true);
			for (let i = 0; i < logicContent().length; i += 1)
			{
				oLogicDocument.Set_CurrentElement(i, true);
				oLogicDocument.SetParagraphStyle("Heading 1");
			}
			oLogicDocument.MoveCursorToStartPos();
			const props = new Asc.CTableOfContentsPr();
			props.put_OutlineRange(1, 9);
			props.put_Hyperlink(true);
			props.put_ShowPageNumbers(true);
			props.put_RightAlignTab(true);
			props.put_TabLeader(Asc.c_oAscTabLeader.Dot);
			editor.asc_AddTableOfContents(null, props);

			oLogicDocument.MoveCursorToEndPos();
			const oParagraph = createParagraphWithText('Hello');
			oLogicDocument.AddToContent(logicContent().length, oParagraph);
			moveToParagraph(oParagraph);
			oLogicDocument.SetParagraphStyle("Heading 1");

			logicContent()[0].SetThisElementCurrent();
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(logicContent()[0].Content.Content.length, 5, 'Check update fields shortcut');
		});

		QUnit.test('Check superscript shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Superscript};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_SuperScript, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_Baseline, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check non breaking hyphen shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.NonBreakingHyphen};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x002D), createNativeEvent(), oAssert, 'Check non breaking hyphen shortcut');
		});

		QUnit.test('Check horizontal ellipsis shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.HorizontalEllipsis};
			checkTextAfterKeyDownHelperEmpty(String.fromCharCode(0x2026), createNativeEvent(), oAssert, 'Check add horizontal ellipsis shortcut');
		});

		QUnit.test('Check subscript shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Subscript};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_SubScript, 'Check center para shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_VertAlign(), AscCommon.vertalign_Baseline, 'Check center para shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check show hyperlink menu shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertHyperlink};
			executeTestWithCatchEvent('asc_onDialogAddHyperlink', () => true, true, createNativeEvent(), oAssert, () =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
				moveToParagraph(oParagraph);
				oGlobalLogicDocument.SelectAll();
			});
		});

		QUnit.test('Check print shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.PrintPreviewAndPrint};
			executeTestWithCatchEvent('asc_onPrint', () => true, true, createNativeEvent(), oAssert);
		});

		QUnit.test('Check save shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.Save};
			const fOldSave = editor._onSaveCallbackInner;
			let bCheck = false;
			editor._onSaveCallbackInner = function ()
			{
				bCheck = true;
				editor.canSave = true;
			};
			onKeyDown(createNativeEvent());
			oAssert.strictEqual(bCheck, true, 'Check save shortcut');
			editor._onSaveCallbackInner = fOldSave;
		});

		QUnit.test('Check increase font size shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.IncreaseFontSize};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), 11, 'Check increase font size shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_FontSize(), 12, 'Check increase font size shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check decrease font size shortcut', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.DecreaseFontSize};
			const fAnotherCheck = checkDirectTextPrAfterKeyDown((oTextPr) => oTextPr.Get_FontSize(), 9, 'Check decrease font size shortcut', createNativeEvent(), oAssert);
			fAnotherCheck((oTextPr) => oTextPr.Get_FontSize(), 8, 'Check decrease font size shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check apply heading 1', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading1};
			checkApplyParagraphStyle('Heading 1', 'Check apply heading 1 shortcut', createNativeEvent(), oAssert);
		});
		QUnit.test('Check apply heading 2', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading2};
			checkApplyParagraphStyle('Heading 2', 'Check apply heading 2 shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check apply heading 3', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.ApplyHeading3};
			checkApplyParagraphStyle('Heading 3', 'Check apply heading 3 shortcut', createNativeEvent(), oAssert);
		});

		QUnit.test('Check insert footnotes now', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertFootnoteNow};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['Hello']);
			oLogicDocument.SelectAll();
			onKeyDown(createNativeEvent());
			const arrFootnotes = oLogicDocument.GetFootnotesList();
			oAssert.deepEqual(arrFootnotes.length, 1, 'Check insert footnote shortcut');
		});

		QUnit.test('Check insert equation', (oAssert) =>
		{
			editor.getShortcut = function () {return c_oAscDocumentShortcutType.InsertEquation};
			const {oLogicDocument} = getLogicDocumentWithParagraphs(['']);
			onKeyDown(createNativeEvent());
			const oMath = oLogicDocument.GetCurrentMath();
			oAssert.true(!!oMath, 'Check insert equation shortcut');
		});

		QUnit.module("Test getting desired action by event")
		QUnit.test("Test getting common desired action by event", (oAssert) =>
		{
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(13, true, false, false, false, false, false)), c_oAscDocumentShortcutType.InsertPageBreak, 'Check getting c_oAscDocumentShortcutType.InsertPageBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(13, false, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertLineBreak, 'Check getting c_oAscDocumentShortcutType.InsertLineBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(13, true, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertColumnBreak, 'Check getting c_oAscDocumentShortcutType.InsertColumnBreak action');
			oAssert.strictEqual(editor.getShortcut(createEvent(32, true, false, false, false, false, false)), c_oAscDocumentShortcutType.ResetChar, 'Check getting c_oAscDocumentShortcutType.ResetChar action');
			oAssert.strictEqual(editor.getShortcut(createEvent(32, true, true, false, false, false, false)), c_oAscDocumentShortcutType.NonBreakingSpace, 'Check getting c_oAscDocumentShortcutType.NonBreakingSpace action');
			oAssert.strictEqual(editor.getShortcut(createEvent(56, true, true, false, false, false, false)), c_oAscDocumentShortcutType.ShowAll, 'Check getting c_oAscDocumentShortcutType.ShowAll action');
			oAssert.strictEqual(editor.getShortcut(createEvent(65, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditSelectAll, 'Check getting c_oAscDocumentShortcutType.EditSelectAll action');
			oAssert.strictEqual(editor.getShortcut(createEvent(66, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Bold, 'Check getting c_oAscDocumentShortcutType.Bold action');
			oAssert.strictEqual(editor.getShortcut(createEvent(67, true, false, true, false, false, false)), c_oAscDocumentShortcutType.CopyFormat, 'Check getting c_oAscDocumentShortcutType.CopyFormat action');
			oAssert.strictEqual(editor.getShortcut(createEvent(71, true, false, true, false, false, false)), c_oAscDocumentShortcutType.CopyrightSign, 'Check getting c_oAscDocumentShortcutType.CopyrightSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(68, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEndnoteNow, 'Check getting c_oAscDocumentShortcutType.InsertEndnoteNow action');
			oAssert.strictEqual(editor.getShortcut(createEvent(69, true, false, false, false, false, false)), c_oAscDocumentShortcutType.CenterPara, 'Check getting c_oAscDocumentShortcutType.CenterPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(69, true, false, true, false, false, false)), c_oAscDocumentShortcutType.EuroSign, 'Check getting c_oAscDocumentShortcutType.EuroSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(73, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Italic, 'Check getting c_oAscDocumentShortcutType.Italic action');
			oAssert.strictEqual(editor.getShortcut(createEvent(74, true, false, false, false, false, false)), c_oAscDocumentShortcutType.JustifyPara, 'Check getting c_oAscDocumentShortcutType.JustifyPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(75, true, false, false, false, false, false)), c_oAscDocumentShortcutType.InsertHyperlink, 'Check getting c_oAscDocumentShortcutType.InsertHyperlink action');
			oAssert.strictEqual(editor.getShortcut(createEvent(76, true, true, false, false, false, false)), c_oAscDocumentShortcutType.ApplyListBullet, 'Check getting c_oAscDocumentShortcutType.ApplyListBullet action');
			oAssert.strictEqual(editor.getShortcut(createEvent(76, true, false, false, false, false, false)), c_oAscDocumentShortcutType.LeftPara, 'Check getting c_oAscDocumentShortcutType.LeftPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(77, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Indent, 'Check getting c_oAscDocumentShortcutType.Indent action');
			oAssert.strictEqual(editor.getShortcut(createEvent(77, true, true, false, false, false, false)), c_oAscDocumentShortcutType.UnIndent, 'Check getting c_oAscDocumentShortcutType.UnIndent action');
			oAssert.strictEqual(editor.getShortcut(createEvent(80, true, false, false, false, false, false)), c_oAscDocumentShortcutType.PrintPreviewAndPrint, 'Check getting c_oAscDocumentShortcutType.PrintPreviewAndPrint action');
			oAssert.strictEqual(editor.getShortcut(createEvent(80, true, true, false, false, false, false)), c_oAscDocumentShortcutType.InsertPageNumber, 'Check getting c_oAscDocumentShortcutType.InsertPageNumber action');
			oAssert.strictEqual(editor.getShortcut(createEvent(82, true, false, false, false, false, false)), c_oAscDocumentShortcutType.RightPara, 'Check getting c_oAscDocumentShortcutType.RightPara action');
			oAssert.strictEqual(editor.getShortcut(createEvent(82, true, false, true, false, false, false)), c_oAscDocumentShortcutType.RegisteredSign, 'Check getting c_oAscDocumentShortcutType.RegisteredSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(83, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Save, 'Check getting c_oAscDocumentShortcutType.Save action');
			oAssert.strictEqual(editor.getShortcut(createEvent(84, true, false, true, false, false, false)), c_oAscDocumentShortcutType.TrademarkSign, 'Check getting c_oAscDocumentShortcutType.TrademarkSign action');
			oAssert.strictEqual(editor.getShortcut(createEvent(85, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Underline, 'Check getting c_oAscDocumentShortcutType.Underline action');
			oAssert.strictEqual(editor.getShortcut(createEvent(86, true, false, true, false, false, false)), c_oAscDocumentShortcutType.PasteFormat, 'Check getting c_oAscDocumentShortcutType.PasteFormat action');
			oAssert.strictEqual(editor.getShortcut(createEvent(89, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditRedo, 'Check getting c_oAscDocumentShortcutType.EditRedo action');
			oAssert.strictEqual(editor.getShortcut(createEvent(90, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EditUndo, 'Check getting c_oAscDocumentShortcutType.EditUndo action');
			oAssert.strictEqual(editor.getShortcut(createEvent(109, true, false, false, false, false, false)), c_oAscDocumentShortcutType.EnDash, 'Check getting c_oAscDocumentShortcutType.EnDash action');
			oAssert.strictEqual(editor.getShortcut(createEvent(109, true, false, true, false, false, false)), c_oAscDocumentShortcutType.EmDash, 'Check getting c_oAscDocumentShortcutType.EmDash action');
			oAssert.strictEqual(editor.getShortcut(createEvent(120, false, false, false, false, false, false)), c_oAscDocumentShortcutType.UpdateFields, 'Check getting c_oAscDocumentShortcutType.UpdateFields action');
			oAssert.strictEqual(editor.getShortcut(createEvent(188, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Superscript, 'Check getting c_oAscDocumentShortcutType.Superscript action');
			oAssert.strictEqual(editor.getShortcut(createEvent(189, true, true, false, false, false, false)), c_oAscDocumentShortcutType.NonBreakingHyphen, 'Check getting c_oAscDocumentShortcutType.NonBreakingHyphen action');
			oAssert.strictEqual(editor.getShortcut(createEvent(190, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Subscript, 'Check getting c_oAscDocumentShortcutType.Subscript action');
			oAssert.strictEqual(editor.getShortcut(createEvent(219, true, false, false, false, false, false)), c_oAscDocumentShortcutType.DecreaseFontSize, 'Check getting c_oAscDocumentShortcutType.DecreaseFontSize action');
			oAssert.strictEqual(editor.getShortcut(createEvent(221, true, false, false, false, false, false)), c_oAscDocumentShortcutType.IncreaseFontSize, 'Check getting c_oAscDocumentShortcutType.IncreaseFontSize action');
			editor.Shortcuts = new AscCommon.CShortcuts();
		});

		QUnit.test("Test getting windows desired action by event", (oAssert) =>
		{
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(190, true, false, true, false, false, false)), c_oAscDocumentShortcutType.HorizontalEllipsis, 'Check getting c_oAscDocumentShortcutType.HorizontalEllipsis action');
			oAssert.strictEqual(editor.getShortcut(createEvent(49, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading1, 'Check getting c_oAscDocumentShortcutType.ApplyHeading1 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(50, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading2, 'Check getting c_oAscDocumentShortcutType.ApplyHeading2 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(51, false, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading3, 'Check getting c_oAscDocumentShortcutType.ApplyHeading3 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(70, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertFootnoteNow, 'Check getting c_oAscDocumentShortcutType.InsertFootnoteNow shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(187, false, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEquation, 'Check getting c_oAscDocumentShortcutType.InsertEquation shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(53, true, false, false, false, false, false)), c_oAscDocumentShortcutType.Strikeout, 'Check getting c_oAscDocumentShortcutType.Strikeout action');
			editor.Shortcuts = new AscCommon.CShortcuts();
		});

		QUnit.test("Test getting macOs desired action by event", (oAssert) =>
		{
			const bOldMacOs = AscCommon.AscBrowser.isMacOs;
			AscCommon.AscBrowser.isMacOs = true;
			editor.initDefaultShortcuts();
			oAssert.strictEqual(editor.getShortcut(createEvent(49, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading1, 'Check getting c_oAscDocumentShortcutType.ApplyHeading1 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(50, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading2, 'Check getting c_oAscDocumentShortcutType.ApplyHeading2 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(51, true, false, true, false, false, false)), c_oAscDocumentShortcutType.ApplyHeading3, 'Check getting c_oAscDocumentShortcutType.ApplyHeading3 shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(187, true, false, true, false, false, false)), c_oAscDocumentShortcutType.InsertEquation, 'Check getting c_oAscDocumentShortcutType.InsertEquation shortcut type');
			oAssert.strictEqual(editor.getShortcut(createEvent(88, true, true, false, false, false, false)), c_oAscDocumentShortcutType.Strikeout, 'Check getting c_oAscDocumentShortcutType.Strikeout action');
			editor.Shortcuts = new AscCommon.CShortcuts();
			AscCommon.AscBrowser.isMacOs = bOldMacOs;
		});

		QUnit.module('Test hotkeys module', {
			afterEach: function ()
			{
				resetLogicDocument(oGlobalLogicDocument);
			}
		})
		QUnit.test("test remove back symbol", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				selectAll();
				oAssert.strictEqual(getSelectedText(), 'Hello Worl', 'Test remove back symbol');
			}, oTestTypes.removeBackSymbol);
		});

		QUnit.test("test remove back word", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				selectAll();
				oAssert.strictEqual(getSelectedText(), 'Hello ', 'Test remove back word');
			}, oTestTypes.removeBackWord);
		});

		QUnit.test("test remove shape", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs([''], true);
				moveToParagraph(oParagraph);
				const oDrawing = createShape();
				selectParaDrawing(oDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(oParagraph.GetRunByElement(oDrawing), null, 'Test remove shape');
			}, oTestTypes.removeShape);
		});

		QUnit.test("test remove form", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['']);
				moveToParagraph(oParagraph);
				const oInlineLvlSdt = createComboBox();
				onKeyDown(oEvent);
				oAssert.strictEqual(oParagraph.GetPosByElement(oInlineLvlSdt), null, 'Test remove form');
			}, oTestTypes.removeForm);
		});


		QUnit.test("test move to next form", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs(['']);
				const oInlineSdt1 = createComboBox();
				moveCursorRight();
				const oInlineSdt2 = createComboBox();
				moveCursorRight();
				const oInlineSdt3 = createComboBox();
				setFillingFormsMode(true);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt1, 'Test move to next form');

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt2, 'Test move to next form');

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt3, 'Test move to next form');

				setFillingFormsMode(false);
			}, oTestTypes.moveToNextForm);
		});


		QUnit.test("test move to previous form", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['']);
				const oInlineSdt1 = createComboBox();
				moveCursorRight();
				const oInlineSdt2 = createComboBox();
				moveCursorRight();
				const oInlineSdt3 = createComboBox();
				setFillingFormsMode(true);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt2, 'Test move to next form');

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt1, 'Test move to next form');

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetSelectedElementsInfo().GetInlineLevelSdt(), oInlineSdt3, 'Test move to next form');

				setFillingFormsMode(false);
			}, oTestTypes.moveToPreviousForm);
		});



		QUnit.test("Test handle tab in math", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs([''], true);
				createMath();
				addText('abcd+abcd+abcd');
				moveToParagraph(oParagraph);
				moveCursorLeft();
				moveCursorLeft();
				moveCursorLeft();
				moveCursorLeft();
				moveCursorLeft();
				insertManualBreak();
				onKeyDown(oEvent);
				moveCursorRight();
				const oContentPosition = oGlobalLogicDocument.GetContentPosition();
				const oCurRun = oContentPosition[oContentPosition.length - 1].Class;

				oAssert.strictEqual(oCurRun.MathPrp.Get_AlnAt(), 1, 'Test move to next form');
			}, oTestTypes.handleTab);
		});

		QUnit.test("test move to cell", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const oTable = createTable(3, 3);
				moveToTable(oTable, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(oTable.CurCell.Index, 1, 'Test move to next cell');
			}, oTestTypes.moveToNextCell);
		});

		QUnit.test("test move to cell", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const oTable = createTable(3, 3);
				moveToTable(oTable, true);
				moveCursorRight();
				onKeyDown(oEvent);
				oAssert.strictEqual(oTable.CurCell.Index, 0, 'Test move to previous cell');
			}, oTestTypes.moveToPreviousCell);
		});

		QUnit.test("test select next object", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oFirstParaDrawing = createShape();
				const oSecondParaDrawing = createShape();
				selectParaDrawing(oFirstParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(drawingObjects().selectedObjects.length === 1 && drawingObjects().selectedObjects[0] === oSecondParaDrawing.GraphicObj, true, 'Test select next object');

			}, oTestTypes.selectNextObject);
		});

		QUnit.test("test select previous object", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oFirstParaDrawing = createShape();
				const oSecondParaDrawing = createShape();
				const oThirdParaDrawing = createShape();
				selectParaDrawing(oFirstParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(drawingObjects().selectedObjects.length === 1 && drawingObjects().selectedObjects[0] === oThirdParaDrawing.GraphicObj, true, 'Test select previous object');
				onKeyDown(oEvent);
				oAssert.strictEqual(drawingObjects().selectedObjects.length === 1 && drawingObjects().selectedObjects[0] === oSecondParaDrawing.GraphicObj, true, 'Test select previous object');
			}, oTestTypes.selectPreviousObject);
		});

		QUnit.test("test indent", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs(['Hello world', "Hello world"]);
				const oFirstParagraph = logicContent()[0];
				const oSecondParagraph = logicContent()[1];
				selectAll();
				onKeyDown(oEvent);
				let arrSteps = [];
				moveToParagraph(oFirstParagraph);
				arrSteps.push(directParaPr().GetIndLeft());
				moveToParagraph(oSecondParagraph);
				arrSteps.push(directParaPr().GetIndLeft());
				oAssert.deepEqual(arrSteps, [12.5, 12.5], 'Test indent');

				moveToParagraph(oFirstParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual();
			}, oTestTypes.testIndent);
		});

		QUnit.test("test unindent", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs(['hello', 'hello']);
				const oFirstParagraph = logicContent()[0];
				const oSecondParagraph = logicContent()[1];
				oFirstParagraph.Set_Ind({Left: 12.5});
				oSecondParagraph.Set_Ind({Left: 12.5});
				selectAll();
				onKeyDown(oEvent);

				const arrSteps = [];
				moveToParagraph(oFirstParagraph);
				arrSteps.push(directParaPr().GetIndLeft());
				moveToParagraph(oSecondParagraph);
				arrSteps.push(directParaPr().GetIndLeft());

				oAssert.deepEqual(arrSteps, [0, 0], 'Test unindent');
			}, oTestTypes.testUnIndent);
		});


		QUnit.test("test add tab to paragraph", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello World']);
				moveToParagraph(oParagraph, true);
				moveCursorRight();
				onKeyDown(oEvent);
				selectAll();

				oAssert.strictEqual(getSelectedText(), 'H\tello World', 'Test indent');
			}, oTestTypes.addTabToParagraph);
		});

		QUnit.test("test visit hyperlink", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['']);
				addBreakPage();
				createHyperlink();
				moveCursorLeft();
				moveCursorLeft();
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetCurrentParagraph(), logicContent()[0]);
				oAssert.strictEqual(oGlobalLogicDocument.Get_CurPage(), 0);
			}, oTestTypes.visitHyperlink);
		});

		QUnit.test("Test add break line to inlinelvlsdt", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oInlineSdt = createComplexForm();
				setFillingFormsMode(true);
				onKeyDown(oEvent);
				oAssert.strictEqual(oInlineSdt.Lines[0], 2);
				setFillingFormsMode(false);
			}, oTestTypes.addBreakLineInlineLvlSdt);
		});

		QUnit.test("Test create textBoxContent", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const oParaDrawing = createShape();
				selectParaDrawing(oParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(!!oParaDrawing.GraphicObj.textBoxContent, true);
			}, oTestTypes.createTextBoxContent);
		});

		QUnit.test("Test create txBody", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const oParaDrawing = createShape();
				oParaDrawing.GraphicObj.setWordShape(false);
				selectParaDrawing(oParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(!!oParaDrawing.GraphicObj.txBody, true);
			}, oTestTypes.createTextBody);
		});

		QUnit.test("Test add new line to math", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['']);
				createMath(c_oAscMathType.FractionVertical);
				moveCursorLeft();
				moveCursorLeft();
				addText('Hello');
				moveCursorLeft();
				moveCursorLeft();
				onKeyDown(oEvent);
				const oParaMath = oParagraph.GetAllParaMaths()[0];
				const oFraction = oParaMath.Root.GetFirstElement();
				const oNumerator = oFraction.getNumerator();
				const oEqArray = oNumerator.GetFirstElement();
				oAssert.strictEqual(oEqArray.getRowsCount(), 2, 'Check add new line math');
			}, oTestTypes.addNewLineToMath);
		});

		QUnit.test("Test move cursor to start position shape", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oParaDrawing = createShape();
				const oShape = oParaDrawing.GraphicObj;
				oShape.createTextBoxContent();
				selectParaDrawing(oParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(oShape.getDocContent().IsCursorAtBegin(), true);
			}, oTestTypes.moveCursorToStartPositionShapeEnter);
		});

		QUnit.test("Test select all in shape", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true)
				const oParaDrawing = createShape();
				const oShape = oParaDrawing.GraphicObj;
				oShape.createTextBoxContent();
				moveToParagraph(oShape.getDocContent().Content[0]);
				addText('Hello');
				selectParaDrawing(oParaDrawing);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Hello');
			}, oTestTypes.selectAllShapeEnter);
		});

		QUnit.test("Test move cursor to start position chart title", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const oParaDrawing = createChart();
				const oChart = oParaDrawing.GraphicObj;
				const oTitles = oChart.getAllTitles();
				const oContent = AscFormat.CreateDocContentFromString('', drawingObjects().getDrawingDocument(), oTitles[0].txBody);
				oTitles[0].txBody.content = oContent;
				selectParaDrawing(oParaDrawing);

				const oController = drawingObjects();
				oController.selection.chartSelection = oChart;
				oChart.selectTitle(oTitles[0], 0);

				onKeyDown(oEvent);
				oAssert.true(oContent.IsCursorAtBegin(), 'Check move cursor to begin pos in title');
			}, oTestTypes.moveCursorToStartPositionTitleEnter);
		});

		QUnit.test("Test select all in chart title", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oParaDrawing = createChart();
				const oChart = oParaDrawing.GraphicObj;
				selectParaDrawing(oParaDrawing);
				const oTitles = oChart.getAllTitles();
				const oController = drawingObjects();
				oController.selection.chartSelection = oChart;
				oChart.selectTitle(oTitles[0], 0);

				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Diagram Title', 'Check select all title');
			}, oTestTypes.selectAllInChartTitle);
		});

		QUnit.test("Test add new paragraph", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello Text']);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(logicContent().length, 2, 'Test add new paragraph to content');
			}, oTestTypes.addNewParagraphContent);
		});

		QUnit.test("Test add new paragraph", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs(['']);
				createMath();
				addText('abcd');
				moveCursorLeft();
				onKeyDown(oEvent);
				oAssert.strictEqual(logicContent().length, 2, 'Test add new paragraph with math');
			}, oTestTypes.addNewParagraphMath);
		});

		QUnit.test("Test close all window popups", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				executeTestWithCatchEvent('asc_onMouseMoveStart', () => true, true, oEvent, oAssert);
				executeTestWithCatchEvent('asc_onMouseMove', () => true, true, oEvent, oAssert);
				executeTestWithCatchEvent('asc_onMouseMoveEnd', () => true, true, oEvent, oAssert);
			}, oTestTypes.closeAllWindowsPopups);
		});

		QUnit.test("Test reset shape selection", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const arrDrawings = [createShape(), createShape()];
				const oGroup = createGroup(arrDrawings);
				const oChart = createChart();
				selectOnlyObjects([oChart, oGroup, arrDrawings[0]]);
				onKeyDown(oEvent);
				oAssert.strictEqual(drawingObjects().getSelectedArray().length, 0, "Test reset shape selection");
			}, oTestTypes.resetShapeSelection);
		});

		QUnit.test("Test reset add shape", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				editor.StartAddShape('rect');
				onKeyDown(oEvent);
				oAssert.strictEqual(editor.isStartAddShape, false, "Test reset add shape");
			}, oTestTypes.resetStartAddShape);
		});

		QUnit.test("Test reset formatting by example", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				editor.SetPaintFormat(AscCommon.c_oAscFormatPainterState.kOn);
				onKeyDown(oEvent);
				oAssert.strictEqual(editor.isFormatPainterOn(), false, "Test reset formatting by example");
			}, oTestTypes.resetFormattingByExample);
		});

		QUnit.test("Test reset", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				editor.SetMarkerFormat(true, true, 0, 0, 0);
				onKeyDown(oEvent);
				oAssert.strictEqual(editor.isMarkerFormat, false, "Test reset marker");
			}, oTestTypes.resetMarkerFormat);

		});

		QUnit.test("Test reset drag'n'drop", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(['Hello Hello'], true);
				moveToParagraph(oParagraph, true);
				moveCursorRight(true, true);
				mouseDown(5, 10, 0, false, 1);
				mouseMove(35, 10, 0, false, 1);
				onKeyDown(oEvent);
				oAssert.strictEqual(editor.WordControl.m_oDrawingDocument.IsTrackText(), false, "Test reset drag'n'drop");
				mouseUp(35, 10, 0, false, 1);
			}, oTestTypes.resetDragNDrop);
		});

		QUnit.test("Test end editing form", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs([''], true);
				const oCheckBox = createCheckBox();
				oCheckBox.MoveCursorToContentControl(true);
				setFillingFormsMode(true);
				onKeyDown(oEvent);
				const oSelectedInfo = oGlobalLogicDocument.GetSelectedElementsInfo();
				oAssert.strictEqual(!!oSelectedInfo.GetInlineLevelSdt(), false, "Test end editing form");
				setFillingFormsMode(false);

				editor.GoToHeader(0);
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Test end editing footer");
				editor.asc_RemoveHeader(0);

				editor.GoToFooter(0);
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), AscCommonWord.docpostype_Content, "Test end editing footer");
				editor.asc_RemoveFooter(0);
			}, oTestTypes.endEditing);
		});

		QUnit.test("Test toggle checkbox", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				getLogicDocumentWithParagraphs(['']);
				const oInlineSdt = createCheckBox();
				setFillingFormsMode(true);
				onKeyDown(oEvent);
				oAssert.strictEqual(oInlineSdt.IsCheckBoxChecked(), true);
				setFillingFormsMode(false);
			}, oTestTypes.toggleCheckBox);
		});

		QUnit.test("Test actions to page up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);

				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 125);
			}, oTestTypes.moveToPreviousPage);
		});

		QUnit.test("Test actions to page up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 90);
			}, oTestTypes.moveToStartPreviousPage);
		});

		QUnit.test("Test move to previous header or footer", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				editor.GoToHeader(2);
				editor.GoToFooter(2);
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[2].Header);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[1].Footer);
			}, oTestTypes.moveToPreviousHeaderFooter);
			editor.asc_RemoveHeader(2);
			editor.asc_RemoveFooter(2);
		});

		QUnit.test("Test actions to page up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);

				editor.GoToHeader(2);
				editor.GoToFooter(2);
				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[1].Header);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[0].Header);
			}, oTestTypes.moveToPreviousHeader);
			editor.asc_RemoveHeader(2);
			editor.asc_RemoveFooter(2);
		});
		function drawingDocument()
		{
			return editor.WordControl.m_oDrawingDocument;
		}
		QUnit.test("Test select to previous page", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);

				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), ' World Hello World Hello World Hello World Hello World Hello World Hello World Hello World');
			}, oTestTypes.selectToPreviousPage);
		});


		QUnit.test("Test select to start previous page", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World');
			}, oTestTypes.selectToStartPreviousPage);
		});

		QUnit.test("Test select to start of next page", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello ', "Test select to begin of next page");
			}, oTestTypes.selectToStartNextPage);
		});
		QUnit.test("Test move to start of next page", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 90, "Test move to begin of next page");
			}, oTestTypes.moveToStartNextPage);
		});
		QUnit.test("Test select to next page", (oAssert) =>
		{
			startTest((oEvent) =>
			{

				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				moveCursorRight();
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'ello World Hello World Hello World Hello World Hello World Hello World Hello World Hello W', "Test select to next page");
			}, oTestTypes.selectToNextPage);
		});
		QUnit.test("Test move to next page", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				moveCursorRight();
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 91, "Test move to next page");
			}, oTestTypes.moveToNextPage);
		});
		QUnit.test("Test move to next header/footer", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				editor.GoToFooter(0);
				editor.GoToHeader(0);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[0].Footer);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[1].Header);
			}, oTestTypes.moveToNextHeaderFooter);
		});
		QUnit.test("Test move to next header", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				editor.GoToFooter(0);
				editor.GoToHeader(0);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[1].Header);

				onKeyDown(oEvent);
				oAssert.strictEqual(oGlobalLogicDocument.GetDocPosType(), docpostype_HdrFtr);
				oAssert.strictEqual(oGlobalLogicDocument.Controller.HdrFtr.CurHdrFtr, oGlobalLogicDocument.Controller.HdrFtr.Pages[2].Header);
			}, oTestTypes.moveToNextHeader);
		});

		QUnit.test("Test actions to end", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 107, "Test move to end of document");
			}, oTestTypes.moveToEndDocument);
		});

		QUnit.test("Test actions to end", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 18, "Test move to end of line");
			}, oTestTypes.moveToEndLine);
		});

		QUnit.test("Test actions to end", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), "Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World", "Test select to end of document");

			}, oTestTypes.selectToEndDocument);
		});

		QUnit.test("Test actions to end", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), "Hello World Hello ", "Test select to end of line");
			}, oTestTypes.selectToEndLine);
		});

		QUnit.test("Test actions to home", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), "World Hello World", "Test select to home of line");
			}, oTestTypes.selectToStartLine);
		});

		QUnit.test("Test actions to home", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), "Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World", "Test select to home of document");

			}, oTestTypes.selectToStartDocument);

		});

		QUnit.test("Test actions to home", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 90, "Test move to home of line");
			}, oTestTypes.moveToStartLine);
		});

		QUnit.test("Test actions to home", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph);

				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 0, "Test move to home of document");
			}, oTestTypes.moveToStartDocument);
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'World', "Test select to previous word");
			}, oTestTypes.selectLeftWord);
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 18, "Test move to previous word");
			}, oTestTypes.moveToLeftWord);
			let oEvent;
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World"], true);
				moveToParagraph(oParagraph);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'd', "Test select to previous symbol");
			}, oTestTypes.selectLeftSymbol);
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World"], true);
				moveToParagraph(oParagraph);

				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 22, "Test move to previous symbol");
			}, oTestTypes.moveToLeftChar);
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(-1, 0, oAssert, oEvent);
			}, oTestTypes.littleMoveGraphicObjectLeft);
		});

		QUnit.test("Test actions to left", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(-5, 0, oAssert, oEvent);
			}, oTestTypes.bigMoveGraphicObjectLeft);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(1, 0, oAssert, oEvent);
			}, oTestTypes.littleMoveGraphicObjectRight);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(5, 0, oAssert, oEvent);
			}, oTestTypes.bigMoveGraphicObjectRight);
		});

		QUnit.test("Test actions to up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(0, -1, oAssert, oEvent);
			}, oTestTypes.littleMoveGraphicObjectUp);
		});

		QUnit.test("Test actions to up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(0, -5, oAssert, oEvent);
			}, oTestTypes.bigMoveGraphicObjectUp);
		});

		QUnit.test("Test actions to down", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(0, 1, oAssert, oEvent);
			}, oTestTypes.littleMoveGraphicObjectDown);
		});

		QUnit.test("Test actions to down", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				moveShapeHelper(0, 5, oAssert, oEvent);
			}, oTestTypes.bigMoveGraphicObjectDown);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(contentPosition(), 1, "Test move to next symbol");
			}, oTestTypes.moveToRightChar);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), "H", "Test select to next symbol");
			}, oTestTypes.selectRightChar);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.deepEqual(contentPosition(), 6, "Test move to next word");
			}, oTestTypes.moveToRightWord);
		});

		QUnit.test("Test actions to right", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Hello ', "Test select to next word");
			}, oTestTypes.selectRightWord);
		});

		QUnit.test("Test actions to up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				moveCursorDown();
				onKeyDown(oEvent);
				oAssert.deepEqual(contentPosition(), 0, "Test move to upper line");
			}, oTestTypes.moveUp);
		});

		QUnit.test("Test actions to up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				moveCursorDown();
				oEvent = createNativeEvent(38, false, true, false, false);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Hello World Hello ', "Test select to upper line");
			}, oTestTypes.selectUp);
		});

		QUnit.test("Test actions to up", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				clean();
				getLogicDocumentWithParagraphs(['']);
				createComboBox();
				setFillingFormsMode(true);
				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(logicContent()[0]), 'Hello', "Test select previous option in combo box");

				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(logicContent()[0]), 'yo', "Test select previous option in combo box");
				setFillingFormsMode(false);
			}, oTestTypes.previousOptionComboBox);
		});

		QUnit.test("Test actions to down", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.deepEqual(contentPosition(), 18, "Test move to down line");
			}, oTestTypes.moveDown);
		});

		QUnit.test("Test actions to down", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World Hello World"], true);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), 'Hello World Hello ', "Test select to down line");
			}, oTestTypes.selectDown);
		});

		QUnit.test("Test actions to down", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				clean();
				getLogicDocumentWithParagraphs(['']);
				createComboBox();
				setFillingFormsMode(true);
				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(logicContent()[0]), 'Hello', "Test select next option in combo box");
				onKeyDown(oEvent);
				oAssert.strictEqual(AscTest.GetParagraphText(logicContent()[0]), 'World', "Test select next option in combo box");
				setFillingFormsMode(false);
			}, oTestTypes.nextOptionComboBox);
		});

		QUnit.test("Test remove front", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World"]);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				selectAll();
				oAssert.strictEqual(getSelectedText(), 'ello World', 'Test remove front symbol');
			}, oTestTypes.removeFrontSymbol);
		});

		QUnit.test("Test remove front", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello World"]);
				moveToParagraph(oParagraph, true);
				onKeyDown(oEvent);
				selectAll();
				oAssert.strictEqual(getSelectedText(), 'World', 'Test remove front word');
			}, oTestTypes.removeFrontWord);
		});

		QUnit.test("Test replace unicode code to symbol", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["2601"]);
				moveToParagraph(oParagraph, true);
				moveCursorRight(true, true);
				onKeyDown(oEvent);
				oAssert.strictEqual(getSelectedText(), '☁', 'Test replace unicode code to symbol');
			}, oTestTypes.unicodeToChar);
		});

		QUnit.test("Test show context menu", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				const {oParagraph} = getLogicDocumentWithParagraphs(["Hello Text"]);
				moveToParagraph(oParagraph, true);

				oEvent = createNativeEvent(93, false, false, false, false);
				executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);

				AscCommon.AscBrowser.isOpera = true;
				oEvent = createNativeEvent(57351, false, false, false, false);
				executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
				AscCommon.AscBrowser.isOpera = false;

				oEvent = createNativeEvent(121, false, true, false, false);
				executeTestWithCatchEvent('asc_onContextMenu', () => true, true, oEvent, oAssert);
			}, oTestTypes.showContextMenu);

		});

		QUnit.test("Test disable numlock", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				oEvent = createNativeEvent(144, false, false, false, false);
				onKeyDown(oEvent);
				oAssert.strictEqual(oEvent.isDefaultPrevented, true, 'Test prevent default on numlock');
			}, oTestTypes.disableNumLock);
		});

		QUnit.test("Test disable scroll lock", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				oEvent = createNativeEvent(145, false, false, false, false);
				onKeyDown(oEvent);
				oAssert.strictEqual(oEvent.isDefaultPrevented, true, 'Test prevent default on scroll lock');
			}, oTestTypes.disableScrollLock);
		});

		QUnit.test("Test add SJK test", (oAssert) =>
		{
			startTest((oEvent) =>
			{
				checkTextAfterKeyDownHelperEmpty(' ', oEvent, oAssert, 'Check add space after SJK space');
			}, oTestTypes.addSJKSpace);
		});
	});
})(window);
