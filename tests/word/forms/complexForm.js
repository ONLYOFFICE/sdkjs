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

$(function () {

	let logicDocument = CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);

	let formsManager = logicDocument.GetFormsManager();

	let p1 = new AscWord.CParagraph(editor.WordControl);
	let p2 = new AscWord.CParagraph(editor.WordControl);

	logicDocument.AddToContent(0, p1);
	logicDocument.AddToContent(1, p2);

	let r1 = new AscWord.CRun();
	p1.AddToContent(0, r1);
	r1.AddText("Hello Word!");

	let r2 = new AscWord.CRun();
	p2.AddToContent(0, r2);
	r2.AddText("Абракадабра");


	QUnit.module("Check forms");


	QUnit.test("Test: \"Complex Form\"", function (assert)
	{
		let complexForm = logicDocument.AddComplexForm();
		complexForm.SetFormPr(new AscWord.CSdtFormPr());

		assert.strictEqual(formsManager.GetAllForms().length, 1, "Add complex form to document (check forms count)");

		complexForm.MoveCursorToStartPos();
		complexForm.SetThisElementCurrent();



	});
});
