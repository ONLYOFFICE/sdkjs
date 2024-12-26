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


$(function ()
{
	let pdfDoc = AscTest.CreatePdfDocument();
	pdfDoc.AddPage();
	
	AscPDF.CTextField.prototype.UpdateScroll = function(){};
	
	function CreateTextForm(name)
	{
		let oField = pdfDoc.AddField(name, AscPDF.FIELD_TYPES.text, 0, [20, 20, 50, 20]);
		return oField; 
	}
	function EnterTextToForm(form, text)
	{
		let chars = text.codePointsArray();
		AscTest.Editor.DocumentRenderer.getPDFDoc().activeForm = form;
		form.EnterText(chars);
		form.SetDrawHighlight(false);
		pdfDoc.EnterDownActiveField();
	}
	function AddJsAction(form, trigger, script)
	{
		form.SetAction(trigger, script);
	}
	
	QUnit.module("PDF form actions test");
	
	QUnit.test("Test calculate action", function (assert)
	{
		let textForm1 = CreateTextForm("TextForm1");
		let textForm2 = CreateTextForm("TextForm2");
		let textForm3 = CreateTextForm("TextForm3");
		
		textForm1.GetFormApi().value = "1";
		textForm2.GetFormApi().value = "2";
		textForm3.GetFormApi().value = "3";
		
		assert.strictEqual(textForm1.GetValue(), "1", "Check form1 value");
		assert.strictEqual(textForm2.GetValue(), "2", "Check form2 value");
		assert.strictEqual(textForm3.GetValue(), "3", "Check form3 value");
		
		AddJsAction(textForm1, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm2').value += 1");
		AddJsAction(textForm2, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm3').value += 1");
		AddJsAction(textForm3, AscPDF.FORMS_TRIGGERS_TYPES.Calculate, "this.getField('TextForm1').value += 1");
		
		textForm2.MoveCursorRight();
		EnterTextToForm(textForm2, "2");
		assert.strictEqual(textForm1.GetValue(), "2", "Check form1 value");
		assert.strictEqual(textForm2.GetValue(), "22", "Check form2 value");
		assert.strictEqual(textForm3.GetValue(), "4", "Check form3 value");

		textForm3.MoveCursorRight();
		EnterTextToForm(textForm3, "3");
		
		assert.strictEqual(textForm1.GetValue(), "3", "Check form1 value");
		assert.strictEqual(textForm2.GetValue(), "23", "Check form2 value");
		assert.strictEqual(textForm3.GetValue(), "43", "Check form3 value");
	});
});
