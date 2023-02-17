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
	
	let PluginsApi = AscTest.Editor;
	
	let logicDocument = AscTest.CreateLogicDocument();
	logicDocument.RemoveFromContent(0, logicDocument.GetElementsCount(), false);
	
	function MoveToNewParagraph()
	{
		let p = AscTest.CreateParagraph();
		logicDocument.AddToContent(logicDocument.GetElementsCount(), p);
		p.SetThisElementCurrent();
		return p;
	}
	
	QUnit.module("Test plugins api");
	
	QUnit.test("Test work with addin fields", function (assert)
	{
		MoveToNewParagraph();
		
		assert.strictEqual(PluginsApi.pluginMethod_GetAllAddinFields().length, 0, "Check addin fields in empty document");
		
		MoveToNewParagraph();
		PluginsApi.pluginMethod_AddAddinField({"Value" : "Test addin", "Content" : 123});
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				AscWord.CAddinFieldData.FromObject({"FieldId" : "1", "Value" : "Test addin", "Content" : "123"})
			],
			"Add addin field and check get function");
		
		MoveToNewParagraph();
		assert.strictEqual(logicDocument.GetAllFields().length, 1, "Check the number of all fields");
		logicDocument.AddFieldWithInstruction("PAGE");
		assert.strictEqual(logicDocument.GetAllFields().length, 2, "Add PAGE field and check the number of all fields");
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields().length, 1, "Check the number of addin fields");
		
		MoveToNewParagraph();
		PluginsApi.pluginMethod_AddAddinField({"Value" : "Addin №2", "Content" : "This is the second addin field"});
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				AscWord.CAddinFieldData.FromObject({"FieldId" : "1", "Value" : "Test addin", "Content" : "123"}),
				AscWord.CAddinFieldData.FromObject({"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"})
			],
			"Add addin field and check get function");
		
		PluginsApi.pluginMethod_UpdateAddinFields(
			[
				AscWord.CAddinFieldData.FromObject({"FieldId" : "1", "Value" : "Addin №1", "Content" : "This is the first addin field"}),
			],
			"Add addin field and check get function");
		
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				AscWord.CAddinFieldData.FromObject({"FieldId" : "1", "Value" : "Addin №1", "Content" : "This is the first addin field"}),
				AscWord.CAddinFieldData.FromObject({"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"})
			],
			"Change the first adding and check get function");
		
		logicDocument.RemoveFromContent(1, 1);
		assert.deepEqual(PluginsApi.pluginMethod_GetAllAddinFields(),
			[
				AscWord.CAddinFieldData.FromObject({"FieldId" : "3", "Value" : "Addin №2", "Content" : "This is the second addin field"})
			],
			"Remove the paragraph with the first field and check get addin function");
	});
	
	QUnit.test("Test RemoveFieldWrapper", function(assert)
	{
		AscTest.ClearDocument();
		MoveToNewParagraph();
		assert.strictEqual(logicDocument.GetAllFields().length, 0, "Check the number of all fields in the empty document");
		
		MoveToNewParagraph();
		logicDocument.AddFieldWithInstruction("PAGE");
		
		let p = MoveToNewParagraph();
		let field = logicDocument.AddFieldWithInstruction("PAGE");
		assert.strictEqual(logicDocument.GetAllFields().length, 2, "Add two PAGE fields and check count of all fields");
		
		logicDocument.UpdateFields(false);
		
		assert.strictEqual(p.GetText(), "1 ", "Check the text of the third paragraph");
		
		let fieldId = field.GetFieldId();
		PluginsApi.pluginMethod_RemoveFieldWrapper(fieldId);
		assert.strictEqual(logicDocument.GetAllFields().length, 1, "Remove field wrapper from second field and check number of fields");
		assert.strictEqual(p.GetText(), "1 ", "Check the text of the third paragraph");
	});
	
	QUnit.test("Test SetEditingRestrictions", function(assert)
	{
		AscTest.ClearDocument();
		MoveToNewParagraph();
		
		assert.strictEqual(logicDocument.CanEdit(), true, "Check if we can edit new document");
		
		PluginsApi.pluginMethod_SetEditingRestrictions("readOnly");
		assert.strictEqual(logicDocument.CanEdit(), false, "Set read only restriction and check if we can edit document");
	});
	
	
});
