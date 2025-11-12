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

$(function () {
	
	QUnit.module("Test writeTo/readFrom Xml");
	
	
	QUnit.test("UserMaster:", function(assert)
	{
		function Test(user, msg)
		{
			let writer = AscTest.GetXmlWriter();
			user.toXml(writer);
			
			let reader = AscTest.GetXmlReader(writer);
			let _user  = AscOForm.CUserMaster.fromXml(reader);
			assert.strictEqual(user.isEqual(_user), true, msg);
		}
		
		let user = new AscOForm.CUserMaster(true);
		Test(user, "Check empty user master");
		
		user.setRole("RoleName");
		Test(user, "Check user with role=" + "RoleName");
		
		user.setColor(4, 8, 15);
		Test(user, "Check user with color");
	});
	
	QUnit.test("FieldGroup:", function(assert)
	{
		function Test(fieldGroup, msgHeader)
		{
			let writer = AscTest.GetXmlWriter();
			fieldGroup.toXml(writer);
			
			let reader = AscTest.GetXmlReader(writer);
			
			if (!reader.ReadNextNode() || "fieldGroup" !== reader.GetNameNoNS())
			{
				assert.strictEqual(false, true, msgHeader + ": bad xml");
				return;
			}
			
			let fG     = AscOForm.CFieldGroup.fromXml(reader);
			
			assert.strictEqual(fG.getWeight(), fieldGroup.getWeight(), msgHeader + " check weight");
			assert.strictEqual(fG.getUserCount(), fieldGroup.getUserCount(), msgHeader + ": check number of users");
			
			if (fieldGroup.getUserCount() === fG.getUserCount())
			{
				for (let index = 0, count = fieldGroup.getUserCount(); index < count; ++index)
				{
					assert.strictEqual(fG.getUser(index), fieldGroup.getUser(index), msgHeader + ": check user[" + index + "]");
				}
			}
			
			assert.strictEqual(fieldGroup.getFieldCount(), fieldGroup.getFieldCount(), msgHeader + ": check number of fields");
			
			if (fieldGroup.getFieldCount() === fG.getFieldCount())
			{
				for (let index = 0, count = fieldGroup.getFieldCount(); index < count; ++index)
				{
					assert.strictEqual(fG.getField(index), fieldGroup.getField(index), msgHeader + ": check field[" + index + "]");
				}
			}
		}
		
		let fieldGroup = new AscOForm.CFieldGroup();
		Test(fieldGroup, "Empty group with undefined weight");
		
		fieldGroup.setWeight(12);
		Test(fieldGroup, "Empty group with weight=12");
	});
});
