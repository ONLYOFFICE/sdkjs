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

	QUnit.module("Unit-tests for AscWord.CParagraphContentPos");


	QUnit.test("Test:", function (assert)
	{
		let oPos = new AscWord.CParagraphContentPos();
		assert.strictEqual(oPos.GetDepth(), -1, "Create new pos and check the depth (must be negative)");

		oPos.Add(4);
		oPos.Add(8);
		oPos.Add(15);
		oPos.Add(16);
		oPos.Add(23);
		oPos.Add(42);

		assert.strictEqual(oPos.GetDepth(), 5, "Fill the pos and check the depth");

		assert.strictEqual(oPos.Get(0), 4, "[0]");
		assert.strictEqual(oPos.Get(1), 8, "[1]");
		assert.strictEqual(oPos.Get(2), 15, "[2]");
		assert.strictEqual(oPos.Get(3), 16, "[3]");
		assert.strictEqual(oPos.Get(4), 23, "[4]");
		assert.strictEqual(oPos.Get(5), 42, "[5]");

		let oPos2 = oPos.Copy();

		assert.strictEqual(oPos2.Compare(oPos), 0, "Make a copy and check for equality");

		oPos2.Update2(20, 3);
		assert.strictEqual(oPos2.Get(3), 20, "Check [3] after update");
		assert.strictEqual(oPos2.GetDepth(), 5, "Check depth after Update2");
		assert.strictEqual(oPos2.Compare(oPos), 1, "Compare pos +1");

		oPos2.Update2(2, 3);
		assert.strictEqual(oPos2.Compare(oPos), -1, "Compare pos -1");

		oPos2.DecreaseDepth(3);
		assert.strictEqual(oPos2.GetDepth(), 2, "Check depth after decreasing by 3");
		assert.strictEqual(oPos2.Compare(oPos), -1, "Compare decreased pos and original pos");
		assert.strictEqual(oPos2.IsPartOf(oPos), true, "Check decreased pos as part of original");

		oPos2.Update2(2, 2);
		assert.strictEqual(oPos2.IsPartOf(oPos), false, "Spoil decreased pos and check it as part of original");

		let oPos3 = oPos.Copy();
		oPos3.DecreaseDepth(1);
		assert.strictEqual(oPos3.Compare(oPos), -1, "Make a copy and decrease pos and check for equality");
	});
});
