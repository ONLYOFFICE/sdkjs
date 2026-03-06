/*
 * (c) Copyright Ascensio System SIA 2010-2026
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
	var Api = AscBuilder.Cell.Api;
	var ws  = AscTest.JsApi.GetActiveSheet();

	// ======= CLEANUP HELPERS =======

	function clearAllComments() {
		// Doc-level (workbook) comments
		Asc.editor.wbModel.aComments.length = 0;
		// Active-worksheet comments
		Asc.editor.wbModel.getActiveWs().aComments.length = 0;
	}

	// ======= MODULE 1: Api.AddComment — basics =======

	QUnit.module("Api.AddComment — basics", { beforeEach: clearAllComments });

	QUnit.test("Returns ApiComment for valid text and author", function (assert) {
		var c = Api.AddComment("Hello world", "Author");
		assert.ok(c, "Returns an object");
		assert.strictEqual(c.GetClassType(), "comment", "ClassType is 'comment'");
	});

	QUnit.test("GetText matches input text", function (assert) {
		var c = Api.AddComment("My text", "Author");
		assert.strictEqual(c.GetText(), "My text", "GetText matches");
	});

	QUnit.test("GetAuthorName matches provided author", function (assert) {
		var c = Api.AddComment("Hello", "CustomAuthor");
		assert.strictEqual(c.GetAuthorName(), "CustomAuthor", "Author set correctly");
	});

	QUnit.test("Returns null for empty string", function (assert) {
		assert.strictEqual(Api.AddComment("", "Author"), null, "Empty string → null");
	});

	QUnit.test("Returns null for whitespace-only string", function (assert) {
		assert.strictEqual(Api.AddComment("   ", "Author"), null, "Whitespace-only → null");
	});

	QUnit.test("Returns null for non-string input", function (assert) {
		assert.strictEqual(Api.AddComment(42,        "Author"), null, "Number → null");
		assert.strictEqual(Api.AddComment(null,      "Author"), null, "null → null");
		assert.strictEqual(Api.AddComment(undefined, "Author"), null, "undefined → null");
	});

	// ======= MODULE 2: Api.GetComments — workbook-level =======

	QUnit.module("Api.GetComments — workbook-level", { beforeEach: clearAllComments });

	QUnit.test("Returns empty array initially", function (assert) {
		var comments = Api.GetComments();
		assert.ok(Array.isArray(comments), "Returns an array");
		assert.strictEqual(comments.length, 0, "No comments initially");
	});

	QUnit.test("Reflects added doc-level comments", function (assert) {
		Api.AddComment("First", "Author");
		Api.AddComment("Second", "Author");
		assert.strictEqual(Api.GetComments().length, 2, "Two comments returned");
	});

	QUnit.test("Returns ApiComment instances with correct text", function (assert) {
		Api.AddComment("Doc comment", "Author");
		var comments = Api.GetComments();
		assert.strictEqual(comments[0].GetClassType(), "comment", "Instance is ApiComment");
		assert.strictEqual(comments[0].GetText(), "Doc comment", "Text matches");
	});

	QUnit.test("Worksheet-level comments not included in GetComments()", function (assert) {
		ws.GetRange("A1").AddComment("Sheet only", "Author");
		assert.strictEqual(Api.GetComments().length, 0, "GetComments returns only doc-level comments");
	});

	QUnit.test("Comments property getter matches GetComments()", function (assert) {
		Api.AddComment("Prop test", "Author");
		assert.strictEqual(Api.Comments.length, Api.GetComments().length, "Property matches method");
	});

	// ======= MODULE 3: Api.GetCommentById =======

	QUnit.module("Api.GetCommentById", { beforeEach: clearAllComments });

	QUnit.test("Finds an added doc-level comment by ID", function (assert) {
		var c = Api.AddComment("Find me", "Author");
		var id = c.GetId();
		assert.ok(id, "Comment has an ID");
		var found = Api.GetCommentById(id);
		assert.ok(found, "Found comment is truthy");
		assert.strictEqual(found.GetText(), "Find me", "Text matches");
	});

	QUnit.test("Finds an added worksheet-level comment by ID", function (assert) {
		var c = ws.GetRange("A1").AddComment("Sheet find me", "Author");
		var id = c.GetId();
		var found = Api.GetCommentById(id);
		assert.ok(found, "Found worksheet comment");
		assert.strictEqual(found.GetText(), "Sheet find me", "Text matches");
	});

	QUnit.test("Returns null for unknown ID", function (assert) {
		assert.strictEqual(Api.GetCommentById("nonexistent-id-xyz"), null, "Unknown ID → null");
	});

	// ======= MODULE 4: Api.GetAllComments =======

	QUnit.module("Api.GetAllComments", { beforeEach: clearAllComments });

	QUnit.test("Returns empty array initially", function (assert) {
		assert.strictEqual(Api.GetAllComments().length, 0, "Empty initially");
	});

	QUnit.test("Includes doc-level comments", function (assert) {
		Api.AddComment("Workbook comment", "Author");
		var all = Api.GetAllComments();
		assert.ok(all.length >= 1, "Doc-level comment is included");
	});

	QUnit.test("Includes worksheet-level comments", function (assert) {
		ws.GetRange("A1").AddComment("Sheet comment", "Author");
		var all = Api.GetAllComments();
		assert.ok(all.length >= 1, "Sheet-level comment is included");
	});

	QUnit.test("Combines doc-level and sheet-level comments", function (assert) {
		Api.AddComment("Doc level", "Author");
		ws.GetRange("B2").AddComment("Sheet level", "Author");
		var all = Api.GetAllComments();
		assert.ok(all.length >= 2, "Both doc-level and sheet-level comments present");
	});

	QUnit.test("AllComments property getter matches GetAllComments()", function (assert) {
		Api.AddComment("Prop check", "Author");
		assert.strictEqual(Api.AllComments.length, Api.GetAllComments().length, "Property matches method");
	});

	// ======= MODULE 5: ApiComment — properties =======

	QUnit.module("ApiComment — properties", { beforeEach: clearAllComments });

	QUnit.test("Id: non-empty string, getter matches GetId()", function (assert) {
		var c = Api.AddComment("Test", "Author");
		var id = c.GetId();
		assert.strictEqual(typeof id, "string", "Id is a string");
		assert.ok(id.length > 0, "Id is non-empty");
		assert.strictEqual(c.Id, id, "Id property matches GetId()");
	});

	QUnit.test("Text: get/set, property, ignores empty/whitespace", function (assert) {
		var c = Api.AddComment("Original", "Author");
		assert.strictEqual(c.GetText(), "Original", "GetText initial");
		c.SetText("Updated");
		assert.strictEqual(c.GetText(), "Updated", "GetText after SetText");
		assert.strictEqual(c.Text, "Updated", "Text property after SetText");
		c.SetText("");
		assert.strictEqual(c.GetText(), "Updated", "Empty string ignored");
		c.SetText("   ");
		assert.strictEqual(c.GetText(), "Updated", "Whitespace ignored");
	});

	QUnit.test("AuthorName: get/set, property", function (assert) {
		var c = Api.AddComment("Test", "Author1");
		assert.strictEqual(c.GetAuthorName(), "Author1", "GetAuthorName initial");
		c.SetAuthorName("Author2");
		assert.strictEqual(c.GetAuthorName(), "Author2", "GetAuthorName after SetAuthorName");
		assert.strictEqual(c.AuthorName, "Author2", "AuthorName property");
	});

	QUnit.test("UserId: get/set, property", function (assert) {
		var c = Api.AddComment("Test", "Author");
		c.SetUserId("user-42");
		assert.strictEqual(c.GetUserId(), "user-42", "GetUserId after set");
		assert.strictEqual(c.UserId, "user-42", "UserId property");
	});

	QUnit.test("Solved: false initially, get/set, property", function (assert) {
		var c = Api.AddComment("Test", "Author");
		assert.strictEqual(c.IsSolved(), false, "Initially not solved");
		c.SetSolved(true);
		assert.strictEqual(c.IsSolved(), true, "Solved after SetSolved(true)");
		assert.strictEqual(c.Solved, true, "Solved property is true");
		c.SetSolved(false);
		assert.strictEqual(c.Solved, false, "Solved property is false after reset");
	});

	QUnit.test("TimeUTC: get/set, property", function (assert) {
		var c = Api.AddComment("Test", "Author");
		c.SetTimeUTC(1700000000000);
		assert.strictEqual(c.GetTimeUTC(), 1700000000000, "GetTimeUTC after set");
		assert.strictEqual(c.TimeUTC, 1700000000000, "TimeUTC property");
	});

	QUnit.test("TimeUTC: invalid string yields 0", function (assert) {
		var c = Api.AddComment("Test", "Author");
		c.SetTimeUTC("not-a-number");
		assert.strictEqual(c.GetTimeUTC(), 0, "Invalid string → 0");
	});

	QUnit.test("RepliesCount: 0 initially, property matches method", function (assert) {
		var c = Api.AddComment("Test", "Author");
		assert.strictEqual(c.GetRepliesCount(), 0, "No replies initially");
		assert.strictEqual(c.RepliesCount, 0, "RepliesCount property is 0");
	});

	// ======= MODULE 6: ApiComment — replies =======

	QUnit.module("ApiComment — replies", { beforeEach: clearAllComments });

	QUnit.test("AddReply increments RepliesCount", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("Reply 1", "ReplyAuthor");
		assert.strictEqual(c.GetRepliesCount(), 1, "Count is 1 after first reply");
		c.AddReply("Reply 2", "ReplyAuthor2");
		assert.strictEqual(c.GetRepliesCount(), 2, "Count is 2 after second reply");
	});

	QUnit.test("GetReply returns correct text and author", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("Hello reply", "ReplyAuthor");
		var r = c.GetReply(0);
		assert.ok(r, "GetReply(0) returns an object");
		assert.strictEqual(r.GetClassType(), "commentReply", "Reply ClassType is 'commentReply'");
		assert.strictEqual(r.Text, "Hello reply", "Reply text matches");
		assert.strictEqual(r.AuthorName, "ReplyAuthor", "Reply author matches");
	});

	QUnit.test("GetReply with out-of-range index falls back to first reply", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("Only reply", "ReplyAuthor");
		var r = c.GetReply(999);
		assert.ok(r, "Returns an object for out-of-range index");
		assert.strictEqual(r.Text, "Only reply", "Falls back to first reply");
	});

	QUnit.test("AddReply ignores empty text", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("", "ReplyAuthor");
		assert.strictEqual(c.GetRepliesCount(), 0, "Empty text → reply not added");
	});

	QUnit.test("AddReply at specific position", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("R1", "ReplyAuthor");
		c.AddReply("R2", "ReplyAuthor");
		c.AddReply("Middle", "Mid", undefined, 1);
		assert.strictEqual(c.GetRepliesCount(), 3, "Three replies");
		assert.strictEqual(c.GetReply(1).Text, "Middle", "Middle reply inserted at position 1");
	});

	QUnit.test("RemoveReplies removes by count from position", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("R1", "ReplyAuthor");
		c.AddReply("R2", "ReplyAuthor");
		c.AddReply("R3", "ReplyAuthor");
		c.RemoveReplies(0, 2);
		assert.strictEqual(c.GetRepliesCount(), 1, "One reply remains");
		assert.strictEqual(c.GetReply(0).Text, "R3", "R3 is the remaining reply");
	});

	QUnit.test("RemoveReplies with bRemoveAll removes all", function (assert) {
		var c = Api.AddComment("Parent", "Author");
		c.AddReply("R1", "ReplyAuthor");
		c.AddReply("R2", "ReplyAuthor");
		c.RemoveReplies(0, 1, true);
		assert.strictEqual(c.GetRepliesCount(), 0, "All replies removed");
	});

	// ======= MODULE 7: ApiComment.Delete =======

	QUnit.module("ApiComment.Delete", { beforeEach: clearAllComments });

	QUnit.test("Delete removes comment from Api.GetComments()", function (assert) {
		Api.AddComment("Keep", "Author");
		var toDelete = Api.AddComment("Delete me", "Author");
		assert.strictEqual(Api.GetComments().length, 2, "Two comments before delete");
		toDelete.Delete();
		var after = Api.GetComments();
		assert.strictEqual(after.length, 1, "One comment after delete");
		assert.strictEqual(after[0].GetText(), "Keep", "Remaining comment is correct");
	});

	QUnit.test("Delete worksheet comment removes it from ApiWorksheet.GetComments()", function (assert) {
		ws.GetRange("A1").AddComment("Sheet keep", "Author");
		var toDelete = ws.GetRange("B2").AddComment("Sheet delete", "Author");
		assert.strictEqual(ws.GetComments().length, 2, "Two before delete");
		toDelete.Delete();
		assert.strictEqual(ws.GetComments().length, 1, "One after delete");
		assert.strictEqual(ws.GetComments()[0].GetText(), "Sheet keep", "Correct comment remains");
	});

	// ======= MODULE 8: ApiWorksheet.GetComments =======

	QUnit.module("ApiWorksheet.GetComments", { beforeEach: clearAllComments });

	QUnit.test("Returns empty array initially", function (assert) {
		var comments = ws.GetComments();
		assert.ok(Array.isArray(comments), "Returns an array");
		assert.strictEqual(comments.length, 0, "No worksheet comments initially");
	});

	QUnit.test("Range.AddComment adds to worksheet comments", function (assert) {
		ws.GetRange("A1").AddComment("Sheet comment 1", "Author");
		ws.GetRange("B2").AddComment("Sheet comment 2", "Author");
		var comments = ws.GetComments();
		assert.strictEqual(comments.length, 2, "Two worksheet comments");
	});

	QUnit.test("Returns ApiComment instances with correct text", function (assert) {
		ws.GetRange("C3").AddComment("Check text", "Author");
		var comments = ws.GetComments();
		assert.strictEqual(comments[0].GetClassType(), "comment", "Instance is ApiComment");
		assert.strictEqual(comments[0].GetText(), "Check text", "Text matches");
	});

	QUnit.test("Doc-level comments not included in ApiWorksheet.GetComments()", function (assert) {
		Api.AddComment("Doc only", "Author");
		assert.strictEqual(ws.GetComments().length, 0, "Doc-level comment not in worksheet");
	});

	QUnit.test("Comments property getter matches GetComments()", function (assert) {
		ws.GetRange("D4").AddComment("Prop test", "Author");
		assert.strictEqual(ws.Comments.length, ws.GetComments().length, "Property matches method");
	});

	// ======= MODULE 9: ApiRange.AddComment / GetComment =======

	QUnit.module("ApiRange.AddComment / GetComment", { beforeEach: clearAllComments });

	QUnit.test("AddComment returns ApiComment with correct text", function (assert) {
		var c = ws.GetRange("A1").AddComment("Range comment", "Author");
		assert.ok(c, "Returns an object");
		assert.strictEqual(c.GetClassType(), "comment", "ClassType is 'comment'");
		assert.strictEqual(c.GetText(), "Range comment", "Text matches");
	});

	QUnit.test("AddComment with author sets author name", function (assert) {
		var c = ws.GetRange("B1").AddComment("With author", "RangeAuthor");
		assert.strictEqual(c.GetAuthorName(), "RangeAuthor", "Author set correctly");
	});

	QUnit.test("AddComment returns null for empty text", function (assert) {
		assert.strictEqual(ws.GetRange("A1").AddComment("", "Author"), null, "Empty text → null");
	});

	QUnit.test("GetComment returns the comment on a cell", function (assert) {
		ws.GetRange("A1").AddComment("Find me", "Author");
		var c = ws.GetRange("A1").GetComment();
		assert.ok(c, "GetComment returns an object");
		assert.strictEqual(c.GetText(), "Find me", "Text matches");
	});

	QUnit.test("GetComment returns null for cell without a comment", function (assert) {
		assert.strictEqual(ws.GetRange("Z99").GetComment(), null, "No comment → null");
	});

	QUnit.test("GetComment returns null for multi-cell range", function (assert) {
		ws.GetRange("A1").AddComment("Single cell", "Author");
		assert.strictEqual(ws.GetRange("A1:B2").GetComment(), null, "Multi-cell range → null");
	});

	QUnit.test("Comments property on ApiRange matches GetComment()", function (assert) {
		ws.GetRange("A1").AddComment("Prop test", "Author");
		var viaProp   = ws.GetRange("A1").Comments;
		var viaMethod = ws.GetRange("A1").GetComment();
		assert.ok(viaProp, "Comments property is truthy");
		assert.strictEqual(viaProp.GetText(), viaMethod.GetText(), "Both return same comment text");
	});
});
