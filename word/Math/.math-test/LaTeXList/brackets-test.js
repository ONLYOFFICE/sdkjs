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
function brackets (test) {
	test(
		"(a)[b]\\{c\\}|d|\\|e\\|\\langlef\\rangle\\lfloorg\\rfloor\\lceilh\\rceil\\ulcorneri\\urcorner/j\\backslash",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": "BracketBlock",
					"left": "(",
					"right": ")",
					"value": {
						"type": "CharLiteral",
						"value": "a"
					}
				},
				{
					"type": "BracketBlock",
					"left": "[",
					"right": "]",
					"value": {
						"type": "CharLiteral",
						"value": "b"
					}
				},
				{
					"type": "BracketBlock",
					"left": "\\{",
					"right": "\\}",
					"value": {
						"type": "CharLiteral",
						"value": "c"
					}
				},
				{
					"type": "BracketBlock",
					"left": "|",
					"right": "|",
					"value": {
						"type": "CharLiteral",
						"value": "d"
					}
				},
				{
					"type": "BracketBlock",
					"left": "\\|",
					"right": "\\|",
					"value": {
						"type": "CharLiteral",
						"value": "e"
					}
				},
				{
					"type": "BracketBlock",
					"left": "⟨",
					"right": "⟩",
					"value": {
						"type": "CharLiteral",
						"value": "f"
					}
				},
				{
					"type": "BracketBlock",
					"left": "⌊",
					"right": "⌋",
					"value": {
						"type": "CharLiteral",
						"value": "g"
					}
				},
				{
					"type": "BracketBlock",
					"left": "⌈",
					"right": "⌉",
					"value": {
						"type": "CharLiteral",
						"value": "h"
					}
				},
				{
					"type": "BracketBlock",
					"left": "┌",
					"right": "┐",
					"value": {
						"type": "CharLiteral",
						"value": "i"
					}
				},

				//Word doesn't implement it brackets
				// {
				// 	"type": "BracketBlock",
				// 	"left": "/",
				// 	"right": "\\",
				// 	"value": {
				// 		"type": "CharLiteral",
				// 		"value": "j"
				// 	}
				// }
			]
		},
		"Check brackets (a)[b]\\{c\\}|d|\\|e\\|\\langlef\\rangle\\lfloorg\\rfloor\\lceilh\\rceil\\ulcorneri\\urcorner/j\\backslash"
	);
	test(
		"(2+1]",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "BracketBlock",
				"left": "(",
				"right": "]",
				"value": [
					{
						"type": "NumberLiteral",
						"value": "2"
					},
					{
						"type": "OperatorLiteral",
						"value": "+"
					},
					{
						"type": "NumberLiteral",
						"value": "1"
					}
				]
			}
		},
		"Check (2+1]"
	);
	//Word doesn't support \backslash bracket
	// test(
	// 	"\\{2+1\\backslash",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": {
	// 			"type": "BracketBlock",
	// 			"left": "\\{",
	// 			"right": "\\",
	// 			"value": [
	// 				{
	// 					"type": "NumberLiteral",
	// 					"value": "2"
	// 				},
	// 				{
	// 					"type": "OperatorLiteral",
	// 					"value": "+"
	// 				},
	// 				{
	// 					"type": "NumberLiteral",
	// 					"value": "1"
	// 				}
	// 			]
	// 		}
	// 	},
	// 	"Check \\{2+1\\backslash"
	// );
	test(
		"\\left.1+2\\right)",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "BracketBlock",
				"left": ".",
				"right": ")",
				"value": [
					{
						"type": "NumberLiteral",
						"value": "1"
					},
					{
						"type": "OperatorLiteral",
						"value": "+"
					},
					{
						"type": "NumberLiteral",
						"value": "2"
					}
				]
			}
		},
		"Check \\left.1+2\\right)"
	);
	test(
		"|2|+\\{1\\}+|2|",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": "BracketBlock",
					"left": "|",
					"right": "|",
					"value": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				{
					"type": "OperatorLiteral",
					"value": "+"
				},
				{
					"type": "BracketBlock",
					"left": "\\{",
					"right": "\\}",
					"value": {
						"type": "NumberLiteral",
						"value": "1"
					}
				},
				{
					"type": "OperatorLiteral",
					"value": "+"
				},
				{
					"type": "BracketBlock",
					"left": "|",
					"right": "|",
					"value": {
						"type": "NumberLiteral",
						"value": "2"
					}
				}
			]
		},
		"Check |2|+\\{1\\}+|2|"
	);
}
window["AscMath"].brackets = brackets;
