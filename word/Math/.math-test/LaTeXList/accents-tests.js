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
function accents(test) {
	test(
		"\\dot{a}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "a"
				},
				"value": "̇"
			}
		},
		"Check \\dot{a}"
	);
	test(
		"\\ddot{b}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "b"
				},
				"value": "̈"
			}
		},
		"Check \\ddot{b}"
	);
	test(
		"\\acute{c}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "c"
				},
				"value": "́"
			}
		},
		"Check \\acute{c}"
	);
	test(
		"\\grave{d}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "d"
				},
				"value": "̀"
			}
		},
		"Check \\grave{d}"
	);
	test(
		"\\check{e}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "e"
				},
				"value": "̌"
			}
		},
		"Check \\check{e}"
	);
	test(
		"\\breve{f}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "f"
				},
				"value": "̆"
			}
		},
		"Check \\breve{f}"
	);
	test(
		"\\tilde{g}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "g"
				},
				"value": "̃"
			}
		},
		"Check \\tilde{g}"
	);
	test(
		"\\bar{h}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "h"
				},
				"value": "̅"
			}
		},
		"Check \\bar{h}"
	);
	test(
		"\\widehat{j}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "j"
				},
				"value": "̂"
			}
		},
		"Check \\widehat{j}"
	);
	test(
		"\\vec{k}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "CharLiteral",
					"value": "k"
				},
				"value": "⃗"
			}
		}
		,
		"Check \\vec{k}"
	);
	//doesn't implement in word
	// test(
	// 	"\\not{l}",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": {
	// 			"type": "AccentLiteral",
	// 			"base": {
	// 				"type": "CharLiteral",
	// 				"value": "l"
	// 			},
	// 			"value": 824
	// 		}
	// 	},
	// 	"Check \\not{l}"
	// );
	test(
		"\\vec \\frac{k}{2}",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "FractionLiteral",
					"up": {
						"type": "CharLiteral",
						"value": "k"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "2"
					}
				},
				"value": "⃗"
			}
		},
		"Check \\vec \\frac{k}{2}"
	);
	// test(
	// 	"\\not\\notl2",
	// 	{
	// 		"type": "LaTeXEquation",
	// 		"body": [
	// 			{
	// 				"type": "AccentLiteral",
	// 				"base": {
	// 					"type": "AccentLiteral",
	// 					"base": {
	// 						"type": "CharLiteral",
	// 						"value": "l"
	// 					},
	// 					"value": 824
	// 				},
	// 				"value": 824
	// 			},
	// 			{
	// 				"type": "NumberLiteral",
	// 				"value": "2"
	// 			}
	// 		]
	// 	},
	// 	"Check \\notl"
	// );
	test(
		"5''",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "NumberLiteral",
					"value": "5"
				},
				"value": "''"
			}
		},
		"Check 5''"
	);
	test(
		"\\frac{4}{5}''",
		{
			"type": "LaTeXEquation",
			"body": {
				"type": "AccentLiteral",
				"base": {
					"type": "FractionLiteral",
					"up": {
						"type": "NumberLiteral",
						"value": "4"
					},
					"down": {
						"type": "NumberLiteral",
						"value": "5"
					}
				},
				"value": "''"
			}
		},
		"Check \\frac{4}{5}''"
	);
}
window["AscMath"].accents = accents;
