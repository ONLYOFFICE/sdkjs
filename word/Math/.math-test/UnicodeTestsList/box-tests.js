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
function boxTests(test) {
	test(
		`□(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка box: □(1+2)"
	);
	test(
		`□(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка box: □(1+2)"
	);
	test(
		`□1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка box: □1"
	);
	test(
		`□1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "boxLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка box: □1/2"
	);
	test(
		`▭(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка rect: ▭(1+2)"
	);
	test(
		`▭1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка rect: ▭1"
	);
	test(
		`▭1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "rectLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка rect: ▭1/2"
	);
	test(
		`▁(1+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка underbar: ▁(1+2)"
	);

	test(
		`▁1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка underbar: ▁1"
	);
	test(
		`▁1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "underbarLiteral",
					value: {
						NumberLiteral: "1",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка underbar: ▁1/2"
	);
	test(
		` ̄(1+2)`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "1",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "2",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка underbar:  ̄(1+2)"
	);

	test(
		` ̄1`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
					NumberLiteral: "1",
				},
			},
		},
		"Проверка underbar:  ̄1"
	);
	test(
		`(1+2)̂`.trim(),
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "DiacriticLiteral",
				"base": {
					"type": "expBracketLiteral",
					"exp": [
						[
							{
								"type": "NumberLiteral",
								"value": "1"
							}
						],
						{
							"type": "OperatorLiteral",
							"value": "+"
						},
						[
							{
								"type": "NumberLiteral",
								"value": "2"
							}
						]
					],
					"open": "(",
					"close": ")"
				},
				"value": "̂"
			}
		},
		"Проверка underbar: (1+2)̂"
	);
}

window["AscMath"].box = boxTests;

