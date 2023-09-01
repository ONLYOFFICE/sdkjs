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
function bracketsTests(test) {
	test(
		`(1+2)+2`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
				{
					Operator: "+",
				},
				{
					NumberLiteral: "2",
				},
			],
		},
		"Проверка скобок: (1+2)+2"
	);
	test(
		`{1+2}-X`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "{",
					close: "}",
				},
				{
					Operator: "-",
				},
				{
					CharLiteral: "X",
				},
			],
		},
		"Проверка скобок: {1+2}-X"
	);
	test(
		`[1+2]*i`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "[",
					close: "]",
				},
				{
					Operator: "*",
				},
				{
					CharLiteral: "i",
				},
			],
		},
		"Проверка скобок: [1+2]*i"
	);
	test(
		`|1+2|-89/2`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "|",
					close: "|",
				},
				{
					Operator: "-",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "89",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			],
		},
		"Проверка скобок: |1+2|-89/2"
	);
	test(
		`|1+2|-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "|",
					close: "|",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: |1+2|-〖89/2〗"
	);
	test(
		`⌈1+2⌉-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "⌈",
					close: "⌉",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: ⌈1+2⌉-〖89/2〗"
	);
	test(
		`⌊1+2⌋-〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
					open: "⌊",
					close: "⌋",
				},
				{
					Operator: "-",
				},
				{
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			],
		},
		"Проверка скобок: ⌊1+2⌋-〖89/2〗"
	);
	test(
		`〖89/2〗/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка скобок:〖89/2〗/2"
	);
	test(
		`√〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
			},
		},
		"Проверка скобок:√〖89/2〗"
	);
	test(
		`〖89/2〗_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка скобок:〖89/2〗_2"
	);
	test(
		`〖89/2〗^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "89",
						},
						opOver: "/",
						down: {
							NumberLiteral: "2",
						},
					},
					open: "〖",
					close: "〗",
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Проверка скобок:〖89/2〗^2"
	);
	test(
		`2_〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2_〖89/2〗"
	);
	test(
		`2^〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2^〖89/2〗"
	);

	test(
		`2_〖89/2〗_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubscript",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2_〖89/2〗_2"
	);
	test(
		`2^〖89/2〗^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expSuperscript",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						up: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2^〖89/2〗^2"
	);

	test(
		`2┴〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2┴〖89/2〗"
	);
	test(
		`2┴〖89/2〗┴2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expAbove",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						up: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2┴〖89/2〗┴2"
	);
	test(
		`2┬〖89/2〗`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								NumberLiteral: "89",
							},
							opOver: "/",
							down: {
								NumberLiteral: "2",
							},
						},
						open: "〖",
						close: "〗",
					},
				},
			},
		},
		"Проверка скобок: 2┬〖89/2〗"
	);
	test(
		`2┬〖89/2〗┬2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBelow",
						base: {
							type: "expBracketLiteral",
							exp: {
								type: "fractionLiteral",
								up: {
									NumberLiteral: "89",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
							open: "〖",
							close: "〗",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
					},
				},
			},
		},
		"Проверка скобок: 2┬〖89/2〗┬2"
	);
	test(
		"├]a+b┤[",
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				open: "]",
				close: "[",
				value: [
					{
						type: "CharLiteral",
						value: "a"
					},
					{
						type: "OperatorLiteral",
						value: "+"
					},
					{
						type: "CharLiteral",
						value: "b"
					}
				]
			}
		},
		"Check: ├]a+b┤["
	)
}

window["AscMath"].bracket = bracketsTests;
