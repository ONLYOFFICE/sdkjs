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
function sqrtTests(test) {
	test(
		`√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check √5"
	);
	test(
		`√a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check √a"
	);
	test(
		`√a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Check √a/2"
	);
	test(
		`√(2&a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					NumberLiteral: "2",
				},
				content: [
					{
						CharLiteral: "a",
					},
					{
						Operator: "-",
					},
					{
						NumberLiteral: "4",
					},
				],
			},
		},
		"Check √(2&a-4)"
	);
	test(
		`∛5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check ∛5"
	);
	test(
		`∛a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check ∛a"
	);
	test(
		`∛a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "cubertLiteral",
					value: {
						CharLiteral: "a",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Check ∛a/2"
	);
	test(
		`∛(a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "cubertLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "a",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "4",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Check ∛(a-4)"
	);
	test(
		`∜5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					NumberLiteral: "5",
				},
			},
		},
		"Check ∜5"
	);
	test(
		`∜a`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					CharLiteral: "a",
				},
			},
		},
		"Check ∜a"
	);
	test(
		`∜a/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "fourthrtLiteral",
					value: {
						CharLiteral: "a",
					},
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Check ∜a/2"
	);
	test(
		`∜(a-4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fourthrtLiteral",
				value: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "a",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "4",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Check ∜(a-4)"
	);
	test(
		`√(10&a/4)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					NumberLiteral: "10",
				},
				content: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						NumberLiteral: "4",
					},
				},
			},
		},
		"Check √(10&a/4)"
	);
	test(
		`√(10^2&a/4+2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "nthrtLiteral",
				index: {
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "10",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				content: [
					{
						type: "fractionLiteral",
						up: {
							CharLiteral: "a",
						},
						opOver: "/",
						down: {
							NumberLiteral: "4",
						},
					},
					{
						Operator: "+",
					},
					{
						NumberLiteral: "2",
					},
				],
			},
		},
		"Check √(10^2&a/4+2)"
	);
	test(
		`√5^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			},
		},
		"Check √5^2"
	);
	test(
		`√5_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubscript",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			},
		},
		"Check √5_2"
	);
	test(
		`√5^2_x`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "x",
						},
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
		"Check √5^2_x"
	);
	test(
		`√5_2^x`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "x",
						},
					},
				},
			},
		},
		"Check √5_2^x"
	);
	test(
		`(_5^2)√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "5",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "2",
					},
				},
			},
		},
		"Check (_5^2)√5"
	);
	test(
		`√5┴exp1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expAbove",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "exp",
							},
							{
								NumberLiteral: "1",
							},
						],
					},
				},
			},
		},
		"Check √5┴exp1"
	);
	test(
		`√5┬exp1`,
		{
			type: "UnicodeEquation",
			body: {
				type: "sqrtLiteral",
				value: {
					type: "expBelow",
					base: [
						{
							NumberLiteral: "5",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "exp",
							},
							{
								NumberLiteral: "1",
							},
						],
					},
				},
			},
		},
		"Check √5┬exp1"
	);
	test(
		`(√5┬exp1]`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "sqrtLiteral",
					value: {
						type: "expBelow",
						base: [
							{
								NumberLiteral: "5",
							},
						],
						down: {
							type: "soperandLiteral",
							operand: [
								{
									CharLiteral: "exp",
								},
								{
									NumberLiteral: "1",
								},
							],
						},
					},
				},
				open: "(",
				close: "]",
			},
		},
		"Check (√5┬exp1]"
	);
	test(
		`□√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "boxLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check □√5"
	);
	test(
		`▭√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "rectLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ▭√5"
	);
	test(
		`▁√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "underbarLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ▁√5"
	);
	test(
		` ̄√5`.trim(),
		{
			type: "UnicodeEquation",
			body: {
				type: "overbarLiteral",
				value: {
					type: "sqrtLiteral",
					value: {
						NumberLiteral: "5",
					},
				},
			},
		},
		"Check ̄√5"
	);
	test(
		`∑_√5^√5`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					type: "opNary",
					value: "∑",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "sqrtLiteral",
						value: {
							type: "expSuperscript",
							base: [
								{
									NumberLiteral: "5",
								},
							],
							up: {
								type: "soperandLiteral",
								operand: {
									type: "sqrtLiteral",
									value: {
										NumberLiteral: "5",
									},
								},
							},
						},
					},
				},
			},
		},
		"Check ∑_√5^√5"
	);
	// test(
	// 	`\\root n+1\\of(b+c)+x`,
	// 	{},
	// 	"Check \\root n+1\\of(b+c)+x"
	// );
}
window["AscMath"].sqrt = sqrtTests;
