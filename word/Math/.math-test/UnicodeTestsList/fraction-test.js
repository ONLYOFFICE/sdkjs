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
function fractionTests(test) {
	test(
		`1/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					NumberLiteral: "1",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: 1/2"
	);
	test(
		`x/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					CharLiteral: "x",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: x/2"
	);
	test(
		`x+5/2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "x",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "5",
					},
					opOver: "/",
					down: {
						NumberLiteral: "2",
					},
				},
			],
		},
		"Проверка работы базового деления: x+5/2"
	);
	test(
		`x+5/x+2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "x",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "5",
					},
					opOver: "/",
					down: {
						CharLiteral: "x",
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
		"Проверка работы базового деления: x+5/x+2"
	);
	test(
		`1∕2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					NumberLiteral: "1",
				},
				opOver: "∕",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы базового деления: 1∕2"
	);
	test(
		`(x+5)/2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "5",
						},
					],
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы деления: (x+5)/2"
	);
	test(
		`x/(2+1)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					CharLiteral: "x",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "1",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: x/(2+1)"
	);
	test(
		`(x-5)/(2+1)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "-",
						},
						{
							NumberLiteral: "5",
						},
					],
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							NumberLiteral: "1",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: (x-5)/(2+1)"
	);
	test(
		`1+3/2/3`,
		{
			type: "UnicodeEquation",
			body: [
				{
					NumberLiteral: "1",
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						NumberLiteral: "3",
					},
					opOver: "/",
					down: {
						type: "fractionLiteral",
						up: {
							NumberLiteral: "2",
						},
						opOver: "/",
						down: {
							NumberLiteral: "3",
						},
					},
				},
			],
		},
		"Проверка работы цепи деления: 1+3/2/3"
	);
	test(
		`(𝛼_2^3)/(𝛽_2^3+𝛾_2^3)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "expSubsup",
						base: {
							type: "anOther",
							value: "𝛼",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "2",
							},
						},
						up: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "3",
							},
						},
					},
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							type: "expSubsup",
							base: {
								type: "anOther",
								value: "𝛽",
							},
							down: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "3",
								},
							},
						},
						{
							Operator: "+",
						},
						{
							type: "expSubsup",
							base: {
								type: "anOther",
								value: "𝛾",
							},
							down: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "3",
								},
							},
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы деления: (𝛼_2^3)/(𝛽_2^3+𝛾_2^3)"
	);

	test(
		`(a/(b+c))/(d/e + f)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expBracketLiteral",
					exp: {
						type: "fractionLiteral",
						up: {
							CharLiteral: "a",
						},
						opOver: "/",
						down: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "b",
								},
								{
									Operator: "+",
								},
								{
									CharLiteral: "c",
								},
							],
							open: "(",
							close: ")",
						},
					},
					open: "(",
					close: ")",
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
						{
							type: "fractionLiteral",
							up: {
								CharLiteral: "d",
							},
							opOver: "/",
							down: {
								CharLiteral: "e",
							},
						},
						{
							type: "SpaceLiteral",
							value: " ",
						},
						{
							Operator: "+",
						},
						{
							type: "SpaceLiteral",
							value: " ",
						},
						{
							CharLiteral: "f",
						},
					],
					open: "(",
					close: ")",
				},
			},
		},
		"Проверка работы вложенного деления: (a/(b+c))/(d/e + f)"
	);

	test(
		`(a/(c/(z/x)))`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "fractionLiteral",
					up: {
						CharLiteral: "a",
					},
					opOver: "/",
					down: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								CharLiteral: "c",
							},
							opOver: "/",
							down: {
								type: "expBracketLiteral",
								exp: {
									type: "fractionLiteral",
									up: {
										CharLiteral: "z",
									},
									opOver: "/",
									down: {
										CharLiteral: "x",
									},
								},
								open: "(",
								close: ")",
							},
						},
						open: "(",
						close: ")",
					},
				},
				open: "(",
				close: ")",
			},
		},
		"Проверка работы деления: (a/(c/(z/x)))"
	);

	test(
		`1¦2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "binomLiteral",
				numerator: {
					NumberLiteral: "1",
				},
				operand: {
					NumberLiteral: "2",
				},
			},
		},
		"Проверка работы деления: 1¦2"
	);
	test(
		`(1¦2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBracketLiteral",
				exp: {
					type: "binomLiteral",
					numerator: {
						NumberLiteral: "1",
					},
					operand: {
						NumberLiteral: "2",
					},
				},
				open: "(",
				close: ")",
			},
		},
		"Проверка работы деления: (1¦2)"
	);
}
window["AscMath"].fraction = fractionTests;
