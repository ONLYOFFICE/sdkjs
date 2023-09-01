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
function specialTest(test) {
	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹`,
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
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2⁰¹²³⁴⁵⁶⁷⁸⁹`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
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
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "x",
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubscript",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
				},
				{
					NumberLiteral: "56",
				},
			],
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎56`"
	);
	test(
		`z₂₃₄₊₍₆₇₋₀₌₆₇₎56`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubscript",
					base: {
						CharLiteral: "z",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
				},
				{
					NumberLiteral: "56",
				},
			],
		},
		"Проверка `z₂₃₄₊₍₆₇₋₀₌₆₇₎56`"
	);

	test(
		`2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2⁰¹²³⁴⁵⁶⁷⁸⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎`"
	);
	test(
		`2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`"
	);
	test(
		`x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: {
						CharLiteral: "x",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹₂₃₄₊₍₆₇₋₀₌₆₇₎+45`"
	);

	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "0123456789",
					},
				},
			},
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁰¹²³⁴⁵⁶⁷⁸⁹`"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: [
					{
						NumberLiteral: "2",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "234",
						},
						{
							Operator: "+",
						},
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "67",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "0",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "67",
								},
							],
							open: "(",
							close: ")",
						},
					],
				},
				up: {
					type: "soperandLiteral",
					operand: [
						[
							{
								NumberLiteral: "4",
							},
							{
								CharLiteral: "in",
							},
						],
						{
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "5",
								},
								{
									Operator: "-",
								},
								{
									NumberLiteral: "6",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "7",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "8",
								},
							],
							open: "(",
							close: ")",
						},
						{
							NumberLiteral: "9",
						},
					],
				},
			},
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹`"
	);
	test(
		`2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `2₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
	test(
		`x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSubsup",
					base: {
						CharLiteral: "x",
					},
					down: {
						type: "soperandLiteral",
						operand: [
							{
								NumberLiteral: "234",
							},
							{
								Operator: "+",
							},
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "67",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "0",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "67",
									},
								],
								open: "(",
								close: ")",
							},
						],
					},
					up: {
						type: "soperandLiteral",
						operand: [
							[
								{
									NumberLiteral: "4",
								},
								{
									CharLiteral: "in",
								},
							],
							{
								type: "expBracketLiteral",
								exp: [
									{
										NumberLiteral: "5",
									},
									{
										Operator: "-",
									},
									{
										NumberLiteral: "6",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "7",
									},
									{
										Operator: "=",
									},
									{
										NumberLiteral: "8",
									},
								],
								open: "(",
								close: ")",
							},
							{
								NumberLiteral: "9",
							},
						],
					},
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "45",
				},
			],
		},
		"Проверка `x₂₃₄₊₍₆₇₋₀₌₆₇₎⁴ⁱⁿ⁽⁵⁻⁶⁺⁷⁼⁸⁾⁹+45`"
	);
}
window["AscMath"].special = specialTest;
