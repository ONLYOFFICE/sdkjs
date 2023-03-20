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
function testsAboveBelow(test) {
	test(
		`base┴2+2`,
		{
			body: [
				{
					base: {
						type: "CharLiteral",
						value: "base"
					},
					down: undefined,
					type: "expAbove",
					up: [
						{
							type: "NumberLiteral",
							value: "2"
						}
					]
				},
				{
					type: "OperatorLiteral",
					value: "+"
				},
				[
					{
						type: "NumberLiteral",
						value: "2"
					}
				]
			],
			type: "UnicodeEquation"
		},
		"Check base┴2+2"
	);
	test(
		`base┴2┴x+2`,
		{
			"body": [
				{
					"base": {
						"type": "CharLiteral",
						"value": "base"
					},
					"down": undefined,
					"type": "expAbove",
					"up": {
						"base": [
							{
								"type": "NumberLiteral",
								"value": "2"
							}
						],
						"down": undefined,
						"type": "expAbove",
						"up": {
							"type": "CharLiteral",
							"value": "x"
						}
					}
				},
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
			"type": "UnicodeEquation"
		},
		"Check base┴2┴x+2"
	);
	test(
		`base┴2┴(x/y+6)+2`,
		{
			"body": [
				{
					"base": {
						"type": "CharLiteral",
						"value": "base"
					},
					"down": undefined,
					"type": "expAbove",
					"up": {
						"base": [
							{
								"type": "NumberLiteral",
								"value": "2"
							}
						],
						"down": undefined,
						"type": "expAbove",
						"up": {
							"close": ")",
							"exp": [
								{
									"down": {
										"type": "CharLiteral",
										"value": "y"
									},
									"type": "FractionLiteral",
									"up": {
										"type": "CharLiteral",
										"value": "x"
									}
								},
								{
									"type": "OperatorLiteral",
									"value": "+"
								},
									[
									{
										"type": "NumberLiteral",
										"value": "6"
									}
									]
							],
							"open": "(",
							"type": "expBracketLiteral"
						}
					}
				},
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
			"type": "UnicodeEquation"
		},
		"Check base┴2┴(x/y+6)+2"
	);
	test(
		`x^23┴2/y`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expSuperscript",
					base: {
						CharLiteral: "x",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							type: "expAbove",
							base: [
								{
									NumberLiteral: "23",
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
				opOver: "/",
				down: {
					CharLiteral: "y",
				},
			},
		},
		"Check x^23┴2/y"
	);
	test(
		`(x^23)┴2/y`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expAbove",
					base: {
						type: "expBracketLiteral",
						exp: {
							type: "expSuperscript",
							base: {
								CharLiteral: "x",
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "23",
								},
							},
						},
						open: "(",
						close: ")",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				opOver: "/",
				down: {
					CharLiteral: "y",
				},
			},
		},
		"Check (x^23)┴2/y"
	);

	test(
		`4┴2+2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expAbove",
					base: [
						{
							NumberLiteral: "4",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
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
		"Check 4┴2+2"
	);
	test(
		`base┴exp*xz`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expAbove",
					base: {
						CharLiteral: "base",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "exp",
						},
					},
				},
				{
					Operator: "*",
				},
				{
					CharLiteral: "xz",
				},
			],
		},
		"Check base┴exp*xz"
	);
	test(
		`2┴exp-p`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expAbove",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "exp",
						},
					},
				},
				{
					Operator: "-",
				},
				{
					CharLiteral: "p",
				},
			],
		},
		"Check 2┴exp-p"
	);
	test(
		`base┬2*x`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: {
						CharLiteral: "base",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "*",
				},
				{
					CharLiteral: "x",
				},
			],
		},
		"Check base┬2*x"
	);
	test(
		`4┬2+x/y`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: [
						{
							NumberLiteral: "4",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "+",
				},
				{
					type: "fractionLiteral",
					up: {
						CharLiteral: "x",
					},
					opOver: "/",
					down: {
						CharLiteral: "y",
					},
				},
			],
		},
		"Check 4┬2+x/y"
	);
	test(
		`base┬exp*x^2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: {
						CharLiteral: "base",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "exp",
						},
					},
				},
				{
					Operator: "*",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "x",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
			],
		},
		"Check base┬exp*x^2"
	);
	test(
		`2┬exp-x_i`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "exp",
						},
					},
				},
				{
					Operator: "-",
				},
				{
					type: "expSubscript",
					base: {
						CharLiteral: "x",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "i",
						},
					},
				},
			],
		},
		"Check 2┬exp-x_i"
	);
	test(
		`2┬(exp+2)+(2+1)`,
		{
			type: "UnicodeEquation",
			body: [
				{
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
							exp: [
								{
									CharLiteral: "exp",
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
				{
					Operator: "+",
				},
				{
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
			],
		},
		"Check 2┬(exp+2)+(2+1)"
	);
	test(
		`2┬(exp+2+x/2)^2 - 1`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: [
						{
							NumberLiteral: "2",
						},
					],
					down: {
						type: "soperandLiteral",
						operand: {
							type: "expSuperscript",
							base: {
								type: "expBracketLiteral",
								exp: [
									{
										CharLiteral: "exp",
									},
									{
										Operator: "+",
									},
									{
										NumberLiteral: "2",
									},
									{
										Operator: "+",
									},
									{
										type: "fractionLiteral",
										up: {
											CharLiteral: "x",
										},
										opOver: "/",
										down: {
											NumberLiteral: "2",
										},
									},
								],
								open: "(",
								close: ")",
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
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "-",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					NumberLiteral: "1",
				},
			],
		},
		"Check 2┬(exp+2+x/2)^2 - 1"
	);
	test(
		`(2+x)┬exp`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "x",
						},
					],
					open: "(",
					close: ")",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "exp",
					},
				},
			},
		},
		"Check `(2+x)┬exp`"
	);
	test(
		`(2+y)┬(exp+2+x/2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "y",
						},
					],
					open: "(",
					close: ")",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "exp",
							},
							{
								Operator: "+",
							},
							{
								NumberLiteral: "2",
							},
							{
								Operator: "+",
							},
							{
								type: "fractionLiteral",
								up: {
									CharLiteral: "x",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Check `(2+y)┬(exp+2+x/2)`"
	);
	test(
		`(2+y^2)┬(exp_3+2+x/2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expBelow",
				base: {
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "2",
						},
						{
							Operator: "+",
						},
						{
							type: "expSuperscript",
							base: {
								CharLiteral: "y",
							},
							up: {
								type: "soperandLiteral",
								operand: {
									NumberLiteral: "2",
								},
							},
						},
					],
					open: "(",
					close: ")",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								type: "expSubscript",
								base: {
									CharLiteral: "exp",
								},
								down: {
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
								NumberLiteral: "2",
							},
							{
								Operator: "+",
							},
							{
								type: "fractionLiteral",
								up: {
									CharLiteral: "x",
								},
								opOver: "/",
								down: {
									NumberLiteral: "2",
								},
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Check `(2+y^2)┬(exp_3+2+x/2)`"
	);
}
window["AscMath"].aboveBelow = testsAboveBelow;
