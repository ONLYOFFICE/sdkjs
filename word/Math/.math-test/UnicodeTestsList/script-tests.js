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
function scriptTests(test) {
	test(
		`2^2 + 2`,
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
						operand: {
							NumberLiteral: "2",
						},
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
					NumberLiteral: "2",
				},
			],
		},
		"Check script/index 2^2 + 2"
	);
	test(
		`x^2+2`,
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
		"Check script/index x^2+2"
	);
	test(
		`x^(256+34)*y`,
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
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "256",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "34",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
				{
					Operator: "*",
				},
				{
					CharLiteral: "y",
				},
			],
		},
		"Check script/index: x^(256+34)*y"
	);
	test(
		`(x+34)^(256+34)-y/x`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "x",
							},
							{
								Operator: "+",
							},
							{
								NumberLiteral: "34",
							},
						],
						open: "(",
						close: ")",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									NumberLiteral: "256",
								},
								{
									Operator: "+",
								},
								{
									NumberLiteral: "34",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
				{
					Operator: "-",
				},
				{
					type: "fractionLiteral",
					up: {
						CharLiteral: "y",
					},
					opOver: "/",
					down: {
						CharLiteral: "x",
					},
				},
			],
		},
		"Check script/index: (x+34)^(256+34)-y/x"
	);
	test(
		`𝛿_(𝜇 + 𝜈)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					type: "anOther",
					value: "𝛿",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								type: "anOther",
								value: "𝜇",
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
								type: "anOther",
								value: "𝜈",
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Check script/index: 𝛿_(𝜇 + 𝜈)"
	);
	test(
		`a_b_c`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					CharLiteral: "a",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubscript",
						base: {
							CharLiteral: "b",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								CharLiteral: "c",
							},
						},
					},
				},
			},
		},
		"Check script/index: a_b_c"
	);
	test(
		`1_2_3`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "1",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubscript",
						base: [
							{
								NumberLiteral: "2",
							},
						],
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "3",
							},
						},
					},
				},
			},
		},
		"Check script/index: 1_2_3"
	);

	test(
		`A^5b^i`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSuperscript",
				base: {
					CharLiteral: "A",
				},
				up: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "5",
						},
						{
							type: "expSuperscript",
							base: {
								CharLiteral: "b",
							},
							up: {
								type: "soperandLiteral",
								operand: {
									CharLiteral: "i",
								},
							},
						},
					],
				},
			},
		},
		"Check script/index: A^5b^i"
	);
	test(
		`a_b_c^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					CharLiteral: "a",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubsup",
						base: {
							CharLiteral: "b",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								CharLiteral: "c",
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
		},
		"Check script/index: a_b_c^2"
	);

	test(
		`a_b_c^2^2^2^2^2^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					CharLiteral: "a",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubsup",
						base: {
							CharLiteral: "b",
						},
						down: {
							type: "soperandLiteral",
							operand: {
								CharLiteral: "c",
							},
						},
						up: {
							type: "soperandLiteral",
							operand: {
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
										base: [
											{
												NumberLiteral: "2",
											},
										],
										up: {
											type: "soperandLiteral",
											operand: {
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
														base: [
															{
																NumberLiteral:
																	"2",
															},
														],
														up: {
															type: "soperandLiteral",
															operand: {
																type: "expSuperscript",
																base: [
																	{
																		NumberLiteral:
																			"2",
																	},
																],
																up: {
																	type: "soperandLiteral",
																	operand: {
																		NumberLiteral:
																			"2",
																	},
																},
															},
														},
													},
												},
											},
										},
									},
								},
							},
						},
					},
				},
			},
		},
		"Check script/index: a_b_c^2^2^2^2^2^2"
	);

	test(
		`1_2_3^2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: [
					{
						NumberLiteral: "1",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubsup",
						base: [
							{
								NumberLiteral: "2",
							},
						],
						down: {
							type: "soperandLiteral",
							operand: {
								NumberLiteral: "3",
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
		},
		"Check script/index: 1_2_3^2"
	);

	test(
		`a_(b_c)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubscript",
				base: {
					CharLiteral: "a",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "expSubscript",
							base: {
								CharLiteral: "b",
							},
							down: {
								type: "soperandLiteral",
								operand: {
									CharLiteral: "c",
								},
							},
						},
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Check script/index: a_(b_c)"
	);

	test(
		`a^b_c`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: {
					CharLiteral: "a",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "c",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "b",
					},
				},
			},
		},
		"Check script/index: a^b_c"
	);
	test(
		`sin^2 x`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "sin",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					CharLiteral: "x",
				},
			],
		},
		"Check script/index: 'sin^2 x'"
	);
	test(
		`𝑊^3𝛽_𝛿_1𝜌_1𝜎_2`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: {
					CharLiteral: "𝑊",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expSubscript",
						base: {
							type: "anOther",
							value: "𝛿",
						},
						down: {
							type: "soperandLiteral",
							operand: [
								{
									NumberLiteral: "1",
								},
								{
									type: "expSubscript",
									base: {
										type: "anOther",
										value: "𝜌",
									},
									down: {
										type: "soperandLiteral",
										operand: [
											{
												NumberLiteral: "1",
											},
											{
												type: "expSubscript",
												base: {
													type: "anOther",
													value: "𝜎",
												},
												down: {
													type: "soperandLiteral",
													operand: {
														NumberLiteral: "2",
													},
												},
											},
										],
									},
								},
							],
						},
					},
				},
				up: {
					type: "soperandLiteral",
					operand: [
						{
							NumberLiteral: "3",
						},
						{
							type: "anOther",
							value: "𝛽",
						},
					],
				},
			},
		},
		"Check script/index: '𝑊^3𝛽_𝛿_1𝜌_1𝜎_2'"
	);
	test(
		`(_23^4)45`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: [
					{
						NumberLiteral: "45",
					},
				],
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "23",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "4",
					},
				},
			},
		},
		"Check script/index: '(_23^4)45'"
	);
	test(
		`(_x^y)45`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: [
					{
						NumberLiteral: "45",
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
						CharLiteral: "y",
					},
				},
			},
		},
		"Check script/index: '(_x^y)45'"
	);
	test(
		`(_x^y)zyu`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					CharLiteral: "zyu",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "x",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						CharLiteral: "y",
					},
				},
			},
		},
		"Check script/index: (_x^y)zyu"
	);
	test(
		`(_453^56)zyu`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					CharLiteral: "zyu",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "453",
					},
				},
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "56",
					},
				},
			},
		},
		"Check script/index: (_453^56)zyu"
	);
	test(
		`(_(453+2)^56)zyu`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					CharLiteral: "zyu",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								NumberLiteral: "453",
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
				up: {
					type: "soperandLiteral",
					operand: {
						NumberLiteral: "56",
					},
				},
			},
		},
		"Check script/index: '(_(453+2)^56)zyu'"
	);
	test(
		`(_(453+2)^(345432+y+x/z))zyu`,
		{
			type: "UnicodeEquation",
			body: {
				type: "prescriptSubsup",
				base: {
					CharLiteral: "zyu",
				},
				down: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								NumberLiteral: "453",
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
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								NumberLiteral: "345432",
							},
							{
								Operator: "+",
							},
							{
								CharLiteral: "y",
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
									CharLiteral: "z",
								},
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Check script/index: '(_(453+2)^(345432+y+x/z))zyu'"
	);
}
window["AscMath"].script = scriptTests;
