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
function literalTests(test) {

	test(
		"×",
		{
			type: "UnicodeEquation",
			body: {
				type: "OperatorLiteral",
				value: "×"
			}
		},
		"Check operator: ×"
	);
	test(
		"⋅",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⋅"
			}
		},
		"Check operator: ⋅"
	);
	test(
		"∈",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∈"
			}
		},
		"Check operator: ∈"
	);
	test(
		"∋",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∋"
			}
		},
		"Check operator: ∋"
	);
	test(
		"∼",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∼"
			}
		},
		"Check operator: ∼"
	);
	test(
		"≃",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≃"
			}
		},
		"Check operator: ≃"
	);
	test(
		"≅",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≅"
			}
		},
		"Check operator: ≅"
	);
	test(
		"≈",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≈"
			}
		},
		"Check operator: ≈"
	);
	test(
		"≍",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≍"
			}
		},
		"Check operator: ≍"
	);
	test(
		"≡",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≡"
			}
		},
		"Check operator: ≡"
	);
	test(
		"≤",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≤"
			}
		},
		"Check operator: ≤"
	);
	test(
		"≥",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≥"
			}
		},
		"Check operator: ≥"
	);
	test(
		"≶",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≶"
			}
		},
		"Check operator: ≶"
	);
	test(
		"≷",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≷"
			}
		},
		"Check operator: ≷"
	);
	test(
		"≽",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≽"
			}
		},
		"Check operator: ≽"
	);
	test(
		"≺",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≺"
			}
		},
		"Check operator: ≺"
	);
	test(
		"≻",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≻"
			}
		},
		"Check operator: ≻"
	);
	test(
		"≼",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "≼"
			}
		},
		"Check operator: ≼"
	);
	test(
		"⊂",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊂"
			}
		},
		"Check operator: ⊂"
	);
	test(
		"⊃",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊃"
			}
		},
		"Check operator: ⊃"
	);
	test(
		"⊆",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊆"
			}
		},
		"Check operator: ⊆"
	);
	test(
		"⊇",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊇"
			}
		},
		"Check operator: ⊇"
	);
	test(
		"⊑",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊑"
			}
		},
		"Check operator: ⊑"
	);
	test(
		"⊒",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊒"
			}
		},
		"Check operator: ⊒"
	);
	test(
		"+",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "+"
			}
		},
		"Check operator: +"
	);
	test(
		"-",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "-"
			}
		},
		"Check operator: -"
	);
	test(
		"=",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "="
			}
		},
		"Check operator: ="
	);
	test(
		"*",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "*"
			}
		},
		"Check operator: *"
	);

	test(
		"∃",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∃"
			}
		},
		"Check logic operator: ∃"
	);
	test(
		"∀",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∀"
			}
		},
		"Check logic operator: ∀"
	);
	test(
		"¬",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "¬"
			}
		},
		"Check logic operator: ¬"
	);
	test(
		"∧",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∧"
			}
		},
		"Check logic operator: ∧"
	);
	test(
		"∨",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "∨"
			}
		},
		"Check logic operator: ∨"
	);
	test(
		"⇒",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⇒"
			}
		},
		"Check logic operator: ⇒"
	);
	test(
		"⇔",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⇔"
			}
		},
		"Check logic operator: ⇔"
	);
	test(
		"⊕",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊕"
			}
		},
		"Check logic operator: ⊕"
	);
	test(
		"⊤",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊤"
			}
		},
		"Check logic operator: ⊤"
	);
	test(
		"⊥",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊥"
			}
		},
		"Check logic operator: ⊥"
	);
	test(
		"⊢",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⊢"
			}
		},
		"Check logic operator: ⊢"
	);

	test(
		"⨯",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⨯"
			}
		},
		"Check db operator: ⨯"
	);
	test(
		"⟕",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⟕"
			}
		},
		"Check db operator: ⟕"
	);
	test(
		"⟖",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⟖"
			}
		},
		"Check db operator: ⟖"
	);
	test(
		"⟗",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⟗"
			}
		},
		"Check db operator: ⟗"
	);
	test(
		"⋉",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⋉"
			}
		},
		"Check db operator: ⋉"
	);
	test(
		"⋊",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⋊"
			}
		},
		"Check db operator: ⋊"
	);
	test(
		"▷",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "▷"
			}
		},
		"Check db operator: ▷"
	);
	test(
		"÷",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "÷"
			}
		},
		"Check db operator: ÷"
	);

	test(
		"⁡",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⁡"
			}
		},
		"Check invisible function application operator: ⁡"
	);
	test(
		"⁢",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⁢"
			}
		},
		"Check invisible times operator: ⁢"
	);
	test(
		"⁣",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⁣"
			}
		},
		"Check invisible separator operator: ⁣"
	);
	test(
		"⁤",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "OperatorLiteral",
				"value": "⁤"
			}
		},
		"Check invisible plus operator: ⁤"
	);
	test(
		"​",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": "​"
			}
		},
		"Check zero-width space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 1/18em space (very very thin math space)"
	);
	test(
		"  ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": "  ",
			}
		},
		"Check 2/18em space (very thin math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 3/18em space (thin math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 5/18em space (thick math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 6/18em space (very thick math space)"
	);
	test(
		"  ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": "  ",
			}
		},
		"Check 7/18em space (very very thick math space)"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 9/18em space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check 1em space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check digit-width space"
	);
	test(
		" ",
		{
			"type": "UnicodeEquation",
			"body": {
				"type": "SpaceLiteral",
				"value": " ",
			}
		},
		"Check space-with space (non-breaking space)"
	);

	test(
		`a`,
		{
			"body": {
				"data": "a",
				"type": "CharLiteral",
			},
			"type": "UnicodeEquation"
		},
		"Check: a"
	);
	test(
		`abcdef`,
		{
			"body": {
				"data": "abcdef",
				"type": "CharLiteral",
			},
			"type": "UnicodeEquation"
		},
		"Check: abcdef"
	);
	test(
		`1`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "NumberLiteral",
					data: "1"
				}
			]
		},
		"Check: 1"
	);
	test(
		`1234`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "NumberLiteral",
					data: "1234"
				}
			]
		},
		"Check: 1234"
	);
	test(
		`1+2`,
		{
			type: "UnicodeEquation",
			body: [
				[
					{
						type: "NumberLiteral",
						data: "1"
					}
				],
				{
					type: "OperatorLiteral",
					value: "+"
				},
				[
					{
						type: "NumberLiteral",
						data: "2"
					}
				]
			]
		},
		"Check: 1+2"
	);
	test(
		`1+2+3`,
		{
			type: "UnicodeEquation",
			body: [
				[
					{
						type: "NumberLiteral",
						data: "1"
					}
				],
				{
					type: "OperatorLiteral",
					value: "+"
				},
				[
					{
						type: "NumberLiteral",
						data: "2"
					}
				],
				{
					type: "OperatorLiteral",
					value: "+"
				},
				[
					{
						type: "NumberLiteral",
						data: "3"
					}
				]
			]
		},
		"Check: 1+2+3"
	);

	test(
		`ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω`,
		{
			body: {
				data: "ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω",
				type: "OtherLiteral"
			},
			type: "UnicodeEquation"
		},
		"Check greek letters: ΑαΒβΓγΔδΕεΖζΗηΘθΙιΚκΛλΜμΝνΞξΟοΠπΡρΣσΤτΥυΦφΧχΨψΩω"
	);
	test(
		"abc123def",
		{
			type: "UnicodeEquation",
			body: [
				[
					{
						CharLiteral: "abc",
					},
					{
						NumberLiteral: "123",
					},
				],
				{
					CharLiteral: "def",
				},
			],
		},
		"Проверка простого литерала: abc123def"
	);
	test(
		"abc+123+def",
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "abc",
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "123",
				},
				{
					Operator: "+",
				},
				{
					CharLiteral: "def",
				},
			],
		},
		"Проверка простого литерала: abc+123+def"
	);
	test(
		"𝐀𝐁𝐂𝐨𝐹",
		{
			type: "UnicodeEquation",
			body: {
				CharLiteral: "𝐀𝐁𝐂𝐨𝐹",
			},
		},
		"Проверка простого литерала: 𝐀𝐁𝐂𝐨𝐹"
	);

	//spaces
	test(
		"   𝐀𝐁𝐂𝐨𝐹   ",
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					CharLiteral: "𝐀𝐁𝐂𝐨𝐹",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
			],
		},
		"Проверка простого литерала - пробелы: '   𝐀𝐁𝐂𝐨𝐹   '"
	);

	//spaces & tabs
	test(
		" 	𝐀𝐁𝐂𝐨𝐹  	 ",
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: "\t",
				},
				{
					CharLiteral: "𝐀𝐁𝐂𝐨𝐹",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "SpaceLiteral",
					value: "\t",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
			],
		},
		"Проверка простого литерала - пробелы и табуляция: ' 	𝐀𝐁𝐂𝐨𝐹  	 '"
	);

	test(
		`1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5`,
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
					CharLiteral: "fbnd",
				},
				{
					Operator: "+",
				},
				{
					type: "expBracketLiteral",
					exp: [
						{
							NumberLiteral: "3",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "𝐀𝐁𝐂𝐨𝐹",
						},
					],
					open: "(",
					close: ")",
				},
				{
					Operator: "+",
				},
				{
					CharLiteral: "c",
				},
				{
					Operator: "+",
				},
				{
					NumberLiteral: "5",
				},
			],
		},
		"Проверка простого литерала - пробелы и табуляция: '1+fbnd+(3+𝐀𝐁𝐂𝐨𝐹)+c+5'"
	);

	// test(
	// 	`1/3.1416`,
	// 	{
	// 		type: "UnicodeEquation",
	// 		body: {
	// 			type: "expLiteral",
	// 			value: [
	// 				{
	// 					type: "fractionLiteral",
	// 					numerator: {
	// 						type: "numeratorLiteral",
	// 						value: [
	// 							{
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "1",
	// 									},
	// 								],
	// 							},
	// 						],
	// 					},
	// 					opOver: {
	// 						type: "opOver",
	// 						value: "/",
	// 					},
	// 					operand: [
	// 						{
	// 							type: "numberLiteral",
	// 							number: {
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "3",
	// 									},
	// 								],
	// 							},
	// 							decimal: ".",
	// 							after: {
	// 								type: "digitsLiteral",
	// 								value: [
	// 									{
	// 										type: "NumericLiteral",
	// 										value: "1416",
	// 									},
	// 								],
	// 							},
	// 						},
	// 					],
	// 				},
	// 			],
	// 		},
	// 	},
	// 	"Проверка простого литерала - пробелы и табуляция: '1/3.1416'"
	// );


	test(
		"1\\above2",
		{
			body: {
				base: [
					{
						data: "1",
						type: "NumberLiteral"
					}
				],
				down: undefined,
				type: "expAbove",
				up: {
					type: "soperandLiteral",
					value: [
						{
							data: "2",
							type: "NumberLiteral"
						}
					]
				}
			},
			type: "UnicodeEquation"
		},
		"Check: 1\\above2"
	)
	test(
		"\\\\above",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "┴"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\above"
	)
	test(
		"1\\acute2",
		{
			"type": "UnicodeEquation",
			"body": [
				{
					"type": "atomLiteral",
					"base": {
						"type": "DiacriticBaseLiteral",
						"data": [
							{
								"type": "NumberLiteral",
								"data": "1"
							}
						],
						"isAn": true
					},
					"diacritic": {
						"type": "DiacriticLiteral",
						"value": "́"
					}
				},
				[
					{
						"type": "NumberLiteral",
						"data": "2"
					}
				]
			]
		},
		"Check: 1\\acute2"
	)
	test(
		"\\\\acute",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "́"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\acute"
	)

	test(
		"\\\\aleph",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ℵ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\aleph"
	)
	test(
		"\\\\alpha",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "α"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\alpha"
	)
	test(
		"\\\\amalg",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∐"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\amalg"
	);
	test(
		"\\\\angle",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∠"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\angle"
	)
	test(
		"\\\\aoint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∳"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\aoint"
	)
	test(
		"\\\\approx",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≈"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\approx"
	)
	test(
		"\\\\asmash",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⬆"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\asmash"
	)
	test(
		"\\\\ast",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∗"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ast"
	)
	test(
		"\\\\asymp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≍"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\asymp"
	)
	test(
		"\\\\atop",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "¦"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\atop"
	)


	test(
		"\\\\Bar",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̿"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Bar"
	)
	test(
		"\\\\bar",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̅"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bar"
	)
	test(
		"\\\\because",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∵"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\because"
	)
	test(
		"\\\\begin",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "〖"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\begin"
	)
	test(
		"\\\\below",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "┬"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\below"
	)
	test(
		"\\\\beta",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "β"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\beta"
	)
	test(
		"\\\\beth",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ℶ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\beth"
	)
	test(
		"\\\\bot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⊥"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bot"
	)

	test(
		"\\\\bigcap",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋂"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigcap"
	)
	test(
		"\\\\bigcup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋂"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigcup"
	)
	test(
		"\\\\bigodot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⨀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigodot"
	)

	test(
		"\\\\bigoplus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⨁"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigoplus"
	)
	test(
		"\\\\bigotimes",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⨂"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigotimes"
	)
	test(
		"\\\\bigsqcup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⨆"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigsqcup"
	)
	test(
		"\\\\biguplus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⨄"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\biguplus"
	)
	test(
		"\\\\bigvee",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋁"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigvee"
	)
	test(
		"\\\\bigwedge",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bigwedge"
	)
	test(
		"\\\\bowtie",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋈"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bowtie"
	)
	test(
		"\\\\box",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "□"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\box"
	)
	test(
		"\\\\bra",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⟨"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bra"
	)
	test(
		"\\\\breve",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̆"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\breve"
	)
	test(
		"\\\\bullet",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∙"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\bullet"
	)
	test(
		"\\\\boxdot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⊡"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\boxdot"
	)
	test(
		"\\\\boxminus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⊟"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\boxminus"
	)
	test(
		"\\\\boxplus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⊞"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\boxplus"
	)
	test(
		"\\\\cap",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∩"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cap"
	)
	test(
		"\\\\cbrt",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∛"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cbrt"
	)
	test(
		"\\\\cdots",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋯"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cdots"
	)
	test(
		"\\\\cdot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋅"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cdot"
	)
	test(
		"\\\\check",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̌"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\check"
	)
	test(
		"\\\\chi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "χ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\chi"
	)
	test(
		"\\\\circ",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∘"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\circ"
	)
	test(
		"\\\\close",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "┤"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\close"
	)
	test(
		"\\\\clubsuit",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "♣"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\clubsuit"
	)
	test(
		"\\\\coint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∲"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\coint"
	)
	test(
		"\\\\cong",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≅"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cong"
	)
	test(
		"\\\\contain",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∋"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\contain"
	)
	test(
		"\\\\cup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∪"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\cup"
	)


	test(
		"\\\\daleth",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ℸ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\daleth"
	)
	test(
		"\\\\dashv",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⊣"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dashv"
	)
	test(
		"\\\\dd",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ⅆ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dd"
	)
	test(
		"\\\\ddddot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⃜"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ddddot"
	)
	test(
		"\\\\dddot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⃛"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dddot"
	)
	test(
		"\\\\ddot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̈"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ddot"
	)
	test(
		"\\\\ddots",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋱"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ddots"
	)
	test(
		"\\\\degree",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "°"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\degree"
	)
	test(
		"\\\\Delta",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "Δ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Delta"
	)
	test(
		"\\\\delta",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "δ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\delta"
	)
	test(
		"\\\\diamond",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⋄"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\diamond"
	)

	test(
		"\\\\diamondsuit",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "♢"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\diamondsuit"
	)
	test(
		"\\\\div",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "÷"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\div"
	)
	test(
		"\\\\dot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̇"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dot"
	)
	test(
		"\\\\doteq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≐"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\doteq"
	)
	test(
		"\\\\dots",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "…"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dots"
	)
	test(
		"\\\\downarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "↓"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\downarrow"
	)
	test(
		"\\\\dsmash",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⬇"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\dsmash"
	)

	test(
		"\\\\degc",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "℃"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\degc"
	)
	test(
		"\\\\degf",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "℉"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\degf"
	)


	test(
		"\\\\ee",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ⅇ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ee"
	)
	test(
		"\\\\ell",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ℓ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ell"
	)
	test(
		"\\\\emptyset",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∅"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\emptyset"
	)
	test(
		"\\\\emsp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": " "
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\emsp"
	)
	test(
		"\\\\end",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "〗"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\end"
	)
	test(
		"\\\\ensp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": " "
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ensp"
	)
	test(
		"\\\\epsilon",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ϵ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\epsilon"
	)
	test(
		"\\\\eqarray",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "█"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\eqarray"
	)
	test(
		"\\\\eqno",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "#"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\eqno"
	)
	test(
		"\\\\equiv",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≡"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\equiv"
	)
	test(
		"\\\\eta",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "η"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\eta"
	)
	test(
		"\\\\exists",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∃"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\exists"
	)


	test(
		"\\\\forall",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "∀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\forall"
	)
	test(
		"\\\\funcapply",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⁡"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\funcapply"
	)
	test(
		"\\\\frown",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "⌑"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\frown"
	)

	test(
		"\\\\Gamma",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "Γ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Gamma"
	)
	test(
		"\\\\gamma",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "γ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\gamma"
	)
	test(
		"\\\\ge",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≥"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ge"
	)
	test(
		"\\\\geq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≥"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\geq"
	)
	test(
		"\\\\gets",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "←"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\gets"
	)
	test(
		"\\\\gg",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "≫"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\gg"
	)
	test(
		"\\\\gimel",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "ℷ"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\gimel"
	)
	test(
		"\\\\grave",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\grave"
	)

	test(
		"\\\\hairsp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hairsp"
	)
	test(
		"\\\\hat",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hat"
	)
	test(
		"\\\\hbar",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hbar"
	)
	test(
		"\\\\heartsuit",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\heartsuit"
	)
	test(
		"\\\\hookleftarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hookleftarrow"
	)

	test(
		"\\\\hphantom",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hphantom"
	)


	test(
		"\\\\hsmash",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hsmash"
	)
	test(
		"\\\\hvec",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\hvec"
	)


	test(
		"\\\\Im",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Im"
	)
	test(
		"\\\\iiiint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\iiiint"
	)
	test(
		"\\\\iiint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\iiint"
	)
	test(
		"\\\\iint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\iint"
	)
	test(
		"\\\\ii",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ii"
	)
	test(
		"\\\\int",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\int"
	)
	test(
		"\\\\imath",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\imath"
	)
	test(
		"\\\\inc",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\inc"
	)
	test(
		"\\\\infty",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\infty"
	)
	test(
		"\\\\in",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\in"
	)
	test(
		"\\\\iota",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\iota"
	)
	test(
		"\\\\jj",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\jj"
	)
	test(
		"\\\\jmath",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\jmath"
	)
	test(
		"\\\\kappa",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\kappa"
	)
	test(
		"\\\\ket",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ket"
	)
	test(
		"\\\\Longleftrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Longleftrightarrow"
	)
	test(
		"\\\\Longrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Longrightarrow"
	)
	test(
		"\\\\Lambda",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Lambda"
	)

	test(
		"\\\\lambda",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\lambda"
	)
	test(
		"\\\\langle",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\langle"
	)
	test(
		"\\\\lbrack",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\lbrack"
	)

	test(
		"\\\\ldiv",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ldiv"
	)
	test(
		"\\\\ldots",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ldots"
	)
	test(
		"\\\\le",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\le"
	)
	test(
		"\\\\Leftarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Leftarrow"
	)
	test(
		"\\\\leftarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\leftarrow"
	)
	test(
		"\\\\leftharpoondown",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\leftharpoondown"
	)
	test(
		"\\\\leftharpoonup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\leftharpoonup"
	)
	test(
		"\\\\Leftrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Leftrightarrow"
	)
	test(
		"\\\\leftrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\leftrightarrow"
	)
	test(
		"\\\\leq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\leq"
	)
	test(
		"\\\\lfloor",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\lfloor"
	)
	test(
		"\\\\ll",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ll"
	)
	test(
		"\\\\Longleftarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Longleftarrow"
	)
	test(
		"\\\\longleftarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\longleftarrow"
	)
	test(
		"\\\\longleftrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\longleftrightarrow"
	)
	test(
		"\\\\longrightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\longrightarrow"
	)

	test(
		"\\\\lmoust",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\lmoust"
	)

	test(
		"\\\\mapsto",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\mapsto"
	)
	test(
		"\\\\matrix",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\matrix"
	)
	test(
		"\\\\medsp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\medsp"
	)
	test(
		"\\\\mid",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\mid"
	)
	test(
		"\\\\models",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\models"
	)
	test(
		"\\\\mp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\mp"
	)
	test(
		"\\\\mu",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\mu"
	)
	test(
		"\\\\nabla",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\nabla"
	)
	test(
		"\\\\naryand",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\naryand"
	)
	test(
		"\\\\nbsp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\nbsp"
	)
	test(
		"\\\\ndiv",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ndiv"
	)
	test(
		"\\\\ne",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ne"
	)
	test(
		"\\\\nearrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\nearrow"
	)
	test(
		"\\\\neg",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\neg"
	)
	test(
		"\\\\neq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\neq"
	)
	test(
		"\\\\ni",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ni"
	)
	test(
		"\\\\norm",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\norm"
	)
	test(
		"\\\\nu",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\nu"
	)
	test(
		"\\\\nwarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\nwarrow"
	)

	test(
		"\\\\Omega",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Omega"
	)
	test(
		"\\\\odot",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\odot"
	)
	test(
		"\\\\of",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\of"
	)
	test(
		"\\\\oiiint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\oiiint"
	)
	test(
		"\\\\oiint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\oiint"
	)
	test(
		"\\\\oint",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\oint"
	)
	test(
		"\\\\omega",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\omega"
	)
	test(
		"\\\\ominus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ominus"
	)
	test(
		"\\\\open",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\open"
	)
	test(
		"\\\\oplus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\oplus"
	)

	test(
		"\\\\oslash",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\oslash"
	)
	test(
		"\\\\otimes",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\otimes"
	)
	test(
		"\\\\over",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\over"
	)
	test(
		"\\\\overbar",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\overbar"
	)
	test(
		"\\\\overbrace",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\overbrace"
	)
	test(
		"\\\\overbracket",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\overbracket"
	)
	test(
		"\\\\overparen",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\overparen"
	)
	test(
		"\\\\overshell",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\overshell"
	)
	test(
		"\\\\over",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\over"
	)
	test(
		"\\\\Pi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Pi"
	)
	test(
		"\\\\Phi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Phi"
	)
	test(
		"\\\\Psi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Psi"
	)
	test(
		"\\\\parallel",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\parallel"
	)
	test(
		"\\\\partial",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\partial"
	)
	test(
		"\\\\perp",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\perp"
	)
	test(
		"\\\\phantom",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\phantom"
	)
	test(
		"\\\\phi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\phi"
	)
	test(
		"\\\\pi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\pi"
	)
	test(
		"\\\\pm",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\pm"
	)
	test(
		"\\\\pppprime",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\pppprime"
	)
	test(
		"\\\\ppprime",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ppprime"
	)
	test(
		"\\\\pprime",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\pprime"
	)
	test(
		"\\\\prcue",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\prcue"
	)
	test(
		"\\\\prec",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\prec"
	)
	test(
		"\\\\preceq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\preceq"
	)
	test(
		"\\\\preccurlyeq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\preccurlyeq"
	)
	test(
		"\\\\prime",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\prime"
	)
	test(
		"\\\\propto",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\propto"
	)
	test(
		"\\\\psi",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\psi"
	)
	test(
		"\\\\qdrt",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\qdrt"
	)
	test(
		"\\\\Re",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Re"
	)
	test(
		"\\\\Rightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Rightarrow"
	)
	test(
		"\\\\rangle",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rangle"
	)
	test(
		"\\\\ratio",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\ratio"
	)
	test(
		"\\\\rbrace",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rbrace"
	)
	test(
		"\\\\rbrack",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rbrack"
	)
	test(
		"\\\\rceil",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rceil"
	)
	test(
		"\\\\rddots",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rddots"
	)
	test(
		"\\\\rect",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rect"
	)
	test(
		"\\\\rfloor",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rfloor"
	)
	test(
		"\\\\rho",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rho"
	)
	test(
		"\\\\right",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\right"
	)
	test(
		"\\\\rightarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rightarrow"
	)
	test(
		"\\\\rightharpoondown",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rightharpoondown"
	)
	test(
		"\\\\rightharpoonup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rightharpoonup"
	)
	test(
		"\\\\rmoust",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rmoust"
	)
	test(
		"\\\\rrect",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\rrect"
	)
	test(
		"\\\\root",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\root"
	)
	test(
		"\\\\Sigma",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\Sigma"
	)
	test(
		"\\\\sdiv",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sdiv"
	)
	test(
		"\\\\searrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\searrow"
	)
	test(
		"\\\\setminus",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\setminus"
	)
	test(
		"\\\\sigma",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sigma"
	)
	test(
		"\\\\sim",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sim"
	)
	test(
		"\\\\simeq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\simeq"
	)
	test(
		"\\\\smash",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\smash"
	)
	test(
		"\\\\smile",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\smile"
	)
	test(
		"\\\\spadesuit",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\spadesuit"
	)
	test(
		"\\\\sqcap",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sqcap"
	)
	test(
		"\\\\sqcup",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sqcup"
	)
	test(
		"\\\\sqrt",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sqrt"
	)
	test(
		"\\\\sqsubseteq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sqsubseteq"
	)
	test(
		"\\\\sqsuperseteq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sqsuperseteq"
	)
	test(
		"\\\\star",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\star"
	)
	test(
		"\\\\subset",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\subset"
	)
	test(
		"\\\\subseteq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\subseteq"
	)
	test(
		"\\\\succeq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\succeq"
	)
	test(
		"\\\\succ",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\succ"
	)
	test(
		"\\\\sum",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\sum"
	)
	test(
		"\\\\superset",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\superset"
	)

	test(
		"\\\\superseteq",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\superseteq"
	)
	test(
		"\\\\swarrow",
		{
			"body": {
				"type": "OperatorLiteral",
				"value": "̀"
			},
			"type": "UnicodeEquation"
		},
		"Check: \\\\swarrow"
	)
}
window["AscMath"].literal = literalTests;
