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
function complexTest(test) {
	test(
		`(a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "a",
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
								CharLiteral: "b",
							},
						],
						open: "(",
						close: ")",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∑",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "k",
								},
								{
									Operator: "=",
								},
								{
									NumberLiteral: "0",
								},
							],
							open: "(",
							close: ")",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: {
								type: "binomLiteral",
								numerator: {
									CharLiteral: "n",
								},
								operand: {
									CharLiteral: "k",
								},
							},
							open: "(",
							close: ")",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "a",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "k",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "b",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "n",
								},
								{
									Operator: "-",
								},
								{
									CharLiteral: "k",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
			],
		},
		"Проверка простого литерала: (a + b)^n = ∑_(k=0)^n▒(n¦k) a^k b^(n-k),"
	);
	test(
		`∑_2^2▒(n/23)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expSubsup",
				base: {
					type: "opNary",
					value: "∑",
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
						NumberLiteral: "2",
					},
				},
				thirdSoperand: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: {
							type: "fractionLiteral",
							up: {
								CharLiteral: "n",
							},
							opOver: "/",
							down: {
								NumberLiteral: "23",
							},
						},
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Проверка простого литерала: ∑_2^2▒(n/23)"
	);
	test(
		`⏞(x+⋯+x)^(k "times")`,
		{
			type: "UnicodeEquation",
			body: {
				type: "hbrackLiteral",
				operand: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "x",
						},
						{
							Operator: "+",
						},
						{
							CharLiteral: "⋯",
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
				up: {
					type: "soperandLiteral",
					operand: {
						type: "expBracketLiteral",
						exp: [
							{
								CharLiteral: "k",
							},
							{
								type: "SpaceLiteral",
								value: " ",
							},
							{
								CharLiteral: '"times"',
							},
						],
						open: "(",
						close: ")",
					},
				},
			},
		},
		"Проверка простого литерала: ⏞(x+⋯+x)^(k 'times')"
	);
	test(
		`𝐸 = 𝑚𝑐^2`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "𝐸",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "𝑚𝑐",
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
		"Проверка простого литерала: 𝐸 = 𝑚𝑐^2"
	);
	test(
		`∫_0^a▒xⅆx/(x^2+a^2)`,
		{
			type: "UnicodeEquation",
			body: {
				type: "fractionLiteral",
				up: {
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∫",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "0",
						},
					},
					up: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "a",
						},
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "xⅆx",
						},
					},
				},
				opOver: "/",
				down: {
					type: "expBracketLiteral",
					exp: [
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
							type: "expSuperscript",
							base: {
								CharLiteral: "a",
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
			},
		},
		"Проверка простого литерала: ∫_0^a▒xⅆx/(x^2+a^2)"
	);
	test(
		`lim┬(n→∞) a_n`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expBelow",
					base: {
						CharLiteral: "lim",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							type: "expBracketLiteral",
							exp: [
								{
									CharLiteral: "n",
								},
								{
									Operator: "→",
								},
								{
									Operator: "∞",
								},
							],
							open: "(",
							close: ")",
						},
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSubscript",
					base: {
						CharLiteral: "a",
					},
					down: {
						type: "soperandLiteral",
						operand: {
							CharLiteral: "n",
						},
					},
				},
			],
		},
		"Проверка простого литерала: lim┬(n→∞) a_n"
	);
	test(
		`ⅈ²=-1`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "ⅈ",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "=",
				},
				{
					Operator: "-",
				},
				{
					NumberLiteral: "1",
				},
			],
		},
		"Проверка простого литерала: ⅈ²=-1"
	);
	test(
		`E = m⁢c²`,
		{
			type: "UnicodeEquation",
			body: [
				{
					CharLiteral: "E",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					Operator: "=",
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "m⁢c",
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
		"Проверка простого литерала: E = m⁢c²"
	);
	test(
		`a²⋅b²=c²`,
		{
			type: "UnicodeEquation",
			body: [
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "a",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "⋅",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "b",
					},
					up: {
						type: "soperandLiteral",
						operand: {
							NumberLiteral: "2",
						},
					},
				},
				{
					Operator: "=",
				},
				{
					type: "expSuperscript",
					base: {
						CharLiteral: "c",
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
		"Проверка простого литерала: a²⋅b²=c²"
	);
	test(
		`f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx`,
		{
			type: "UnicodeEquation",
			body: [
				[
					{
						CharLiteral: "f̂",
					},
					{
						type: "expBracketLiteral",
						exp: {
							type: "anOther",
							value: "ξ",
						},
						open: "(",
						close: ")",
					},
				],
				{
					Operator: "=",
				},
				{
					type: "expSubsup",
					base: {
						type: "opNary",
						value: "∫",
					},
					down: {
						type: "soperandLiteral",
						operand: "-∞",
					},
					up: {
						type: "soperandLiteral",
						operand: "∞",
					},
					thirdSoperand: {
						type: "soperandLiteral",
						operand: [
							{
								CharLiteral: "f",
							},
							{
								type: "expBracketLiteral",
								exp: {
									CharLiteral: "x",
								},
								open: "(",
								close: ")",
							},
							{
								type: "expSuperscript",
								base: {
									CharLiteral: "ⅇ",
								},
								up: {
									type: "soperandLiteral",
									operand: [
										{
											NumberLiteral: "2",
										},
										{
											type: "anOther",
											value: "π",
										},
										{
											CharLiteral: "ⅈxξ",
										},
									],
									minus: true,
								},
							},
						],
					},
				},
				{
					type: "SpaceLiteral",
					value: " ",
				},
				{
					CharLiteral: "ⅆx",
				},
			],
		},
		"Проверка простого литерала: f̂(ξ)=∫_-∞^∞▒f(x)ⅇ^-2πⅈxξ ⅆx"
	);
	test(
		`(𝑎 + 𝑏)┴→`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: {
					type: "expBracketLiteral",
					exp: [
						{
							CharLiteral: "𝑎",
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
							CharLiteral: "𝑏",
						},
					],
					open: "(",
					close: ")",
				},
				up: {
					Operator: "→",
				},
			},
		},
		"Проверка простого литерала: (𝑎 + 𝑏)┴→"
	);
	test(
		`𝑎┴→`,
		{
			type: "UnicodeEquation",
			body: {
				type: "expAbove",
				base: {
					CharLiteral: "𝑎",
				},
				up: {
					Operator: "→",
				},
			},
		},
		"Проверка простого литерала: 𝑎┴→"
	);
}
window["AscMath"].complex = complexTest;
