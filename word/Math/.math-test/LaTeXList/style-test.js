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
function style(test) {
	test(
		"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\double{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": "CharLiteral",
					"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
				},
				{
					"type": "NumberLiteral",
					"value": "1234567890"
				},
				{
					"type": "MathFontLiteral",
					"fontValue": 12,
					"value": [
						{
							"type": "CharLiteral",
							"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
						},
						{
							"type": "NumberLiteral",
							"value": "1234567890"
						}
					]
				},
				{
					"type": "CharLiteral",
					"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
				},
				{
					"type": "NumberLiteral",
					"value": "1234567890"
				}
			]
		},
		"Check ...\\double{...}..."
	);
	test(
		"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\fraktur{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
		{
			"type": "LaTeXEquation",
			"body": [
				{
					"type": "CharLiteral",
					"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
				},
				{
					"type": "NumberLiteral",
					"value": "1234567890"
				},
				{
					"type": "MathFontLiteral",
					"fontValue": 9,
					"value": [
						{
							"type": "CharLiteral",
							"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
						},
						{
							"type": "NumberLiteral",
							"value": "1234567890"
						}
					]
				},
				{
					"type": "CharLiteral",
					"value": "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
				},
				{
					"type": "NumberLiteral",
					"value": "1234567890"
				}
			]
		},
		"Check ...\\fraktur{...}..."
	);
	test(
		"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\script{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
		{
			body: [
				{
					CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
				},
				{
					NumberLiteral: "1234567890",
				},
				{
					fontValue: 7,
					type: "FontStyleLiteral",
					value: [
						{
							CharLiteral:
								"𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂𝒬𝒲ℰℛ𝒯𝒴𝒰ℐ𝒪𝒫𝒜𝒮𝒟ℱ𝒢ℋ𝒥𝒦ℒ𝒵𝒳𝒞𝒱ℬ𝒩ℳ",
						},
						{
							NumberLiteral: "1234567890",
						},
					],
				},
				{
					CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
				},
				{
					NumberLiteral: "1234567890",
				},
			],
			type: "LaTeXEquation",
		},
		"Check ...\\script{...}..."
	);
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\it{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 1,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚𝑄𝑊𝐸𝑅𝑇𝑌𝑈𝐼𝑂𝑃𝐴𝑆𝐷𝐹𝐺𝐻𝐽𝐾𝐿𝑍𝑋𝐶𝑉𝐵𝑁𝑀",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\it{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathbb{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 12,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝕢𝕨𝕖𝕣𝕥𝕪𝕦𝕚𝕠𝕡𝕒𝕤𝕕𝕗𝕘𝕙𝕛𝕜𝕝𝕫𝕩𝕔𝕧𝕓𝕟𝕞ℚ𝕎𝔼ℝ𝕋𝕐𝕌𝕀𝕆ℙ𝔸𝕊𝔻𝔽𝔾ℍ𝕁𝕂𝕃ℤ𝕏ℂ𝕍𝔹ℕ𝕄",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟙𝟚𝟛𝟜𝟝𝟞𝟟𝟠𝟡𝟘",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathbb{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathbf{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 0,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝐪𝐰𝐞𝐫𝐭𝐲𝐮𝐢𝐨𝐩𝐚𝐬𝐝𝐟𝐠𝐡𝐣𝐤𝐥𝐳𝐱𝐜𝐯𝐛𝐧𝐦𝐐𝐖𝐄𝐑𝐓𝐘𝐔𝐈𝐎𝐏𝐀𝐒𝐃𝐅𝐆𝐇𝐉𝐊𝐋𝐙𝐗𝐂𝐕𝐁𝐍𝐌",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟏𝟐𝟑𝟒𝟓𝟔𝟕𝟖𝟗𝟎",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathbf{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathcal{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 7,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂𝒬𝒲ℰℛ𝒯𝒴𝒰ℐ𝒪𝒫𝒜𝒮𝒟ℱ𝒢ℋ𝒥𝒦ℒ𝒵𝒳𝒞𝒱ℬ𝒩ℳ",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathcal{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathfrak{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 9,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪𝔔𝔚𝔈ℜ𝔗𝔜𝔘ℑ𝔒𝔓𝔄𝔖𝔇𝔉𝔊ℌ𝔍𝔎𝔏ℨ𝔛ℭ𝔙𝔅𝔑𝔐",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathfrak{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathit{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 1,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚𝑄𝑊𝐸𝑅𝑇𝑌𝑈𝐼𝑂𝑃𝐴𝑆𝐷𝐹𝐺𝐻𝐽𝐾𝐿𝑍𝑋𝐶𝑉𝐵𝑁𝑀",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathit{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathsf{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 3,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆𝖰𝖶𝖤𝖱𝖳𝖸𝖴𝖨𝖮𝖯𝖠𝖲𝖣𝖥𝖦𝖧𝖩𝖪𝖫𝖹𝖷𝖢𝖵𝖡𝖭𝖬",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟣𝟤𝟥𝟦𝟧𝟨𝟩𝟪𝟫𝟢",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathsf{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathtt{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 11,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝚚𝚠𝚎𝚛𝚝𝚢𝚞𝚒𝚘𝚙𝚊𝚜𝚍𝚏𝚐𝚑𝚓𝚔𝚕𝚣𝚡𝚌𝚟𝚋𝚗𝚖𝚀𝚆𝙴𝚁𝚃𝚈𝚄𝙸𝙾𝙿𝙰𝚂𝙳𝙵𝙶𝙷𝙹𝙺𝙻𝚉𝚇𝙲𝚅𝙱𝙽𝙼",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟷𝟸𝟹𝟺𝟻𝟼𝟽𝟾𝟿𝟶",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathtt{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mit{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 1,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝑞𝑤𝑒𝑟𝑡𝑦𝑢𝑖𝑜𝑝𝑎𝑠𝑑𝑓𝑔ℎ𝑗𝑘𝑙𝑧𝑥𝑐𝑣𝑏𝑛𝑚𝑄𝑊𝐸𝑅𝑇𝑌𝑈𝐼𝑂𝑃𝐴𝑆𝐷𝐹𝐺𝐻𝐽𝐾𝐿𝑍𝑋𝐶𝑉𝐵𝑁𝑀",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mit{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\oldstyle{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 7,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝓆𝓌ℯ𝓇𝓉𝓎𝓊𝒾ℴ𝓅𝒶𝓈𝒹𝒻ℊ𝒽𝒿𝓀𝓁𝓏𝓍𝒸𝓋𝒷𝓃𝓂𝒬𝒲ℰℛ𝒯𝒴𝒰ℐ𝒪𝒫𝒜𝒮𝒟ℱ𝒢ℋ𝒥𝒦ℒ𝒵𝒳𝒞𝒱ℬ𝒩ℳ",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\oldstyle{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\sf{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 3,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝗊𝗐𝖾𝗋𝗍𝗒𝗎𝗂𝗈𝗉𝖺𝗌𝖽𝖿𝗀𝗁𝗃𝗄𝗅𝗓𝗑𝖼𝗏𝖻𝗇𝗆𝖰𝖶𝖤𝖱𝖳𝖸𝖴𝖨𝖮𝖯𝖠𝖲𝖣𝖥𝖦𝖧𝖩𝖪𝖫𝖹𝖷𝖢𝖵𝖡𝖭𝖬",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟣𝟤𝟥𝟦𝟧𝟨𝟩𝟪𝟫𝟢",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\sf{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathbfit{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 2,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝒒𝒘𝒆𝒓𝒕𝒚𝒖𝒊𝒐𝒑𝒂𝒔𝒅𝒇𝒈𝒉𝒋𝒌𝒍𝒛𝒙𝒄𝒗𝒃𝒏𝒎𝑸𝑾𝑬𝑹𝑻𝒀𝑼𝑰𝑶𝑷𝑨𝑺𝑫𝑭𝑮𝑯𝑱𝑲𝑳𝒁𝑿𝑪𝑽𝑩𝑵𝑴",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathbfit{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathsfbfit{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 6,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝙦𝙬𝙚𝙧𝙩𝙮𝙪𝙞𝙤𝙥𝙖𝙨𝙙𝙛𝙜𝙝𝙟𝙠𝙡𝙯𝙭𝙘𝙫𝙗𝙣𝙢𝙌𝙒𝙀𝙍𝙏𝙔𝙐𝙄𝙊𝙋𝘼𝙎𝘿𝙁𝙂𝙃𝙅𝙆𝙇𝙕𝙓𝘾𝙑𝘽𝙉𝙈",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathsfbfit{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathsfbf{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 4,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝗾𝘄𝗲𝗿𝘁𝘆𝘂𝗶𝗼𝗽𝗮𝘀𝗱𝗳𝗴𝗵𝗷𝗸𝗹𝘇𝘅𝗰𝘃𝗯𝗻𝗺𝗤𝗪𝗘𝗥𝗧𝗬𝗨𝗜𝗢𝗣𝗔𝗦𝗗𝗙𝗚𝗛𝗝𝗞𝗟𝗭𝗫𝗖𝗩𝗕𝗡𝗠",
	// 					},
	// 					{
	// 						NumberLiteral: "𝟭𝟮𝟯𝟰𝟱𝟲𝟳𝟴𝟵𝟬",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathsfbf{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathsfit{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 5,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝘲𝘸𝘦𝘳𝘵𝘺𝘶𝘪𝘰𝘱𝘢𝘴𝘥𝘧𝘨𝘩𝘫𝘬𝘭𝘻𝘹𝘤𝘷𝘣𝘯𝘮𝘘𝘞𝘌𝘙𝘛𝘠𝘜𝘐𝘖𝘗𝘈𝘚𝘋𝘍𝘎𝘏𝘑𝘒𝘓𝘡𝘟𝘊𝘝𝘉𝘕𝘔",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathsfit{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathbfcal{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 8,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝓺𝔀𝓮𝓻𝓽𝔂𝓾𝓲𝓸𝓹𝓪𝓼𝓭𝓯𝓰𝓱𝓳𝓴𝓵𝔃𝔁𝓬𝓿𝓫𝓷𝓶𝓠𝓦𝓔𝓡𝓣𝓨𝓤𝓘𝓞𝓟𝓐𝓢𝓓𝓕𝓖𝓗𝓙𝓚𝓛𝓩𝓧𝓒𝓥𝓑𝓝𝓜",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathbfcal{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\frak{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 9,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝔮𝔴𝔢𝔯𝔱𝔶𝔲𝔦𝔬𝔭𝔞𝔰𝔡𝔣𝔤𝔥𝔧𝔨𝔩𝔷𝔵𝔠𝔳𝔟𝔫𝔪𝔔𝔚𝔈ℜ𝔗𝔜𝔘ℑ𝔒𝔓𝔄𝔖𝔇𝔉𝔊ℌ𝔍𝔎𝔏ℨ𝔛ℭ𝔙𝔅𝔑𝔐",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\frak{...}..."
	// );
	// test(
	// 	"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890\\mathbffrak{qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890}qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890",
	// 	{
	// 		body: [
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 			{
	// 				fontValue: 10,
	// 				type: "FontStyleLiteral",
	// 				value: [
	// 					{
	// 						CharLiteral:
	// 							"𝖖𝖜𝖊𝖗𝖙𝖞𝖚𝖎𝖔𝖕𝖆𝖘𝖉𝖋𝖌𝖍𝖏𝖐𝖑𝖟𝖝𝖈𝖛𝖇𝖓𝖒𝕼𝖂𝕰𝕽𝕿𝖄𝖀𝕴𝕺𝕻𝕬𝕾𝕯𝕱𝕲𝕳𝕵𝕶𝕷𝖅𝖃𝕮𝖁𝕭𝕹𝕸",
	// 					},
	// 					{
	// 						NumberLiteral: "1234567890",
	// 					},
	// 				],
	// 			},
	// 			{
	// 				CharLiteral: "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM",
	// 			},
	// 			{
	// 				NumberLiteral: "1234567890",
	// 			},
	// 		],
	// 		type: "LaTeXEquation",
	// 	},
	// 	"Check ...\\mathbffrak{...}..."
	// );

	test(
		"\\doubleAB",
		{
			body: [
				{
					fontValue: 12,
					type: "FontStyleLiteral",
					value: {
						CharLiteral: "𝔸",
					},
				},
				{
					CharLiteral: "B",
				},
			],
			type: "LaTeXEquation",
		},
		"Check \\doubleA"
	);
	test(
		"\\frakturAB",
		{
			body: [
				{
					fontValue: 9,
					type: "FontStyleLiteral",
					value: {
						CharLiteral: "𝔄",
					},
				},
				{
					CharLiteral: "B",
				},
			],
			type: "LaTeXEquation",
		},
		"Check \\frakturA"
	);
	test(
		"\\fraktur AB",
		{
			body: [
				{
					fontValue: 9,
					type: "FontStyleLiteral",
					value: {
						CharLiteral: "𝔄",
					},
				},
				{
					CharLiteral: "B",
				},
			],
			type: "LaTeXEquation",
		},
		"Check \\frakturA"
	);

	test(
		"\\dd",
		{
			body: {
				CharLiteral: "𝑑",
			},
			type: "LaTeXEquation",
		},
		"Check \\dd"
	);
	test(
		"\\Dd",
		{
			body: {
				CharLiteral: "𝐷",
			},
			type: "LaTeXEquation",
		},
		"Check \\Dd"
	);
	test(
		"\\ee",
		{
			body: {
				CharLiteral: "𝑒",
			},
			type: "LaTeXEquation",
		},
		"Check \\ee"
	);
	test(
		"\\hbar",
		{
			body: {
				CharLiteral: "ℏ",
			},
			type: "LaTeXEquation",
		},
		"Check \\hbar"
	);
	test(
		"\\ii",
		{
			body: {
				CharLiteral: "𝑖",
			},
			type: "LaTeXEquation",
		},
		"Check \\ii"
	);
	test(
		"\\Im",
		{
			body: {
				CharLiteral: "𝕴",
			},
			type: "LaTeXEquation",
		},
		"Check \\Im"
	);
	test(
		"\\imath",
		{
			body: {
				CharLiteral: "𝚤",
			},
			type: "LaTeXEquation",
		},
		"Check \\imath"
	);
	test(
		"\\j",
		{
			body: {
				CharLiteral: "𝐽𝑎𝑦",
			},
			type: "LaTeXEquation",
		},
		"Check \\j"
	);
	test(
		"\\jj",
		{
			body: {
				CharLiteral: "𝑗",
			},
			type: "LaTeXEquation",
		},
		"Check \\jj"
	);
	test(
		"\\jmath",
		{
			body: {
				CharLiteral: "𝐽",
			},
			type: "LaTeXEquation",
		},
		"Check \\jmath"
	);
	test(
		"\\partial",
		{
			body: {
				CharLiteral: "∂",
			},
			type: "LaTeXEquation",
		},
		"Check \\partial"
	);
	test(
		"\\Re",
		{
			body: {
				CharLiteral: "ℜ",
			},
			type: "LaTeXEquation",
		},
		"Check \\Re"
	);
	test(
		"\\wp",
		{
			body: {
				CharLiteral: "℘",
			},
			type: "LaTeXEquation",
		},
		"Check \\wp"
	);
	test(
		"\\aleph",
		{
			body: {
				CharLiteral: "ℵ",
			},
			type: "LaTeXEquation",
		},
		"Check \\aleph"
	);
	test(
		"\\bet",
		{
			body: {
				CharLiteral: "ℶ",
			},
			type: "LaTeXEquation",
		},
		"Check \\bet"
	);
	test(
		"\\beth",
		{
			body: {
				CharLiteral: "ℶ",
			},
			type: "LaTeXEquation",
		},
		"Check \\beth"
	);
	test(
		"\\gimel",
		{
			body: {
				CharLiteral: "ℷ",
			},
			type: "LaTeXEquation",
		},
		"Check \\gimel"
	);
	test(
		"\\dalet",
		{
			body: {
				CharLiteral: "ℸ",
			},
			type: "LaTeXEquation",
		},
		"Check \\dalet"
	);
	test(
		"\\daleth",
		{
			body: {
				CharLiteral: "ℸ",
			},
			type: "LaTeXEquation",
		},
		"Check \\daleth"
	);

	test(
		"\\alpha",
		{
			body: {
				CharLiteral: "α",
			},
			type: "LaTeXEquation",
		},
		"Check \\alpha"
	);
}

window["AscMath"].style = style;
